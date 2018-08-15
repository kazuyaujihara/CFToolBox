using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using MSWord = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Ujihara.Chemistry.MSOffice;
using Ujihara.Chemistry.IO;
using OX = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ujihara.Chemistry.MergeSF
{
    public class SciFinderReferencesExtractor
        : IEnumerable<ReferenceInfo>
    {
        private static object missing = Type.Missing;

        private const string GroupName_Text = "text";
        private const string titleLinePattern = @"^\d+\.\s(?<" + GroupName_Text + ">.*)$";
        private static readonly Regex reIsTitleTable = new Regex(titleLinePattern, RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex reTitleLine = new Regex(titleLinePattern, RegexOptions.Multiline | RegexOptions.Compiled);

        const string C_By = "By";
        const string C_Assignee = "Assignee";
        const string C_PatentInformation = "Patent Information";
        const string C_Source = "Source";
        const string C_AccessionNumber = "Accession Number";
        const string C_Language = "Language";
        const string C_Company_Organization = "Company/Organization";
        const string C_Publisher = "Publisher";
        private static readonly Regex reBy = MakeRE(C_By, ".*", true);
        private static readonly Regex reAssignee = MakeRE(C_Assignee, ".*", true);
        private static readonly Regex rePatentInformation = MakeRE(C_PatentInformation, ".*", true);
        private static readonly Regex reSource = MakeRE(C_Source, ".*", true);
        private static readonly Regex reAccessionNumber = MakeRE(C_AccessionNumber, @"\d\d\d\d\:?\d+", false);
        private static readonly Regex reLanguage = MakeRE(C_Language, ".*", true);
        private static readonly Regex reCompany = MakeRE(C_Company_Organization.Replace("/", @"\/"), ".*", true);
        private static readonly Regex rePublisher = MakeRE(C_Publisher, ".*", true);

        private const string Text_Abstract = "Abstract";
        private const string Text_PatentInfomation = "Patent Information";
        private const string Text_PriorityApplication = "Priority Application";
        private const string Text_Indexingn = "Indexing";
        private const string Text_Concepts = "Concepts";
        private const string Text_Substances = "Substances";
        private const string Text_SupplementaryTerms = "Supplementary Terms";
        private static readonly Regex reIsAbstractTable = new Regex(@"^" + Text_Abstract + "$", RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex reIsPatentInfomation = new Regex(@"^" + Text_PatentInfomation + "$", RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex reIsJournal = new Regex(@" Journal(\, \d\d\d\d\,|\; [ A-Za-z0-9]*\, \d\d\d\d\,)", RegexOptions.Compiled);

        private const string GroupName_CASRN = "casrn";
        private const string GroupName_Info = "info";
        private static readonly Regex reCASRN_A = new Regex(@"\s*(?<" + GroupName_CASRN + @">\d+\-\d\d\-\d)[A-Z]\s*", RegexOptions.Compiled);
        private static readonly Regex reSubstanceLine = new Regex(
            @"^(?<" + GroupName_CASRN + @">\d+\-\d\d\-\d)[A-Z]?(\s*(?<" + GroupName_Info + @">.*))?$", RegexOptions.Compiled);

        private static Regex MakeRE(string tag, string pat, bool isWholeLine)
        {
            return new Regex("^" + tag + @"\:\s*(?<" + GroupName_Text + @">" + pat + ")" + (isWholeLine ? "$" : ""), RegexOptions.Multiline | RegexOptions.Compiled);
        }

        private string source;
        public string Source
        {
            get { return source; }
        }

        public SciFinderReferencesExtractor(string source)
        {
            this.source = Path.GetFullPath(source);
        }

        private WordprocessingDocument wd;
        private TempDirectory tempdir = null;

        ~SciFinderReferencesExtractor()
        {
            if (tempdir != null)
                tempdir.Dispose();
        }

        public IEnumerator<ReferenceInfo> GetEnumerator()
        {
            if (this.Source == null)
                yield break;

            string docx;
            if (Path.GetExtension(Source) == ".docx")
                docx = Source;
            else
            {
                tempdir = new TempDirectory();
                docx = Path.Combine(tempdir.Directory.FullName, Path.GetFileNameWithoutExtension(Source) + ".docx");
                WordUtility.ConvertToDocx(Source, docx);
            }

            // TODO: Handle Mac char to Windows char

            using (wd = WordprocessingDocument.Open(docx, false))
            {
                var body = wd.MainDocumentPart.Document.Body;

                if (TotalNumbrSetter != null)
                    TotalNumbrSetter(body.Elements().Count());

                foreach (var table in body.Elements<W.Table>())
                {
                    var firstRow = table.Elements<W.TableRow>().FirstOrDefault();
                    if (firstRow != null)
                    {
                        var text = firstRow.InnerText;
                        if (reIsTitleTable.IsMatch(text))
                            yield return CreateReferenceInfo(table);
                    }
                    if (ProgressIncrementer != null)
                        ProgressIncrementer();
                }
            }
            yield break;
        }

        private static string ExtractPropertyInAbb(Regex regex, string text)
        {
            var ma = regex.Match(text);
            if (ma.Success)
                return ma.Groups[GroupName_Text].Value;
            return null;
        }

        enum EK
        {
            Nil = 0,
            Unknown, 
            Table, 
            Paragraph,
            
            BibTable, 
            Abstract,
            AbstractImage,
            PatentInformation,
            SubstancesLine,
            Substance,
            SubstanceImage,
            Term,
            Copyright,
        }

        private ReferenceInfo CreateReferenceInfo(DocumentFormat.OpenXml.Wordprocessing.Table bibTable)
        {
            var ri = new AReferenceInfo();

            {
                var firstRow = bibTable.Elements<W.TableRow>().FirstOrDefault();
                if (firstRow == null)
                    throw new FormatException();
                ri.Title = ExtractPropertyInAbb(reTitleLine, firstRow.InnerText);

                foreach (var row in bibTable.Elements<W.TableRow>().Skip(1))
                {
                    var text = row.InnerText;

                    if (text.StartsWith(C_AccessionNumber + ":"))
                        ri.AccessionNumber = ExtractPropertyInAbb(reAccessionNumber, text);
                    if (text.StartsWith(C_Assignee + ":"))
                        ri.PatentAssignee = ExtractPropertyInAbb(reAssignee, text);
                    if (text.StartsWith(C_Company_Organization + ":"))
                        ri.CorporateSource = ExtractPropertyInAbb(reCompany, text);
                    if (text.StartsWith(C_By + ":"))
                        ri.By = ExtractPropertyInAbb(reBy, text);
                    if (text.StartsWith(C_Publisher + ":"))
                        ri.Publisher = ExtractPropertyInAbb(rePublisher, text);
                    if (text.StartsWith(C_Language + ":"))
                        ri.Language = ExtractPropertyInAbb(reLanguage, text);
                    if (text.StartsWith(C_Source + ":"))
                        ri.Source = ExtractPropertyInAbb(rePatentInformation, text);
                    if (ri.Source != null)
                    {
                        ri.DocumentType = CmpdDbManager.DocumentType_Patent;
                    }
                    else
                    {
                        var s = ExtractPropertyInAbb(reSource, text);
                        if (s != null)
                        {
                            ri.Source = s;
                            if (reIsJournal.IsMatch(s))
                            {
                                ri.DocumentType = CmpdDbManager.DocumentType_Journal;
                            }
                        }
                    }
                }
            }

            var prev = EK.BibTable;
            for (var elm = bibTable.NextSibling(); elm != null; elm = elm.NextSibling())
            {
                if (elm is W.Table)
                {
                    var table = (W.Table)elm;
                    var rows = table.Elements<W.TableRow>();
                    var firstRow = rows.FirstOrDefault();
                    if (firstRow == null)
                    {
                        prev = EK.Table;
                        continue;
                    }
                    var text = firstRow.InnerText.Trim();
                    if (reTitleLine.IsMatch(text))
                        break; // next doc found
                    switch (text)
                    {
                        case Text_Abstract:
                            {
                                var row = rows.Skip(1).FirstOrDefault();
                                if (row == null)
                                {
                                    prev = EK.Table;
                                    continue;
                                }
                                var cell = row.Elements<W.TableCell>().FirstOrDefault();
                                if (cell == null)
                                {
                                    prev = EK.Table;
                                    continue;
                                }
                                ri.Abstract = WordUtility.ToRtf(cell, "Meiryo UI", 17);
                                prev = EK.Abstract;
                                continue;
                            }
                        case Text_PatentInfomation:
                            {
                                var sb = new StringBuilder();
                                foreach (var row in rows.Skip(1))
                                {
                                    sb.Append(row.InnerText.Trim()).Append('\n');
                                }
                                ri.PatentInfomation = sb.ToString();
                                prev = EK.PatentInformation;
                                continue;
                            }
                        default:
                            prev = EK.Table;
                            continue;
                    }
                }
                else if (elm is W.Paragraph)
                {
                    var paragraph = (W.Paragraph)elm;
                    var run = paragraph.Elements<W.Run>().FirstOrDefault();
                    if (run == null)
                    {
                        prev = EK.Paragraph;
                        continue;
                    }
                    var text = run.InnerText.Trim();
                    var jc = paragraph.ParagraphProperties.Justification;
                    if (text == "" && jc != null && jc.Val == "center" && prev == EK.Abstract)
                    {
                        var g = paragraph.Descendants<A.Graphic>().FirstOrDefault();
                        if (g != null)
                        {
                            ri.Container = wd.MainDocumentPart;
                            ri.Graphic = g;
                            prev = EK.AbstractImage;
                            continue;
                        }
                    }
                    if (text == Text_Substances)
                    {
                        ri.SubstancesInfo = GetSubstances(paragraph);
                        prev = EK.SubstancesLine;
                        continue;
                    }
                    if (text.StartsWith("Copyright "))
                    {
                        ri.Copyright = text;
                        prev = EK.Copyright;
                        break;
                    }
                }
            }
            if (ri.SubstancesInfo == null)
                ri.SubstancesInfo = new SubstanceInfo[0];
            return ri;
        }

        private IEnumerable<SubstanceInfo> GetSubstances(W.Paragraph substancesLine)
        {
            var elm = substancesLine.NextSibling();
            for (; elm != null; elm = elm.NextSibling())
            {
                if (!(elm is DocumentFormat.OpenXml.Wordprocessing.Paragraph))
                    break;
            }
            if (elm == null)
                throw new FormatException();

            string currTerm = null;
            A.Graphic currG = null;
            EK prev = EK.Nil;
            for (elm = elm.PreviousSibling(); elm != null && elm != substancesLine; elm = elm.PreviousSibling())
            {
                if (ProgressIncrementer != null)
                    ProgressIncrementer();

                var paragraph = (W.Paragraph)elm;
                var text = paragraph.InnerText.Trim();
                var jc = paragraph.ParagraphProperties.Justification;
                if (jc != null && jc.Val == "center")
                {
                    var g = paragraph.Descendants<A.Graphic>().FirstOrDefault();
                    if (g != null)
                    {
                        currG = g;
                        prev = EK.SubstanceImage;
                        continue;
                    }
                    prev = EK.Paragraph;
                    continue;
                }
                if (text == "")
                {
                    prev = EK.Paragraph;
                    continue;
                }

                var ma = reSubstanceLine.Match(text);
                if (ma.Success)
                {
                    var sub = new ASubstanceInfo();
                    sub.CASRN = ma.Groups[GroupName_CASRN].Value;
                    sub.Name = ma.Groups[GroupName_Info].Value;
                    if (currTerm != null)
                        sub.Keywords = currTerm;
                    if (currG != null)
                    {
                        sub.Container = wd.MainDocumentPart;
                        sub.Graphic = currG;
                    }
                    yield return sub;
                    currG = null;
                    prev = EK.Substance;
                }
                else
                {
                    if (prev != EK.Term || currTerm == null)
                    {
                        currTerm = text;
                    }
                    else
                    {
                        currTerm = text + "; " + currTerm;
                    }
                    prev = EK.Term;
                }
            }
            yield break;
        }

        public Action ProgressIncrementer = null;
        public Action<int> TotalNumbrSetter = null;

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}

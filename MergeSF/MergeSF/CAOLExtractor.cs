using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Ujihara.Chemistry.MSOffice;
using OX = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ujihara.Chemistry.MergeSF
{
    class CAOLExtractor
        : IEnumerable<ReferenceInfo>
    {
        private static object missing = Type.Missing;

        private const string GroupName_Prop = "prop";
        private const string GroupName_Text = "text";
        private const string GroupName_CASRN = "casrn";
        private const string GroupName_Copyright = "copyright";
        private static readonly Regex regexRNLine = new Regex("^RN {3}(?<" + GroupName_Text + @">\d+\-\d\d\-\d)", RegexOptions.Compiled);
        private static readonly Regex regexINFirstLine = new Regex("^(?<" + GroupName_Prop + ">IN|CN) {3}(?<" + GroupName_Text + ">.*)$", RegexOptions.Compiled);
        private static readonly Regex regexContinurousLine = new Regex("^ {5}(?<" + GroupName_Text + ">.*)$", RegexOptions.Compiled);
        private static readonly Regex regexHeaderLine = new Regex(@"^L\d+\s+\d+\s+ANSWERS\s+(?<" + GroupName_Copyright + ">[A-Za-z0-9][A-Za-z0-9 ]*)", RegexOptions.Compiled);   //L6   439 ANSWERS   REGISTRY  COPYRIGHT 2014 ACS on STN 
        private static readonly Regex regexRefHeaderLine = new Regex(@"^L\d+\s+ANSWER\s+\d+\s+OF\s+\d+\s+(?<" + GroupName_Copyright + ">[A-Za-z0-9][A-Za-z0-9 ]*)", RegexOptions.Compiled);  //L23  ANSWER 1 OF 2  HCAPLUS  COPYRIGHT 2013 ACS on STN 
        private static readonly Regex regexOneSpaceLine = new Regex(@"^ $", RegexOptions.Compiled);

        private static readonly Regex regexAccessionNumber = new Regex(@"^" + PadForRefExtractor("ACCESSION NUMBER") + @"(?<" + GroupName_Text + @">\d{3,4}\:?\d+)\s+", RegexOptions.Compiled);
        private static readonly Regex regexTitleLine = CreateRegexForRefExtraction("TITLE");
        private static readonly Regex regexInventor = CreateRegexForRefExtraction("INVENTOR(S)");
        private static readonly Regex regexPatentAssignee = CreateRegexForRefExtraction("PATENT ASSIGNEE(S)");
        private static readonly Regex regexAuthor = CreateRegexForRefExtraction("AUTHOR(S)");
        private static readonly Regex regexCorporateSource = CreateRegexForRefExtraction("CORPORATE SOURCE");
        private static readonly Regex regexSource = CreateRegexForRefExtraction("SOURCE");
        private static readonly Regex regexPublisher = CreateRegexForRefExtraction("PUBLISHER");
        private static readonly Regex regexLanguage = CreateRegexForRefExtraction("LANGUAGE");
        private static readonly Regex regexDocType = CreateRegexForRefExtraction("DOCUMENT TYPE");
        private static readonly Regex regexPatentInformation = new Regex(@"^PATENT INFORMATION\:\s*$", RegexOptions.Compiled);
        private static readonly Regex regexAbstract = new Regex(@"^ABSTRACT\:\s*$", RegexOptions.Compiled);
        private static readonly Regex regexGraphicsImage = new Regex(@"^GRAPHIC IMAGE\:\s*$", RegexOptions.Compiled);

        private static string PadForRefExtractor(string tag)
        {
            return (tag + ":" + new string(' ', 24 - tag.Length)).Replace("(", @"\(").Replace(")", @"\)").Replace(":", @"\:");
        }

        private static Regex CreateRegexForRefExtraction(string tag)
        {
            var tagLength = tag.Length;
            if (tagLength >= 25)
                throw new ArgumentException();
            tag = PadForRefExtractor(tag);
            return new Regex(tag + @"(?<" + GroupName_Text + @">[^ ].*)$", RegexOptions.Compiled);
        }

        public string Source
        {
            get;
            private set;
        }

        public CAOLExtractor(string source)
        {
            this.Source = Path.GetFullPath(source);
        }

        public Action ProgressIncrementer = null;

        private WordprocessingDocument wd;
        private TempDirectory tempdir = null;

        ~CAOLExtractor()
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

            wd = WordprocessingDocument.Open(docx, false);
            var body = wd.MainDocumentPart.Document.Body;

            var ri = new AReferenceInfo();
            var substances = new List<SubstanceInfo>();
            var si = new ASubstanceInfo();

            foreach (var paragraph in body.Elements<W.Paragraph>())
            {
                var text = paragraph.InnerText; // Do not trim here.
                if (text == "")
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
            }
        }



        private ReferenceInfo GetReferenceInfo(MSWord.Selection selection, ref int cursor, ref string copyright)
        {
            if (cursor < 0)
                return null;

            selection.SetRange(cursor, cursor);

            while (cursor >= 0)
            {
                ReferenceInfo ri = new ReferenceInfo();
                var substances = new List<SubstanceInfo>();
                SubstanceInfo si = new SubstanceInfo();

                for (; ; )
                {
                    {
                        int nextPos = WordUtility.SelectLine(selection);
                        if (nextPos < 0)
                        {
                            cursor = -1;
                            break;
                        }
                        {
                            bool hasShapes = EvalAsBitmapLineAndSet(selection, ri, si);
                            if (hasShapes)
                            {
                                selection.SetRange(nextPos, nextPos);
                                continue;
                            }
                        }
                    }

                    string lineText;
                    lineText = WordUtility.GetLine(selection);
                    if (lineText == null)
                    {
                        // reaching EOF
                        cursor = -1;
                        break;
                    }

                    if (CheckAndReadHeaderPart(selection, lineText, ref cursor, out copyright))
                        break;

                    if (CheckAndReadReferenceInfoPart(selection, lineText, ri)) 
                        continue;
                    
                    if (CheckAndReadSubstanceInfoPart(selection, lineText, substances, ref si)) 
                        continue;
                }

                if (!IsEmpty(si))
                {
                    substances.Add(si);

                    if (ProgressIncrementer != null)
                        ProgressIncrementer();
                }
                ri.SubstancesInfo = substances;
                if (!IsEmpty(ri))
                {
                    ri.Copyright = copyright;
                    return ri;
                }
            }
            return null;
        }

        private bool CheckAndReadHeaderPart(MSWord.Selection selection, string lineText, ref int cursor, out string copyright)
        {
            Match ma;
            ma = regexRefHeaderLine.Match(lineText);
            if (ma.Success)
            {
                copyright = ma.Groups[GroupName_Copyright].Value;
                cursor = selection.Start;
                return true;
            }

            if (regexHeaderLine.IsMatch(lineText))
            {
                copyright = ma.Groups[GroupName_Copyright].Value;
                cursor = selection.Start;
                return true;
            }

            copyright = null;
            return false;
        }

        private bool CheckAndReadReferenceInfoPart(MSWord.Selection selection, string lineText, ReferenceInfo ri)
        {
            if (ri.AccessionNumber == null)
            {
                var ma = regexAccessionNumber.Match(lineText);
                if (ma.Success)
                {
                    ri.AccessionNumber = ma.Groups[GroupName_Text].Value;
                    return true;
                }
            }

            if (Read5SpacedLinesProp(selection, lineText, ref ri._Title, regexTitleLine)) return true;
            if (Read5SpacedLinesProp(selection, lineText, ref ri._By, regexInventor)) return true;
            if (Read5SpacedLinesProp(selection, lineText, ref ri._PatentAssignee, regexPatentAssignee)) return true;
            if (Read5SpacedLinesProp(selection, lineText, ref ri._By, regexAuthor)) return true;
            if (Read5SpacedLinesProp(selection, lineText, ref ri._CorporateSource, regexCorporateSource)) return true;
            if (Read5SpacedLinesProp(selection, lineText, ref ri._Source, regexSource)) return true;
            if (Read5SpacedLinesProp(selection, lineText, ref ri._Publisher, regexPublisher)) return true;
            if (Read5SpacedLinesProp(selection, lineText, ref ri._Language, regexLanguage)) return true;
            if (ReadPatentInfomation(selection, lineText, ri)) return true;
            if (ReadAbstractLikeProp(selection, lineText, ref ri._Abstract, regexAbstract)) return true;

            return false;
        }

        private bool CheckAndReadSubstanceInfoPart(MSWord.Selection selection, string text, List<SubstanceInfo> substances, ref SubstanceInfo si)
        {
            var ma = regexRNLine.Match(text);
            if (ma.Success)
            {
                if (!IsEmpty(si))
                {
                    substances.Add(si);

                    if (ProgressIncrementer != null)
                        ProgressIncrementer();

                    si = new SubstanceInfo();
                }
                si.CASRN = ma.Groups[GroupName_Text].Value.Trim();
                return true;
            }

            if (Read5SpacedLinesProp(selection, text, ref si._CAIndexName, regexINFirstLine))
            {
                si._CAIndexName = NormalizeCASIN(si._CAIndexName);
                return true;
            }

            return false;
        }

        private static bool IsEmpty(ReferenceInfo ri)
        {
            //return ri.AccessionNumber == null && !(ri.SubstancesInfo != null && ri.SubstancesInfo.Count() != 0);
            return ri.AccessionNumber == null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="si"></param>
        /// <returns><value>true</value> if <paramref name="selection"/> has some InlineShapes.</returns>
        private static bool EvalAsBitmapLineAndSet(MSWord.Selection selection, ReferenceInfo ri, SubstanceInfo si)
        {
            MSWord.InlineShapes inlineShapes = selection.InlineShapes;
            try
            {
                if (inlineShapes.Count > 0)
                {
                    byte[] bitmap;
                    var inlineShape = inlineShapes[1];
                    try
                    {
                        bitmap = WordUtility.ReadAsImage(inlineShape);
                    }
                    finally
                    {
                        Utility.ReleaseComObject(inlineShape);
                    }

                    if (IsEmpty(si))
                    {
                        if (!IsEmpty(ri) && ri.AbstractImage == null)
                        {
                            ri.AbstractImage = bitmap;
                        }
                    }
                    else
                    {
                        if (si.Bitmap == null)
                        {
                            si.Bitmap = bitmap;
                        }
                    }
                    return true;
                }
                return false;
            }
            finally
            {
                Utility.ReleaseComObject(inlineShapes);
            }
        }

        private static bool ReadAbstractLikeProp(MSWord.Selection selection, string text, ref string p, Regex regex)
        {
            if (p == null)
            {
                var ma = regex.Match(text);
                if (ma.Success)
                {
                    var la = new LineAppender();
                    for (; ; )
                    {
                        string line = WordUtility.GetLine(selection);
                        if (line == null || regexOneSpaceLine.IsMatch(line))
                            break;
                        la.Append(line);
                    }
                    p = la.ToString();
                    return true;
                }
            }
            return false;
        }

        private static bool ReadPatentInfomation(MSWord.Selection selection, string text, ReferenceInfo ri)
        {
            if (ri.PatentInfomation == null)
            {
                var ma = regexPatentInformation.Match(text);
                if (ma.Success)
                {
                    string line = WordUtility.GetLine(selection);
                    if (!regexOneSpaceLine.IsMatch(line))
                        return true;

                    var pi = ReadPILines(selection);
                    ri.PatentInfomation = LinesToString(pi);
                    if (pi.Count >= 3)
                        ri.Source = pi[2];

                    return true;
                }
            }
            return false;
        }

        private static bool Read5SpacedLinesProp(MSWord.Selection selection, string text, ref string p, Regex regex)
        {
            if (p == null)
            {
                var ma = regex.Match(text);
                if (ma.Success)
                {
                    var s = ReadINLines(selection, ma);
                    p = s;
                    return true;
                }
            }
            return false;
        }

        private static IEnumerable<string> ReadContLines(MSWord.Selection selection, string firstLine)
        {
            if (firstLine != null)
                yield return firstLine;
            for (; ; )
            {
                var saveStart = selection.Start;
                var positionOfNextLine = WordUtility.SelectLine(selection);
                if (positionOfNextLine < 0)
                    break; // reaching EOF

                var line = selection.Text.TrimEnd();
                var lineMatch = regexContinurousLine.Match(line);

                if (lineMatch.Success)
                {
                    selection.SetRange(positionOfNextLine, positionOfNextLine);
                    yield return lineMatch.Groups[GroupName_Text].Value.Trim();
                }
                else
                {
                    selection.SetRange(saveStart, saveStart);
                    break;
                }
            }
            yield break;
        }

        private static string LinesToString(IEnumerable<string> lines)
        {
            var sb = new StringBuilder();
            foreach (var line in lines)
                sb.Append(line).Append('\n');
            return sb.ToString();
        }

        private static IList<string> ReadPILines(MSWord.Selection selection)
        {
            return ReadContLines(selection, null).ToList();
        }

        internal class LineAppender
        {
            private StringBuilder sb;
            private string prevLine;

            public LineAppender()
            {
                sb = new StringBuilder();
                prevLine = null;
            }

            public void Append(IEnumerable<string> lines)
            {
                foreach (var line in lines)
                    this.Append(line);
            }

            public void Append(string line)
            {
                if (line == null)
                    return;
                if (prevLine != null && !prevLine.EndsWith("-"))
                    sb.Append(" ");
                sb.Append(line);
                prevLine = line;
            }

            public override string ToString()
            {
                return sb.ToString();
            }
        }

        private static string ReadINLines(MSWord.Selection selection, Match firstLineMatch)
        {
            var firstLine = firstLineMatch.Groups[GroupName_Text].Value.Trim();
            var lines = ReadContLines(selection, firstLine);
            var la = new LineAppender();
            la.Append(lines);
            return la.ToString();
        }

        private static string NormalizeCASIN(string s)
        {
            if (s == "INDEX NAME NOT YET ASSIGNED")
            {
                s = "";
            }
            else if (s.EndsWith("(CA INDEX NAME)"))
            {
                s = s.Remove(s.Length - "(CA INDEX NAME)".Length).TrimEnd();
            }
            return s;
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

    }
}

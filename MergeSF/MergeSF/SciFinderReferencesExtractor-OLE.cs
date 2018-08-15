using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using MSWord = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Ujihara.Chemistry.MSOffice;

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
        private static readonly Regex reBy = MakeRE("By", ".*", true);
        private static readonly Regex reAssignee = MakeRE("Assignee", ".*", true);
        private static readonly Regex rePatentInformation = MakeRE("Patent Information", ".*", true);
        private static readonly Regex reSource = MakeRE("Source", ".*", true);
        private static readonly Regex reAccessionNumber = MakeRE("Accession Number", @"\d\d\d\d\:?\d+", false);
        private static readonly Regex reLanguage = MakeRE("Language", ".*", true);
        private static readonly Regex reCompany = MakeRE(@"Company\/Organization", ".*", true);
        private static readonly Regex rePublisher = MakeRE("Publisher", ".*", true);
        private const string Text_Abstract = "Abstract";
        private const string Text_PatentInfomation = "Patent Information";
        private static readonly Regex reIsAbstractTable = new Regex(@"^" + Text_Abstract + "$", RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex reIsPatentInfomation = new Regex(@"^" + Text_PatentInfomation + "$", RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex reIsJournal = new Regex(@" Journal(\, \d\d\d\d\,|\; [ A-Za-z0-9]*\, \d\d\d\d\,)", RegexOptions.Compiled);

        private const string GroupName_CASRN = "casrn";
        private static readonly Regex reCASRN_A = new Regex(@"\s*(?<" + GroupName_CASRN + @">\d+\-\d\d\-\d)[A-Z]\s*", RegexOptions.Compiled);

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

        /// <summary>
        /// Used in ExtractSubstances method.
        /// </summary>
        private static bool IsCompoundLine(string line)
        {
            return char.IsDigit(line.Length > 0 ? line[0] : '\0');
        }

        /// <summary>
        /// Used in ExtractSubstanceLines method.
        /// </summary>
        private enum SubstanceLinesReadMode
        {
            Compound,
            Subject,
            Class,
        }

        private enum ThreeState
        {
            ItsNextRef,
            TableFound,
            NotTable,
        }
       
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cursor"><value>-1</value>means already found eof.</param>
        /// <returns></returns>
        private ReferenceInfo GetReferenceInfo(MSWord.Selection selection, ref int cursor)
        {
            if (cursor < 0)
                return null;

            ReferenceInfo ri = new ReferenceInfo();
            var lines = new List<object>();

            selection.SetRange(cursor, cursor);

            for (; ; )
            {
                {
                    var t = AnalyzeInTable(selection, ri, ref cursor);
                    if (t == ThreeState.ItsNextRef)
                        break;
                    if (t == ThreeState.TableFound)
                        continue;
                }

                string lineText;
                {
                    var l = WordUtility.SelectLine(selection);
                    if (l < 0)
                    {
                        cursor = -1;
                        break;
                    }

                    if (ri.AbstractImage == null)
                    {
                        var im = GetImageIfCenteredImage(selection);
                        if (im != null)
                        {
                            ri.AbstractImage = im;
                            selection.SetRange(l, l);
                            continue;
                        }
                    }

                    lineText = selection.Text;
                    selection.SetRange(l, l);
                }

                if (lineText.StartsWith("Copyright"))
                {
                    // All records end with copyright line.
                    ri.Copyright = lineText;
                    cursor = selection.Start;
                    break;
                }
                else if (lineText == "Substances")
                {
                    for (; ; )
                    {
                        if (HasTable(selection))
                        {
                            cursor = selection.Start;
                            break;
                        }
                        if (AddLine(selection, lines) < 0)
                        {
                            cursor = -1;
                            break;
                        }
                    }
                }
            }
            if (ri.Title == null)
                return null;

            ri.SubstancesInfo = ToSubstancesInfo(lines);

            return ri;
        }

        private static byte[] GetImageIfCenteredImage(MSWord.Selection selection)
        {
            var pf = selection.ParagraphFormat;
            try
            {
                var al = pf.Alignment;
                if (al == Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter)
                {
                    MSWord.InlineShapes inlineShapes = selection.InlineShapes;
                    try
                    {
                        if (inlineShapes.Count > 0)
                        {
                            var inlineShape = inlineShapes[1];
                            try
                            {
                                return WordUtility.ReadAsImage(inlineShape);
                            }
                            finally
                            {
                                Utility.ReleaseComObject(inlineShape);
                            }
                        }
                    }
                    finally
                    {
                        Utility.ReleaseComObject(inlineShapes);
                    }
                }
            }
            finally
            {
                Utility.ReleaseComObject(pf);
            }
            return null;
        }

        private int AddLine(MSWord.Selection selection, List<object> lines)
        {
            var l = WordUtility.SelectLine(selection);
            if (l < 0)
                return -1;

            // All picture related objects are centered 
            var pf = selection.ParagraphFormat;
            try
            {
                var al = pf.Alignment;
                if (al == MSWord.WdParagraphAlignment.wdAlignParagraphCenter)
                {
                    MSWord.InlineShapes inlineShapes = selection.InlineShapes;
                    try
                    {
                        if (inlineShapes.Count > 0)
                        {
                            var inlineShape = inlineShapes[1];
                            try
                            {
                                lines.Add(WordUtility.ReadAsImage(inlineShape));
                            }
                            finally
                            {
                                Utility.ReleaseComObject(inlineShape);
                            }
                        }
                    }
                    finally
                    {
                        Utility.ReleaseComObject(inlineShapes);
                    }
                }
                else
                {
                    lines.Add(selection.Text);

                    if (ProgressIncrementer != null)
                        ProgressIncrementer();
                }
            }
            finally
            {
                Utility.ReleaseComObject(pf);
            }
            selection.SetRange(l, l);

            return l;
        }

        private static Regex reIsCasRNInTable = new Regex(@"^\d+\-\d\d\-\d\r", RegexOptions.Compiled);

        private static bool HasTable(MSWord.Selection selection)
        {
            var tables = selection.Tables;
            try
            {
                if (tables.Count > 0)
                {
                    MSWord.Table table9 = null;
                    MSWord.Range range_table9 = null;
                    MSWord.Cells cells = null;
                    MSWord.Cell cell = null;
                    MSWord.Range cell_range = null;
                    try
                    {
                        table9 = tables[tables.Count];
                        range_table9 = table9.Range;
                        cells = range_table9.Cells;
                        cell = cells[1];
                        cell_range = cell.Range;
                        var text = cell_range.Text.Replace("\r\a", "\n");
                        if (reIsCasRNInTable.IsMatch(text))
                        {
                            // skip multi components
                            selection.SetRange(range_table9.End, range_table9.End);
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        Utility.ReleaseComObject(cell_range);
                        Utility.ReleaseComObject(cell);
                        Utility.ReleaseComObject(cells);
                        Utility.ReleaseComObject(range_table9);
                        Utility.ReleaseComObject(table9);
                    }
                }
                return false;
            }
            finally
            {
                Utility.ReleaseComObject(tables);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="ri"></param>
        /// <returns>Found eof</returns>
        private static ThreeState AnalyzeInTable(MSWord.Selection selection, ReferenceInfo ri, ref int cursor)
        {
            MSWord.Tables tables = selection.Tables;
            try
            {
                if (tables.Count == 0)
                {
                    cursor = selection.Start;
                    return ThreeState.NotTable;
                }

                MSWord.Table lastTable = tables[tables.Count];
                var lastTableRange = lastTable.Range;
                try
                {
                    var textInTable = lastTableRange.Text.Replace("\r\a", "\n");
                    if (reIsTitleTable.IsMatch(textInTable))
                    {
                        if (ri.Title != null)
                        {
                            // it mean cursor is reaching to next ref.
                            cursor = selection.Start;
                            return ThreeState.ItsNextRef;
                        }
                        SetValuesInAbbevTable(ri, textInTable);
                    }
                    else if (reIsAbstractTable.IsMatch(textInTable))
                    {
                        int n = (Text_Abstract + "\n\n").Length;
                        ri.Abstract = textInTable.Substring(n);
                    }
                    else if (reIsPatentInfomation.IsMatch(textInTable))
                    {
                        int n = (Text_PatentInfomation + "\n\n").Length;
                        ri.PatentInfomation = textInTable.Substring(n);
                    }

                    selection.SetRange(lastTableRange.End, lastTableRange.End + 1);
                    cursor = selection.Start;
                }
                finally
                {
                    Utility.ReleaseComObject(lastTableRange);
                    Utility.ReleaseComObject(lastTable);
                }
            }
            finally
            {
                Utility.ReleaseComObject(tables);
            }
            return ThreeState.TableFound;
        }

        private static void SetValuesInAbbevTable(ReferenceInfo ri, string textInTable)
        {
            ri.Title = ExtractPropertyInAbb(reTitleLine, textInTable);
            ri.AccessionNumber = ExtractPropertyInAbb(reAccessionNumber, textInTable);
            ri.PatentAssignee = ExtractPropertyInAbb(reAssignee, textInTable);
            ri.CorporateSource = ExtractPropertyInAbb(reCompany, textInTable);
            ri.By = ExtractPropertyInAbb(reBy, textInTable);
            ri.Publisher = ExtractPropertyInAbb(rePublisher, textInTable);
            ri.Language = ExtractPropertyInAbb(reLanguage, textInTable);
            ri.Source = ExtractPropertyInAbb(rePatentInformation, textInTable);
            if (ri.Source != null)
            {
                ri.DocumentType = CmpdDbManager.DocumentType_Patent;
            }
            else
            {
                var s = ExtractPropertyInAbb(reSource, textInTable);
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

        private static string ExtractPropertyInAbb(Regex regex,  string text)
        {
            var ma = regex.Match(text);
            if (ma.Success)
                return ma.Groups[GroupName_Text].Value;
            return null;
        }

        /// <summary>
        /// Substances Info. Empty if <paramref name="filename"/> is null.
        /// </summary>
        private static IEnumerable<SubstanceInfo> ToSubstancesInfo(IList<object> lines)
        {
            var substancesInfo = new List<SubstanceInfo>();

            SubstanceLinesReadMode mode = SubstanceLinesReadMode.Class;
            string currClass = "";
            byte[] currBitmap = null;
            foreach (var raw_line_o in lines.Reverse<object>())
            {
                if (raw_line_o is byte[])
                {
                    currBitmap = (byte[])raw_line_o;
                }
                else if (raw_line_o is string)
                {
                    var line = ((string)raw_line_o).Trim();
                EvalAgain:
                    if (line == "" || line == "Double bond geometry as shown.")
                        continue;
                    var isCompoundLine = IsCompoundLine(line);
                    switch (mode)
                    {
                        case SubstanceLinesReadMode.Class:
                            currClass = line;
                            mode = SubstanceLinesReadMode.Subject;
                            break;
                        case SubstanceLinesReadMode.Subject:
                            mode = SubstanceLinesReadMode.Compound;
                            break;
                        case SubstanceLinesReadMode.Compound:
                            if (isCompoundLine)
                            {
                                var elm = new SubstanceInfo();
                                string c_casrn = null;
                                var firstSpaceIndex = line.IndexOf(' ');
                                if (firstSpaceIndex < 0)
                                {
                                    c_casrn = line;
                                    elm.Name = "";
                                }
                                else
                                {
                                    c_casrn = line.Substring(0, firstSpaceIndex);
                                    elm.Name = line.Substring(firstSpaceIndex + 1);
                                }
                                {
                                    var ma = reCASRN_A.Match(c_casrn);
                                    if (ma.Success)
                                    {
                                        elm.CASRN = ma.Groups[GroupName_CASRN].Value;
                                        elm.Keywords = currClass;
                                        elm.Bitmap = currBitmap;
                                        currBitmap = null;
                                        substancesInfo.Add(elm);
                                    }
                                }
                            }
                            else
                            {
                                mode = SubstanceLinesReadMode.Class;
                                goto EvalAgain;
                            }
                            break;
                        default:
                            break;
                    }
                }
            }

            {
                var currNumber = substancesInfo.Count;
                foreach (var elm in substancesInfo)
                {
                    elm.Order = currNumber;
                    currNumber--;
                }
            }
            return substancesInfo;
        }

        public Action ProgressIncrementer = null;

        public IEnumerator<ReferenceInfo> GetEnumerator()
        {
            if (this.Source == null)
                yield break;

            MSWord.Application app = null;
            MSWord.Documents docs = null;
            MSWord.Document doc = null;

            app = new MSWord.Application();
            //app.Visible = true;
            docs = app.Documents;
            object filename = Path.GetFullPath(this.Source);
            doc = docs.Open(ref filename, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            try
            {
                doc.Activate();
                MSWord.Selection selection = app.Selection;
                try
                {
                    WordUtility.NormalizeMacSymbol(app);

                    int cursor = 0;
                    for (; ; )
                    {
                        var refInfo = GetReferenceInfo(selection, ref cursor);
                        if (refInfo == null)
                            yield break;
                        yield return refInfo;
                    }
                }
                finally
                {
                    Utility.ReleaseComObject(selection);
                }
            }
            finally
            {
#pragma warning disable 467
                if (doc != null)
                {
                    doc.Close(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(doc);
                }
                Marshal.ReleaseComObject(docs);
                if (app != null)
                {
                    app.Quit(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(app);
                }
#pragma warning restore 467
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}

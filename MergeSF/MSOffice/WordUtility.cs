using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSWord = Microsoft.Office.Interop.Word;
using System.Xml;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ujihara.Chemistry.MSOffice
{
    public class WordUtility
    {
        private static object unitWdParagraph = MSWord.WdUnits.wdParagraph;
        private static object unitWdLine = MSWord.WdUnits.wdLine;
        private static object missing = Type.Missing;

        public static byte[] ExtractGraphicPart(OpenXmlPartContainer Container, A.Graphic Graphic)
        {
            if (Graphic != null)
            {
                var blip = Graphic.Descendants<A.Blip>().FirstOrDefault();
                if (!(blip == null || !blip.Embed.HasValue))
                {
                    var part = Container.GetPartById(blip.Embed.Value);
                    using (var mem = new MemoryStream())
                    {
                        using (var srm = part.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            int c;
                            while ((c = srm.ReadByte()) != -1)
                                mem.WriteByte((byte)c);
                        }
                        return mem.ToArray();
                    }
                }
            }
            return null;
        }

        public static string ToRtf(OpenXmlElement cell)
        {
            return ToRtf(cell, null, 0);
        }

        public static string ToRtf(OpenXmlElement cell, string fontName, int size2)
        {
            var sb = new StringBuilder();
            sb.Append(@"{\rtf1");
            if (fontName != null)
                sb.Append(@"\deff0{\fonttbl{\f0 ").Append(fontName).Append(@";}}\f0 ");
            if (size2 != 0)
                sb.Append(@"\fs").Append(size2).Append(" ");
            foreach (var paragraph in cell.Elements<W.Paragraph>())
            {
                foreach (var run in paragraph.Elements<W.Run>())
                {
                    sb.Append(ToRtfTag(run.RunProperties, false));
                    foreach (var c in run.InnerText)
                    {
                        switch (c)
                        {
                            case '"':
                                sb.Append(c).Append(c);
                                break;
                            case '{':
                            case '}':
                            case '\\':
                                sb.Append('\\').Append(c);
                                break;
                            default:
                                if (0x20 <= (int)c && (int)c <= 0x7e)
                                    sb.Append(c);
                                else
                                    sb.Append(@"\u").Append((uint)c).Append('?');
                                break;
                        }
                    }
                    sb.Append(ToRtfTag(run.RunProperties, true));
                }
                sb.Append(@"\par ");

            }
            sb.Append('}');
            return sb.ToString();
        }

        private static string ToRtfTag(W.RunProperties prop, bool isClose)
        {
            if (prop == null) return "";
            var l = new List<string>();
            if (prop.Italic != null) l.Add("i");
            if (prop.Underline != null) l.Add("ul");
            if (prop.Bold != null) l.Add("b");
            if (prop.VerticalTextAlignment != null && prop.VerticalTextAlignment.Val.HasValue)
            {
                switch (prop.VerticalTextAlignment.Val.Value)
                {
                    case W.VerticalPositionValues.Superscript:
                        l.Add("super");
                        break;
                    case W.VerticalPositionValues.Subscript:
                        l.Add("sub");
                        break;
                    default:
                        break;
                }
            }
            if (isClose) l.Reverse();
            var sb = new StringBuilder();
            foreach (var e in l)
            {
                if (isClose)
                {
                    sb.Append('}');
                }
                else
                {
                    sb.Append(@"{\").Append(e).Append(' ');
                }
            }
            return sb.ToString();
        }

        private static string ToHtmlTag(W.RunProperties prop, bool isClose)
        {
            if (prop == null) return "";
            var l = new List<string>();
            if (prop.Italic != null) l.Add("i");
            if (prop.Underline != null) l.Add("u");
            if (prop.Bold != null) l.Add("b");
            if (prop.Emphasis != null) l.Add("em");
            if (prop.VerticalTextAlignment != null && prop.VerticalTextAlignment.Val.HasValue)
            {
                switch (prop.VerticalTextAlignment.Val.Value)
                {
                    case W.VerticalPositionValues.Superscript:
                        l.Add("sup");
                        break;
                    case W.VerticalPositionValues.Subscript:
                        l.Add("sub");
                        break;
                    default:
                        break;
                }
            }
            if (isClose) l.Reverse();
            var sb = new StringBuilder();
            foreach (var e in l)
            {
                sb.Append('<');
                if (isClose) sb.Append('/');
                sb.Append(e).Append('>');
            }
            return sb.ToString();
        }

        public static void ConvertToDocx(string original, string docx)
        {
            object filename;

            MSWord.Application app = null;
            MSWord.Documents docs = null;
            MSWord.Document doc = null;

            app = new MSWord.Application();
            //app.Visible = true;
            docs = app.Documents;
            filename = Path.GetFullPath(original);
            doc = docs.Open(ref filename, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);
            try
            {
                doc.Activate();
                filename = docx;
                object format = MSWord.WdSaveFormat.wdFormatXMLDocument;
                object false_ = false;
                object true_ = true;
                object empty = "";
                object forteen = 14;
                doc.SaveAs2(ref filename, ref format, ref false_,
                    ref empty, ref true_, ref empty, ref false_, ref false_,
                    ref false_, ref false_, ref false_, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref forteen);
            }
            finally
            {
#pragma warning disable 467
                if (doc != null)
                {
                    doc.Close(ref missing, ref missing, ref missing);
                    Utility.ReleaseComObject(doc);
                }
                Utility.ReleaseComObject(docs);
                if (app != null)
                {
                    app.Quit(ref missing, ref missing, ref missing);
                    Utility.ReleaseComObject(app);
                }
#pragma warning restore 467
            }
        }

        public static void NormalizeMacSymbol(MSWord.Application app)
        {
            var selection = app.Selection;
            try
            {
                var saveStart = selection.Start;
                var saveEnd = selection.End;

                selection.Start = 0;
                selection.End = 0;
                NormalizeMacSymbol(selection);

                selection.SetRange(saveStart, saveEnd);
            }
            finally
            {
                Utility.ReleaseComObject(selection);
            }
        }

        private const int MacRightArrow = 61614;
        private const int MacSmallAlphaCodeInMSWord = 61537;    //61537 is charactor code of alpha in Microsoft Word for Mac
        private static void NormalizeMacSymbol(MSWord.Selection selection)
        {
            var find = selection.Find;
            var replacement = find.Replacement;
            try
            {
                find.ClearFormatting();
                find.ClearAllFuzzyOptions();
                find.MatchCase = true;
                for (var i = 0; i < 24; i++)
                {
                    FindAndReplace(find, replacement, new string((char)(MacSmallAlphaCodeInMSWord + i), 1), new string((char)((int)'α' + i), 1));
                }
                FindAndReplace(find, replacement, new string((char)MacRightArrow, 1), "→");
            }
            finally
            {
                Utility.ReleaseComObject(replacement);
                Utility.ReleaseComObject(find);
            }
        }

        private static void FindAndReplace(MSWord.Find find, MSWord.Replacement replacement, string findchar, string replaced)
        {
            find.Text = findchar;
            replacement.Text = replaced;
            object replace = MSWord.WdReplace.wdReplaceAll;
            find.Execute(ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref replace,
                ref missing, ref missing, ref missing, ref missing);
        }

        public static byte[] ReadAsImage(MSWord.InlineShape inlineShape)
        {
            const string ns_pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

            XmlDocument xmlShape = new XmlDocument();
            xmlShape.LoadXml(inlineShape.Range.WordOpenXML);
            XmlNamespaceManager nm = new XmlNamespaceManager(xmlShape.NameTable);
            nm.AddNamespace("pkg", ns_pkg);

            var parts = xmlShape.DocumentElement.SelectNodes("//pkg:part[@pkg:contentType]", nm);
            foreach (XmlElement n in parts)
            {
                var ct = n.GetAttribute("contentType", ns_pkg);
                if (ct != null && ct.StartsWith("image/"))
                {
                    var bd = n.SelectSingleNode("//pkg:binaryData", nm);
                    if (bd == null)
                        return null;

                    return Convert.FromBase64String(bd.InnerText);
                }
            }
            return null;
        }


        private static object StaticObjectOne = 1;

        /// <summary>
        /// Select line on <paramref name="selection"/>.
        /// </summary>
        /// <param name="selection"></param>
        /// <returns>Position of next line. <value>-1</value> on reaching EOF.</returns>
        public static int SelectLine(MSWord.Selection selection)
        {
            var saveStart = selection.Start;
            var ll = selection.MoveDown(ref unitWdParagraph, ref StaticObjectOne, ref missing);
            if (ll == 0)
                return -1; // reaching EOF
            if (!(selection.Start > saveStart))
                return -1;
            var saveEnd = selection.End;
            selection.SetRange(saveStart, saveEnd - 1);
            return saveEnd;
        }

        /// <summary>
        /// Returns current line as text and <paramref name="selection"/> moved to the top of next line.
        /// </summary>
        /// <param name="selection"></param>
        /// <returns>Text of current line. <value>null</value> on reaching EOF.</returns>
        public static string GetLine(MSWord.Selection selection)
        {
            var l = SelectLine(selection);
            if (l < 0)
                return null;
            var text = selection.Text;
            selection.SetRange(l, l);
            return text;
        }
    }
}

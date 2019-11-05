using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ChemFinder = ChemFinder19;
using MolServer = MolServer19;

namespace Ujihara.Chemistry
{
    public class CfxManager
    {
        public const string DefaultTableName = "MolTable";

        public const string FieldName_MolID = "Mol_ID";
        public const string FieldName_Structure = "Structure";

        public const string ControlType_DataBox = "Data Box";
        public const string ControlType_RichText = "Rich Text";
        public const string ControlType_Structure = "Structure";
        public const string ControlType_Picture = "Picture";
        public const string ControlType_Subform = "Subform";

        public const string ElementName_form = "form";
        public const string ElementName_formprops = "formprops";
        public const string ElementName_boxes = "boxes";
        public const string ElementName_connections = "connections";
        public const string ElementName_box = "box";
        public const string ElementName_font = "font";
        public const string ElementName_tabs = "tabs";
        public const string ElementName_tab = "tab";
        public const string AttrivuteName_nomstupgrade = "nomstupgrade";
        public const string AttributeName_dsn = "dsn";
        public const string AttributeName_tablename = "tablename";
        public const string AttributeName_filename = "filename";
        public const string AttributeName_mstname = "mstname";
        public const string AttributeName_orasettings = "orasettings";
        public const string AttributeName_type = "type";
        public const string AttributeName_rect = "rect";
        public const string AttributeName_field = "field";
        public const string AttributeName_linktocol = "linktocol";
        public const string AttributeName_name = "name";
        public const string AttributeName_size = "size";
        public const string AttributeName_viewtype = "viewtype";
        public const string AttributeName_grid = "grid";
        public const string AttributeName_tabno = "tabno";

        const int HDiffOfBox = 8;
        const int VDiffOfBox = 4;
        const int WidthOfLabel = 144;
        const int WidthOfBox = 480;
        const int HeightOfBox = 32;
        const int HeightOfPictureBox = 120;
        const int HeightOfSubform = 360;
        const int LeftOfSubform = 0;

        const int TopOfFirstBox = VDiffOfBox;
        const int LeftOfLabel = HDiffOfBox;
        const int LeftOfBox = LeftOfLabel + WidthOfLabel + HDiffOfBox;
        const int DiffLabelAndBox = LeftOfBox - LeftOfLabel - WidthOfLabel;
        const int WidthOfSubForm = LeftOfBox + WidthOfBox;

        /// <summary>
        /// 
        /// </summary>
        /// <remarks>
        /// "Microsoft.Jet.OLEDB.4.0" supports only 32 bit but ChemFinder supports only Jet database.
        /// </remarks>
        public const string DefaultProviderName = "Microsoft.Jet.OLEDB.4.0";
        public string ProviderName { get; private set; }

        public XDocument Document { get; private set; }

        public CfxManager()
        {
            this.ProviderName = DefaultProviderName;
        }

        public void Create(string path)
        {
            var cfxFullPath = Path.GetFullPath(path);

            ChemFinder.Application app = null;
            ChemFinder.Documents docs = null;
            ChemFinder.Document doc = null;
            try
            {
                app = new ChemFinder.Application();
                docs = app.Documents;
                doc = docs.Add();
                doc.SaveAs(cfxFullPath);
            }
            finally
            {
                if (doc != null)
                    doc.Close(Type.Missing, Type.Missing);
                if (app != null)
                    app.Quit();
                Utility.ReleaseComObject(doc);
                Utility.ReleaseComObject(docs);
                Utility.ReleaseComObject(app);
            }
        }

        public void CreateWithDefault(string path)
        {
            var cfxFullPath = Path.GetFullPath(path);
            var cfxDirName = Path.GetDirectoryName(cfxFullPath);

            var cfxFileName = Path.GetFileName(cfxFullPath);
            var theName = Path.GetFileNameWithoutExtension(Path.GetFileName(cfxFileName));
            var mdbFileName = theName + ".mdb";
            var mstFileName = theName + ".mst";
            var mdbFullName = Path.Combine(cfxDirName, mdbFileName);
            var mstFullName = Path.Combine(cfxDirName, mstFileName);

            CreateDefaultMdb(mdbFullName);
            CreateDefaultMst(mstFullName);  // MolServer 13 does not delete cfx file but MolServer 14 does 

            Create(cfxFullPath);

            var man = new CfxManager();
            man.Load(cfxFullPath);
            man.BindDatabase(mdbFileName, DefaultTableName, mstFileName, man.GetRootForm());
            man.AddField(FieldName_Structure, ControlType_Structure);
            man.Save(cfxFullPath);
        }

        private static void CreateDefaultMst(string mstPath)
        {
            var mstFullName = Path.GetFullPath(mstPath);
            var doc = new MolServer.Document();
            try
            {
                doc.Create(mstFullName, 80);    //max supported mst in ChemFinder 18
                doc.Close();
            }
            finally
            {
                Utility.ReleaseComObject(doc);
            }
        }

        private void CreateDefaultMdb(string mdbPath)
        {
            var mdbFullPath = Path.GetFullPath(mdbPath);
            {
                var cat = new ADOX.Catalog();
                try
                {
                    cat.Create("Provider=" + ProviderName + ";"
                            + "Data Source="
                            + mdbFullPath);
                }
                finally
                {
                    Utility.ReleaseComObject(cat);
                }
            }

            using (var conn = new System.Data.OleDb.OleDbConnection())
            {
                conn.ConnectionString =
                     "Provider=" + ProviderName + ";"
                   + "Data Source=" + mdbFullPath;
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    string sql = "CREATE TABLE "
                    + DefaultTableName + " "
                    + "( "
                    + FieldName_MolID + " INTEGER, "
                    + FieldName_Structure + " IMAGE "
                    + ") ";
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        public void Load(string path)
        {
            this.Document = XDocument.Load(Path.GetFullPath(path));
            if (this.Document.Root.Name != ElementName_form)
                throw new FormatException("Bad image of '" + path + "'.");
        }

        public void Save(string path)
        {
            var fullPath = Path.GetFullPath(path);
            using (var srm = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
            {
                Encoding encoding;
                switch (this.Document.Declaration.Encoding.ToUpper())
                {
                    case "UTF-8":
                        // ChemFinder does not support BOM.
                        encoding = new UTF8Encoding(false);
                        break;
                    default:
                        encoding = Encoding.GetEncoding(this.Document.Declaration.Encoding);
                        break;
                }
                using (var w = new StreamWriter(srm, encoding))
                {
                    w.Write(this.Document.Declaration.ToString());
                    w.Write(this.Document.ToString());
                }
            }
        }

        public XElement AddField(string fieldName)
        {
            return AddField(fieldName, ControlType_DataBox);
        }

        public XElement GetRootForm()
        {
            var form = this.Document.Root;
            if (form.Name.LocalName != ElementName_form)
                throw new InvalidDataException();
            return form;
        }

        public XElement AddField(string fieldName, string type)
        {
            var form = GetRootForm();
            return AddField(fieldName, type, form);
        }

        public XElement AddSubform(string fieldName, XElement formParent, string linktocol)
        {
            var boxSubform = AddField(fieldName, ControlType_Subform, formParent);
            boxSubform.SetAttributeValue(AttributeName_linktocol, linktocol);
            //  <box type="Subform" rect="l:118,t:320,r:727,b:713" field="Accession_Number" linktocol="Accession_Number">
            var form = new XElement(ElementName_form);
            boxSubform.Add(form);
            var formprops = new XElement(ElementName_formprops);
            formprops.SetAttributeValue(AttributeName_type, "sub");
            form.Add(formprops);
            var boxes = new XElement(ElementName_boxes);
            form.Add(boxes);

            return form;
        }

        public XElement SetFont(XElement box, string name)
        {
            var elmFont = new XElement(ElementName_font);
            elmFont.SetAttributeValue(AttributeName_name, name);
            elmFont.SetAttributeValue(AttributeName_size, "8");
            box.Add(elmFont);
            return elmFont;
        }

        public XElement AddField(string fieldName, string type, XElement form)
        {
            return AddField(fieldName, type, form, 0);
        }

        /// <summary>
        /// Add field to ChemFinder form.
        /// </summary>
        /// <param name="fieldName">Field name in database table.</param>
        /// <param name="type"><value>Data Box</value>, <value>Rich Text</value> or <value>Structurte</value>.</param>
        /// <param name="form"></param>
        /// <param name="tabno"></param>
        /// <returns></returns>
        public XElement AddField(string fieldName, string type, XElement form, int tabno)
        {
            int heightOfControl;
            int widthOfControl = WidthOfBox;
            switch (type)
            {
                case ControlType_Structure:
                case ControlType_Picture:
                    heightOfControl = HeightOfPictureBox;
                    break;
                case ControlType_Subform:
                    heightOfControl = HeightOfSubform;
                    widthOfControl = WidthOfSubForm;
                    break;
                default:
                    heightOfControl = HeightOfBox;
                    break;
            }

            var boxes = form.Element(ElementName_boxes);
            int l, t, r, b;
            var boxs = boxes.Elements(ElementName_box).Where(n => n.Attribute(AttributeName_type).Value != "Plain Text");

            var boxsInTab = boxs.Where(n => tabno == 0 ?
                    (n.Attributes(AttributeName_tabno).Count() == 0 ||
                     n.Attribute(AttributeName_tabno).Value == "0") :
                    (n.Attributes(AttributeName_tabno).Count() != 0 &&
                     n.Attribute(AttributeName_tabno).Value == tabno.ToString()));
            var lastBox = boxsInTab.LastOrDefault();
            // Subform does not have a caption box
            var lastNonSubform = boxsInTab.LastOrDefault(n => n.Attribute(AttributeName_type).Value != "Subform");

            // l, r
            if (lastNonSubform == null)
            {
                l = LeftOfBox;
                r = l + widthOfControl;
            }
            else
            {
                var lastRect = lastNonSubform.Attribute(AttributeName_rect).Value;
                GetLTRBAttributeValue(lastRect, out l, out t, out r, out b);
                //l = l;
                r = l + widthOfControl;
            }

            // t, b
            if (lastBox == null || lastNonSubform == null)
            {
                t = TopOfFirstBox;
                b = t + heightOfControl;
            }
            else
            {
                var lastRect = lastNonSubform.Attribute(AttributeName_rect).Value;
                GetLTRBAttributeValue(lastRect, out l, out t, out r, out b);
                t = b + VDiffOfBox;
                b = t + heightOfControl;
            }

            if (type == ControlType_Subform)
            {
                l = 0;
            }

            if (type != ControlType_Subform)
            {
                var elementLabel = new XElement(ElementName_box);
                elementLabel.SetAttributeValue(AttributeName_type, "Plain Text");
                elementLabel.SetAttributeValue(AttributeName_rect, CreateLTRBAttributeValue(
                    l - DiffLabelAndBox - WidthOfLabel, 
                    t,
                    l - DiffLabelAndBox, 
                    b));
                elementLabel.SetAttributeValue("text", fieldName);
                elementLabel.SetAttributeValue("dtype", "fixed");
                SetFont(elementLabel, "Meiryo UI");
                elementLabel.SetAttributeValue(CfxManager.AttributeName_tabno, tabno.ToString());
                boxes.Add(elementLabel);
            }

            var elementBox = new XElement(ElementName_box);
            elementBox.SetAttributeValue(AttributeName_type, type);
            elementBox.SetAttributeValue(AttributeName_rect, CreateLTRBAttributeValue(l, t, r, b));
            elementBox.SetAttributeValue(AttributeName_field, fieldName);
            SetFont(elementBox, "Meiryo UI");
            elementBox.SetAttributeValue(CfxManager.AttributeName_tabno, tabno.ToString());
            boxes.Add(elementBox);

            return elementBox;
        }

        private static Regex LTRBRegex = new Regex(@"l\:(?<l>\d+)\,t\:(?<t>\d+)\,r\:(?<r>\d+)\,b\:(?<b>\d+)", RegexOptions.Compiled);

        public static void GetLTRBAttributeValue(string value, out int l, out int t, out int r, out int b)
        {
            var match = LTRBRegex.Match(value);
            if (!match.Success)
                throw new ArgumentException();
            l = int.Parse(match.Groups["l"].Value);
            t = int.Parse(match.Groups["t"].Value);
            r = int.Parse(match.Groups["r"].Value);
            b = int.Parse(match.Groups["b"].Value);
        }

        public static string CreateLTRBAttributeValue(int l, int t, int r, int b)
        {
            return "l:" + l.ToString() + ","
                 + "t:" + t.ToString() + ","
                 + "r:" + r.ToString() + ","
                 + "b:" + b.ToString();
        }

        public XElement FormProps
        {
            get { return Document.Root.Element(ElementName_formprops); }
        }

        public XElement Boxes
        {
            get { return Document.Root.Element(ElementName_boxes); }
        }

        public XElement Connections
        {
            get { return Document.Root.Element(ElementName_connections); }
        }

        public string DatabaseFileName
        {
            get { return Connections.Attribute(AttributeName_filename).Value; }
        }

        public string TableName
        {
            get 
            {
                var name = Connections.Attribute(AttributeName_tablename).Value;
                Utility.CheckSQLName(name);
                return name;
            }
        }

        /// <summary>
        /// Mst name. Returns <see langword="null"/> if mst is not specified.
        /// </summary>
        public string MstName
        {
            get
            {
                return Connections.Attribute(AttributeName_mstname)?.Value;
            }
        }

        private XElement RequestConnectionsElement(XElement formParent)
        {
            var elm = formParent.Elements(ElementName_connections).FirstOrDefault();
            if (elm == null)
            {
                elm = new XElement(ElementName_connections);
                formParent.Add(elm);
            }
            return elm;
        }

        public void BindDatabase(string mdbName, string tableName, XElement formParent)
        {
            var elm = RequestConnectionsElement(formParent);
            var dsn = Path.GetFileNameWithoutExtension(Path.GetFileName(mdbName));
            elm.SetAttributeValue(AttributeName_dsn, dsn);
            elm.SetAttributeValue(AttributeName_tablename, tableName);
            elm.SetAttributeValue(AttributeName_filename, mdbName);
            elm.SetAttributeValue(AttributeName_orasettings, "4");   // "4" is magic number for me.
        }

        public void BindMst(string mstName, XElement formParent)
        {
            var elm = RequestConnectionsElement(formParent);
            elm.SetAttributeValue(AttributeName_mstname, mstName);

            FormProps.SetAttributeValue(AttrivuteName_nomstupgrade, true);
        }

        public void BindDatabase(string mdbName, string tableName, string mstName, XElement formParent)
        {
            BindDatabase(mdbName, tableName, formParent);
            BindMst(mstName, formParent);
        }
    }
}

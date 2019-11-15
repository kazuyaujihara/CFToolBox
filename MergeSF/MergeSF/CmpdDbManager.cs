using System;
using System.IO;
using System.Xml.Linq;

namespace Ujihara.Chemistry.MergeSF
{
    public class CmpdDbManager
    {
        // Field name
        public const string CASRN_FieldName = "cas_rn";
        public const string MolID_FieldName = "Mol_ID";
        public const string Structure_FieldName = "Structure";
        public const string StructureSource_FieldName = "StructureSource";
        public const string MolFileName_FieldName = "MolFileName";
        public const string MolecularFormula_FieldName = "molecular_formula";
        public const string Copyright_FieldName = "copyright";
        public const string CasIndexName_FieldName = "cas_index_name";
        public const string LocalID_FieldName = "Local_ID";
        public const string StructureBmp_FieldName = "StructureBmp";
        public const string Smiles_FieldName = "SMILES";
        public const string InChi_FieldName = "InChi";
        public const string OtherNames_FieldName = "Other_Names";
        public const string ClassIdentifier_FieldName = "Class_Identifier";

        public const string AccessionNumber_FieldName = "Accession_Number";
        public const string Title_FieldName = "Title";
        //public const string Inventor_FieldName = "Inventor";
        //public const string PatentAssignee_FieldName = "Patent_Assignee";
        public const string Author_FieldName = "Author";
        public const string CorporateSource_FieldName = "Corporate_Source";
        public const string Source_FieldName = "Source";
        public const string Publisher_FieldName = "Publisher";
        public const string DocumentType_FieldName = "DocumentType";
        public const string Language_FieldName = "Language";
        public const string PatentInfomation_FieldName = "Patent_Infomation";
        public const string Abstract_FieldName = "Abstract";
        public const string AbstractBitmap_FieldName = "Abstract_Bitmap";

        public const string OrderInDoc_FieldName = "Order_In_Doc";
        public const string NameInDoc_FieldName = "Name_In_Doc";
        public const string Keywords_FieldName = "Keywords";
        public const string Tag_FieldName = "Tag";
        public const string Name_FieldName = "Name";

        // Field size
        public const int SizeOfCASRNField = 20;    // Base64 encoded GUID can be stored in CASRN field
        public const int SizeOfAccessionNumberField = 20;
        public const int SizeOfCopyrightField = 20;
        public const int SizeOfStructureSource = 20;

        // Table name

        public const string MolTable_TableName = "MolTable";
        public const string Substances_TableName = "Substances";
        public const string Documents_TableName = "Documents";
        public const string OtherNames_TableName = "OtherNames";

        // View name

        public const string MolTableWithInfo_ViewName = "MolTableWithInfo";
        public const string SubstancesWithInfo_ViewName = "SubstancesWithInfo";

        // Values in cfx 

        private const string FontName_Monospaced = "Consolas";

        // Values for DocumentType

        public const string DocumentType_Unknown = "Unknown";
        public const string DocumentType_Patent = "Patent";
        public const string DocumentType_Journal = "Journal";

        // Values for Structure Source

        public const string StructureSource_SDF = "SDF";
        public const string StructureSource_InChi = "InChi";
        public const string StructureSource_SMILES = "SMILES";
        public const string StructureSource_CAIndexName = "CAIndexName";
        public const string StructureSource_OtherName = "OtherName";
        public const string StructureSource_OSRA = "OSRA";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename">ChemFinder file name (.cfx)</param>
        public static void MakeNew(string filename)
        {
            var man = new CfxManager();

            string cfxFullPath = Path.GetFullPath(filename);

            man.CreateWithDefault(cfxFullPath);
            ConstructTemplateMdb(cfxFullPath);

            man.Load(cfxFullPath);

            string mdbName = man.DatabaseFileName;
            string mstName = man.MstName;

            man.BindDatabase(mdbName, CmpdDbManager.MolTable_TableName, man.GetRootForm());
            var rootFormProps = man.GetRootForm().Element(CfxManager.ElementName_formprops);
            if (rootFormProps != null)
            {
                // remove grid property to switch grid off
                var a = rootFormProps.Attributes(CfxManager.AttributeName_grid);
                if (a != null)
                    a.Remove();
            }
            man.AddField(CmpdDbManager.CASRN_FieldName);
            man.AddField(CmpdDbManager.CasIndexName_FieldName);
            man.AddField(CmpdDbManager.StructureBmp_FieldName, CfxManager.ControlType_Picture);
            man.AddField(CmpdDbManager.OtherNames_FieldName, CfxManager.ControlType_DataBox);

            var rf = man.GetRootForm();
            man.AddField(CmpdDbManager.MolID_FieldName, CfxManager.ControlType_DataBox, rf, 1);
            man.AddField(CmpdDbManager.Structure_FieldName, CfxManager.ControlType_Structure, rf, 1);
            foreach (var li in MolTableFieldsInfo)
            {
                var controlType = CfxManager.ControlType_DataBox;
                if (li[1].StartsWith("IMAGE"))
                    controlType = CfxManager.ControlType_Picture;
                man.AddField(li[0], controlType, rf, 1);
            }

            {
                var tabs = man.GetRootForm().Element(CfxManager.ElementName_tabs);
                var tab1 = new XElement(CfxManager.ElementName_tab);
                tabs.Add(tab1);
            }

            var subform = man.AddSubform(CmpdDbManager.CASRN_FieldName, man.GetRootForm(), CmpdDbManager.CASRN_FieldName);
            man.BindDatabase(mdbName, CmpdDbManager.SubstancesWithInfo_ViewName, subform);
            man.AddField(CmpdDbManager.AccessionNumber_FieldName, CfxManager.ControlType_DataBox, subform);
            man.AddField(CmpdDbManager.Title_FieldName, CfxManager.ControlType_DataBox, subform);
            man.AddField(CmpdDbManager.Source_FieldName, CfxManager.ControlType_DataBox, subform);
            man.AddField(CmpdDbManager.Author_FieldName, CfxManager.ControlType_DataBox, subform);
            man.AddField(CmpdDbManager.CorporateSource_FieldName, CfxManager.ControlType_DataBox, subform);
            man.AddField(CmpdDbManager.Keywords_FieldName, CfxManager.ControlType_DataBox, subform);
            {
                var tab = new XElement(CfxManager.ElementName_tab);
                var tabs = new XElement(CfxManager.ElementName_tabs, tab);
                man.GetRootForm().Add(tabs);
            }

            man.Save(cfxFullPath);
        }

        public static void MakeDocumentBaseForm(string filename, string mdbName, string mstName)
        {
            var man = new CfxManager();

            string cfxFullPath = Path.GetFullPath(filename);
            man.Create(cfxFullPath);
            man.Load(cfxFullPath);

            man.BindDatabase(mdbName, CmpdDbManager.Documents_TableName, man.GetRootForm());
            var rootFormProps = man.GetRootForm().Element(CfxManager.ElementName_formprops);
            if (rootFormProps != null)
            {
                var a = rootFormProps.Attributes(CfxManager.AttributeName_grid);
                if (a != null)
                    a.Remove();
            }
            man.AddField(CmpdDbManager.AccessionNumber_FieldName);
            man.AddField(CmpdDbManager.Title_FieldName);
            man.AddField(CmpdDbManager.Author_FieldName);
            man.AddField(CmpdDbManager.CorporateSource_FieldName);
            man.AddField(CmpdDbManager.Source_FieldName);
            {
                var pi_box = man.AddField(CmpdDbManager.Abstract_FieldName);
                var ltrb = pi_box.Attribute(CfxManager.AttributeName_rect).Value;
                CfxManager.GetLTRBAttributeValue(ltrb, out int l, out int t, out int r, out int b);
                b = (b - t) * 4 + t;
                ltrb = CfxManager.CreateLTRBAttributeValue(l, t, r, b);
                pi_box.SetAttributeValue(CfxManager.AttributeName_rect, ltrb);
            }
            man.AddField(CmpdDbManager.AbstractBitmap_FieldName, CfxManager.ControlType_Picture);

            var subform = man.AddSubform(CmpdDbManager.AccessionNumber_FieldName, man.GetRootForm(), CmpdDbManager.AccessionNumber_FieldName);
            man.BindDatabase(mdbName, CmpdDbManager.MolTableWithInfo_ViewName, subform);
            man.AddField(CmpdDbManager.Structure_FieldName, CfxManager.ControlType_Structure, subform);
            man.AddField(CmpdDbManager.CASRN_FieldName, CfxManager.ControlType_DataBox, subform);
            man.AddField(CmpdDbManager.NameInDoc_FieldName, CfxManager.ControlType_DataBox, subform);
            man.AddField(CmpdDbManager.Keywords_FieldName, CfxManager.ControlType_DataBox, subform);
            {
                var tab = new XElement(CfxManager.ElementName_tab);
                tab.SetAttributeValue(CfxManager.AttributeName_viewtype, "1");
                var tabs = new XElement(CfxManager.ElementName_tabs, tab);
                subform.Add(tabs);
            }

            man.Save(cfxFullPath);
        }

        private static string GenerateUniqString(int n)
        {
            return Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Substring(0, n);
        }

        public static string GeneratePseudoCASRNBase()
        {
            return "^" + GenerateUniqString(10);
        }

        public static string GeneratePseudoCASRN()
        {
            return "~" + GenerateUniqString(15);
        }

        public static string GeneratePseudoCASRN(string baseid, int n)
        {
            return baseid + "-" + n.ToString("D5");
        }

        private static string[][] MolTableFieldsInfo = {
            new[] { CmpdDbManager.StructureSource_FieldName,  "TEXT(" + CmpdDbManager.SizeOfStructureSource.ToString() + ")" },
            new[] { CmpdDbManager.CASRN_FieldName, "TEXT(" + CmpdDbManager.SizeOfCASRNField.ToString() + ")" + "UNIQUE" },
            new[] { CmpdDbManager.CasIndexName_FieldName, "TEXT" },
            new[] { CmpdDbManager.MolecularFormula_FieldName, "TEXT" },
            new[] { CmpdDbManager.LocalID_FieldName, "TEXT" }, 
            new[] { CmpdDbManager.StructureBmp_FieldName, "IMAGE" }, 
            new[] { CmpdDbManager.Smiles_FieldName, "TEXT" }, 
            new[] { CmpdDbManager.InChi_FieldName, "TEXT" }, 
            new[] { CmpdDbManager.OtherNames_FieldName, "TEXT", },
            new[] { CmpdDbManager.ClassIdentifier_FieldName, "TEXT", },
            new[] { CmpdDbManager.Copyright_FieldName, "TEXT", }, 
         };

        private static string[][] SubstancesFieldsInfo = {
            new[] { CmpdDbManager.AccessionNumber_FieldName, "TEXT(" + CmpdDbManager.SizeOfAccessionNumberField + ")" },
            new[] { CmpdDbManager.OrderInDoc_FieldName, "INTEGER" },
            new[] { CmpdDbManager.CASRN_FieldName, "TEXT(" + CmpdDbManager.SizeOfCASRNField + ")" },
            new[] { CmpdDbManager.NameInDoc_FieldName, "TEXT" },
            new[] { CmpdDbManager.Keywords_FieldName, "TEXT" }, 
            new[] { CmpdDbManager.Tag_FieldName, "TEXT" }, 
         };

        private static string[][] DocumentsFieldsInfo = {
            new[] { CmpdDbManager.AccessionNumber_FieldName, "TEXT(" + CmpdDbManager.SizeOfAccessionNumberField + ")" },
            new[] { CmpdDbManager.Title_FieldName, "TEXT" },
            new[] { CmpdDbManager.Author_FieldName, "TEXT" },
            new[] { CmpdDbManager.CorporateSource_FieldName, "TEXT" },
            new[] { CmpdDbManager.Source_FieldName, "TEXT" },
            new[] { CmpdDbManager.Publisher_FieldName, "TEXT" },
            new[] { CmpdDbManager.DocumentType_FieldName, "TEXT" },
            new[] { CmpdDbManager.Language_FieldName, "TEXT" },
            new[] { CmpdDbManager.PatentInfomation_FieldName, "TEXT" },
            new[] { CmpdDbManager.Abstract_FieldName, "TEXT" },
            new[] { CmpdDbManager.AbstractBitmap_FieldName, "IMAGE" },
            new[] { CmpdDbManager.Copyright_FieldName, "TEXT", }, 
         };

        private static string[][] OtherNamesFieldsInfo = 
        {
            new[] { CmpdDbManager.CASRN_FieldName, "TEXT(" + CmpdDbManager.SizeOfCASRNField.ToString() + ")" },
            new[] { CmpdDbManager.Name_FieldName, "TEXT", }, 
        };

        private static void AddFields(ChemFinderStructureDb db, string tableName, string[][] info)
        {
            foreach (var fieldInfo in info)
            {
                db.CreateField(tableName, fieldInfo[0], fieldInfo[1]);
            }
        }

        private static void ConstructTemplateMdb(string cfxPath)
        {
            using (var db = new ChemFinderStructureDb(cfxPath, FileAccess.ReadWrite))
            {
                AddFields(db, db.TableName, MolTableFieldsInfo);
                db.CreateTable(CmpdDbManager.Documents_TableName);
                AddFields(db, CmpdDbManager.Documents_TableName, DocumentsFieldsInfo);
                db.CreateTable(CmpdDbManager.Substances_TableName);
                AddFields(db, CmpdDbManager.Substances_TableName, SubstancesFieldsInfo);
                db.CreateTable(CmpdDbManager.OtherNames_TableName);
                AddFields(db, CmpdDbManager.OtherNames_TableName, OtherNamesFieldsInfo);

                {
                    // TODO: Fixing to secure 
                    var sql = "CREATE VIEW "
                        + CmpdDbManager.MolTableWithInfo_ViewName + " AS "
                        + "SELECT " + MolTable_TableName + ".*, "
                        + Substances_TableName + "." + AccessionNumber_FieldName + ", "
                        + Substances_TableName + "." + OrderInDoc_FieldName + ", "
                        + Substances_TableName + "." + NameInDoc_FieldName + ", "
                        + Substances_TableName + "." + Keywords_FieldName + " "
                        + "FROM " + MolTable_TableName + " "
                        + "LEFT JOIN " + Substances_TableName + " "
                        + "ON " + MolTable_TableName + "." + CASRN_FieldName + " = "
                                + Substances_TableName + "." + CASRN_FieldName;
                    db.ExecuteNonQuery(sql);
                }

                {
                    // TODO: Fixing to secure 
                    var sql = "CREATE VIEW "
                        + SubstancesWithInfo_ViewName + " AS "
                        + "SELECT " + Substances_TableName + ".*, "
                        + Documents_TableName + "." + Title_FieldName + ", "
                        + Documents_TableName + "." + Author_FieldName + ", "
                        + Documents_TableName + "." + CorporateSource_FieldName + ", "
                        + Documents_TableName + "." + Source_FieldName + " "
                        + "FROM " + Substances_TableName + " "
                        + "LEFT JOIN " + Documents_TableName + " " 
                        + "ON " + Substances_TableName + "." + AccessionNumber_FieldName + " = "
                                + Documents_TableName + "." + AccessionNumber_FieldName;
                    db.ExecuteNonQuery(sql);
                }
            }
        }
    }
}

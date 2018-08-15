using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CambridgeSoft.ChemScript16;
using MolServer = MolServer16;
using Ujihara.Chemistry.IO;

namespace Ujihara.Chemistry.CfxUtility
{
    public class Program
        : IDisposable
    {
        const string Mol_ID_FieldName = "Mol_ID";
        const string DefaultChemNameFieldName = "cas_index_name";
        public const string DefaultLocalIDFieldName = "Local_ID";
        public const string DefaultSmiles_FieldName = "SMILES";
        public const string DefaultInChi_FieldName = "InChi";

        static string[] SupportedScaffordExtensions = new[] { ".cdx" };
        static string[] SupportedInputFileExtensions = new[] { ".cfx" };
        static string[] SupportedLocalDatabaseExtensions = new[] { ".cfx" };

        public bool GenerateStructureFlag { get; set; } // = false;
        public bool CleanupStructureFlag { get; set; } // = false;
        public bool GenerateSmilesFlag { get; set; } // = false;
        public bool GenerateStructureFromInChi { get; set; }
        public bool GenerateStructureFromSmiles { get; set; }
        public bool FillNoStructure { get; set; }
        
        public string ScaffordCdx { get; set; } // = null;

        public string CfxFilePath { get; set; } // the file to manipulate

        private IList<string> _ChemNameFieldNames = new List<string>(new string[] { DefaultChemNameFieldName });
        public IList<string> ChemNameFieldNames
        {
            get { return _ChemNameFieldNames; }
            set { _ChemNameFieldNames = value; }
        }
        public void SetChemNameFieldNames(string semi_colon_sep_names)
        {
            this.ChemNameFieldNames = Utility.SemiColonSeparatedStringToList(semi_colon_sep_names);
        }

        private StructureData _ScaffordCSMol;
        private bool scafGenerated = false;
        private StructureData ScaffordCSMol
        {
            get 
            {
                if (scafGenerated)
                    return _ScaffordCSMol;
                if (ScaffordCdx != null)
                {
                    _ScaffordCSMol = StructureData.LoadFile(ScaffordCdx);
                }
                scafGenerated = true;
                return _ScaffordCSMol; 
            }
        }

        public string LocalDatabasePath { get; set; }
        public string FieldName_LocalID_Input { get; set; }
        private string _FieldName_LocalID_Output = DefaultLocalIDFieldName;
        public string FieldName_LocalID_Output
        {
            get { return _FieldName_LocalID_Output; }
            set { _FieldName_LocalID_Output = value; }
        }

        private string _FieldName_Smiles = DefaultSmiles_FieldName;
        public string FieldName_Smiles
        {
            get { return _FieldName_Smiles; }
            set { _FieldName_Smiles = value; }
        }

        private string _FieldName_InChi = DefaultInChi_FieldName;
        public string FieldName_InChi
        {
            get { return _FieldName_InChi; }
            set { _FieldName_InChi = value; }
        }

        private ChemFinderStructureDb _DbToManipurate = null;
        private ChemFinderStructureDb DbToManipurate
        {
            get
            {
                if (_DbToManipurate == null)
                    _DbToManipurate = new ChemFinderStructureDb(CfxFilePath, FileAccess.ReadWrite);
                return _DbToManipurate;
            }
        }

        private ChemFinderStructureDb _LocalDb;
        private ChemFinderStructureDb LocalDb
        {
            get
            {
                if (_LocalDb == null)
                {
                    if (LocalDatabasePath == null || LocalDatabasePath == "")
                        _LocalDb = null;
                    else
                        _LocalDb = new ChemFinderStructureDb(LocalDatabasePath, FileAccess.Read);
                }
                return _LocalDb;
            }
        }

        private ChemDraw.Application _ChemDrawApp = null;
        /// <summary>
        /// Utility 
        /// </summary>
        private ChemDraw.Application ChemDrawApp
        {
            get
            {
                if (_ChemDrawApp == null)
                    _ChemDrawApp = new ChemDraw.Application();
                return _ChemDrawApp;
            }
        }

        private ChemDraw.Documents _ChemDrawDocs = null;
        /// <summary>
        /// Utility
        /// </summary>
        private ChemDraw.Documents ChemDrawDocs
        {
            get
            {
                if (_ChemDrawDocs == null)
                {
                    _ChemDrawDocs = ChemDrawApp.Documents;
                }
                return _ChemDrawDocs;
            }
        }

        public void Dispose()
        {
            Utility.ReleaseComObject(_ChemDrawDocs);
            if (_ChemDrawApp != null)
                _ChemDrawApp.Quit();
            Utility.ReleaseComObject(_ChemDrawApp);

            if (_LocalDb != null)
            {
                _LocalDb.Dispose();
                _LocalDb = null;
            }
            if (_DbToManipurate != null)
            {
                _DbToManipurate.Dispose();
                _DbToManipurate = null;
            }
        }

        private static void CheckSupportedExtension(IEnumerable<string> extensions, string path)
        {
            if (!extensions.Contains(Path.GetExtension(path)))
                throw new UnsupportedExtensionException(path);
        }

        private void GenerateSmiles(StructureDb.Recordset cursor)
        {
            object missing = Type.Missing;

            var mol = (MolServer.Molecule)(cursor.GetValue(CfxManager.FieldName_MolID));
            if (mol != null)
            {
                try
                {
                    using (var tempCdx = new TempFile(".cdx"))
                    {
                        mol.Write(tempCdx.Path, null, null);
                        var doc = ChemDrawDocs.Open(tempCdx.Path, ref missing, ref missing);
                        try
                        {
                            using (var tempTxt = new TempFile(".txt"))
                            {
                                object filename = tempTxt.Path;
                                object format = ChemDraw.CDFormat.kCDFormatSMILES;
                                doc.SaveAs(ref filename, ref format, ref missing, ref missing, ref missing);

                                using (var f = new FileStream(tempTxt.Path, FileMode.Open, FileAccess.Read))
                                {
                                    using (var t = new StreamReader(f, Encoding.ASCII))
                                    {
                                        var smiles = t.ReadToEnd().Trim();
                                        cursor.SetValue(FieldName_Smiles, smiles);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            doc.Close(ref missing, ref missing);
                            Utility.ReleaseComObject(doc);
                        }
                    }
                }
                finally
                {
                    Utility.ReleaseComObject(mol);
                }
            }
        }

        private void NameLocalID(StructureDb.Recordset cursor)
        {
            if (LocalDb == null)
                return;

            var mol = (MolServer.Molecule)(cursor.GetValue(CfxManager.FieldName_MolID));
            try
            {
                using (var sr = LocalDb.Search(Mol_ID_FieldName, mol))
                {
                    var localIDAsString = ToString(sr);
                    if (localIDAsString != null)
                    {
                        cursor.SetValue(FieldName_LocalID_Output, localIDAsString);
                        //((Ujihara.Chemistry.ChemFinderStructureDb.ChemFinderDbRecordset)cursor).Update();
                    }
                }
            }
            finally
            {
                Utility.ReleaseComObject(mol);
            }
        }

        private string ToString(StructureDb.Recordset hitList)
        {
            var localIDsAsSB = new StringBuilder();
            while (!hitList.EOF)
            {
                var localCode = hitList.GetValue(FieldName_LocalID_Input);
                if (localCode == null)
                {
                    // localCode can be null when the structure is existing in mst but not in mdb.
                }
                else
                {
                    if (localIDsAsSB.Length != 0)
                        localIDsAsSB.Append(";");
                    localIDsAsSB.Append(localCode);
                }
                hitList.MoveNext();
            }
            string comment;
            if (hitList.Properties.TryGetValue("Comment", out comment))
            {
                if (localIDsAsSB.Length != 0)
                    localIDsAsSB.Append(" ");
                localIDsAsSB.Append("(").Append(comment).Append(")");
            }
            return localIDsAsSB.ToString();
        }

        private void ManipulateStructure(StructureDb.Recordset cursor, List<string> invalidFields)
        {
            {
                var molid = cursor.GetRawValue(CfxManager.FieldName_MolID) as int?;

                StructureData csmol = this.GetCSMol(DbToManipurate, molid,
                    () =>
                    {
                        StructureData csmolFromName = null;
                        if (GenerateStructureFlag)
                        {   
                            foreach (var chemNameFieldName in ChemNameFieldNames)
                            {
                                if (!invalidFields.Contains(chemNameFieldName))
                                {
                                    string chemname = null;
                                    try
                                    {
                                        chemname = cursor.GetValue(chemNameFieldName) as string;
                                    }
                                    catch (Exception)
                                    {
                                        invalidFields.Add(chemNameFieldName);
                                    }

                                    if (chemname != null)
                                    {
                                        csmolFromName = ChemScriptUtility.StructureDataFromName(chemname);
                                        if (csmolFromName != null)
                                            break;
                                    }
                                }
                            }
                        }
                        return csmolFromName;
                    });
                if (csmol != null)
                {
                    using (var cdx = new TempFile(".cdx"))
                    {
                        csmol.WriteFile(cdx.Path, "chemical/x-cdx");
                        cursor.SetValue(CfxManager.FieldName_MolID, cdx.Path);
                    }
                }
            }

            if (GenerateStructureFromInChi)
            {
                GenerateStructureFromField(cursor, FieldName_InChi, ChemDraw.CDFormat.kCDFormatInChI);
            }

            if (GenerateStructureFromSmiles)
            {
                GenerateStructureFromField(cursor, FieldName_Smiles, ChemDraw.CDFormat.kCDFormatSMILES);
            }

            if (FillNoStructure)
            {
                var molid = cursor.GetRawValue(CfxManager.FieldName_MolID) as int?;
                if (molid == null)
                {
                    object missing = Type.Missing;
                    var doc = ChemDrawDocs.Add();
                    try
                    {
                        using (var tempCdx = new TempFile(".cdx"))
                        {
                            object filename = tempCdx.Path;
                            object formatCdx = ChemDraw.CDFormat.kCDFormatCDX;
                            doc.SaveAs(ref filename, ref formatCdx, ref missing, ref missing, ref missing);

                            cursor.SetValue(CfxManager.FieldName_MolID, tempCdx.Path);
                        }
                    }
                    finally
                    {
                        doc.Close(ref missing, ref missing);
                        Utility.ReleaseComObject(doc);
                    }
                }
            }
        }

        private void GenerateStructureFromField(StructureDb.Recordset cursor, string fieldName, ChemDraw.CDFormat cdFormat)
        {
            var molid = cursor.GetRawValue(CfxManager.FieldName_MolID) as int?;
            if (molid != null)
                return;

            var fieldValue = cursor.GetValue(fieldName) as string;
            if (fieldValue != null && fieldValue.Length > 0)
            {
                using (var tempTxt = new TempFile(".txt"))
                using (var tempCdx = new TempFile(".cdx"))
                {
                    using (var ts = new StreamWriter(tempTxt.Path))
                    {
                        ts.Write(fieldValue);
                    }
                    object missing = Type.Missing;
                    object format = cdFormat;
                    var doc = ChemDrawDocs.Open(tempTxt.Path, ref missing, ref format);
                    try
                    {
                        if (doc != null)
                        {
                            object filename = tempCdx.Path;
                            object formatCdx = ChemDraw.CDFormat.kCDFormatCDX;
                            doc.SaveAs(ref filename, ref formatCdx, ref missing, ref missing, ref missing);

                            cursor.SetValue(CfxManager.FieldName_MolID, tempCdx.Path);
                        }
                    }
                    finally
                    {
                        doc.Close(ref missing, ref missing);
                        Utility.ReleaseComObject(doc);
                    }
                }
            }
        }

        private StructureData GetCSMol(ChemFinderStructureDb db, int? molid, Func<StructureData> csmolGenerator)
        {
            StructureData csmol = null;

            if (GenerateStructureFlag)
            {
                if (molid == null)
                {
                    csmol = csmolGenerator();
                }
            }

            if (ScaffordCSMol != null)
            {
                if (csmol == null && molid != null)
                {
                    var mol = db.MstDocument.GetMol((int)molid);
                    if (mol == null)
                        return null;
                    try
                    {
                        csmol = ChemScriptUtility.ToStructureData(mol);
                    }
                    finally
                    {
                        Utility.ReleaseComObject(mol);
                    }
                }
                if (csmol != null)
                {
                    csmol.ScaffoldCleanup(ScaffordCSMol);
                }
            }

            return csmol;
        }

        public void ManipurateAll()
        {
            using (var db = new ChemFinderStructureDb(this.CfxFilePath, FileAccess.ReadWrite))
            {
                if (LocalDb != null)
                    db.CreateField(FieldName_LocalID_Output, "TEXT");
                var cursor = db.ObtainRecordset();
                var invalidFields = new List<string>();
                while (!cursor.EOF)
                {
                    ManipulateStructure(cursor, invalidFields);

                    if (CleanupStructureFlag)
                        CleanStructure(cursor);
                    
                    if (GenerateSmilesFlag)
                        GenerateSmiles(cursor);
                    
                    NameLocalID(cursor);

                    cursor.MoveNext();
                }
            }
        }

        private void CleanStructure(StructureDb.Recordset cursor)
        {
            var mol = cursor.GetValue(CfxManager.FieldName_MolID) as MolServer.Molecule;
            if (mol == null)
                return;

            try
            {
                using (var dir = new TempDirectory())
                {
                    object missing = Type.Missing;

                    var original = Path.Combine(dir.Directory.FullName, "original.cdx");
                    var cleanedup = Path.Combine(dir.Directory.FullName, "cleanedup.cdx");

                    ChemDraw.Document doc = null;
                    ChemDraw.Objects objs = null;
                    try
                    {
                        mol.Write(original, Type.Missing, Type.Missing);
                        doc = ChemDrawDocs.Open(original, ref missing, ref missing);
                        objs = doc.Objects;
                        objs.Clean(true);

                        object filename = cleanedup;
                        doc.SaveAs(ref filename, ref missing, ref missing, ref missing, ref missing);
                        mol.Read(cleanedup);
                    }
                    finally
                    {
                        Utility.ReleaseComObject(objs);
                        if (doc != null)
                            doc.Close(ref missing, ref missing);
                        Utility.ReleaseComObject(doc);
                    }
                }

                cursor.SetValue(CfxManager.FieldName_MolID, mol);
            }
            finally
            {
                Utility.ReleaseComObject(mol);
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                using (var program = new Program())
                {
                    ParseCommandLine(args, program);
                    program.ManipurateAll();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
            finally
            {
                System.GC.Collect();                  
                System.GC.WaitForPendingFinalizers(); 
                System.GC.Collect(); 
            }
        }

        private static void ParseCommandLine(string[] args, Program program)
        {
            var indexArgs = 0;
            for (; indexArgs < args.Length; indexArgs++)
            {
                var arg = args[indexArgs];
                if (arg[0] == '-')
                {
                    switch (arg)
                    {
                        case "-c":
                            program.CleanupStructureFlag = true;
                            break;
                        case "-g":
                            program.GenerateStructureFlag = true;
                            break;
                        case "-n":
                            indexArgs++;
                            program.ChemNameFieldNames = Utility.SemiColonSeparatedStringToList(args[indexArgs]);
                            break;
                        case "-s":
                            indexArgs++;
                            program.ScaffordCdx = args[indexArgs];
                            CheckSupportedExtension(SupportedScaffordExtensions, program.ScaffordCdx);
                            break;
                        case "-d":
                            indexArgs++;
                            program.LocalDatabasePath = args[indexArgs];
                            CheckSupportedExtension(SupportedLocalDatabaseExtensions, program.LocalDatabasePath);
                            break;
                        case "-f":
                            indexArgs++;
                            program.FieldName_LocalID_Input = args[indexArgs];
                            break;
                        case "-help":
                            throw new ApplicationException(
                                "usege: CfxUtility [options] path\n"
                                + "-c\tCleanup structure.\n"
                                + "-g\tGenerate structure from chem-name field.\n"
                                + "-n chem-name\tSpecify {chem-name field}. Default is " + DefaultChemNameFieldName + ".\n"
                                + "-s cdx-path\tSpecify scafford template structure file name.\n"
                                + "-d cfx-path\n"
                                + "-f field-name\n"
                            );
                        default:
                            throw new ApplicationException("Unknown option '" + arg + "'.");
                    }
                }
                else
                    break;
            }

            program.CfxFilePath = args[indexArgs++];
            CheckSupportedExtension(SupportedInputFileExtensions, program.CfxFilePath);

            if (indexArgs < args.Length)
                throw new ApplicationException("Only one input file is supported");
        }
    }

    class UnsupportedExtensionException
        : ApplicationException
    {
        public UnsupportedExtensionException(string path)
            : base("Extension of '" + path + "' is not supported.")
        {
            this.Path = path;
        }

        public string Path { get; private set; }
    }
}
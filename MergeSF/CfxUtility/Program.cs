using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CambridgeSoft.ChemScript16;
using MolServer = MolServer16;
using Ujihara.Chemistry.IO;
using Ujihara.Chemistry.MergeSF;

namespace Ujihara.Chemistry.CfxUtility
{
    public class Program
        : IDisposable, IRunnable
    {
        static string[] SupportedScaffordExtensions = new[] { ".cdx" };
        static string[] SupportedInputFileExtensions = new[] { ".cfx" };
        static string[] SupportedLocalDatabaseExtensions = new[] { ".cfx" };

        public bool GenerateStructureFlag { get; set; } // = false;

        public bool CleanupStructureFlag
        {
            get
            {
                return CleanupStructureShrethold != 0;
            }
            set
            {
                if (CleanupStructureShrethold == 0)
                {
                    CleanupStructureShrethold = 600; // This value is daterminded not to clean C60 up 
                }
            }
        } 
        public double CleanupStructureShrethold { get; set; } 

        public bool GenerateSmilesFlag { get; set; } // = false;
        public bool GenerateStructureFromInChi { get; set; }
        public bool GenerateStructureFromSmiles { get; set; }
        public bool GenerateStructureFromImage { get; set; }
        public bool FillNoStructure { get; set; }
        
        public string ScaffordCdx { get; set; } // = null;

        public string CfxFilePath { get; set; } // the file to manipulate

        private IList<string> _ChemNameFieldNames = new List<string>(new string[] { CmpdDbManager.CasIndexName_FieldName });
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

        public string GetProgress()
        {
            return "";
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
                                        cursor.SetValue(CmpdDbManager.Smiles_FieldName, smiles);
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
                using (var sr = LocalDb.Search(CmpdDbManager.MolID_FieldName, mol))
                {
                    var localIDAsString = ToString(sr);
                    if (localIDAsString != null)
                    {
                        cursor.SetValue(CmpdDbManager.LocalID_FieldName, localIDAsString);
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

                var csmols = this.GetCSMol(DbToManipurate, molid, cursor, invalidFields);

                if (csmols != null)
                {
                    using (var cdx = new TempFile(".cdx"))
                    {
                        csmols.Structure.WriteFile(cdx.Path, "chemical/x-cdx");
                        cursor.SetValue(CfxManager.FieldName_MolID, cdx.Path);
                        if (csmols.Source != null)
                            cursor.SetValue(CmpdDbManager.StructureSource_FieldName, csmols.Source);
                    }
                }
                else if (molid == null)
                {
#if SUPPORT_OSRA
                    if (GenerateStructureFromImage)
                    {
                        var image = (byte[])cursor.GetValue(CmpdDbManager.StructureBmp_FieldName);
                        if (image != null)
                        {
                            using (var ms = new MemoryStream(image))
                            using (var im = System.Drawing.Image.FromStream(ms))
                            {
                                string ext = "";
                                if (im.RawFormat == System.Drawing.Imaging.ImageFormat.Bmp)
                                    ext = ".bmp";
                                else if (im.RawFormat == System.Drawing.Imaging.ImageFormat.Gif)
                                    ext = ".gif";
                                else if (im.RawFormat == System.Drawing.Imaging.ImageFormat.Jpeg)
                                    ext = ".jpg";
                                else if (im.RawFormat == System.Drawing.Imaging.ImageFormat.Png)
                                    ext = ".png";
                                using (var tf = new TempFile(ext))
                                using (var mf = new TempFile(ext))
                                {
                                    im.Save(tf.Path);
                                    if (OsraAPI.ImageToMol(tf.Path, mf.Path))
                                    {
                                        cursor.SetValue(CfxManager.FieldName_MolID, mf.Path);
                                        cursor.SetValue(CmpdDbManager.StructureSource_FieldName, CmpdDbManager.StructureSource_OSRA);
                                    }
                                }
                            }
                        }
                    }
#endif
                }
            }

            if (GenerateStructureFromInChi)
            {
                GenerateStructureFromField(cursor, CmpdDbManager.LocalID_FieldName, ChemDraw.CDFormat.kCDFormatInChI);
                cursor.SetValue(CmpdDbManager.StructureSource_FieldName, CmpdDbManager.StructureSource_InChi);
            }

            if (GenerateStructureFromSmiles)
            {
                GenerateStructureFromField(cursor, CmpdDbManager.Smiles_FieldName, ChemDraw.CDFormat.kCDFormatSMILES);
                cursor.SetValue(CmpdDbManager.StructureSource_FieldName, CmpdDbManager.StructureSource_SMILES);
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

        class StructureDataAndSource
        {
            public StructureDataAndSource(StructureData structure, string source)
            {
                this.Structure = structure;
                this.Source = source;
            }

            public StructureData Structure;
            public string Source;
        }

        private StructureDataAndSource GetCSMol(ChemFinderStructureDb db, int? molid, StructureDb.Recordset cursor, List<string> invalidFields)
        {
            StructureDataAndSource csmolands = null;

            if (molid == null)
            {
                if (GenerateStructureFlag)
                {
                    string source = null;
                    StructureData csmolFromName = null;
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
                                {
                                    source = chemNameFieldName;
                                    break;
                                }
                            }
                        }
                    }
                    csmolands = csmolFromName == null ? null : new StructureDataAndSource(csmolFromName, source);
                }
            }

            if (ScaffordCSMol != null)
            {
                if (csmolands == null && molid != null)
                {
                    var mol = db.MstDocument.GetMol((int)molid);
                    if (mol == null)
                        return null;
                    StructureData csmol = null;
                    try
                    {
                        csmol = MolServerUtility.ToStructureData(mol);
                    }
                    finally
                    {
                        Utility.ReleaseComObject(mol);
                    }
                    if (csmol != null)
                    {
                        csmolands = new StructureDataAndSource(csmol, null);
                    }
                }
                if (csmolands != null)
                {
                    csmolands.Structure.ScaffoldCleanup(ScaffordCSMol);
                }
            }

            return csmolands;
        }

        public void Run()
        {
            using (var db = new ChemFinderStructureDb(this.CfxFilePath, FileAccess.ReadWrite))
            {
                if (LocalDb != null)
                    db.CreateField(CmpdDbManager.LocalID_FieldName, "TEXT");
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
                        if (objs.MolecularWeight < CleanupStructureShrethold)
                        {
                            objs.Clean(true);

                            object filename = cleanedup;
                            doc.SaveAs(ref filename, ref missing, ref missing, ref missing, ref missing);
                            cursor.SetValue(CfxManager.FieldName_MolID, cleanedup);
                        }
                    }
                    finally
                    {
                        Utility.ReleaseComObject(objs);
                        if (doc != null)
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

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                using (var program = new Program())
                {
                    ParseCommandLine(args, program);
                    program.Run();
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
                                + "-n chem-name\tSpecify {chem-name field}. Default is " + CmpdDbManager.CasIndexName_FieldName + ".\n"
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
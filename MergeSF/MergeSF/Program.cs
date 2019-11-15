using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Ujihara.Chemistry.IO;
using MolServer = MolServer19;

namespace Ujihara.Chemistry.MergeSF
{
    public interface IRunnable
    {
        void Run();
    }

    public interface IProgressReporter
    {
        string GetReport();
    }

    public class Program
        : IDisposable, IRunnable, IProgressReporter
    {
        private string outputPath = "a.cfx";
        public string OutputPath
        {
            get { return outputPath; }
            set { outputPath = value; }
        }

        public bool AppendFlag { get; set; }
        public bool OverWriteFlag { get; set; }
        public bool LoadSubstanceImageFlag { get; set; }

        private IList<string> _ReferenceFiles = new List<string>();
        /// <summary>
        /// Reference file exported in SciFinder, <value>null</value> if reference file is not available. 
        /// </summary>
        public IList<string> ReferenceFiles
        {
            get { return _ReferenceFiles; }
            set { _ReferenceFiles = value; }
        }

        IList<string> _SDFiles = new List<string>();
        /// <summary>
        /// SD files exported in SciFinder.
        /// </summary>
        public IList<string> SDFiles
        {
            get { return _SDFiles; }
            set { _SDFiles = value; }
        }

        IList<string> _CSVFiles = new List<string>();
        public IList<string> CSVFiles
        {
            get { return _CSVFiles; }
            set { _CSVFiles = value; }
        }

        IList<string> _ListFiles = new List<string>();
        public IList<string> ListFiles
        {
            get { return _ListFiles; }
            set { _ListFiles = value; }
        }

        IList<string> _CasOnlineFiles = new List<string>();
        public IList<string> CasOnlineFiles
        {
            get { return _CasOnlineFiles; }
            set { _CasOnlineFiles = value; }
        }

        IList<string> _CfxFiles = new List<string>();
        public IList<string> CfxFiles
        {
            get { return _CfxFiles; }
            set { _CfxFiles = value; }
        }

        IList<string> _SmilesListFiles = new List<string>();
        public IList<string> SmilesListFiles
        {
            get { return _SmilesListFiles; }
            set { _SmilesListFiles = value; }
        }

        IList<string> _ImageFiles = new List<string>();
        public IList<string> ImageFiles
        {
            get { return _ImageFiles; }
            set { _ImageFiles = value; }
        }

        private string _CfxOutput = null;
        private string CfxOutput
        {
            get
            {
                if (_CfxOutput == null)
                {
                    _CfxOutput = CfxOutputWithoutExtension + ".cfx";
                }
                return _CfxOutput;
            }
        }

        private string _CfxOutputWithoutExtension = null;
        private string CfxOutputWithoutExtension
        {
            get
            {
                if (_CfxOutputWithoutExtension == null)
                {
                    _CfxOutputWithoutExtension = Path.GetFileNameWithoutExtension(Path.GetFileName(OutputPath));
                }
                return _CfxOutputWithoutExtension;
            }
        }

        public void Run()
        {
            if (AppendFlag)
            {
                BuildInnerCfx(OutputPath);
            }
            else
            {
                using (var tempFilesCreator = new TempDirectory())
                {
                    var cfxOutputFullPath = Path.Combine(tempFilesCreator.Directory.FullName, CfxOutput);
                    CmpdDbManager.MakeNew(cfxOutputFullPath);

                    BuildInnerCfx(cfxOutputFullPath);

                    var man = new CfxManager();
                    man.Load(cfxOutputFullPath);
                    var cfxDocBase = Path.Combine(tempFilesCreator.Directory.FullName, CfxOutputWithoutExtension + "_doc.cfx");
                    CmpdDbManager.MakeDocumentBaseForm(cfxDocBase, man.DatabaseFileName, man.MstName);

                    var body = Path.GetFileNameWithoutExtension(Path.GetFileName(OutputPath));
                    var directoryName = Path.GetDirectoryName(OutputPath);
                    foreach (var f in tempFilesCreator.Directory.GetFiles())
                    {
                        var dest = Path.Combine(directoryName, Path.GetFileName(f.Name));
                        if (Path.GetExtension(f.Name) != ".ldb")
                        {
                            f.CopyTo(dest, true);
                        }
                    }
                }
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                var program = new Program();
                ParseCommandLine(args, program);

                program.Run();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
            }
            finally
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                System.GC.Collect();
            }
        }

        private static string CreateCommaDelimitted(IEnumerable<string> strings)
        {
            var sb = new StringBuilder();
            bool isNotFirst = false;
            foreach (var s in strings)
            {
                if (isNotFirst)
                    sb.Append(", ");
                sb.Append(s);
                isNotFirst = true;
            }
            return sb.ToString();
        }

        private ChemFinderStructureDb db;
        private volatile string curr_file_name = null;
        private volatile int curr_number = 0;
        private volatile int total_number = 0;

        public string GetReport()
        {
            var sb = new StringBuilder();
            if (curr_file_name != null)
            {
                sb.Append(curr_file_name).Append(" ");
            }
            if (curr_number != 0)
            {
                sb.Append(curr_number);
                if (total_number != 0)
                {
                    sb.Append("/").Append(total_number);
                }
            }
            return sb.ToString();
        }

        private void SetFileName(string value)
        {
            curr_file_name = value;
        }

        private void SetTotalNumber(int n)
        {
            total_number = n;
        }

        private void IncProgressCount()
        {
            curr_number++;
        }

        private void ResetProgressCount()
        {
            curr_number = 0;
            total_number = 0;
        }

        private void ResetProgress()
        {
            curr_file_name = null;
            ResetProgressCount();
        }

        private void BuildInnerCfx(string cfxPath)
        {
            ResetProgress();

            using (var cfx = new ChemFinderStructureDb(cfxPath, FileAccess.ReadWrite))
            {
                db = cfx;
                LoadSDFs(this.SDFiles);
                LoadSciFinderReferencesFiles(this.ReferenceFiles);
                LoadSmilesFromList(this.SmilesListFiles);
                LoadMoleculesFromCompoundNameList(this.ListFiles);
                LoadSubstancesFromCSV(this.CSVFiles);
                LoadCAOnlineFiles(this.CasOnlineFiles);
                LoadCfxFiles(this.CfxFiles);
#if SUPPORT_OSRA
                LoadImageFiles(this.ImageFiles);
#endif
                //AddMolsNotInMolTable();
            }
            db = null;
        }

        private void AddMolsNotInMolTable()
        {
            var sql = "INSERT INTO "
                + CmpdDbManager.MolTable_TableName + " "
                + "( "
                + CmpdDbManager.CASRN_FieldName + " ) "
                + "SELECT DISTINCT "
                + CmpdDbManager.CASRN_FieldName + " "
                + "FROM "
                + CmpdDbManager.Substances_TableName + " "
                + "WHERE (((Substances.cas_rn) Not In (SELECT cas_rn FROM MolTable )))";
            db.ExecuteNonQuery(sql);
        }

#if SUPPORT_OSRA
        private void LoadImageFiles(IList<string> filenames)
        {
            using (var rstDocuments = db.ObtainRecordset(CmpdDbManager.Documents_TableName))
            using (var rstSubstances = db.ObtainRecordset(CmpdDbManager.Substances_TableName))
            using (var rstMolTable = db.ObtainRecordset(CmpdDbManager.MolTable_TableName))
            {
                foreach (var filename in filenames)
                {
                    LoadImageFile(rstDocuments, rstMolTable, rstSubstances, filename);
                }
                ResetProgress();
            }
        }

        private void LoadImageFile(
            ChemFinderStructureDb.ChemFinderDbRecordset rstDocuments,
            ChemFinderStructureDb.ChemFinderDbRecordset rstMolTable,
            ChemFinderStructureDb.ChemFinderDbRecordset rstSubstances,
            string imagefile)
        {
            string casrnbase = CmpdDbManager.GeneratePseudoCASRNBase();

            var theFileName = Path.GetFileName(imagefile);
            SetFileName(theFileName);
            var refInfo = new ReferenceInfo();
            var pseudoAN = "~" + theFileName.GetHashCode().ToString("X8");
            refInfo.AccessionNumber = pseudoAN;
            refInfo.DocumentType = CmpdDbManager.DocumentType_Unknown;
            refInfo.Source = theFileName;

            RegisterReferenceInfo(rstDocuments, refInfo);

            using (var sdf = new TempFile(".sdf"))
            using (var molExtractedDir = new TempDirectory())
            using (var imageExtractedDir = new TempDirectory())
            {
                const string prefix = "a";
                if (OsraAPI.ImageToSdf(imagefile, sdf.Path,
                    Path.Combine(imageExtractedDir.Directory.FullName, prefix)))
                {
                    using (var sr = new StreamReader(sdf.Path))
                    using (var sdfReader = new SDFReader(sr))
                    {
                        int idx = 0;
                        IDictionary<string, string> sdrecord;
                        while ((sdrecord = sdfReader.Read()) != null)
                        {
                            string extractedImageFileName = Path.Combine(imageExtractedDir.Directory.FullName, prefix + idx.ToString() + ".png");
                            string molInString;
                            // Empty means the MOL
                            if (sdrecord.TryGetValue("", out molInString))
                            {
                                var entry = new SubstanceInfo();
                                entry.CASRN = CmpdDbManager.GeneratePseudoCASRN(casrnbase, idx);
                                var tempMolFile = Path.Combine(molExtractedDir.Directory.FullName, prefix + idx.ToString() + ".mol");
                                Utility.GenerateFileFromString(molInString, tempMolFile);
                                if (File.Exists(extractedImageFileName))
                                {
                                    entry.Bitmap =  Utility.FileToArray(extractedImageFileName);
                                }
                                RegisterSubstanceAndCmpdFileToMolTable(rstMolTable, entry, tempMolFile, CmpdDbManager.StructureSource_OSRA);
                                RegisterSubstanceInfo(rstSubstances, refInfo, entry);
                            }
                            idx++;
                        }
                    }
                }
            }
        }

        private static void LoadImageFile(ChemFinderStructureDb.ChemFinderDbRecordset rstTo, string imagefile)
        {
            string casrnbase = CmpdDbManager.GeneratePseudoCASRNBase();

            using (var sdf = new TempFile(".sdf"))
            using (var molExtractedDir = new TempDirectory())
            using (var imageExtractedDir = new TempDirectory())
            {
                const string prefix = "a";
                if (OsraAPI.ImageToSdf(imagefile, sdf.Path,
                    Path.Combine(imageExtractedDir.Directory.FullName, prefix)))
                {
                    using (var sr = new StreamReader(sdf.Path))
                    using (var sdfReader = new SDFReader(sr))
                    {
                        int idx = 0;
                        IDictionary<string, string> sdrecord;
                        while ((sdrecord = sdfReader.Read()) != null)
                        {
                            var entry = new SubstanceInfo();
                            entry.CASRN = CmpdDbManager.GeneratePseudoCASRN(casrnbase, idx);
                            
                            string extractedImageFileName = Path.Combine(imageExtractedDir.Directory.FullName, prefix + idx.ToString() + ".png");
                            string molInString;
                            // Empty means the MOL
                            if (sdrecord.TryGetValue("", out molInString))
                            {
                                var tempMolFile = Path.Combine(molExtractedDir.Directory.FullName, prefix + idx.ToString() + ".mol");
                                Utility.GenerateFileFromString(molInString, tempMolFile);

                                var irec = new StructureDb.Record();
                                irec.SetValue(CmpdDbManager.MolID_FieldName, tempMolFile);
                                irec.SetValue(CmpdDbManager.StructureSource_FieldName, CmpdDbManager.StructureSource_OSRA);
                                irec.SetValue(CmpdDbManager.CASRN_FieldName, entry.CASRN);

                                if (File.Exists(extractedImageFileName))
                                {
                                    irec.SetValue(CmpdDbManager.StructureBmp_FieldName, Utility.FileToArray(extractedImageFileName)); 
                                }
                                rstTo.Add(irec);
                            }
                            idx++;
                        }
                    }
                }
            }
        }
#endif

        private void LoadCfxFiles(IList<string> cfxFiles)
        {
            foreach (var cfx in CfxFiles)
            {
                using (var indb = new ChemFinderStructureDb(cfx, FileAccess.Read))
                {
                    using (var rstTo = db.ObtainRecordset(CmpdDbManager.MolTable_TableName))
                    using (var rstFrom = indb.ObtainRecordset(CmpdDbManager.MolTable_TableName))
                    {
                        while (!rstFrom.EOF)
                        {
                            var si = new SubstanceInfo();
                            si.CASRN = (string)rstFrom.GetValue(CmpdDbManager.CASRN_FieldName);
                            si.CAIndexName = (string)rstFrom.EGetValue(CmpdDbManager.CasIndexName_FieldName);
                            si.Name = (string)rstFrom.EGetValue(CmpdDbManager.OtherNames_FieldName);
                            si.MolecularFormula = (string)rstFrom.EGetValue(CmpdDbManager.MolecularFormula_FieldName);
                            si.ClassIdentifier = (string)rstFrom.EGetValue(CmpdDbManager.ClassIdentifier_FieldName);
                            si.Bitmap = (byte[])rstFrom.EGetValue(CmpdDbManager.StructureBmp_FieldName);
                            si.Copyright = (string)rstFrom.EGetValue(CmpdDbManager.Copyright_FieldName);

                            RegisterSubstanceToMolTable(rstTo, si);

                            var cfmol = rstFrom.GetValue(CmpdDbManager.MolID_FieldName) as MolServer.Molecule;
                            if (cfmol != null)
                            {
                                try
                                {
                                    rstTo.Filter = CmpdDbManager.CASRN_FieldName + "='" + si.CASRN + "'";
                                    rstTo.SetValue(CmpdDbManager.MolID_FieldName, cfmol);
                                }
                                finally
                                {
                                    Utility.ReleaseComObject(cfmol);
                                }
                            }
                            
                            rstFrom.MoveNext();
                        }
                    }

                    using (var rstTo = db.ObtainRecordset(CmpdDbManager.Documents_TableName))
                    using (var rstFrom = indb.ObtainRecordset(CmpdDbManager.Documents_TableName))
                    {
                        while (!rstFrom.EOF)
                        {
                            var ri = new ReferenceInfo();
                            ri.AccessionNumber = (string)rstFrom.GetValue(CmpdDbManager.AccessionNumber_FieldName);
                            ri.Title = (string)rstFrom.EGetValue(CmpdDbManager.Title_FieldName);
                            ri.By = (string)rstFrom.EGetValue(CmpdDbManager.Author_FieldName);
                            ri.CorporateSource = (string)rstFrom.EGetValue(CmpdDbManager.CorporateSource_FieldName);
                            ri.PatentAssignee = (string)rstFrom.EGetValue(CmpdDbManager.CorporateSource_FieldName);
                            ri.Source = (string)rstFrom.EGetValue(CmpdDbManager.Source_FieldName);
                            ri.Publisher = (string)rstFrom.EGetValue(CmpdDbManager.Publisher_FieldName);
                            ri.DocumentType = (string)rstFrom.EGetValue(CmpdDbManager.DocumentType_FieldName);
                            ri.Language = (string)rstFrom.EGetValue(CmpdDbManager.Language_FieldName);
                            ri.PatentInfomation = (string)rstFrom.EGetValue(CmpdDbManager.PatentInfomation_FieldName);
                            ri.Abstract = (string)rstFrom.EGetValue(CmpdDbManager.Abstract_FieldName);
                            ri.AbstractImage = (byte[])rstFrom.EGetValue(CmpdDbManager.AbstractBitmap_FieldName);
                            ri.Copyright = (string)rstFrom.EGetValue(CmpdDbManager.Copyright_FieldName);

                            RegisterReferenceInfo(rstTo, ri);

                            rstFrom.MoveNext();
                        }
                    }

                    using (var rstTo = db.ObtainRecordset(CmpdDbManager.Substances_TableName))
                    using (var rstFrom = indb.ObtainRecordset(CmpdDbManager.Substances_TableName))
                    {
                        while (!rstFrom.EOF)
                        {
                            var ri = new ReferenceInfo();
                            var si = new SubstanceInfo();

                            ri.AccessionNumber = (string)rstFrom.GetValue(CmpdDbManager.AccessionNumber_FieldName);
                            si.Order = (int?)rstFrom.EGetValue(CmpdDbManager.OrderInDoc_FieldName);
                            si.CASRN = (string)rstFrom.GetValue(CmpdDbManager.CASRN_FieldName);
                            si.Name = (string)rstFrom.EGetValue(CmpdDbManager.NameInDoc_FieldName);
                            si.Keywords = (string)rstFrom.EGetValue(CmpdDbManager.Keywords_FieldName);

                            RegisterSubstanceInfo(rstTo, ri, si);

                            rstFrom.MoveNext();
                        }
                    }
                }
            }
        }

        private void LoadSDFs(IList<string> SDFiles)
        {
            var rst = db.ObtainRecordset(CmpdDbManager.MolTable_TableName);
            foreach (var sdf in SDFiles)
            {
                SetFileName(Path.GetFileName(sdf));

                using (var textReader = new StreamReader(sdf))
                using (var sdfReader = new SDFReader(textReader, true))
                {
                    IDictionary<string, string> sdrecord;
                    while ((sdrecord = sdfReader.Read()) != null)
                    {
                        string molInString, casrn, index_name, molecular_formula, copyright;
                        if (!sdrecord.TryGetValue("", out molInString))
                            molInString = null;
                        if (!sdrecord.TryGetValue(CmpdDbManager.CasIndexName_FieldName, out index_name))
                            index_name = null;
                        index_name = NormalizeCASIN(index_name);

                        if (!sdrecord.TryGetValue(CmpdDbManager.CASRN_FieldName, out casrn))
                            casrn = null;
                        if (casrn == null)
                            casrn = CmpdDbManager.GeneratePseudoCASRN();
                        else
                            casrn = NormalizeCASRN(casrn);

                        if (!sdrecord.TryGetValue(CmpdDbManager.MolecularFormula_FieldName, out molecular_formula))
                            molecular_formula = null;

                        if (!sdrecord.TryGetValue(CmpdDbManager.Copyright_FieldName, out copyright))
                            copyright = null;

                        rst.Filter = CmpdDbManager.CASRN_FieldName + "='" + casrn + "'";

                        bool willRegistStructrue = true;
                        IGetSetValue irec;
                        if (rst.EOF)
                            irec = new StructureDb.Record();
                        else
                        {
                            irec = rst;
                            if (!this.OverWriteFlag)
                                willRegistStructrue = (rst.GetValue(CmpdDbManager.MolID_FieldName) == null);
                        }

                        irec.SetValue(CmpdDbManager.CASRN_FieldName, casrn);
                        irec.SetValue(CmpdDbManager.CasIndexName_FieldName, index_name);
                        irec.SetValue(CmpdDbManager.MolecularFormula_FieldName, molecular_formula);
                        irec.SetValue(CmpdDbManager.Copyright_FieldName, copyright);

                        TempFile tempMolFile = null;
                        if (willRegistStructrue)
                        {
                            tempMolFile = new TempFile(".mol");
                            Utility.GenerateFileFromString(molInString, tempMolFile.Path);
                            irec.SetValue(CmpdDbManager.MolID_FieldName, tempMolFile.Path);
                        }

                        if (irec is StructureDb.Record)
                            rst.Add((StructureDb.Record)irec);

                        if (tempMolFile != null)
                            tempMolFile.Dispose();

                        rst.Filter = "";

                        IncProgressCount();
                    }
                    ResetProgressCount();
                }
            }
            ResetProgress();
        }

        private static string NormalizeCASRN(string input)
        {
            return input.StartsWith("CAS-") ? input.Substring(4) : input;
        }

        private static string NormalizeCASIN(string input)
        {
            return input == "INDEX NAME NOT YET ASSIGNED" ? "" : input;
        }

        private void LoadSubstancesFromCSV(IEnumerable<string> filenames)
        {
            using (var rstMolTable = db.ObtainRecordset(CmpdDbManager.MolTable_TableName))
            using (var rstOtherNamesTable = db.ObtainRecordset(CmpdDbManager.OtherNames_TableName))
            {
                foreach (var filename in filenames)
                {
                    var theFileName = Path.GetFileName(filename);
                    SetFileName(theFileName);

                    var extractor = new SubstancesCSVExtractor(filename);
                    foreach (var info in extractor.GetSubstancesInfo())
                    {
                        if (info.CASRN == null)
                            info.CASRN = CmpdDbManager.GeneratePseudoCASRN();
                        RegisterSubstanceToMolTable(rstMolTable, rstOtherNamesTable, info);

                        IncProgressCount();
                    }
                    ResetProgressCount();
                }
                ResetProgress();
            }
        }

        private void LoadCAOnlineFiles(IEnumerable<string> referenceFiles)
        {
            LoadReferenceFiles(referenceFiles, referenceFileName => 
                {
                    var tor = new CAOLExtractor(referenceFileName);
                    tor.ProgressIncrementer = IncProgressCount;
                    return tor;
                });
        }

        private void LoadSciFinderReferencesFiles(IEnumerable<string> referenceFiles)
        {
            LoadReferenceFiles(referenceFiles, referenceFileName => 
                {
                    var tor = new SciFinderReferencesExtractor(referenceFileName);
                    tor.ProgressIncrementer = IncProgressCount;
                    tor.TotalNumbrSetter = SetTotalNumber;
                    return tor;
                });
        }

        private void LoadReferenceFiles(IEnumerable<string> referenceFiles, Func<string, IEnumerable<ReferenceInfo>> extractorCreator)
        {
            using (var rstDocuments = db.ObtainRecordset(CmpdDbManager.Documents_TableName))
            using (var rstSubstances = db.ObtainRecordset(CmpdDbManager.Substances_TableName))
            using (var rstMolTable = db.ObtainRecordset(CmpdDbManager.MolTable_TableName))
            {
                foreach (var referenceFile in referenceFiles)
                {
                    SetFileName(Path.GetFileName(referenceFile));

                    var refsExtr = extractorCreator(referenceFile);

                    foreach (var refInfo in refsExtr)
                    {
                        if (refInfo.AccessionNumber != null)
                        {
                            if (!RegisterReferenceInfo(rstDocuments, refInfo))
                                continue;
                        }

                        foreach (var entry in refInfo.SubstancesInfo)
                        {
                            if (entry.CASRN == null)
                                entry.CASRN = CmpdDbManager.GeneratePseudoCASRN();
                            RegisterSubstanceInfo(rstSubstances, refInfo, entry);
                            RegisterSubstanceToMolTable(rstMolTable, entry);
                        }

                        rstDocuments.Filter = "";

                        IncProgressCount();
                    }
                    ResetProgressCount();
                }
                ResetProgress();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="docRst"></param>
        /// <param name="refInfo"></param>
        /// <returns><value>true</value> if we will add reference info</returns>
        private bool RegisterReferenceInfo(ChemFinderStructureDb.ChemFinderDbRecordset docRst, ReferenceInfo refInfo)
        {
            docRst.Filter = CmpdDbManager.AccessionNumber_FieldName + "= '" + refInfo.AccessionNumber + "'";
            IGetSetValue recDoc;
            if (docRst.EOF)
                recDoc = new StructureDb.Record();
            else
            {
                if (!this.OverWriteFlag)
                    return false;
                recDoc = docRst;
            }

            SetValueIfNotEmpty(recDoc, CmpdDbManager.AccessionNumber_FieldName, refInfo.AccessionNumber);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.Title_FieldName, refInfo.Title);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.Author_FieldName, refInfo.By);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.CorporateSource_FieldName, refInfo.CorporateSource);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.CorporateSource_FieldName, refInfo.PatentAssignee);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.Source_FieldName, refInfo.Source);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.Publisher_FieldName, refInfo.Publisher);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.DocumentType_FieldName, refInfo.DocumentType);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.Language_FieldName, refInfo.Language);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.PatentInfomation_FieldName, refInfo.PatentInfomation);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.Abstract_FieldName, refInfo.Abstract);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.AbstractBitmap_FieldName, refInfo.AbstractImage);
            SetValueIfNotEmpty(recDoc, CmpdDbManager.Copyright_FieldName, refInfo.Copyright);
            if (recDoc is StructureDb.Record)
                docRst.Add((StructureDb.Record)recDoc);

            return true;
        }

        private static void RegisterSubstanceInfo(
            ChemFinderStructureDb.ChemFinderDbRecordset rst, 
            ReferenceInfo refInfo, SubstanceInfo entry)
        {
            if (refInfo.AccessionNumber != null)
            {
                rst.Filter = CmpdDbManager.AccessionNumber_FieldName + "= '" + refInfo.AccessionNumber + "'" + " AND "
                           + CmpdDbManager.CASRN_FieldName + "= '" + entry.CASRN + "'";
                IGetSetValue recCmp;
                if (rst.EOF)
                    recCmp = new StructureDb.Record();
                else
                    recCmp = rst;
                SetValueIfNotEmpty(recCmp, CmpdDbManager.AccessionNumber_FieldName, refInfo.AccessionNumber);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.OrderInDoc_FieldName, entry.Order);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.CASRN_FieldName, entry.CASRN);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.NameInDoc_FieldName, entry.Name);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.Keywords_FieldName, entry.Keywords);
                if (recCmp is StructureDb.Record)
                    rst.Add((StructureDb.Record)recCmp);

                rst.Filter = "";
            }
        }

        private void RegisterSubstanceToMolTable(
            ChemFinderStructureDb.ChemFinderDbRecordset rstMol,
            ChemFinderStructureDb.ChemFinderDbRecordset rstOtherNames,
            SubstanceInfo entry)
        {
            RegisterSubstanceAndCmpdFileToMolTable(rstMol, rstOtherNames, entry, null, null);
        }

        private void RegisterSubstanceToMolTable(
            ChemFinderStructureDb.ChemFinderDbRecordset rstMol,
            SubstanceInfo entry)
        {
            RegisterSubstanceToMolTable(rstMol, null, entry);
        }

        private void RegisterSubstanceAndCmpdFileToMolTable(
            ChemFinderStructureDb.ChemFinderDbRecordset rstMol, 
            SubstanceInfo entry, 
            string cmpdfile, string structsource)
        {
            RegisterSubstanceAndCmpdFileToMolTable(rstMol, null, entry, cmpdfile, structsource);
        }

        private void RegisterSubstanceAndCmpdFileToMolTable(
            ChemFinderStructureDb.ChemFinderDbRecordset rstMol,
            ChemFinderStructureDb.ChemFinderDbRecordset rstOtherNames,
            SubstanceInfo entry,
            string cmpdfile, string structsource)
        {
            if (!(entry.CAIndexName == null && entry.MolecularFormula == null && entry.Bitmap == null && entry.Name == null && entry.OtherNames == null && entry.Smiles == null))
            {
                rstMol.Filter = CmpdDbManager.CASRN_FieldName + "= '" + entry.CASRN + "'";
                IGetSetValue recCmp;
                if (rstMol.EOF)
                    recCmp = new StructureDb.Record();
                else
                    recCmp = rstMol;
                SetValueIfNotEmpty(recCmp, CmpdDbManager.CASRN_FieldName, entry.CASRN);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.CasIndexName_FieldName, entry.CAIndexName);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.OtherNames_FieldName, entry.Name);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.MolecularFormula_FieldName, entry.MolecularFormula);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.ClassIdentifier_FieldName, entry.ClassIdentifier);
                if (LoadSubstanceImageFlag)
                    SetValueIfNotEmpty(recCmp, CmpdDbManager.StructureBmp_FieldName, entry.Bitmap);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.Copyright_FieldName, entry.Copyright);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.MolID_FieldName, cmpdfile);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.StructureSource_FieldName, structsource);
                SetValueIfNotEmpty(recCmp, CmpdDbManager.Smiles_FieldName, entry.Smiles);

                if (recCmp is StructureDb.Record)
                    rstMol.Add((StructureDb.Record)recCmp);
                rstMol.Filter = "";
            }
            if (rstOtherNames != null && entry.OtherNames != null)
            {
                try
                {
                    foreach (var name in entry.OtherNames)
                    {
                        rstOtherNames.Filter =
                            CmpdDbManager.CASRN_FieldName + "= '" + entry.CASRN + "' AND " +
                            CmpdDbManager.Name_FieldName + "= '" + name + "'";
                        IGetSetValue recCmp;
                        if (rstOtherNames.EOF)
                        {
                            recCmp = new StructureDb.Record();
                            SetValueIfNotEmpty(recCmp, CmpdDbManager.CASRN_FieldName, entry.CASRN);
                            SetValueIfNotEmpty(recCmp, CmpdDbManager.Name_FieldName, name);
                            rstOtherNames.Add((StructureDb.Record)recCmp);
                        }
                    }
                }
                finally
                {
                    rstOtherNames.Filter = "";
                }
            }
        }

        private void LoadSmilesFromList(IEnumerable<string> filenames)
        {
            using (var rstDocuments = db.ObtainRecordset(CmpdDbManager.Documents_TableName))
            using (var rstSubstances = db.ObtainRecordset(CmpdDbManager.Substances_TableName))
            using (var rstMolTable = db.ObtainRecordset(CmpdDbManager.MolTable_TableName))
            {
                foreach (var filename in filenames)
                {
                    var theFileName = Path.GetFileName(filename);
                    SetFileName(theFileName);

                    var refInfo = new ReferenceInfo();
                    var pseudoAN = "~" + theFileName.GetHashCode().ToString("X8");
                    refInfo.AccessionNumber = pseudoAN;
                    refInfo.DocumentType = CmpdDbManager.DocumentType_Unknown;
                    refInfo.Source = theFileName;
                    if (refInfo.AccessionNumber != null)
                        RegisterReferenceInfo(rstDocuments, refInfo);

                    var extractor = new SmilesListExtractor(filename);
                    var base_id = CmpdDbManager.GeneratePseudoCASRNBase();
                    int n = 0;
                    foreach (var entry in extractor.GetSubstancesInfo())
                    {
                        if (entry.CASRN == null)
                            entry.CASRN = CmpdDbManager.GeneratePseudoCASRN(base_id, n++);
                        RegisterSubstanceToMolTable(rstMolTable, entry);
                        RegisterSubstanceInfo(rstSubstances, refInfo, entry);

                        IncProgressCount();
                    }
                }
                ResetProgress();
            }
        }

        private void LoadMoleculesFromCompoundNameList(IEnumerable<string> filenames)
        {
            using (var rstDocuments = db.ObtainRecordset(CmpdDbManager.Documents_TableName))
            using (var rstSubstances = db.ObtainRecordset(CmpdDbManager.Substances_TableName))
            using (var rstMolTable = db.ObtainRecordset(CmpdDbManager.MolTable_TableName))
            {
                foreach (var filename in filenames)
                {
                    var theFileName = Path.GetFileName(filename);
                    SetFileName(theFileName);

                    var refInfo = new ReferenceInfo();
                    var pseudoAN = "~" + theFileName.GetHashCode().ToString("X8");
                    refInfo.AccessionNumber = pseudoAN;
                    refInfo.DocumentType = CmpdDbManager.DocumentType_Unknown;
                    refInfo.Source = theFileName;
                    if (refInfo.AccessionNumber != null)
                        RegisterReferenceInfo(rstDocuments, refInfo);

                    var extractor = new ListExtractor(filename);
                    var base_id = CmpdDbManager.GeneratePseudoCASRNBase();
                    int n = 0;
                    foreach (var entry in extractor.GetSubstancesInfo())
                    {
                        if (entry.CASRN == null)
                                entry.CASRN = CmpdDbManager.GeneratePseudoCASRN(base_id, n++);
                        RegisterSubstanceToMolTable(rstMolTable, entry);
                        RegisterSubstanceInfo(rstSubstances, refInfo, entry);
                    }
                }
                ResetProgress();
            }
        }

        private static void SetValueIfNotEmpty(IGetSetValue irec, string fieldname, object value)
        {
            if (value != null)
                irec.SetValue(fieldname, value);
        }

        private static void SetValueIfNotEmpty(IGetSetValue irec, string fieldname, string value)
        {
            if (value != null && value != "")
                irec.SetValue(fieldname, value);
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
                        case "-a":
                            program.AppendFlag = true;
                            break;
                        case "-o":
                            indexArgs++;
                            program.OutputPath = args[indexArgs];
                            break;
                        case "-O":
                            program.OverWriteFlag = true;
                            break;
                        case "-C":
                            indexArgs++;
                            {
                                var filename = args[indexArgs];
                                switch (Path.GetExtension(filename))
                                {
                                    case ".rtf":
                                    case ".doc":
                                    case ".docx":
                                    case ".docm":
                                        program.CasOnlineFiles.Add(filename);
                                        break;
                                    default:
                                        throw new ApplicationException("Unknown extension '" + filename + "'.");
                                }
                            }
                            break;
                        case "-I":
                            program.LoadSubstanceImageFlag = true;
                            break;
                        case "-?":
                        case "--help":
                            throw new ApplicationException(
                                  "usege: MergeSF [option]... ([+]filename)\n"
                                + "-o file-name\tSpecify output CFX file name\n"
                                + "-a \tAppend them into CFX\n"
                                + "-C file-name\tInput filename from CAS Online\n"
                                + "-I import substance images.\n"
                                + "filename from SciFinder"
                            );
                        default:
                            throw new ApplicationException("Unknown option '" + arg + "'.");
                    }
                }
                else
                {
                    var filename = args[indexArgs];
                    program.AddInputPath(filename);
                }
            }
        }

        private void AddInputPath(string filename)
        {
            switch (Path.GetExtension(filename))
            {
                case ".rtf":
                case ".doc":
                case ".docx":
                case ".docm":
                    this.ReferenceFiles.Add(filename);
                    break;
                case ".lst":
                    this.ListFiles.Add(filename);
                    break;
                case ".sdf":
                    this.SDFiles.Add(filename);
                    break;
                case ".smi":
                case ".rsmi":
                    this.SmilesListFiles.Add(filename);
                    break;
                case ".txt":
                case ".csv":
                    this.CSVFiles.Add(filename);
                    break;
                case ".cfx":
                    this.CfxFiles.Add(filename);
                    break;
                default:
                    throw new ApplicationException("Unknown extension '" + filename + "'.");
            }
        }

        public IEnumerable<string> InputPaths
        {
            set
            {
                foreach (var filename in value)
                    AddInputPath(filename);
            }
        }

        public void Dispose()
        {
        }
    }
}
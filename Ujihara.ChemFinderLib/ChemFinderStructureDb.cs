using CambridgeSoft.ChemScript19;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Ujihara.Chemistry.IO;
using MolServer = MolServer19;

namespace Ujihara.Chemistry
{
    public class ChemFinderStructureDb
        : StructureDb, IDisposable
    {
        private static Regex IsMolIDFieldNameRE = new Regex("^Mol_ID(?<num>\\d*)$", RegexOptions.Compiled);
        public static bool IsMolIDFieldName(string fieldName)
        {
            return IsMolIDFieldNameRE.IsMatch(fieldName);
        }

        public static bool IsMolIDFieldName(string fieldName, out int num)
        {
            var match = IsMolIDFieldNameRE.Match(fieldName);
            if (match.Success)
            {
                var s = match.Groups["num"].Value;
                if (s == "")
                    num = -1;
                else
                    num = int.Parse(s);
                return true;
            }
            else
            {
                num = -1;
                return false;
            }
        }

        public static string GetStructureFieldName()
        {
            return "Structure"; 
        }

        public static string GetStructureFieldName(int num)
        {
            if (num < 0)
                return GetStructureFieldName();
            return GetStructureFieldName() + num.ToString();
        }

        protected string OleDbConnectionString
        {
            get
            {
                return "Provider=" + CfxManager.DefaultProviderName + ";Data Source=" + RecordSource;
            }
        }
        
        // We consider ADODB is better than System.Data at this time becuase table can have no key field.
        private ADODB.Connection _Connection = null;
        protected ADODB.Connection Connection
        {
            get
            {
                if (_Connection == null)
                {
                    _Connection = new ADODB.Connection();
                    _Connection.Open(OleDbConnectionString, null, null, 0);
                }
                return _Connection;
            }
        }

        private MolServer.Document _MstDocument = null;
        public MolServer.Document MstDocument
        {
            get
            {
                if (_MstDocument == null)
                {
                    _MstDocument = new MolServer.Document();
                    _MstDocument.Open(this.MstFullPath, (int)MolServerUtility.ToMSOpenModes(this.Access), "");
                }
                return _MstDocument;
            }
        }

        /// <summary>
        /// Timeout in millisecond. Default value is <value>-1</value> (no timeout).
        /// </summary>
        public int Timeout { get; set; }
        private const int DefaultTimeout = -1;

        private string RecordSource { get; set; }
        public string TableName { get; private set; }
        private string MstFullPath { get; set; }
        private FileAccess Access { get; set; }

        public ChemFinderStructureDb(string path, FileAccess access)
            : base()
        {
            var cfxFullPath = Path.GetFullPath(path);
            var cfxDirName = Path.GetDirectoryName(cfxFullPath);
            
            var cfxMan = new CfxManager();
            cfxMan.Load(cfxFullPath);
            this.RecordSource = Path.Combine(cfxDirName, cfxMan.DatabaseFileName);
            this.TableName = cfxMan.TableName;
            this.MstFullPath = Path.Combine(cfxDirName, cfxMan.MstName);

            this.Access = access;
            this.Timeout = DefaultTimeout;
        }

        public ChemFinderDbRecordset ObtainRecordset()
        {
            return new ChemFinderDbRecordset(this);
        }

        public ChemFinderDbRecordset ObtainRecordset(string query)
        {
            return new ChemFinderDbRecordset(this, query);
        }

        public bool QueryRecordExist(string sql)
        {
            var cmd = new ADODB.Command();
            try
            {
                cmd.ActiveConnection = this.Connection;
                cmd.CommandText = sql;

                object missingType = Type.Missing;
                var rst = cmd.Execute(out missingType, ref missingType, 0);
                try
                {
                    return !rst.EOF;
                }
                finally
                {
                    Utility.ReleaseComObject(rst);
                }
            }
            finally
            {
                Utility.ReleaseComObject(cmd);
            }
        }

        public void ExecuteNonQuery(string sql)
        {
            var cmd = new ADODB.Command();
            try
            {
                cmd.ActiveConnection = this.Connection;
                cmd.CommandText = sql;

                object missingType = Type.Missing;
                var rst = cmd.Execute(out missingType, ref missingType, 0);
                Utility.ReleaseComObject(rst);
            }
            finally
            {
                Utility.ReleaseComObject(cmd);
            }
        }

        public void CreateTable(string tableName)
        {
            Utility.CheckSQLName(tableName);

            // FIXME: need safer code
            ExecuteNonQuery("CREATE TABLE [" + tableName + "]");
        }

        public void CreateField(string name, string type)
        {
            CreateField(this.TableName, name, type);
        }

        public void CreateField(string tableName, string name, string type)
        {
            Utility.CheckSQLName(tableName);
            Utility.CheckSQLName(name);
            Utility.CheckSQLName(type);

            object restriction = new object[] { null, null, tableName, name};
            var schema = this.Connection.OpenSchema(
                ADODB.SchemaEnum.adSchemaColumns,
                restriction, Type.Missing);
            try
            {
                if (schema.BOF && schema.EOF) 
                {
                    // fileld is not exist
                    var cmd = new ADODB.Command();
                    try
                    {
                        cmd.ActiveConnection = this.Connection;
                        // FIXME: need safer code
                        cmd.CommandText = "ALTER TABLE "
                            + "[" + tableName + "] "
                            + "ADD [" + name + "] " + type;

                        object missingType = Type.Missing;
                        var rst = cmd.Execute(out missingType, ref missingType, 0);
                        Utility.ReleaseComObject(rst);
                    }
                    finally
                    {
                        Utility.ReleaseComObject(cmd);
                    }
                }
            }            
            finally
            {
                Utility.ReleaseComObject(schema);
            }
        }

        private static string BuildFieldCList(IEnumerable<string> names)
        {
            StringBuilder sb = null;
            foreach (var name in names)
            {
                Utility.CheckSQLName(name);
                if (sb == null)
                    sb = new StringBuilder();
                else
                    sb.Append(',');
                sb.Append('[').Append(name).Append(']');
            }
            return sb.ToString();
        }

        private object GetValueEx(Record record, int i)
        {
            var key = record.Keys[i];
            object value = record.Values[i];

            if (IsMolIDFieldName(key))
            {
                if (value == null || value is MolServer.Molecule)
                {
                    var mol = (MolServer.Molecule)value;

                    this.MstDocument.Lock();

                    var molID = this.MstDocument.AssignID(true);
                    if (mol != null)
                        this.MstDocument.PutMol(mol, molID);
                    value = molID;

                    this.MstDocument.Unlock();
                }
            }
            return value;
        }

        // Get/set with handling Mol_ID on rst. 
        private static object GetValueEx(MolServer.Document mstDoc, ADODB.Recordset rst, string fieldName)
        {
            var value = GetValue(rst, fieldName);
            
            if (value is DBNull)
                return null;

            if (IsMolIDFieldName(fieldName))
            {
                var molID = value as int?;
                if (molID != null)
                    return mstDoc.GetMol((int)value);
            }
            return value;
        }

        private static MolServer.Molecule CreateMoleculeFromFile(string value)
        {
            if (value == null)
                return null;

            var mol = new MolServer.Molecule();
            mol.Read(value);
            return mol;
        }

        private static void SetValueEx(MolServer.Document mstDoc, ADODB.Recordset rst, string fieldName, object value)
        {
            if (rst == null || fieldName == null)
                throw new ArgumentException();

            int num;
            
            if (IsMolIDFieldName(fieldName, out num))
            {
                if (value == null || value is string || value is MolServer.Molecule)
                {
                    MolServer.Molecule mol = null;
                    if (value is MolServer.Molecule)
                        mol = (MolServer.Molecule)value;
                    else if (value is string)
                        mol = CreateMoleculeFromFile((string)value);
                    try
                    {
                        var field_value = GetValue(rst, fieldName);
                        if (field_value == null || field_value is DBNull)
                            field_value = mstDoc.AssignID(true);
                        var molID = field_value as int?;
                        if (molID != null)
                        {
                            mstDoc.PutMol(mol, (int)molID);
                            value = molID;

                            SetValue(rst, GetStructureFieldName(num), null);
                        }
                    }
                    finally
                    {
                        Utility.ReleaseComObject(mol);
                    }
                }
            }
            SetValue(rst, fieldName, value);
        }

        // Utilities for ADODB
        
        private static object GetValue(ADODB.Recordset rst, string fieldName)
        {
            var fields = rst.Fields;
            try
            {
                var field = fields[fieldName];
                try
                {
                    return field.Value;
                }
                finally
                {
                    Utility.ReleaseComObject(field);
                }
            }
            finally
            {
                Utility.ReleaseComObject(fields);
            }
        }

        private static void SetValue(ADODB.Recordset rst, string fieldName, object value)
        {
            var fields = rst.Fields;
            var field = fields[fieldName];
            try
            {
                field.Value = value;
            }
            finally
            {
                Utility.ReleaseComObject(field);
                Utility.ReleaseComObject(fields);
            }
        }

        public void Register(StructureData csmol, int molID)
        {
            using (var cdx = new TempFile(".cdx"))
            {
                csmol.WriteFile(cdx.Path, "chemical/x-cdx");
                var mol = new MolServer.Molecule();
                try
                {
                    mol.Read(cdx.Path);
                    this.Register(mol, molID);
                }
                finally
                {
                    Utility.ReleaseComObject(mol);
                }
            }
        }

        public void Register(MolServer.Molecule mol, int molID)
        {
            this.MstDocument.PutMol(mol, molID);
        }

        public abstract class ChemFinderDbRecordsetBase
            : StructureDb.Recordset
        {
            protected ChemFinderStructureDb Db { get; set; }
        }

        public class ChemFinderDbRecordset
            : ChemFinderDbRecordsetBase
        {
            ADODB.Recordset Recordset { get; set; }

            public ChemFinderDbRecordset(ChemFinderStructureDb db)
                : this(db, db.TableName)
            {
            }

            public ChemFinderDbRecordset(ChemFinderStructureDb db, string query)
            {
                this.Db = db;
                this.Recordset = new ADODB.Recordset();
                this.Recordset.Open
                    (query, this.Db.Connection,
                    ADODB.CursorTypeEnum.adOpenForwardOnly,
                    ADODB.LockTypeEnum.adLockOptimistic, 0);
            }

            public void DeleteCurrent()
            {
                this.Recordset.Delete(ADODB.AffectEnum.adAffectCurrent);
            }

            public string Filter
            {
                get
                {
                    return this.Recordset.Filter as string;
                }

                set
                {
                    this.Recordset.Filter = value;
                }
            }

            public override void MoveNext()
            {
                Recordset.MoveNext();
            }

            public override bool EOF
            {
                get { return Recordset.EOF; }
            }

            public  void Update()
            {
                Recordset.Update(Type.Missing, Type.Missing);
            }

            public override object GetRawValue(string fieldName)
            {
                return ChemFinderStructureDb.GetValue(this.Recordset, fieldName);
            }

            public override object GetValue(string fieldName)
            {
                return GetValueEx(this.Db.MstDocument, this.Recordset, fieldName);
            }

            public object EGetValue(string fieldName)
            {
                try
                {
                    return GetValue(fieldName);
                }
                catch (Exception)
                {
                    return null;
                }
            }

            public override void SetValue(string fieldName, object value)
            {
                SetValueEx(this.Db.MstDocument, this.Recordset, fieldName, value);
            }

            public void Add(Record record)
            {
                this.Recordset.AddNew(Type.Missing, Type.Missing);

                var fields = this.Recordset.Fields;
                try
                {
                    for (var i = 0; i < record.Count; i++)
                    {
                        var key = record.Keys[i];
                        var value = record.Values[i];

                        if (IsMolIDFieldName(key))
                        {
                            if (value == null || value is string)
                            {
                                var mol = CreateMoleculeFromFile((string)value);

                                try
                                {
                                    this.Db.MstDocument.Lock();
                                    var molID = this.Db.MstDocument.AssignID(true);
                                    if (mol != null)
                                    {
                                        this.Db.MstDocument.PutMol(mol, molID);
                                    }
                                    this.Db.MstDocument.Unlock();

                                    value = molID;
                                }
                                finally
                                {
                                    Utility.ReleaseComObject(mol);
                                }
                            }
                        }
                        fields[record.Keys[i]].Value = value;
                    }
                }
                finally
                {
                    Utility.ReleaseComObject(fields);
                }

                
                this.Recordset.Update(Type.Missing, Type.Missing);
            }

            private bool disposed = false;

            public override void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!disposed)
                {
                    if (disposing)
                    {
                        // manageed resources
                    }

                    Utility.ReleaseComObject(Recordset);

                    disposed = true;
                }
            }

            ~ChemFinderDbRecordset()
            {
                Dispose(false);
            }
        }

        private Dictionary<string, MolServer.Hitlist> FieldNameToDomains = new Dictionary<string, MolServer19.Hitlist>();

        MolServer.Hitlist GetMolIDsInThisDb(string fieldName)
        {
            Utility.CheckSQLName(fieldName);

            MolServer.Hitlist hl;
            if (FieldNameToDomains.TryGetValue(fieldName, out hl))
            {
                return hl;
            }
            else
            {
                var cmd = new ADODB.Command();
                try
                {
                    cmd.ActiveConnection = this.Connection;
                    cmd.CommandText = "SELECT " + fieldName + " FROM " + this.TableName;    // FIXME: need safer code
                    cmd.CommandType = ADODB.CommandTypeEnum.adCmdText;
                    cmd.Prepared = true;

                    object missingType = Type.Missing;
                    var rst = cmd.Execute(out missingType, ref missingType, 0);
                    try
                    {
                        hl = new MolServer.Hitlist();
                        while (!rst.EOF)
                        {
                            hl.AddHit((int)rst.Fields[fieldName].Value);
                            rst.MoveNext();
                        }

                        FieldNameToDomains.Add(fieldName, hl);
                        return hl;
                    }
                    finally
                    {
                        Utility.ReleaseComObject(rst);
                    }
                }
                finally
                {
                    Utility.ReleaseComObject(cmd);
                }
            }
        }

        public override StructureDb.Recordset Search(string fieldName, MolServer.Molecule mol)
        {
            if (mol == null)
                return StructureDb.Recordset.Empty;

            var si = new MolServer.searchInfo();
            try
            {
                ChemFinderMolSearchRecordset sr;

                si.MolQuery = mol;
                si.Domain = GetMolIDsInThisDb(fieldName);
                si.FullStructure = true;
                si.StereoDB = true; // double bond
                si.StereoTetr = true;   // enantio
                si.RelativeTetStereo = true;    // diastereo
                sr = Search(fieldName, si);
                if (!sr.IsEmpty)
                    return sr;
                else
                    sr.Dispose();

                si.StereoDB = false; // double bond
                si.StereoTetr = false;   // enantio
                si.RelativeTetStereo = false;    // diastereo
                sr = Search(fieldName, si);
                if (!sr.IsEmpty)
                {
                    sr.Properties["Comment"] = "stereo ignored";
                    return sr;
                }
                else
                    sr.Dispose();
                return StructureDb.Recordset.Empty;
            }
            finally
            {
                Utility.ReleaseComObject(si);
            }
        }

        public virtual ChemFinderMolSearchRecordset Search(string fieldName, MolServer.searchInfo si)
        {
            return new ChemFinderStructureDb.ChemFinderMolSearchRecordset(this, fieldName, si);
        }

        public class ChemFinderMolSearchRecordset
            : ChemFinderDbRecordsetBase
        {
            private string FieldName { get; set; }
            private MolServer.Search so;

            private bool Started { get; set; }
            private MolServer.Hitlist finalHitlist = null;
            private int Index { get; set; }

            public ChemFinderMolSearchRecordset(ChemFinderStructureDb db, string fieldName, MolServer.searchInfo si)
            {
                Utility.CheckSQLName(fieldName);

                this.Db = db;
                this.FieldName = fieldName;
                this.so = db.MstDocument.Search(si);
                this.Started = false;
                this.Index = 0;
            }

            public bool IsEmpty
            {
                get
                {
                    var id = GetAt(0);
                    return id == -1;
                }
            }

            public override void MoveNext()
            {
                var id = GetAt(Index);
                if (id == -1)
                    throw new InvalidOperationException();

                Index++;
            }

            public override bool EOF
            {
                get 
                {
                    var id = GetAt(Index);
                    return id == -1;
                }
            }

            private ADODB.Recordset GetRst(int molId)
            {
                const string P_ID = "@ID";
                var cmd = new ADODB.Command();
                try
                {
                    cmd.ActiveConnection = this.Db.Connection;
                    cmd.CommandText = "SELECT * FROM "
                        + Db.TableName + " "
                        + "WHERE " + this.FieldName + " = " + P_ID + "";
                    cmd.CommandType = ADODB.CommandTypeEnum.adCmdText;
                    cmd.Prepared = true;

                    var paramz = cmd.Parameters;
                    var param = cmd.CreateParameter(P_ID, ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 0, molId);
                    try
                    {
                        paramz.Append(param);

                        object missingType = Type.Missing;
                        var rst = cmd.Execute(out missingType, ref missingType, 0);
                        return rst;
                    }
                    finally
                    {
                        Utility.ReleaseComObject(param);
                        Utility.ReleaseComObject(paramz);
                    }
                }
                finally
                {
                    Utility.ReleaseComObject(cmd);
                }
            }

            public override object GetRawValue(string fieldName)
            {
                return GetValue(fieldName);
            }

            public override object  GetValue(string fieldName)
            {
                Utility.CheckSQLName(fieldName);

                var rst = GetRst(this.GetAt(this.Index));
                try
                {
                    if (rst.EOF)
                        return null;
                    return GetValueEx(this.Db.MstDocument, rst, fieldName);
                }
                finally
                {
                    Utility.ReleaseComObject(rst);
                }
            }

            public override void SetValue(string fieldName, object value)
            {
                Utility.CheckSQLName(fieldName);

                var rst = GetRst(this.GetAt(this.Index));
                try
                {
                    if (rst.EOF)
                        throw new KeyNotFoundException();
                    SetValueEx(this.Db.MstDocument, rst, fieldName, value);
                }
                finally
                {
                    Utility.ReleaseComObject(rst);
                }
            }

            private int GetAt(int index)
            {
                if (!Started)
                {
                    so.Start();
                    Started = true;
                }

                for (; ; )
                {
                    MolServer.Hitlist hitlistToRelease = null;
                    try
                    {
                        MolServer.Hitlist hitlist = finalHitlist;
                        MolServer.SRStatusFlags status;
                        if (hitlist == null)
                        {
                            so.WaitForCompletion(100); // waits 100 ms
                            status = (MolServer.SRStatusFlags)so.Status;
                            hitlist = so.Hitlist;

                            if (status == (int)MolServer.SRStatusFlags.kSRDone)
                                finalHitlist = hitlist;
                            else
                                hitlistToRelease = hitlist;
                        }
                        else
                        {
                            status = MolServer.SRStatusFlags.kSRDone;
                        }
                        if (hitlist.Count > index)
                        {
                            return hitlist.get_At(index);
                        }
                        if (status == MolServer.SRStatusFlags.kSRDone)
                            return -1;
                    }
                    finally
                    {
                        Utility.ReleaseComObject(hitlistToRelease);
                    }
                }
            }

            private bool disposed = false;

            public override void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!disposed)
                {
                    if (disposing)
                    {
                        // manageed resources
                    }

                    if (so.Status != (int)MolServer.SRStatusFlags.kSRDone)
                    {
                        so.Stop();
                        so.WaitForCompletion(DefaultTimeout);
                    }
                    Utility.ReleaseComObject(finalHitlist);
                    finalHitlist = null;
                    Utility.ReleaseComObject(so);
                    so = null;

                    disposed = true;
                }
            }

            ~ChemFinderMolSearchRecordset()
            {
                Dispose(false);
            }
        }

        private bool disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // no managed objects.
                }
                if (_MstDocument != null)
                {
                    _MstDocument.Close();
                    Utility.ReleaseComObject(_MstDocument);
                    _MstDocument = null;
                }

                if (_Connection != null)
                {
                    _Connection.Close();
                    Utility.ReleaseComObject(_Connection);
                    _Connection = null;
                }

                disposed = true;
            }
        }

        ~ChemFinderStructureDb()
        {
            Dispose(false);
        }

        static Dictionary<Type, ADODB.DataTypeEnum> TypeToDataTypeEnumDic;
        static ChemFinderStructureDb()
        {
            TypeToDataTypeEnumDic = new Dictionary<Type, ADODB.DataTypeEnum>();
            TypeToDataTypeEnumDic.Add(typeof(string), ADODB.DataTypeEnum.adBSTR);
            TypeToDataTypeEnumDic.Add(typeof(bool), ADODB.DataTypeEnum.adBoolean);
            TypeToDataTypeEnumDic.Add(typeof(byte), ADODB.DataTypeEnum.adUnsignedTinyInt);
            TypeToDataTypeEnumDic.Add(typeof(sbyte), ADODB.DataTypeEnum.adTinyInt);
            TypeToDataTypeEnumDic.Add(typeof(Int16), ADODB.DataTypeEnum.adSmallInt);
            TypeToDataTypeEnumDic.Add(typeof(UInt16), ADODB.DataTypeEnum.adUnsignedSmallInt);
            TypeToDataTypeEnumDic.Add(typeof(Int32), ADODB.DataTypeEnum.adInteger);
            TypeToDataTypeEnumDic.Add(typeof(UInt32), ADODB.DataTypeEnum.adUnsignedInt);
            TypeToDataTypeEnumDic.Add(typeof(Int64), ADODB.DataTypeEnum.adBigInt);
            TypeToDataTypeEnumDic.Add(typeof(UInt64), ADODB.DataTypeEnum.adUnsignedBigInt);
            TypeToDataTypeEnumDic.Add(typeof(float), ADODB.DataTypeEnum.adSingle);
            TypeToDataTypeEnumDic.Add(typeof(double), ADODB.DataTypeEnum.adDouble);
            TypeToDataTypeEnumDic.Add(typeof(byte[]), ADODB.DataTypeEnum.adBinary);
        }

        static void ToDataTypeEnum(ref object o, out ADODB.DataTypeEnum dataType, out int length)
        {
            if (o == null)
            {
                dataType = ADODB.DataTypeEnum.adInteger;
                length = 0;
                
                return;
            }
            Type type = o.GetType();
            ADODB.DataTypeEnum v;
            if (!TypeToDataTypeEnumDic.TryGetValue(type, out v))
                throw new ArgumentException();

            dataType = v;
            if (type == typeof(byte[]))
                length = ((byte[])o).Length;
            else
                length = 0;
        }
    }
}

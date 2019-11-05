using System;
using System.Collections.Generic;
using MolServer = MolServer19;
using System.Linq;

namespace Ujihara.Chemistry
{
    public abstract class StructureDb
    {
        public static StructureDb Empty { get { return EmptyStructureDb.Singleton; } }

        public abstract StructureDb.Recordset Search(string fieldName, MolServer.Molecule m);

        private sealed class EmptyStructureDb
            : StructureDb
        {
            private static EmptyStructureDb singleton = new EmptyStructureDb();
            public static EmptyStructureDb Singleton { get { return singleton; } }

            EmptyStructureDb()
                : base()
            {
            }

            public override StructureDb.Recordset Search(string fieldName, MolServer19.Molecule m)
            {
                return null;
            }
        }

        public abstract class Recordset
            : IGetSetValue, IDisposable
        {
            public static Recordset Empty { get { return EmptyRecordset.Singleton; } }
            private Dictionary<string, string> _Properties = new Dictionary<string, string>();
            public virtual IDictionary<string, string> Properties { get { return _Properties; } }

            public abstract void MoveNext();
            public abstract bool EOF { get; }
            public abstract object GetRawValue(string fieldName);
            public abstract object GetValue(string fieldName);
            public abstract void SetValue(string fieldName, object value);
            //public abstract void Update();

            private sealed class EmptyRecordset
                : Recordset
            {
                private static Recordset singleton = new EmptyRecordset();
                public static Recordset Singleton { get { return singleton; } }
                public override void MoveNext() { throw new InvalidOperationException(); }

                public override object GetRawValue(string fieldName)
                {
                    throw new InvalidOperationException();
                }

                public override object GetValue(string fieldName)
                { 
                    throw new InvalidOperationException(); 
                }

                public override void SetValue(string fieldName, object value)
                {
                    throw new InvalidOperationException();
                }

                public override bool EOF
                {
                    get { return true; }
                }

                public override void Dispose()
                {
                }
            }

            public abstract void Dispose();
        }

        /// <summary>
        /// Record for insertion into StructureDb.
        /// </summary>
        public class Record
            : IGetSetValue
        {
            private List<string> keys = new List<string>();
            private List<object> values = new List<object>();

            public Record()
            {
            }

            public int Count { get { return keys.Count; } }

            private int FindKey(string key)
            {
                return keys.FindIndex(n => n == key);
            }

            public IList<string> Keys { get { return keys; } }
            public IList<object> Values { get { return values; } }

            public object GetValue(int index)
            {
                return values[index];
            }

            public object GetValue(string key)
            {
                var index = FindKey(key);
                if (index == -1)
                    throw new InvalidOperationException();
                return values[index];
            }

            public void SetValue(string key, object value)
            {
                var index = FindKey(key);
                if (index == -1)
                {
                    keys.Add(key);
                    values.Add(value);
                }
                else
                    values[index] = value;
            }
        }
    }
}

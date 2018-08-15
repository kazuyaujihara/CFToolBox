using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Ujihara.Chemistry
{
    public static class Utility
    {
        public static void ReleaseComObject(object o)
        {
            if (o == null)
                return;
            Marshal.ReleaseComObject(o);
        }

        public static void DeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception)
            {
            }
        }

        public static string GetUniqueFileName(string path, string extension)
        {
            return Path.Combine(path, Guid.NewGuid().ToString() + extension);
        }

        public static string GetUniqueFileName(string extension)
        {
            return GetUniqueFileName(Path.GetTempPath(), extension);
        }

        private static char[] NGChars = "`.!'[]\"*?".ToCharArray();
        public static void CheckSQLName(string name)
        {
            foreach (var c in name)
            {
                if (c < 32 || NGChars.Contains(c))
                    throw new ArgumentException();
            }
        }

        public static byte[] FileToArray(string filename)
        {
            using (var fs = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                return ToArray(fs);
            }
        }

        public static byte[] ToArray(Stream stream)
        {
            using (var ms = new MemoryStream())
            {
                int b;
                while ((b = stream.ReadByte()) >= 0)
                    ms.WriteByte((byte)b);
                return ms.ToArray();
            }
        }

        private static Dictionary<Type, DbType> TypeToDbTypeDic { get; set; }

        static Utility()
        {
            TypeToDbTypeDic = new Dictionary<Type, DbType>
            {
                { typeof(string), DbType.String },
                { typeof(bool), DbType.Boolean },
                { typeof(byte), DbType.Byte },
                { typeof(sbyte), DbType.SByte },
                { typeof(Int16), DbType.Int16 },
                { typeof(UInt16), DbType.UInt16 },
                { typeof(Int32), DbType.Int32 },
                { typeof(UInt32), DbType.UInt32 },
                { typeof(Int64), DbType.Int64 },
                { typeof(UInt64), DbType.UInt64 },
                { typeof(DateTime), DbType.DateTime },
                { typeof(Decimal), DbType.Currency },
                { typeof(float), DbType.Double },
                { typeof(double), DbType.Double },
                { typeof(byte[]), DbType.Binary }
            };
        }

        public static DbType ToDbType(Type type)
        {
            if (TypeToDbTypeDic.TryGetValue(type, out DbType dbType))
            {
                return dbType;
            }
            throw new ArgumentException();            
        }

        public static IList<string> SemiColonSeparatedStringToList(string value)
        {
            return value.Split(';').Select(n => n.Trim()).Where(n => n != "").ToList();
        }

        public static string ToSemiColonSeparatedString(IEnumerable<string> strings)
        {
            var sb = new StringBuilder();
            foreach (var s in strings.Select(n => n.Trim()).Where(n => n != ""))
            {
                sb.Append(';').Append(s);
            }
            sb.Remove(0, 1);
            return sb.ToString();
        }

        public static void GenerateFileFromString(string str, string filenameToGenerate)
        {
            using (var t = new StreamWriter(filenameToGenerate))
            {
                t.Write(str);
            }
        }
    }
}

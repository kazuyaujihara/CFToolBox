using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace Ujihara.Chemistry.IO
{
    public class SDFReader : IDisposable
    {
        private const string GroupName_FieldName = "fieldname";
        Regex reSDFDataHeaderLTGT = new Regex(@"^\>.*\<(?<" + GroupName_FieldName + @">[A-Za-z]([A-Za-z0-9\.]*[A-Za-z0-9]|))\>.*$", RegexOptions.Compiled);
        Regex reSDFDataHeaderDTn = new Regex(@"^\>.*\A(?<" + GroupName_FieldName + @">DT\d+)\a.*$", RegexOptions.Compiled);

        public TextReader Input { get; private set; }
        public bool IsReplaceDotToBar { get; private set; }

        public SDFReader(TextReader reader)
            : this(reader, false)
        {
        }

        public SDFReader(TextReader reader, bool replaceDotToBar)
        {
            this.Input = reader;
            this.IsReplaceDotToBar = replaceDotToBar;
        }

        public IDictionary<string, string> Read()
        {
            var dic = new Dictionary<string, string>();

            // Reading MOL part
            {
                var sb = new StringBuilder();
                for (; ; )
                {
                    var line = Input.ReadLine();
                    if (line == null)
                    {
                        var restString = sb.ToString();
                        if (restString.Replace("\n", "") == "")
                            return null;
                        throw new Exception("Incorrect data.");
                    }
                    sb.Append(line).Append('\n');
                    if (line == "M  END")
                        break;
                }
                dic.Add("", sb.ToString());
            }

            for (; ; )
            {
                // reading Header

                string header = Input.ReadLine();
                if (header == null)
                    throw new Exception("$$$$ mark is missing.");
                if (header == "$$$$")
                    break;

                var ma = reSDFDataHeaderLTGT.Match(header);
                if (!ma.Success)
                    ma = reSDFDataHeaderDTn.Match(header);
                if (!ma.Success)
                    throw new Exception("Header format is not correct.");
                var fieldName = ma.Groups[GroupName_FieldName].Value;
                if (IsReplaceDotToBar)
                    fieldName = fieldName.Replace('.', '_');

                string data;
                // Reading data
                // Blanc line is terminator
                {
                    var sb = new StringBuilder();
                    var isFirstLine = true;
                    for (; ; )
                    {
                        var line = Input.ReadLine();
                        if (line == null)
                            throw new Exception("Blank line is missing.");
                        if (line == "")
                            break;
                        if (!isFirstLine)
                            sb.Append('\n');
                        sb.Append(line);
                    }
                    data = sb.ToString();
                }

                dic.Add(fieldName, data);
            }
            return dic;
        }

        public void Close()
        {
            Dispose();
        }

        private bool disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // no managed objects.
                }
                disposed = true;
            }
        }
    }
}

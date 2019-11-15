using System;
using System.Collections.Generic;
using System.IO;
using Ujihara.Chemistry.IO;

namespace Ujihara.Chemistry.MergeSF
{
    internal class SubstancesCSVExtractor
    {
        private const string Registry_Number = "Registry Number";
        private const string CA_Index_Name = "CA Index Name";
        private const string Class_Identifier = "Class Identifier";
        private const string Other_Names = "Other Names";
        private const string Formula = "Formula";
        private const string Copyright = "Copyright";

        public string FileName { get; private set; }
        private static object missing = Type.Missing;

        public SubstancesCSVExtractor(string filename)
        {
            this.FileName = filename;
        }

        /// <summary>
        /// Column number of 'CA Index Name'.
        /// </summary>
        private int ColumnNumberOf_Registry_Number = -1;
        private int ColumnNumberOf_CA_Index_Name = -1;
        private int ColumnNumberOf_Other_Names = -1;
        private int ColumnNumberOf_Formula = -1;
        private int ColumnNumberOf_Class_Identifier = -1;
        private int ColumnNumberOf_Copyright = -1;

        /// <summary>
        /// Substances Info. Empty if <paramref name="filename"/> is null.
        /// </summary>
        public IEnumerable<SubstanceInfo> GetSubstancesInfo()
        {
            string filename = Path.GetFullPath(this.FileName);
            using (var reader = new CsvReader(filename))
            {
                var line = reader.ReadRow();
                ColumnNumberOf_Registry_Number = line.IndexOf(Registry_Number);
                ColumnNumberOf_CA_Index_Name = line.IndexOf(CA_Index_Name);
                ColumnNumberOf_Other_Names = line.IndexOf(Other_Names);
                ColumnNumberOf_Formula = line.IndexOf(Formula);
                ColumnNumberOf_Class_Identifier = line.IndexOf(Class_Identifier);
                ColumnNumberOf_Copyright = line.IndexOf(Copyright);

                while ((line = reader.ReadRow()) != null)
                {
                    var info = new SubstanceInfo();
                    A(ref info._CASRN, ColumnNumberOf_Registry_Number, line);
                    A(ref info._CAIndexName, ColumnNumberOf_CA_Index_Name, line);
                    A(ref info._Name, ColumnNumberOf_Other_Names, line);
                    A(ref info._MolecularFormula, ColumnNumberOf_Formula, line);
                    A(ref info._ClassIdentifier, ColumnNumberOf_Class_Identifier, line);
                    A(ref info._Copyright, ColumnNumberOf_Copyright, line);
                    if (ColumnNumberOf_Other_Names != 0)
                    {
                        var value = line[ColumnNumberOf_Other_Names].Trim();
                        if (value != "")
                        {
                            info.OtherNames = new List<string>();
                            foreach (var name in value.Split(';'))
                            {
                                info.OtherNames.Add(name.Trim());
                            }
                        }
                    }
                    yield return info;
                }
            }
            yield break;
        }

        private static void A(ref string destination, int columnNr, List<string> line)
        {
            if (columnNr == 0)
                return;
            var value = line[columnNr];
            if (value != "")
                destination = value;
        }
    }
}

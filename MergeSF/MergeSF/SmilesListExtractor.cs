using System.Collections.Generic;
using System.IO;

namespace Ujihara.Chemistry.MergeSF
{
    public class SmilesListExtractor
    {
        public string FileName { get; private set; }

        public SmilesListExtractor(string filename)
        {
            this.FileName = filename;
        }

        public IEnumerable<SubstanceInfo> GetSubstancesInfo()
        {
            int nOderInDoc = 1;
            string filename = Path.GetFullPath(this.FileName);
            using (var reader = new StreamReader(filename, true))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line == null)
                        break;
                    line = line.Trim();
                    if (line == "")
                        continue;

                    var info = new SubstanceInfo();
                    info.Smiles = line;
                    info.Order = nOderInDoc++;
                    yield return info;
                }
                yield break;
            }
        }
    }
}

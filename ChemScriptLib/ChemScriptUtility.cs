#if ChemOfficeVersion16
using CambridgeSoft.ChemScript16;
#endif

namespace Ujihara.Chemistry
{
    public static class ChemScriptUtility
    {
        public static StructureData StructureDataFromName(string chemicalName)
        {
            if (chemicalName == null)
                return null;
            var modifiedName = AlphaToDotAlphaDot(chemicalName);
            
            var csmol = StructureData.LoadData(modifiedName, "name");
            return csmol;
        }

        private static readonly string[][] alphaToDotAlphaTable = new string [][] {
                new[] { "α", "alpha", },
                new[] { "β", "beta", },
                new[] { "γ", "gamma", },
                new[] { "δ", "delta", },
                new[] { "ε", "epsilon", },
                new[] { "ζ", "zeta", },
                new[] { "η", "eta", },
                new[] { "θ", "theta", },
                new[] { "ι", "iota", },
                new[] { "κ", "kappa", },
                new[] { "λ", "lambda", },
                new[] { "μ", "mu", },
                new[] { "ν", "nu", },
                new[] { "ξ", "xi", },
                new[] { "ο", "omicron", },
                new[] { "π", "pi", },
                new[] { "ρ", "rho", },
                new[] { "σ", "sigma", },
                new[] { "τ", "tau", },
                new[] { "υ", "upsilon", },
                new[] { "φ", "phi", },
                new[] { "χ", "chi", },
                new[] { "ψ", "psi", },
                new[] { "ω", "omega", },
            };

        private static readonly string[][] replaceTable = new string[][] {
                new[] { "?", " ", },     // EZ in IUPAC name is ? in CAS name. ChemScript handles nil as unknown
            };

        private static string AlphaToDotAlphaDot(string chemicalName)
        {
            if (chemicalName == null)
                return null;

            foreach (var l in alphaToDotAlphaTable)
            {
                chemicalName = chemicalName.Replace(l[0], "." + l[1] + ".");
            }

            foreach (var l in replaceTable)
            {
                chemicalName = chemicalName.Replace(l[0], l[1]);
            }
            return chemicalName;
        }
    }
}

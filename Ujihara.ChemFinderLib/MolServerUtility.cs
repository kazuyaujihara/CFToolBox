using System;
using System.IO;
using MolServer = MolServer16;
using CambridgeSoft.ChemScript16;
using Ujihara.Chemistry.IO;

namespace Ujihara.Chemistry
{
    public static class MolServerUtility
    {
        internal static MolServer.MSOpenModes ToMSOpenModes(FileAccess flag)
        {
            switch (flag)
            {
                case FileAccess.Read:
                    return MolServer.MSOpenModes.kMSReadOnly;
                case FileAccess.ReadWrite:
                case FileAccess.Write:
                    return MolServer.MSOpenModes.kMSNormal;
                default:
                    throw new ArgumentException("Argument value '" + flag.ToString() + "' is not valid.");
            }
        }

        public static StructureData ToStructureData(MolServer.Molecule mol)
        {
            if (mol == null)
                return null;
            using (var cdx = new TempFile(".cdx"))
            {
                mol.Write(cdx.Path, Type.Missing, Type.Missing);
                var csmol = StructureData.LoadFile(cdx.Path);
                return csmol;
            }
        }
    }
}

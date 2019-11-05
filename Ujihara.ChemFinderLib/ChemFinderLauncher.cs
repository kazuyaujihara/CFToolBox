using System;
using System.IO;
using ChemFinder = ChemFinder19;

namespace Ujihara.Chemistry
{
    public class ChemFinderLauncher
        : IDisposable
    {
        public string FullPath { get; private set; }
        public ChemFinder.Application Application { get; private set; }
        public ChemFinder.Documents Documents { get; private set; }
        public ChemFinder.Document Document { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">Path of ChemFinder file.</param>
        public ChemFinderLauncher(string path)
        {
            path = Path.GetFullPath(path);

            this.Application = new ChemFinder.Application();
            this.Documents = this.Application.Documents;
            this.Document = this.Documents.Open(path, Type.Missing);
            this.FullPath = path;
        }

        public void Import(string sdfFileName, ChemFinder.CFTargetAction targetAction, ChemFinder.CFDuplicateAction duplicateAction)
        {
            this.Document.Import(Path.GetFullPath(sdfFileName), this.FullPath, "", targetAction, duplicateAction);
        }

        public void Export(string exportFileName)
        {
            exportFileName = Path.GetFullPath(exportFileName); // ChemFinder does not handle current directory
            Utility.DeleteFile(exportFileName);
            this.Document.Export(exportFileName);
        }
        
        private bool disposed = false;

        public void Dispose()
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
                }

                if (this.Document != null)
                {
                    this.Document.Close(Type.Missing, Type.Missing);
                    Utility.ReleaseComObject(this.Document);
                }
                if (this.Documents != null)
                {
                    this.Documents.Close();
                    Utility.ReleaseComObject(this.Documents);
                }
                if (this.Application != null)
                {
                    this.Application.Quit();
                    Utility.ReleaseComObject(this.Application);
                }

                disposed = true;
            }
        }

        ~ChemFinderLauncher()
        {
            Dispose(false);
        }
    }
}

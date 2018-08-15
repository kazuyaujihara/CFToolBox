using System;

namespace Ujihara.Chemistry.IO
{
    public sealed class TempFile
        : IDisposable
    {
        public string Path { get; private set; }

        public TempFile(string extension)
        {
            this.Path = Utility.GetUniqueFileName(extension);
        }

        public void Dispose()
        {
            Utility.DeleteFile(Path);

            GC.SuppressFinalize(this);
        }

        ~TempFile()
        {
            Dispose();
        }
    }
}

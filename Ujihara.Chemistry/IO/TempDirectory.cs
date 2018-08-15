using System;
using System.IO;

namespace Ujihara.Chemistry.IO
{
    public sealed class TempDirectory
        : IDisposable
    {
        public DirectoryInfo Directory { get; private set; }

        public TempDirectory()
        {
            var directoryName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory = new DirectoryInfo(directoryName);
            Directory.Create();
        }

        public void Dispose()
        {
            Delete(Directory);
            GC.SuppressFinalize(this);
        }

        ~TempDirectory()
        {
            Dispose();
        }

        private static void Delete(DirectoryInfo dirInfo)
        {
            try
            {
                if (!System.IO.Directory.Exists(dirInfo.FullName))
                    return;
                foreach (var file in dirInfo.GetFiles())
                {
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception)
                    {
                    }
                }
                foreach (var dir in dirInfo.GetDirectories())
                {
                    Delete(dir);
                    try
                    {
                        dir.Delete();
                    }
                    catch (Exception)
                    {
                    }
                }
                try
                {
                    dirInfo.Delete();
                }
                catch (Exception)
                {
                }
            }
            catch (Exception)
            {
            }
        }
    }
}

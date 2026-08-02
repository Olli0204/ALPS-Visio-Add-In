using System;
using System.IO;

namespace ALPS_Visio_AddIn_rewrite.Tests
{
    internal static class TestResources
    {
        public static string Fixture(string fileName)
        {
            string path = Path.Combine(
                AppContext.BaseDirectory, "Fixtures", fileName);
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(
                    "The test fixture was not copied to the output directory.",
                    path);
            }

            return path;
        }

        public static string TemporaryFile(string extension)
        {
            return Path.Combine(
                Path.GetTempPath(),
                "alps-visio-tests-" + Guid.NewGuid().ToString("N")
                    + extension);
        }
    }

    internal sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "alps-visio-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, true);
        }
    }
}

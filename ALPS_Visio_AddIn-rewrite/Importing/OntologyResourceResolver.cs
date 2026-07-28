using System;
using System.IO;

namespace ALPS_Visio_AddIn_rewrite.Importing
{
    /// <summary>
    /// Resolves deployed ontology files and materializes embedded fallbacks.
    /// </summary>
    internal sealed class OntologyResourceResolver
    {
        private readonly string assemblyDirectory;
        private readonly string cacheDirectory;

        public OntologyResourceResolver(string assemblyDirectory, string cacheDirectory)
        {
            if (string.IsNullOrWhiteSpace(assemblyDirectory))
                throw new ArgumentException("An assembly directory is required.",
                    nameof(assemblyDirectory));
            if (string.IsNullOrWhiteSpace(cacheDirectory))
                throw new ArgumentException("An ontology cache directory is required.",
                    nameof(cacheDirectory));

            this.assemblyDirectory = assemblyDirectory;
            this.cacheDirectory = cacheDirectory;
        }

        public string Resolve(string fileName, byte[] embeddedContents)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("An ontology file name is required.",
                    nameof(fileName));

            string deployedPath =
                Path.Combine(assemblyDirectory, "Resources", fileName);
            if (File.Exists(deployedPath)) return deployedPath;

            if (embeddedContents == null || embeddedContents.Length == 0)
                throw new FileNotFoundException(
                    "The embedded ontology resource is unavailable.", fileName);

            Directory.CreateDirectory(cacheDirectory);
            string cachedPath = Path.Combine(cacheDirectory, fileName);
            if (NeedsRefresh(cachedPath, embeddedContents.Length))
                File.WriteAllBytes(cachedPath, embeddedContents);

            return cachedPath;
        }

        private static bool NeedsRefresh(string cachedPath, long expectedLength)
        {
            return !File.Exists(cachedPath)
                || new FileInfo(cachedPath).Length != expectedLength;
        }
    }
}

using ALPS_Visio_AddIn_rewrite.Importing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Text;

namespace ALPS_Visio_AddIn_rewrite.Tests.Importing
{
    [TestClass]
    public sealed class OntologyResourceResolverTests
    {
        [TestMethod]
        public void Resolve_DeployedResourceExists_ReturnsDeployedPath()
        {
            using (TemporaryDirectory root = new TemporaryDirectory())
            {
                string assemblyDirectory = Path.Combine(root.Path, "assembly");
                string resourcesDirectory = Path.Combine(
                    assemblyDirectory, "Resources");
                string cacheDirectory = Path.Combine(root.Path, "cache");
                Directory.CreateDirectory(resourcesDirectory);
                string deployedPath = Path.Combine(resourcesDirectory, "pass.owl");
                File.WriteAllText(deployedPath, "deployed");
                OntologyResourceResolver resolver =
                    new OntologyResourceResolver(
                        assemblyDirectory, cacheDirectory);

                string result = resolver.Resolve(
                    "pass.owl", Encoding.UTF8.GetBytes("embedded"));

                Assert.AreEqual(deployedPath, result);
                Assert.IsFalse(Directory.Exists(cacheDirectory));
            }
        }

        [TestMethod]
        public void Resolve_DeployedResourceMissing_WritesEmbeddedFallback()
        {
            using (TemporaryDirectory root = new TemporaryDirectory())
            {
                string assemblyDirectory = Path.Combine(root.Path, "assembly");
                string cacheDirectory = Path.Combine(root.Path, "cache");
                byte[] contents = Encoding.UTF8.GetBytes("embedded ontology");
                OntologyResourceResolver resolver =
                    new OntologyResourceResolver(
                        assemblyDirectory, cacheDirectory);

                string result = resolver.Resolve("pass.owl", contents);

                Assert.AreEqual(
                    Path.Combine(cacheDirectory, "pass.owl"), result);
                CollectionAssert.AreEqual(contents, File.ReadAllBytes(result));
            }
        }

        [TestMethod]
        public void Resolve_CacheLengthChanged_RefreshesCachedResource()
        {
            using (TemporaryDirectory root = new TemporaryDirectory())
            {
                string cacheDirectory = Path.Combine(root.Path, "cache");
                Directory.CreateDirectory(cacheDirectory);
                string cachedPath = Path.Combine(cacheDirectory, "pass.owl");
                File.WriteAllText(cachedPath, "old");
                byte[] expected = Encoding.UTF8.GetBytes("new ontology");
                OntologyResourceResolver resolver =
                    new OntologyResourceResolver(
                        Path.Combine(root.Path, "assembly"), cacheDirectory);

                string result = resolver.Resolve("pass.owl", expected);

                Assert.AreEqual(cachedPath, result);
                CollectionAssert.AreEqual(expected, File.ReadAllBytes(result));
            }
        }

        [TestMethod]
        public void Resolve_NoDeployedOrEmbeddedResource_ThrowsFileNotFound()
        {
            using (TemporaryDirectory root = new TemporaryDirectory())
            {
                OntologyResourceResolver resolver =
                    new OntologyResourceResolver(
                        Path.Combine(root.Path, "assembly"),
                        Path.Combine(root.Path, "cache"));

                Assert.ThrowsExactly<FileNotFoundException>(
                    () => resolver.Resolve("missing.owl", new byte[0]));
            }
        }
    }
}

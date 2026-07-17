using alps.net.api;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using ALPS_Visio_AddIn_rewrite.OWLShapes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;

namespace ALPS_Visio_AddIn_rewrite
{
    /// <summary>
    /// Importer for OWL ALPS files
    /// </summary>
    public class OWLImporter
    {
        /// <summary>
        /// Singleton OWL importer instance
        /// </summary>
        public static readonly OWLImporter Instance = new OWLImporter();

        private readonly IPASSReaderWriter parser = PASSReaderWriter.getInstance();

        private OWLImporter()
        {
            // enable reflection and set ModelElementFactory to assign parsed objects to Visio classes
            ReflectiveEnumerator.addAssemblyToCheckForTypes(Assembly.GetExecutingAssembly());
            parser.setModelElementFactory(new VisioClassFactory());

            string standardOntology = GetOntologyPath(
                "standard_PASS_ont_v_1.1.0.owl", Properties.Resources.standard_PASS_ont_v_1_1_0);
            string alpsOntology = GetOntologyPath(
                "ALPS_ont_v_0.8.0.owl", Properties.Resources.ALPS_ont_v_0_8_0);

            parser.loadOWLParsingStructure(new List<string>
            {
                standardOntology,
                alpsOntology
            });
        }

        private static string GetOntologyPath(string fileName, byte[] embeddedContents)
        {
            string assemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string deployedPath = Path.Combine(assemblyDirectory, "Resources", fileName);
            if (File.Exists(deployedPath)) return deployedPath;

            if (embeddedContents == null || embeddedContents.Length == 0)
                throw new FileNotFoundException("The embedded ontology resource is unavailable.", fileName);

            string cacheDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ALPS-Visio-Add-In", "Ontologies");
            Directory.CreateDirectory(cacheDirectory);

            string cachedPath = Path.Combine(cacheDirectory, fileName);
            if (!File.Exists(cachedPath) || new FileInfo(cachedPath).Length != embeddedContents.Length)
                File.WriteAllBytes(cachedPath, embeddedContents);

            return cachedPath;
        }

        /// <summary>
        /// Parse and import OWL file.
        /// </summary>
        public void Parse(string fileName)
        {
            try
            {
                IList<IPASSProcessModel> passProcessModels = parser.loadModels(new List<string> { fileName });

                VH.openStencil(VH.VisioStencils.SID_STENCIL);
                VH.setVBAListenersRunning(false);

                foreach (IPASSProcessModel processModel in passProcessModels)
                {
                    IVisioExportable exportable = processModel as IVisioExportable;
                    if (exportable != null) exportable.ExportToVisio(null);
                }
            }
            catch (System.Exception exception)
            {
                MessageBox.Show("The OWL model could not be imported.\n\n" + exception.Message,
                    "ALPS/PASS Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                VH.setVBAListenersRunning(true);
            }
        }
    }
}

using alps.net.api;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using ALPS_Visio_AddIn_rewrite.OWLShapes;
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

            string resourcesDirectory = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources");
            string standardOntology = Path.Combine(resourcesDirectory, "standard_PASS_ont_v_1.1.0.owl");
            string alpsOntology = Path.Combine(resourcesDirectory, "ALPS_ont_v_0.8.0.owl");

            if (!File.Exists(standardOntology) || !File.Exists(alpsOntology))
                throw new FileNotFoundException("The bundled ALPS ontology resources could not be found.", resourcesDirectory);

            parser.loadOWLParsingStructure(new List<string>
            {
                standardOntology,
                alpsOntology
            });
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

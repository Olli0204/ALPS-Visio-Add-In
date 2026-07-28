using alps.net.api;
using alps.net.api.parsing;
using ALPS_Visio_AddIn_rewrite.OWLShapes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ALPS_Visio_AddIn_rewrite.Importing
{
    /// <summary>
    /// Creates the production OWL import dependency graph.
    /// </summary>
    internal static class OwlImportComposition
    {
        public static OwlImportService CreateDefault()
        {
            IPASSReaderWriter parser = CreateConfiguredParser();
            return new OwlImportService(
                parser,
                VisioHelper.OpenImportStencils,
                VisioHelper.setVBAListenersRunning,
                Globals.ThisAddIn.RefreshModelFromDrawing);
        }

        private static IPASSReaderWriter CreateConfiguredParser()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            ReflectiveEnumerator.addAssemblyToCheckForTypes(assembly);

            IPASSReaderWriter parser = PASSReaderWriter.getInstance();
            parser.setModelElementFactory(new VisioClassFactory());

            string assemblyDirectory = Path.GetDirectoryName(assembly.Location);
            string cacheDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ALPS-Visio-Add-In",
                "Ontologies");
            OntologyResourceResolver resolver =
                new OntologyResourceResolver(assemblyDirectory, cacheDirectory);

            string standardOntology = resolver.Resolve(
                "standard_PASS_ont_v_1.1.0.owl",
                Properties.Resources.standard_PASS_ont_v_1_1_0);
            string alpsOntology = resolver.Resolve(
                "ALPS_ont_v_0.8.0.owl",
                Properties.Resources.ALPS_ont_v_0_8_0);

            parser.loadOWLParsingStructure(new List<string>
            {
                standardOntology,
                alpsOntology
            });

            return parser;
        }
    }
}

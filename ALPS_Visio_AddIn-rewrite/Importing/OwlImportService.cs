using alps.net.api;
using alps.net.api.StandardPASS;
using ALPS_Visio_AddIn_rewrite.OWLShapes;
using System;
using System.Collections.Generic;

namespace ALPS_Visio_AddIn_rewrite.Importing
{
    /// <summary>
    /// Coordinates parsing and Visio export without owning UI concerns.
    /// </summary>
    internal sealed class OwlImportService
    {
        private readonly IPASSReaderWriter parser;
        private readonly Action openStencils;
        private readonly Action<bool> setVbaListenersRunning;

        public OwlImportService(IPASSReaderWriter parser, Action openStencils,
            Action<bool> setVbaListenersRunning)
        {
            this.parser = parser ?? throw new ArgumentNullException(nameof(parser));
            this.openStencils = openStencils
                ?? throw new ArgumentNullException(nameof(openStencils));
            this.setVbaListenersRunning = setVbaListenersRunning
                ?? throw new ArgumentNullException(nameof(setVbaListenersRunning));
        }

        public int Import(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("An OWL file name is required.",
                    nameof(fileName));

            int exportedModelCount = 0;
            try
            {
                IList<IPASSProcessModel> processModels =
                    parser.loadModels(new List<string> { fileName });

                openStencils();
                setVbaListenersRunning(false);

                foreach (IPASSProcessModel processModel in processModels)
                {
                    IVisioExportable exportable = processModel as IVisioExportable;
                    if (exportable == null) continue;

                    exportable.ExportToVisio(null);
                    exportedModelCount++;
                }

                return exportedModelCount;
            }
            finally
            {
                setVbaListenersRunning(true);
            }
        }
    }
}

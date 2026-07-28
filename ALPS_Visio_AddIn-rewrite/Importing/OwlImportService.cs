using alps.net.api;
using alps.net.api.parsing;
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
        private readonly System.Action openStencils;
        private readonly System.Action refreshModel;

        public OwlImportService(IPASSReaderWriter parser, System.Action openStencils,
            System.Action refreshModel)
        {
            this.parser = parser ?? throw new ArgumentNullException(nameof(parser));
            this.openStencils = openStencils
                ?? throw new ArgumentNullException(nameof(openStencils));
            this.refreshModel = refreshModel
                ?? throw new ArgumentNullException(nameof(refreshModel));
        }

        public int Import(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("An OWL file name is required.",
                    nameof(fileName));

            int exportedModelCount = 0;
            IList<IPASSProcessModel> processModels =
                parser.loadModels(new List<string> { fileName });

            openStencils();

            foreach (IPASSProcessModel processModel in processModels)
            {
                IVisioExportable exportable = processModel as IVisioExportable;
                if (exportable == null) continue;

                exportable.ExportToVisio(null);
                exportedModelCount++;
            }

            refreshModel();
            return exportedModelCount;
        }
    }
}

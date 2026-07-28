using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Creates and configures ALPS/PASS SID and SBD pages.
    /// </summary>
    internal sealed class VisioPageFactory
    {
        private readonly Func<Visio.Application> applicationProvider;

        public VisioPageFactory(Func<Visio.Application> applicationProvider)
        {
            this.applicationProvider = applicationProvider
                ?? throw new ArgumentNullException(nameof(applicationProvider));
        }

        public Visio.Page CreateSidPage(string name, string universalName, string modelUri,
            string extends, string implements, string priority)
        {
            Visio.Application application = GetApplication();
            Visio.Document document = GetOrCreateDrawing(application);

            Visio.Page page = document.Pages.Add();
            string pageName = GetUniquePageName(document, name, "SID");
            string pageNameU = GetUniquePageName(document, universalName, pageName);
            page.Name = pageName;
            page.NameU = pageNameU;

            VisioShapeSheet.SetProperty(page.PageSheet, Constants.Properties.PageType,
                Constants.Properties.SIDPage);
            VisioShapeSheet.SetProperty(page.PageSheet, Constants.Properties.PageModelURI,
                modelUri);
            VisioShapeSheet.SetProperty(page.PageSheet, Constants.Properties.PageModelVersion,
                string.Empty);
            VisioShapeSheet.SetProperty(page.PageSheet, Constants.Properties.PageLayer,
                universalName);
            VisioShapeSheet.SetProperty(page.PageSheet,
                Constants.Properties.Transition.Extends, extends);
            VisioShapeSheet.SetProperty(page.PageSheet,
                Constants.Properties.Transition.Implements, implements);
            VisioShapeSheet.SetProperty(page.PageSheet,
                Constants.Properties.PriorityOrderNumber, priority);
            VisioShapeSheet.SetProperty(document.DocumentSheet,
                Constants.Properties.DocumentType, Constants.Properties.DocumentType);

            return page;
        }

        public Visio.Page CreateSbdPage(Visio.Page sidPage, string name,
            string universalName, Visio.Shape subjectShape)
        {
            Debug.Print("creating new SBD page");
            if (sidPage == null) throw new ArgumentNullException(nameof(sidPage));
            if (subjectShape == null) throw new ArgumentNullException(nameof(subjectShape));

            // Keep behavior pages in the same drawing even if opening a
            // stencil changed Visio's ActiveDocument in the meantime.
            Visio.Document document = sidPage.Document;
            if (document == null
                || document.Type != Visio.VisDocumentTypes.visTypeDrawing)
                throw new InvalidOperationException(
                    "The SID page does not belong to a Visio drawing.");

            Visio.Page page = document.Pages.Add();
            string pageName = GetUniquePageName(document, name, "SBD");
            string pageNameU = GetUniquePageName(document, universalName, pageName);
            page.Name = pageName;
            page.NameU = pageNameU;

            VisioShapeSheet.SetHyperlink(page.PageSheet, Constants.Properties.LinkedSIDPage,
                sidPage.NameU);
            VisioShapeSheet.SetProperty(page.PageSheet, Constants.Properties.PageLayer,
                sidPage.PageSheet.CellsU["Prop." + Constants.Properties.PageLayer].ResultStr[""]);
            VisioShapeSheet.SetProperty(page.PageSheet, Constants.Properties.PageType,
                Constants.Properties.SBDPage);
            VisioShapeSheet.SetProperty(page.PageSheet,
                Constants.Properties.SBDLinkedSubjectID,
                subjectShape.ID.ToString(CultureInfo.InvariantCulture));
            VisioShapeSheet.SetHyperlink(subjectShape, Constants.Properties.LinkedSBD,
                page.NameU);

            return page;
        }

        private Visio.Application GetApplication()
        {
            Visio.Application application = applicationProvider();
            if (application == null)
                throw new InvalidOperationException("The Visio application is unavailable.");

            return application;
        }

        private static Visio.Document GetOrCreateDrawing(Visio.Application application)
        {
            Visio.Document document = application.ActiveDocument;
            if (document != null
                && document.Type == Visio.VisDocumentTypes.visTypeDrawing)
                return document;

            return application.Documents.Add("");
        }

        private static string GetUniquePageName(Visio.IVDocument document,
            string requestedName, string fallback)
        {
            string baseName = string.IsNullOrWhiteSpace(requestedName)
                ? fallback
                : requestedName.Trim();
            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Visio.Page existingPage in document.Pages)
            {
                usedNames.Add(existingPage.Name);
                usedNames.Add(existingPage.NameU);
            }

            if (!usedNames.Contains(baseName)) return baseName;

            for (int suffix = 2; ; suffix++)
            {
                string candidate = baseName + " ("
                    + suffix.ToString(CultureInfo.InvariantCulture) + ")";
                if (!usedNames.Contains(candidate)) return candidate;
            }
        }
    }
}

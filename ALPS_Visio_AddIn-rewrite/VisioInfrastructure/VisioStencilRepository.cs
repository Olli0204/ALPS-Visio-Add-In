using System;
using System.Collections.Generic;
using Visio = Microsoft.Office.Interop.Visio;
using StencilKind = ALPS_Visio_AddIn_rewrite.VisioHelper.VisioStencils;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Resolves, opens, caches, and uses the ALPS/PASS Visio stencils.
    /// </summary>
    internal sealed class VisioStencilRepository
    {
        private readonly Func<Visio.Documents> documentsProvider;
        private readonly Action<string> errorPresenter;
        private readonly IDictionary<StencilKind, Visio.Document> openStencils =
            new Dictionary<StencilKind, Visio.Document>();

        private static readonly ISet<string> SidMasterNames = new HashSet<string>
        {
            Constants.SIDMasters.StandardActor,
            Constants.SIDMasters.InterfaceActor,
            Constants.SIDMasters.CommunicationRestriction,
            Constants.SIDMasters.StandardMessageConnector,
            Constants.SIDMasters.MessageBox,
            Constants.SIDMasters.Message,
            Constants.SIDMasters.StandAloneMacro,
            Constants.SIDMasters.ActorExtension,
            Constants.SIDMasters.AbstractCommunicationChannel,
            Constants.SIDMasters.SystemInterfaceSubject,
            Constants.SIDMasters.SubjectGroup
        };

        private static readonly ISet<string> SbdMasterNames = new HashSet<string>
        {
            Constants.SBDMasters.DoState,
            Constants.SBDMasters.ReceiveState,
            Constants.SBDMasters.SendState,
            Constants.SBDMasters.GenericReturnToOriginReference,
            Constants.SBDMasters.StandardTransition,
            Constants.SBDMasters.ReceiveTransition,
            Constants.SBDMasters.SendingFailedTransition,
            Constants.SBDMasters.SendTransition,
            Constants.SBDMasters.FlowRestrictor,
            Constants.SBDMasters.TimeTransition,
            Constants.SBDMasters.UserCancelTransition
        };

        public VisioStencilRepository(Func<Visio.Documents> documentsProvider,
            Action<string> errorPresenter)
        {
            this.documentsProvider = documentsProvider
                ?? throw new ArgumentNullException(nameof(documentsProvider));
            this.errorPresenter = errorPresenter
                ?? throw new ArgumentNullException(nameof(errorPresenter));
        }

        public Visio.Document Open(StencilKind stencil)
        {
            Visio.Documents documents = documentsProvider();
            if (documents == null)
                throw new InvalidOperationException("The Visio document collection is unavailable.");

            string stencilName = GetStencilName(stencil);
            Visio.Document openDocument = FindOpen(documents, stencil, stencilName);
            if (openDocument != null) return openDocument;

            try
            {
                short flags =
                    (short)Visio.VisOpenSaveArgs.visOpenDocked;
                Visio.Document document = documents.OpenEx(stencilName, (short)flags);
                openStencils[stencil] = document;
                return document;
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                errorPresenter(
                    "Failed to load ALPS/PASS shapes. Expecting file \""
                    + stencilName + "\" to exist in \"My Shapes\".\n"
                    + "Error: " + exception.Message);
                return null;
            }
        }

        public void OpenImportStencils()
        {
            Open(StencilKind.SID_STENCIL);
            Open(StencilKind.SBD_STENCIL);
        }

        public Visio.Shape Place(string masterName, Visio.Page page)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (string.IsNullOrWhiteSpace(masterName))
                throw new ArgumentException("A Visio master name is required.", nameof(masterName));

            Visio.Document stencil = Open(GetStencil(masterName));
            if (stencil == null)
                throw new InvalidOperationException("The required Visio stencil could not be opened.");

            Visio.Master master = stencil.Masters.get_ItemU(masterName);
            return page.Drop(master, 0, 0);
        }

        public StencilKind GetStencil(string masterName)
        {
            if (SidMasterNames.Contains(masterName)) return StencilKind.SID_STENCIL;
            if (SbdMasterNames.Contains(masterName)) return StencilKind.SBD_STENCIL;

            throw new ArgumentException("Unknown ALPS Visio master: " + masterName,
                nameof(masterName));
        }

        private static string GetStencilName(StencilKind stencil)
        {
            switch (stencil)
            {
                case StencilKind.SID_STENCIL:
                    return ShapeFinder.getSIDName();
                case StencilKind.SBD_STENCIL:
                    return ShapeFinder.getSBDName();
                default:
                    throw new ArgumentOutOfRangeException(nameof(stencil));
            }
        }

        private Visio.Document FindOpen(Visio.Documents documents, StencilKind stencil,
            string stencilName)
        {
            if (openStencils.TryGetValue(stencil, out Visio.Document cachedDocument))
            {
                try
                {
                    if (cachedDocument != null
                        && string.Equals(cachedDocument.Name, stencilName,
                            StringComparison.OrdinalIgnoreCase))
                        return cachedDocument;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // The user closed the cached stencil. Search the live
                    // Documents collection before opening it again.
                }

                openStencils.Remove(stencil);
            }

            foreach (Visio.Document document in documents)
            {
                try
                {
                    if (!string.Equals(document.Name, stencilName,
                        StringComparison.OrdinalIgnoreCase))
                        continue;

                    openStencils[stencil] = document;
                    return document;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Ignore a document that is currently being closed.
                }
            }

            return null;
        }
    }
}

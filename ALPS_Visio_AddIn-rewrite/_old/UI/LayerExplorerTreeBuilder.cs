using System.Collections.Generic;
using System.Linq;
using VisioAddIn.Snapping;

namespace VisioAddIn
{
    /// <summary>
    /// Builds the Model Explorer hierarchy and normalizes duplicate SID priorities.
    /// </summary>
    internal sealed class LayerExplorerTreeBuilder
    {
        private readonly ModelController modelController;

        public LayerExplorerTreeBuilder(ModelController modelController)
        {
            this.modelController = modelController;
        }

        public bool TryBuild(
            IDictionary<IVisioProcessModel, IDictionary<SIDPage, IList<SBDPage>>> models,
            out IList<DirectoryTreeViewItem> rootItems)
        {
            rootItems = new List<DirectoryTreeViewItem>();
            if (!EnsurePrioritiesAreUnique(models))
                return false;

            foreach (KeyValuePair<IVisioProcessModel,
                IDictionary<SIDPage, IList<SBDPage>>> modelEntry in models)
            {
                DirectoryTreeViewItem modelItem = new DirectoryTreeViewItem(null)
                {
                    Header = modelEntry.Key.getModelUri(),
                    Tag = modelEntry.Key
                };
                modelItem.DirectoryParent = modelItem;

                foreach (KeyValuePair<SIDPage, IList<SBDPage>> sidEntry
                    in modelEntry.Value.OrderBy(
                        entry => entry.Key.getPriorityOrder()))
                {
                    DirectoryTreeViewItem sidItem =
                        new DirectoryTreeViewItem(modelItem)
                        {
                            Header = sidEntry.Key.getLayer(),
                            Tag = sidEntry.Key,
                            DirectoryParent = modelItem
                        };

                    foreach (SBDPage sbdPage in sidEntry.Value)
                    {
                        sidItem.Items.Add(new DirectoryTreeViewItem(sidItem)
                        {
                            Header = sbdPage.getNameU(),
                            Tag = sbdPage,
                            DirectoryParent = sidItem
                        });
                    }

                    modelItem.Items.Add(sidItem);
                }

                rootItems.Add(modelItem);
            }

            return true;
        }

        private bool EnsurePrioritiesAreUnique(
            IDictionary<IVisioProcessModel, IDictionary<SIDPage, IList<SBDPage>>>
                models)
        {
            bool valid = true;
            foreach (KeyValuePair<IVisioProcessModel,
                IDictionary<SIDPage, IList<SBDPage>>> modelEntry in models)
            {
                LinkedList<int> usedNumbers = new LinkedList<int>();
                foreach (KeyValuePair<SIDPage, IList<SBDPage>> sidEntry
                    in modelEntry.Value.OrderBy(
                        entry => entry.Key.getPriorityOrder()))
                {
                    SIDPage page = sidEntry.Key;
                    int priority = page.getPriorityOrder();
                    if (usedNumbers.Contains(priority))
                    {
                        valid = false;
                        priority = usedNumbers.First.Value + 2;
                        modelController.updatePagePriority(priority, page);
                        usedNumbers.AddFirst(priority);
                    }
                    else if (usedNumbers.Count == 0
                        || priority > usedNumbers.First.Value)
                    {
                        usedNumbers.AddFirst(priority);
                    }
                    else
                    {
                        usedNumbers.AddLast(priority);
                    }
                }
            }

            return valid;
        }
    }
}

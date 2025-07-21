using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Taskbar;

namespace GitBranchTitlebar
{
    public static class JumpListHelper
    {
        // Serializable helper class representing one jumplist entry.
        private class RecentJumpListItem
        {
            public string Path { get; set; }
            public string BranchName { get; set; }
        }

        // Path to the persisted JSON file stored in AppData.
        private static readonly string DataFilePath;

        // Static constructor sets up the AppData folder and file.
        static JumpListHelper()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = System.IO.Path.Combine(appData, "JumpListHelper");

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            DataFilePath = System.IO.Path.Combine(folder, "recentjumplist.json");
        }

        // Loads the recent items from disk every time an update is requested.
        private static List<RecentJumpListItem> LoadRecentItems()
        {
            List<RecentJumpListItem> recentItems = new List<RecentJumpListItem>();
            if (File.Exists(DataFilePath))
            {
                try
                {
                    string json = File.ReadAllText(DataFilePath);
                    recentItems = Newtonsoft.Json.JsonConvert.DeserializeObject<List<RecentJumpListItem>>(json);
                }
                catch (Exception)
                {
                    // ignore
                }
            }

            return recentItems;
        }

        // Saves the list of recent items to the JSON file.
        private static void SaveRecentItems(List<RecentJumpListItem> recentItems)
        {
            try
            {
                string json = JsonConvert.SerializeObject(recentItems);
                File.WriteAllText(DataFilePath, json);
            }
            catch (Exception)
            {
            }
        }

        public static void UpdateSolutionJumpListEntry(string solutionPath, string branchName)
        {
            // Ensure the solution file exists.
            if (!File.Exists(solutionPath))
            {
                return;
            }

            var recentItems = LoadRecentItems();

            var newItem = new RecentJumpListItem
            {
                Path = solutionPath,
                BranchName = branchName
            };

            // Remove any existing entry for the same solution (ignoring case).
            var existingItem = recentItems.SingleOrDefault(
                (RecentJumpListItem x) => x.Path.Equals(newItem.Path, StringComparison.OrdinalIgnoreCase));
            if (existingItem != null)
            {
                recentItems.Remove(existingItem);
            }
            recentItems.Insert(0, newItem);

            SaveRecentItems(recentItems);

            // Group solutions by filename to detect multiple instances
            var solutionGroups = recentItems
                .GroupBy(item => System.IO.Path.GetFileNameWithoutExtension(item.Path), StringComparer.OrdinalIgnoreCase)
                .ToList();

            JumpList jumpList = JumpList.CreateJumpList();

            // Track which solutions have been added to custom categories
            var processedSolutions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Create custom categories for solutions with multiple instances and branches (limit to 3 categories)
            var customCategoriesCount = 0;
            const int maxCustomCategories = 2;
            const int maxItemsPerCategory = 5; // Limit items per category to 3 (make it very obvious)

            foreach (var solutionGroup in solutionGroups.OrderBy(g => recentItems.Where(item => 
                System.IO.Path.GetFileNameWithoutExtension(item.Path).Equals(g.Key, StringComparison.OrdinalIgnoreCase)).Min(item => recentItems.IndexOf(item))))
            {
                // Only create custom category if there are multiple instances AND at least one has a branch
                // And we haven't exceeded the maximum number of custom categories
                if (customCategoriesCount < maxCustomCategories && 
                    solutionGroup.Count() > 1 && 
                    solutionGroup.Any(item => !string.IsNullOrWhiteSpace(item.BranchName)))
                {
                    var categoryName = solutionGroup.Key; // Solution name without extension
                    var solutionCategory = new JumpListCustomCategory(categoryName);

                    var itemsAdded = 0;
                    // Limit to 3 most recent items in this category
                    foreach (var item in solutionGroup.OrderBy(x => recentItems.IndexOf(x)))
                    {
                        if (itemsAdded >= maxItemsPerCategory)
                            break;

                        var title = string.IsNullOrWhiteSpace(item.BranchName) 
                            ? System.IO.Path.GetDirectoryName(item.Path) // Show directory if no branch
                            : $"[{item.BranchName}]";
                        
                        var jumpListLink = new JumpListLink(item.Path, title)
                        {
                            IconReference = new IconReference(item.Path, 0)
                        };

                        solutionCategory.AddJumpListItems(jumpListLink);
                        processedSolutions.Add(item.Path);
                        itemsAdded++;
                    }

                    jumpList.AddCustomCategories(solutionCategory);
                    customCategoriesCount++;
                }
            }

            // Create "Recent" category for remaining solutions (not in custom categories)
            var regularSolutions = recentItems
                .Where(item => !processedSolutions.Contains(item.Path))
                .Take(6) // Limit to 6 items
                .ToList();

            if (regularSolutions.Any())
            {
                var recentCategory = new JumpListCustomCategory("Recent");

                foreach (var item in regularSolutions)
                {
                    var fileName = System.IO.Path.GetFileNameWithoutExtension(item.Path);
                    var title = string.IsNullOrWhiteSpace(item.BranchName) ? fileName : $"{fileName} [{item.BranchName}]";
                    var jumpListLink = new JumpListLink(item.Path, title)
                    {
                        IconReference = new IconReference(item.Path, 0)
                    };

                    recentCategory.AddJumpListItems(jumpListLink);
                }

                jumpList.AddCustomCategories(recentCategory);
            }

            jumpList.Refresh();
        }
    }
}

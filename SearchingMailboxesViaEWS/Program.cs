using System;
using Microsoft.Exchange.WebServices.Data;

namespace SearchingMailboxesViaEWS
{
    class Program
    {
        static void Main(string[] args)
        {
            var smtpAddressOfMailbox = args[0];
            var exchService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            exchService.AutodiscoverUrl(smtpAddressOfMailbox, foo => true);
            if (exchService.Url == null)
            {
                Console.WriteLine("Autodiscover failed");
                return;
            }

            var mailbox = new Mailbox(smtpAddressOfMailbox);
            var inboxFolderId = new FolderId(WellKnownFolderName.Inbox, mailbox);
            var inboxFolder = Folder.Bind(exchService, inboxFolderId);

            FindItemsBySortAndSeek(inboxFolder);
            FindItemsByRestrictedView(inboxFolder);
            FindItemsBySearchFolder(inboxFolder, exchService, smtpAddressOfMailbox);
        }

        public static void FindItemsBySortAndSeek(Folder folder)
        {
            Console.WriteLine("Finding items received today by sorted ranged retrieval.");
            var today = DateTime.Now.Date;
            var offset = 0;
            const int pageSize = 10;
            FindItemsResults<Item> findItemsResults = null;
            var moreInterestingItems = true;
            do
            {
                var view = new ItemView(pageSize, offset);
                view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
                findItemsResults = folder.FindItems(view);
                foreach (var result in findItemsResults.Items)
                {
                    if (result.DateTimeReceived >= today)
                    {
                        Console.WriteLine("Found item: " + result.Subject);
                    }
                    else
                    {
                        moreInterestingItems = false;
                        break;
                    }
                }

                offset += pageSize;
            } 
            while (findItemsResults.MoreAvailable && moreInterestingItems);
        }

        public static void FindItemsByRestrictedView(Folder folder)
        {
            Console.WriteLine("Finding items received today by restricted view.");
            var today = DateTime.Now.Date;
            var offset = 0;
            const int pageSize = 10;
            FindItemsResults<Item> findItemsResults = null;
            var filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, today);
            do
            {
                var view = new ItemView(pageSize, offset);
                findItemsResults = folder.FindItems(filter, view);
                foreach (var item in findItemsResults.Items)
                {
                    Console.WriteLine("Found item: " + item.Subject);
                }

                offset += pageSize;
            } 
            while (findItemsResults.MoreAvailable);
        }

        public static void FindItemsBySearchFolder(Folder folder, ExchangeService exchService, string smtpAddress)
        {
            Console.WriteLine("Finding items received today by search folder.");
            var today = DateTime.Now.Date;
            const string searchFolderName = "MyReceivedAfterSearchFolder";

            // Check if it exists already
            var searchFoldersId = new FolderId(WellKnownFolderName.SearchFolders, new Mailbox(smtpAddress));
            var searchFoldersFolder = Folder.Bind(exchService, WellKnownFolderName.SearchFolders);
            var searchFoldersView = new FolderView(2);
            var searchFoldersFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, searchFolderName);
            var searchFolderResults = searchFoldersFolder.FindFolders(searchFoldersFilter, searchFoldersView);
            SearchFolder mySearchFolder = null;
            if (searchFolderResults.Folders.Count > 1)
            {
                Console.WriteLine("The expected folder name is ambiguous. How did we end up with multiple folders?");
                Console.WriteLine("Dunno, but I'm going to delete all but the first.");
                for (var x = 1; x < searchFolderResults.Folders.Count; x++)
                {
                    searchFolderResults.Folders[x].Delete(DeleteMode.HardDelete);
                }
            }

            if (searchFolderResults.Folders.Count > 0)
            {
                Console.WriteLine("Found existing search folder.");
                mySearchFolder = searchFolderResults.Folders[0] as SearchFolder;
                if (mySearchFolder == null)
                {
                    Console.WriteLine("Somehow this folder isn't a search folder. Deleting it.");
                    searchFolderResults.Folders[0].Delete(DeleteMode.HardDelete);
                }
                else
                {
                    mySearchFolder.Load(new PropertySet(SearchFolderSchema.SearchParameters));
                    // Is the filter for today, or is this old?
                    var filter = mySearchFolder.SearchParameters.SearchFilter as SearchFilter.IsGreaterThanOrEqualTo;
                    if (filter == null)
                    {
                        Console.WriteLine("Somehow this filter isn't what we expected. Deleting the folder.");
                        mySearchFolder.Delete(DeleteMode.HardDelete);
                        mySearchFolder = null;
                    }
                    else if (DateTime.Parse((string)filter.Value) != today)
                    {
                        Console.WriteLine("This search folder is from a previous day. Updating the filter.");
                        mySearchFolder.SearchParameters.SearchFilter =
                            new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, today);
                        mySearchFolder.Save(searchFoldersId);
                    }
                }
            }
            
            if (mySearchFolder == null)
            {
                Console.WriteLine("Creating a new search folder for today.");
                mySearchFolder = new SearchFolder(exchService);
                mySearchFolder.DisplayName = searchFolderName;
                mySearchFolder.SearchParameters.SearchFilter = 
                    new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, today);
                mySearchFolder.SearchParameters.RootFolderIds.Add(folder.Id);
                mySearchFolder.SearchParameters.Traversal = SearchFolderTraversal.Shallow;
                mySearchFolder.Save(searchFoldersId);
            }

            // After all that, we can finally see if anything matches the search
            Console.WriteLine("Retrieving items from search folder.");
            FindItemsResults<Item> findItemsResults = null;
            var offset = 0;
            const int pageSize = 10;
            do
            {
                var view = new ItemView(pageSize, offset);
                findItemsResults = mySearchFolder.FindItems(view);
                foreach (var item in findItemsResults.Items)
                {
                    Console.WriteLine("Found item: " + item.Subject);
                }

                offset += pageSize;
            }
            while (findItemsResults.MoreAvailable);
        }
    }
}

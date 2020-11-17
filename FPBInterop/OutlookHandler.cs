using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using static olinteroplib.Methods;
using olinteroplib.ExtensionMethods;

namespace FPBInterop {
    internal static class OutlookHandler {

        /// PROPERTIES ///
        private static readonly TraceSource Tracer = new TraceSource("FPBInterop.OutlookHandling");
        private static string Desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private static string ExampleOrdersFolder =
            @"C:\Users\decroom\source\repos\OutlookInteropSolution\OutlookInterop\Example Orders";
        //Directory.GetParent(Directory.GetCurrentDirectory()).Parent.GetDirectories("Example Orders").First().FullName

        private static bool _exitHandlerEnabled = false;

        internal const string MagentoFilter = "[SenderEmailAddress]=\"secureorders@fergusonplarre.com.au\"";

        internal static Application OutlookApp;
        internal static Folder RootFolder;
        internal static Folder Inbox;
        internal static Folder DeletedItems;
        internal static Folder OnlineOrders;

        //TESTING SCENARIO 
        private static List<MailItem> _testItems = new List<MailItem>();
        private static bool _testSetup = false;
        private static Folder _testFolderParent;

        internal static Folder TestFolder { get; private set; }

        /// METHODS

        internal static void SetupAppRefs() {
            OutlookApp = new Application();
            RootFolder = OutlookApp.Session.DefaultStore.GetRootFolder() as Folder;
            Inbox = OutlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;
            DeletedItems = OutlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems) as Folder;

            if (!_exitHandlerEnabled) {
                ((ApplicationEvents_11_Event)OutlookApp).Quit += _OutlookHandling_Quit;
                _exitHandlerEnabled = true;
            }
        }

        private static void _OutlookHandling_Quit() {
            Tracer.TraceEvent(TraceEventType.Information,0,"Outlook instance closed, exiting...");
            StopTestEnv();
        }

        //MAIN SEQUENCE ORDER FILING
        internal static List<MailItem> GetFilteredMailList(Folder folder, string restrictFilter) {
            List<MailItem> filteredOrders = new List<MailItem>();
            try {
                Items matches = folder.Items.Restrict(restrictFilter);
                if (matches.Count == 0)
                    throw new ArgumentException();

                for (int i = 0; i < matches.Count; i++) {
                    filteredOrders.Add((MailItem)matches[i + 1]);
                }
            }
            catch (ArgumentException) { Tracer.TraceEvent(TraceEventType.Information, 0, $"No Magento items found in folder {folder.Name}"); }
            catch (System.Exception) { Tracer.TraceEvent(TraceEventType.Error, 0, "Getting Magento order list failed"); }

            return filteredOrders;
        }

        internal static void ReformatMagentoDates(Folder folder) {
            if (folder == null) {
                Tracer.TraceEvent(TraceEventType.Error,0,"Invalid folder or path");
                return;
            }

            Tracer.TraceEvent(TraceEventType.Verbose,0,$"Begin formatting dates in target folder {folder.Name}");

            ProgressBar pbar = new ProgressBar();
            pbar.Report(0);

            List<MailItem> magentoOrdersUnformatted = GetFilteredMailList(folder, MagentoFilter);
            int totalOrders = magentoOrdersUnformatted.Count;
            magentoOrdersUnformatted = magentoOrdersUnformatted.Where(
                    (item, i) => {
                        pbar.Report((((double)i / (double)totalOrders)) / 2);
                        return item.UserProperties.Find(UserPropertyNames.DATE_FORMATTED) == null;
                    }).ToList();
            for (int i = 0; i < magentoOrdersUnformatted.Count; i++) {
                _ReformatDate(magentoOrdersUnformatted[i]);
                UserProperty dateFormatted =
                    magentoOrdersUnformatted[i].UserProperties.Add(
                        UserPropertyNames.DATE_FORMATTED, OlUserPropertyType.olText, true);
                DisableVisiblePrintUserProp(dateFormatted);
                dateFormatted.Value = "Date Formatted";
                magentoOrdersUnformatted[i].Save();
                pbar.Report((double)0.5 + ((double)i / (double)magentoOrdersUnformatted.Count) / (double)2);
            }
            pbar.Dispose();
            Tracer.TraceEvent(TraceEventType.Verbose, 0, "Formatting dates complete");
        }
        private static bool _ReformatDate(MailItem item) {
            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime newDate = DateTime.MinValue;
            string dateMatch;
            Regex regex = new Regex(@"\d\d\/\d\d\/\d\d\d\d");

            if (regex.IsMatch(item.HTMLBody)) {
                dateMatch = regex.Match(item.HTMLBody).Value;
                newDate = DateTime.ParseExact(dateMatch, "dd/MM/yyyy", provider);
            }
            else {
                regex = new Regex(@"((\w){3,6}day), 0\d ((Jan|Febr)uary|Ma(rch|y)|A(pril|ugust)|Ju(ne|ly)|((Sept|Nov|Dec)em|Octo)ber) (\d){4}");
                if (regex.IsMatch(item.HTMLBody)) {
                    dateMatch = regex.Match(item.HTMLBody).Value;
                    newDate = DateTime.ParseExact(dateMatch, "dddd, dd MMMM yyyy", provider);
                }
                else return false;
            }

            string newDateString = newDate.ToString("dddd, d MMMM yyyy");
            item.HTMLBody = item.HTMLBody.Replace(dateMatch, newDateString);
            item.Save();
            return true;
        }

        internal static void RescueMisfiledOrders() {
            DateTime dt = new DateTime(2020, 10, 12);
            Console.WriteLine("Getting orders");
            Items orders = DeletedItems.Items.Restrict(MagentoFilter);
            orders = orders.Restrict("[ReceivedTime]>'" + dt.ToString("g")+"'");

            Console.WriteLine("start process");
            for (int i = orders.Count; i >= 1; i --) {

                MailItem temp = (MailItem)orders[i];
                string subj = temp.Subject;
                if (!subj.StartsWith("Ferguson Plarre: New Order"))
                    continue;
                string orderNum = subj.Remove(0, 27);
                Trace.Write(orderNum);
                bool toBeReturned = HtmlHandler.Magento.ParseOrderSpecial(temp.HTMLBody);
                if (toBeReturned) {
                    Trace.WriteLine(" REDO");
                }
            }
        }
        internal static void ProcessFolder(Folder folder, bool ignoreProcessed = false) {
            ProcessItems(folder.Items, ignoreProcessed);
        }
        internal static void ProcessItems(Items items, bool ignoreProcessed) {
            ProgressBar pbar = new ProgressBar();
            pbar.Report(0);
            int totalItems = items.Count;
            for (int i = totalItems; i > 0; i--) {
                ProcessItem((MailItem)items[i]);
                pbar.Report((double)(totalItems - i) / (double)(totalItems + 1));
            }
            pbar.Dispose();
            Console.WriteLine("Complete");
        }
        internal static void ProcessItem(MailItem item) {
            if (item.SenderEmailAddress == "secureorders@fergusonplarre.com.au") {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, item.Subject.Remove(0,27));
                _ReformatDate(item);
                MagentoOrder order = HtmlHandler.Magento.MagentoBuilder(item.HTMLBody);
                UserProperty parsed = item.UserProperties.Add(UserPropertyNames.PARSED, OlUserPropertyType.olYesNo, true);
                parsed.Value = true;
                DisableVisiblePrintUserProp(parsed);
                if (order.Meta.HasFlag(OrderMetadata.DoNotProcess)) {
                    item.UnRead = false;
                    item.Move(DeletedItems);
                }
                else {
                    item.UnRead = true;
                }
            }
        }

        internal static void ProcessSelectedItem() {
            ProcessItem((MailItem)OutlookApp.ActiveExplorer().Selection[1]);
        }

        //FOLDER FINDING/HANDLING
        internal static Folder GetFolderByPath(string path) {
            string slashType = path.Contains("/") ? "/" : (path.Contains(@"\") ? @"\" : null);
            if (slashType == null) {
                try {
                    Folder target = RootFolder.GetFolder(path);
                    return target;
                }
                catch {
                    Tracer.TraceEvent(TraceEventType.Verbose, 0, "Invalid path");
                    return null;
                }
            }
            char slashChar = slashType.First();
            Folder root = OutlookApp.Session.DefaultStore.GetRootFolder() as Folder;

            if (path.StartsWith(slashType) | path.EndsWith(slashType))
                path = path.Trim(slashChar);

            string[] folders = path.Split(slashChar);

            Folder folder = root;
            try {
                if (folder != null) {
                    for (int i = 0; i <= folders.GetUpperBound(0); i++) {
                        Folders subFolders = folder.Folders;
                        folder = subFolders.GetFolder(folders[i]);
                        if (folder == null) {
                            Tracer.TraceEvent(TraceEventType.Information, 0, 
                                $"Folder not found at path {path}");
                            return null;
                        }
                    }
                }
            }
            catch {
                Tracer.TraceEvent(TraceEventType.Information, 0, 
                    $"Folder not found at path {path}");
                return null;
            }
            return folder;
        }

        //TESTING SCENARIO
        internal static void StopTestEnv() {
            Tracer.TraceEvent(TraceEventType.Verbose, 0, "Begin test scenario cleanup");
            try {
                DeletedItems.Folders[TestFolder].Delete();
            }
            catch { }

            if (TestFolder == null)
                Tracer.TraceEvent(TraceEventType.Verbose, 0, "TestFolder is null, teststartup was not run");

            try {
                TestFolder = (Folder)_testFolderParent.Folders[TestFolder.Name];
            }
            catch {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, "Cannot set test folder as existing test folder");
            }
            if (TestFolder.Items.Count != _testItems.Count)
                foreach (MailItem item in TestFolder.Items) {
                    if (_testItems.Contains(item))
                        _testItems.Remove(item);
                }

            Tracer.TraceEvent(TraceEventType.Verbose, 0, "Cleaning testing scenario");
            while (_testItems.Count > 0) {
                _testItems[0].Delete();
                _testItems.Remove(_testItems[0]);
                Tracer.TraceEvent(TraceEventType.Verbose, 0, "Items remaining " + _testItems.Count);
            }

            Tracer.TraceEvent(TraceEventType.Verbose, 0, "Deleting test folder");
            TestFolder.MoveTo(DeletedItems);
            while (DeletedItems.Folders.Count == 0) { Thread.Sleep(100); }
            try {
                TestFolder.Delete();
                Tracer.TraceEvent(TraceEventType.Verbose, 0, " - complete");
                TestFolder = null;
            }
            catch { Tracer.TraceEvent(TraceEventType.Verbose, 0, " - failed!"); }

            _testSetup = false;
        }
        internal static bool SetupTestEnv(Folder itemSourceFolder, Folder parentFolder, string testFolderName, int maxItems, string sourceItemFilter = null) {
            if (_testSetup)
                return true;

            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"Begin setup test folder and populate with {maxItems} item(s)");
            if (sourceItemFilter != null)
                Tracer.TraceEvent(TraceEventType.Verbose, 0, $"\t - and apply filter {sourceItemFilter}");

            if (itemSourceFolder.Items.Count == 0) {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, "No valid items in source folder. Try another folder.");
                return false;
            }

            Items FilteredItems = itemSourceFolder.Items;
            if (sourceItemFilter != null)
                try {
                    FilteredItems = itemSourceFolder.Items.Restrict(sourceItemFilter);
                    Tracer.TraceEvent(TraceEventType.Verbose, 0, "Filter successful");
                }
                catch (System.Exception e) {
                    Tracer.TraceEvent(TraceEventType.Verbose, 0, "Failed to filter items, " + e.Message);
                    return false;
                };

            _testFolderParent = parentFolder;
            TestFolder = parentFolder.GetFolder(testFolderName);
            bool alreadyExists = false;

            if (TestFolder != null)
                alreadyExists = true;
            else {
                TestFolder = (Folder)parentFolder.Folders.Add(testFolderName);
                while (TestFolder == null) { Thread.Sleep(200); };
            }
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"Folder already exists: {alreadyExists}");

            _testItems = new List<MailItem>();
            List<MailItem> itemsToBeDuplicated = new List<MailItem>();
            maxItems = (maxItems > FilteredItems.Count) ? FilteredItems.Count : maxItems;
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"maxItems set to {maxItems}");

            if (alreadyExists) {
                foreach (MailItem item in TestFolder.Items) {
                    _testItems.Add(item);
                }
                if (_testItems.Count > maxItems)
                    return true;
            }

            TestFolder.ShowItemCount = OlShowItemCount.olShowTotalItemCount;

            int i = 1;
            while (itemsToBeDuplicated.Count < maxItems) {
                itemsToBeDuplicated.Add((MailItem)FilteredItems[i]);
                i++;
            }

            try {
                foreach (MailItem item in itemsToBeDuplicated) {
                    MailItem copy = (MailItem)item.Copy();
                    copy.UnRead = false;
                    _testItems.Add(copy);
                    copy.Move(TestFolder);
                }
                Tracer.TraceEvent(TraceEventType.Verbose, 0, "\t Complete");
                _testSetup = true;
                return true;
            }
            catch {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, "\t Failed");
                return false;
            }
        }

        private static void _WipeCategories(Folder folder) {
            foreach (MailItem item in folder.Items) {
                item.RemoveAllCategories();
            }
        }
        private static void _WipeCategories(Items items) {
            foreach (MailItem item in items) {
                item.RemoveAllCategories();
            }
        }

        internal static void SaveSelected(string filename) {
            SaveHTML((MailItem)OutlookApp.ActiveExplorer().Selection[1], ExampleOrdersFolder, filename);
        }
        internal static void ExampleWufoo() {
            Items wufoo = DeletedItems.Items.Restrict("no-reply@wufoo.com");
            foreach (MailItem item in wufoo) {
                if (item.Subject.Contains("Decorated Cake Order")) {
                    SaveHTML(item, Desktop, "example.html");
                    return;
                }
            }
        }
        internal static void SaveHTML(MailItem item, string filepath, string filename = null) {
            if (filename == null)
                filename = item.Subject;
            if (!filename.Contains(".html"))
                filename += ".html";
            File.WriteAllText($"{filepath}\\{filename}", item.HTMLBody);
        }

        // MISCELLANEOUS METHODS

        private static DateTime _GetFirstSundayAfterDate(DateTime date) {
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"Getting first Sunday after date {date:dd/MM} - ");
            DateTime sunday = date;
            while (sunday.DayOfWeek != 0) {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, sunday.ToString());
                sunday = sunday.AddDays(1);
            }
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $" - complete");
            return sunday;//.ToString("MMM dd").ToUpper(); ;
        }

        internal static void SetupCategories() {
            throw new NotImplementedException();
        }
        internal static void ClearCategories() {
            throw new NotImplementedException();
        }
        internal static void SetupUserProps() {

        }
    }

    struct UserPropertyNames {
        internal const string DATE_FORMATTED = "Date Formatted";
        internal const string PARSED = "Parsed";
    }
}

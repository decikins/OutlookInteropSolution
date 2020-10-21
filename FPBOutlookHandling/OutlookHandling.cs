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
        class OutlookHandling {

        /// PROPERTIES ///

        private static string Desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private static string ExampleOrdersFolder =
            @"C:\Users\decroom\source\repos\OutlookInteropSolution\OutlookInterop\Example Orders";
        //Directory.GetParent(Directory.GetCurrentDirectory()).Parent.GetDirectories("Example Orders").First().FullName

        private static bool _exitHandlerEnabled = false;

        public static Application OutlookApp;
        public static Folder RootFolder;
        public static Folder Inbox;
        public static Folder DeletedItems;
        public static Folder OnlineOrders;

        //TESTING SCENARIO 
        private const string _testFolderName = "Test";
        private static List<MailItem> _testItems = new List<MailItem>();
        private static bool _testSetup = false;
        private static Folder _testFolderParent;

        public static Folder TestFolder { get; private set; }

        //FOLDER ENUMERATION
        private static bool _foldersEnumerated = false;

        public static List<Folder> AllFoldersList = new List<Folder>();

        //FILTERS FOR ITEMS RESTRICTION
        private const string _MagentoSenderName = "[SenderName]=\"Ferguson Plarre Bakehouses\"";
        private const string _WufooSenderEmail = "[SenderEmailAddress] =\"no-reply@wufoo.com\"";

        /// METHODS

        public static void SetupOutlookRefs() {
            OutlookApp = new Application();
            RootFolder = OutlookApp.Session.DefaultStore.GetRootFolder() as Folder;
            Inbox = OutlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;
            DeletedItems = OutlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems) as Folder;

            if (!_exitHandlerEnabled) {
                ((ApplicationEvents_11_Event)OutlookApp).Quit += _OutlookHandling_Quit;
                _exitHandlerEnabled = true;
            }
        }
        private static void CheckUserDefinedProperties() {

        }
        private static void _OutlookHandling_Quit() {
            Trace.Write("Outlook instance closed, exiting...");
            StopTestEnv();
        }

        public static void EnumerateFolders(bool includeHiddenFolders = false) {
            olinteroplib.Methods.EnumerateFolders(AllFoldersList, RootFolder, includeHiddenFolders);
            _foldersEnumerated = true;
        }

        //MAIN SEQUENCE ORDER FILING
        public static List<MailItem> GetMagentoOrders(string folder) {
            return GetMagentoOrders(_GetFolderSwitch(folder));
        }
        public static List<MailItem> GetMagentoOrders(Folder folder) {
            List<MailItem> MagentoOrders = new List<MailItem>();
            try {
                Items matches = folder.Items.Restrict(_MagentoSenderName);
                if (matches.Count == 0)
                    throw new System.ArgumentException();

                for (int i = 0; i < matches.Count; i++) {
                    MailItem currentItem = (MailItem)matches[i + 1];
                    if (currentItem.Subject.Contains("Ferguson Plarre: New Order"))
                        MagentoOrders.Add(currentItem);
                }
            }
            catch (ArgumentException) { Trace.WriteLine($"No Magento items found in folder {folder.Name}"); }
            catch (System.Exception) { Trace.WriteLine("Getting Magento order list failed"); }

            return MagentoOrders;
        }
        public static List<MailItem> GetWufooOrders(Folder folder) {
            List<MailItem> WufooOrders = new List<MailItem>();
            try {
                Items matches = folder.Items.Restrict(_WufooSenderEmail);
                if (matches.Count == 0)
                    throw new ArgumentException();

                for (int i = 0; i < matches.Count; i++) {
                    MailItem currentItem = (MailItem)matches[i + 1];
                    if (currentItem.Subject.Contains("Contact [#"))
                        continue;
                    WufooOrders.Add(currentItem);
                }
            }
            catch (ArgumentException) { Trace.WriteLine($"No Magento items found in folder {folder.Name}"); }
            catch (System.Exception) { Trace.WriteLine("Getting Magento order list failed"); }

            return WufooOrders;
        }

        public static void ReformatMagentoDates(string folder) {
            ReformatMagentoDates(_GetFolderSwitch(folder));
        }
        public static void ReformatMagentoDates(Folder folder) {
            if (folder == null) {
                Console.WriteLine("Invalid folder or path");
                Trace.WriteLine("Invalid folder or path");
                return;
            }

            Trace.WriteLine($"Begin formatting dates in target folder {folder.Name}");

            ProgressBar pbar = new ProgressBar();
            pbar.Report(0);

            List<MailItem> magentoOrdersUnformatted = GetMagentoOrders(folder);
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
            Console.WriteLine("Formatting dates complete");
            Trace.WriteLine("Formatting dates complete");
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

        public static void ParseSelectedOrder() {
            _ParseOrder((MailItem)OutlookApp.ActiveExplorer().Selection[1]);
        }
        public static void ParseOrdersInFolder(string folderPath) {
            ParseOrdersInCollection(_GetFolderSwitch(folderPath).Items);
        }
        public static void ParseOrdersInFolder(Folder folder) {
            ParseOrdersInCollection(folder.Items);
        }
        public static void ParseOrdersInCollection(Items items) {
            ProgressBar pbar = new ProgressBar();
            pbar.Report(0);
            int totalItems = items.Count;
            for (int i = totalItems; i > 1; i--) {
                _ParseOrder((MailItem)items[i]);
                pbar.Report((double)(totalItems - i) / (double)(totalItems + 1));
            }
            pbar.Dispose();
            Console.WriteLine("Complete");
        }
        private static void _ParseOrder(MailItem item) {
            if (item.SenderEmailAddress == OrderTypeInfo.WufooSenderEmail) {

            }
            if (item.SenderEmailAddress == OrderTypeInfo.MagentoSenderEmail) {
                _ReformatDate(item);
                bool toBeProcessed = HTMLHandling.Magento.ParseOrder(item.HTMLBody);
                if (!toBeProcessed) {
                    item.UnRead = false;
                    UserProperty parsed = item.UserProperties.Add(UserPropertyNames.PARSED, OlUserPropertyType.olText, true);
                    DisableVisiblePrintUserProp(parsed);
                    item.Move(DeletedItems);
                }
            }
        }

        //FOLDER FINDING/HANDLING
        private static Folder _GetFolderSwitch(string input) {
            if (_foldersEnumerated && !(input.Contains("/") || input.Contains("\\")))
                return _GetFolderFromMaster(input);
            else
                return _GetFolderByPath(input);
        }
        private static Folder _GetFolderByPath(string path) {
            string slashType = path.Contains("/") ? "/" : (path.Contains(@"\") ? @"\" : null);
            if (slashType == null) {
                try {

                    Folder target = RootFolder.GetFolder(path);
                    return target;
                }
                catch {
                    Trace.WriteLine("Invalid path");
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
                            Trace.WriteLine("Folder not found at path");
                            return null;
                        }
                    }
                }
            }
            catch {
                Trace.WriteLine("Folder not found at path");
                return null;
            }
            return folder;
        }
        private static Folder _GetFolderFromMaster(string folder) {
            List<Folder> matches;
            string pathChar = folder.Contains("/") ? "/" : (folder.Contains(@"\") ? @"\" : null);
            if (pathChar != null) {
                string[] folderChain = folder.Split(pathChar.First());
                matches = AllFoldersList.Where(
                    n => n.Name.ToLower() == folderChain[1].ToLower() &
                    ((Folder)n.Parent).Name.ToLower() == folderChain[0].ToLower()).ToList();
            }
            else {
                matches = AllFoldersList.Where(n => n.Name == folder).ToList();
            }

            if (matches.Count == 0) {
                return null;
            }
            else if (matches.Count == 1) {
                return matches[0];
            }
            else {
                return _GetFolderFromDuplicates(matches);
            }
        }

        private static Folder _GetFolderFromDuplicates(List<Folder> matches) {
            int result = _ConsoleUserDecision("More than one match found by that name.\n" +
                                   "Please select the correct option number or press Escape to return", matches.Select(f => $"{f.Parent().Name}\\{f.Name}").ToArray());
            if (result == -1) {
                return null;
            }
            else {
                try {
                    Trace.WriteLine($"Target folder set as option " +
                        $"{result + 1}, {matches[result].Parent().Name}\\{matches[result].Name}");
                    return matches[result];
                }
                catch (System.Exception e) {
                    Trace.WriteLine($"Failed!\n{e.Message}");
                    return null;
                }
            }
        }
        private static int _ConsoleUserDecision(string message, string[] options) {
            Trace.WriteLine(message);
            int i = 1;
            foreach (string option in options) {
                Console.WriteLine($"[{i}]. {option}");
                i++;
            }

            ConsoleKeyInfo keyPressed;
            bool isInt;
            do {
                keyPressed = Console.ReadKey(true);
                if (keyPressed.Key == ConsoleKey.Escape) {
                    return -1;
                }

                isInt = int.TryParse(keyPressed.KeyChar.ToString(), out int keyChar);

                if (isInt) {
                    if (keyChar > 0 && keyChar <= options.Count()) {
                        Trace.WriteLine($"Option {keyChar}. {options[keyChar - 1]} selected");
                        return keyChar - 1;
                    }
                    else {
                        Console.WriteLine("Not a valid option number, try again");
                    }
                }
                else {
                    Console.WriteLine("Please select an option number with the number keys, or press escape to return");
                }
            } while (true);
        }

        //TESTING SCENARIO
        public static bool SetupDefaultTestEnv(int maxItems, string sourceItemFilter = null) {
            return _SetupTestEnv(DeletedItems, RootFolder, _testFolderName, maxItems, sourceItemFilter);
        }
        public static void StopTestEnv() {
            Trace.WriteLine("Begin test scenario cleanup");
            try {
                DeletedItems.Folders[TestFolder].Delete();
            }
            catch { }

            if (TestFolder == null)
                Trace.WriteLine("TestFolder is null, teststartup was not run");

            try {
                TestFolder = (Folder)_testFolderParent.Folders[TestFolder.Name];
            }
            catch {
                Trace.WriteLine("Cannot set test folder as existing test folder");
            }
            if (TestFolder.Items.Count != _testItems.Count)
                foreach (MailItem item in TestFolder.Items) {
                    if (_testItems.Contains(item))
                        _testItems.Remove(item);
                }

            Trace.WriteLine("Cleaning testing scenario");
            while (_testItems.Count > 0) {
                _testItems[0].Delete();
                _testItems.Remove(_testItems[0]);
                Trace.WriteLine("Items remaining " + _testItems.Count);
            }

            Trace.Write("Deleting test folder");
            TestFolder.MoveTo(DeletedItems);
            while (DeletedItems.Folders.Count == 0) { Thread.Sleep(100); }
            try {
                TestFolder.Delete();
                Trace.WriteLine(" - complete");
                TestFolder = null;
            }
            catch { Trace.WriteLine(" - failed!"); }

            _testSetup = false;
        }
        private static bool _SetupTestEnv(Folder itemSourceFolder, Folder parentFolder, string testFolderName, int maxItems, string sourceItemFilter = null) {
            if (_testSetup)
                return true;

            Trace.WriteLine($"Begin setup test folder and populate with {maxItems} item(s)");
            if (sourceItemFilter != null)
                Trace.WriteLine($"\t - and apply filter {sourceItemFilter}");

            if (itemSourceFolder.Items.Count == 0) {
                Trace.WriteLine("No valid items in source folder. Try another folder.");
                return false;
            }

            Items FilteredItems = itemSourceFolder.Items;
            if (sourceItemFilter != null)
                try {
                    FilteredItems = itemSourceFolder.Items.Restrict(sourceItemFilter);
                    Trace.WriteLine("Filter successful");
                }
                catch (System.Exception e) {
                    Trace.WriteLine("Failed to filter items, " + e.Message);
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
            Trace.WriteLine($"Folder already exists: {alreadyExists}");

            _testItems = new List<MailItem>();
            List<MailItem> itemsToBeDuplicated = new List<MailItem>();
            maxItems = (maxItems > FilteredItems.Count) ? FilteredItems.Count : maxItems;
            Trace.WriteLine($"maxItems set to {maxItems}");

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
                Trace.WriteLine("\t Complete");
                _testSetup = true;
                return true;
            }
            catch {
                Trace.WriteLine("\t Failed");
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

        public static void SaveSelected(string filename) {
            SaveHTML((MailItem)OutlookApp.ActiveExplorer().Selection[1], ExampleOrdersFolder, filename);
        }
        public static void ExampleWufoo() {
            Items wufoo = DeletedItems.Items.Restrict(_WufooSenderEmail);
            foreach (MailItem item in wufoo) {
                if (item.Subject.Contains("Decorated Cake Order")) {
                    SaveHTML(item, Desktop, "example.html");
                    return;
                }
            }
        }
        public static void SaveHTML(MailItem item, string filepath, string filename = null) {
            if (filename == null)
                filename = item.Subject;
            if (!filename.Contains(".html"))
                filename += ".html";
            File.WriteAllText($"{filepath}\\{filename}", item.HTMLBody);
        }

        // MISCELLANEOUS METHODS

        private static DateTime _GetFirstSundayAfterDate(DateTime date) {
            Trace.Write($"Getting first Sunday after date {date:dd/MM} - ");
            DateTime sunday = date;
            while (sunday.DayOfWeek != 0) {
                Trace.Write(sunday.ToString());
                sunday = sunday.AddDays(1);
            }
            Trace.WriteLine($" - complete");
            return sunday;//.ToString("MMM dd").ToUpper(); ;
        }


        internal static void SetupCategories() {
            Trace.Write("Begin setup categories");
            FieldInfo[] categoryList = typeof(OrderTypeInfo).GetFields();
            foreach (FieldInfo prop in categoryList) {

                //string categoryName = prop.GetValue("Tag");
                Trace.WriteLine(prop.ToString());
                /*try
                {
                    Category category = OutlookApp.Session.Categories[prop.GetValue(null)];
                    if (category == null)
                    {
                        OutlookApp.Session.Categories.Add(categoryName);
                        Trace.WriteLine($"Adding category {categoryName}");
                    }
                    else
                        Trace.WriteLine($"Category {categoryName} already exists");
                }
                catch
                {
                    OutlookApp.Session.Categories.Add(categoryName);
                    Trace.WriteLine($"Adding category {categoryName}");
                }*/
            }
            Trace.WriteLine(" - complete");
        }

        internal static void ClearCategories() {
            Trace.WriteLine("Begin clearing categories");

        }
    }


    struct UserPropertyNames {
        public const string DATE_FORMATTED = "Date Formatted";
        public const string PARSED = "Parsed";
    }

    struct OrderTypeFilter {
        public readonly string Tag;
        public readonly string WufooFilter;
        public readonly string MagentoFilter;

        public OrderTypeFilter(string Tag, string WufooFilter, string MagentoFilter = "") {
            this.Tag = Tag;
            this.WufooFilter = WufooFilter;
            this.MagentoFilter = MagentoFilter;
        }
    }

    static class OrderTypeInfo {
        public static readonly string WufooSenderEmail = "no-reply@wufoo.com";
        public static readonly string MagentoSenderEmail = "secureorders@fergusonplarre.com.au";
        public static readonly OrderTypeFilter Decorated = new OrderTypeFilter("Decorated", "Decorated Cake Order");
        public static readonly OrderTypeFilter CustomGeneral = new OrderTypeFilter("Custom General", "Custom Cake Order");
        public static readonly OrderTypeFilter CustomDeluxe = new OrderTypeFilter("Custom Deluxe", "Custom Cake Order");
        public static readonly OrderTypeFilter CustomWedding = new OrderTypeFilter("Custom Wedding", "Custom Cake Order");
        public static readonly OrderTypeFilter FlourlessOrVegan = new OrderTypeFilter("Flourless / Vegan", "Flourless & Vegan Celebration Cake Order Form");
        public static readonly OrderTypeFilter VanillaSlice = new OrderTypeFilter("Vanilla Slice", "Vanilla Slice Cake");
        public static readonly OrderTypeFilter DesignADrip = new OrderTypeFilter("Design-a-Drip", "Design A Drip");
        public static readonly OrderTypeFilter Cookie = new OrderTypeFilter("Cookie", "Cookie Cake Order");
        public static readonly OrderTypeFilter Cupcake = new OrderTypeFilter("Cupcakes", "Decorated Cake Order");
        public static readonly OrderTypeFilter Extras = new OrderTypeFilter("Extras", "Dec Room Extra");
        public static readonly List<OrderTypeFilter> Filters = new List<OrderTypeFilter>() {
        Decorated,
        CustomGeneral,
        CustomDeluxe,
        CustomWedding,
        FlourlessOrVegan,
        VanillaSlice,
        DesignADrip,
        Cookie,
        Cupcake,
        Extras};
    }
}

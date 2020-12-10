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
using System.Text;
using static olinteroplib.Methods;
using olinteroplib.ExtensionMethods;

namespace FPBInterop {
    internal static class OutlookHandler {
        /// PROPERTIES ///
        private static readonly TraceSource Tracer = new TraceSource("FPBInterop.OutlookHandling");
        //Directory.GetParent(Directory.GetCurrentDirectory()).Parent.GetDirectories("Example Orders").First().FullName

        private static bool _exitHandlerEnabled = false;
        private static List<Category> InternalCategories = new List<Category>();

        internal const string MagentoFilter = "[SenderEmailAddress]=\"secureorders@fergusonplarre.com.au\"";

        internal static Application OutlookApp;
        internal static Folder RootFolder;
        internal static Folder Inbox;
        internal static Folder DeletedItems;

        private static List<(string, OlUserPropertyType)> UserProperties =
            new List<(string, OlUserPropertyType)>() {
                ("AutoProcessed",OlUserPropertyType.olYesNo)
            };

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

            SetupCategories();
        }
        private static void _OutlookHandling_Quit() {
            Tracer.TraceEvent(TraceEventType.Information, 0, "Outlook instance closed, exiting...");
            TestHandler.StopTestEnv();
        }

        //MAIN SEQUENCE ORDER FILING
        private static List<MailItem> _GetFilteredMailList(Folder folder, string restrictFilter) {
            List<MailItem> filteredOrders = new List<MailItem>();
            try {
                Items matches = folder.Items.Restrict(restrictFilter);
                if (matches.Count == 0)
                    throw new ArgumentException();

                for (int i = 0; i < matches.Count; i++) {
                    filteredOrders.Add((MailItem)matches[i + 1]);
                }
            }
            catch (ArgumentException) {
                Tracer.TraceEvent(TraceEventType.Information, 0,
                    $"No items matching filter found in folder {folder.Name}");
            }
            catch (System.Exception) {
                Tracer.TraceEvent(TraceEventType.Error, 0, "Getting order list failed");
            }

            return filteredOrders;
        }

        internal static void ReformatMagentoDates(Folder folder) {
            if (folder == null) {
                Tracer.TraceEvent(TraceEventType.Error, 0, "Invalid folder or path");
                return;
            }

            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"Begin formatting dates in target folder {folder.Name}");

            ProgressBar pbar = new ProgressBar();
            pbar.Report(0);

            List<MailItem> magentoOrdersUnformatted = _GetFilteredMailList(folder, MagentoFilter);
            int totalOrders = magentoOrdersUnformatted.Count;
            magentoOrdersUnformatted = magentoOrdersUnformatted.Where(
                    (item, i) => {
                        pbar.Report((((double)i / (double)totalOrders)) / 2);
                        return item.UserProperties.Find("Date Formatted") == null;
                    }).ToList();
            for (int i = 0; i < magentoOrdersUnformatted.Count; i++) {
                _ReformatDate(magentoOrdersUnformatted[i]);
                UserProperty dateFormatted =
                    magentoOrdersUnformatted[i].UserProperties.Add(
                        "Date Formatted", OlUserPropertyType.olText, true);
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

        internal static void ProcessFolder(Folder folder, bool forceProcessAllItems = false, bool fileItemsToFolder=true) {
            ProcessItems(folder.Items, forceProcessAllItems,fileItemsToFolder);
        }
        internal static void ProcessItems(Items items, bool forceProcessAllItems, bool fileItemsToFolder) {
            ProgressBar pbar = new ProgressBar();
            pbar.Report(0);
            string query = "@SQL=(" + @"http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/AutoProcessed" + " IS NULL)";

            if (!forceProcessAllItems) 
                items = items.Restrict(query);
            
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"{items.Count} items unprocessed");
            int totalItems = items.Count;

            for (int i = totalItems; i > 0; i--) {
                ProcessItem((MailItem)items[i],fileItemsToFolder);
                pbar.Report((double)(totalItems - i) / (double)(totalItems + 1));
            }
            pbar.Dispose();
            Console.WriteLine("Complete");
        }
        internal static void ProcessSelectedItem() {
            ProcessItem((MailItem)OutlookApp.ActiveExplorer().Selection[1], false);
        }
        internal static void ProcessItem(MailItem item, bool fileToFolder) {
            if (item.SenderEmailAddress == "secureorders@fergusonplarre.com.au") {
                _ProcessMagento(item, fileToFolder);
            }
            if(item.SenderEmailAddress == "no-reply@wufoo.com") {
                
            }
        }

        private static void _ProcessMagento(MailItem item, bool fileToFolder) {
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"Magento Order: {item.Subject.Remove(0, 27)}");
            _ReformatDate(item);
            MagentoOrder order = HtmlHandler.Magento.MagentoBuilder(item.HTMLBody);
            UserProperty parsed = item.UserProperties.Add("AutoProcessed", OlUserPropertyType.olYesNo, false);
            parsed.Value = true;
            DisableVisiblePrintUserProp(parsed);
            item.Save();
            if (order.Meta.HasFlag(OrderMetadata.DoNotProcess)) {
                item.UnRead = false;
                item.Move(DeletedItems);
            }
            else {
                item.UnRead = true;
                foreach (MagentoProduct product in order.Products) {
                    if (product.ProductType.Categorise)
                        item.AddCategory(InternalCategories.Where(c => c.Name == product.ProductType.Name).First());
                }
                item.Save();
                if (fileToFolder && OrderShouldBeFiled(order))
                    FileItemForFuture(item, order.DeliveryDate, order.OrderPriority);
            }
        }


        //FOLDER FINDING/HANDLING
        internal static Folder GetFolderByPath(string path) {
            StringBuilder msg = new StringBuilder();
            msg.Append($"Get folder from path input '{path}': ");

            if (!path.Contains("/")) {
                try {
                    Folder target = RootFolder.GetFolder(path);
                    return target;
                }
                catch {
                    msg.Append($"Failed - Invalid folder/path");
                    Tracer.TraceEvent(TraceEventType.Verbose, 0, msg.ToString());
                    throw new DirectoryNotFoundException();
                }
            }
            char slashChar = '/';
            Folder root = OutlookApp.Session.DefaultStore.GetRootFolder() as Folder;

            if (path.StartsWith("/") | path.EndsWith("/"))
                path = path.Trim(slashChar);

            string[] folders = path.Split(slashChar);

            Folder folder = root;
            try {
                if (folder != null) {
                    for (int i = 0; i <= folders.GetUpperBound(0); i++) {
                        Folders subFolders = folder.Folders;
                        folder = subFolders.GetFolder(folders[i]);
                        if (folder == null) {
                            msg.Append($"Failed - Folder not found at path");
                            Tracer.TraceEvent(TraceEventType.Information, 0, msg.ToString());
                            throw new DirectoryNotFoundException();
                        }
                    }
                }
            }
            catch (System.Exception e) {
                msg.Append($"Failed - {e.Message}");
                Tracer.TraceEvent(TraceEventType.Information, 0, msg.ToString());
                throw new DirectoryNotFoundException();
            }
            msg.Append($"Success");
            Tracer.TraceEvent(TraceEventType.Information, 0, msg.ToString());
            return folder;
        }
        internal static Folder CreateFolderAtPath(string path, string name = null) {
            string parentPath = path;
            if (name == null) {
                parentPath = path.Substring(0, path.LastIndexOf("/"));
                name = path.Substring(path.LastIndexOf("/"));
            }
            try {
                Folder newFolder = (Folder)GetFolderByPath(parentPath).Folders.Add("name");
                return newFolder;
            }
            catch (DirectoryNotFoundException) {
                throw new DirectoryNotFoundException(
                    $"Directory not found, could not create folder at {path}");
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

        // MISCELLANEOUS METHODS

        private static DateTime _GetFirstSundayAfterDate(DateTime date) {
            DateTime sunday = date;
            while (sunday.DayOfWeek != 0) {
                sunday = sunday.AddDays(1);
            }
            Tracer.TraceEvent(TraceEventType.Verbose, 0,
                $"First Sunday after date {date:dd/MM} is {sunday} ");
            return sunday;//.ToString("MMM dd").ToUpper(); ;
        }
        private static bool OrderShouldBeFiled(MagentoOrder order) {
            if (order.DeliveryDate < _GetFirstSundayAfterDate(DateTime.Now).AddDays(1))
                return false;
            else {
                if (order.OrderPriority == FilingPriority.CUSTOM &&
                    (DateTime.Now.DayOfWeek != DayOfWeek.Thursday
                    || DateTime.Now.DayOfWeek != DayOfWeek.Friday))
                    return false;
                else
                    return true;
            }
        }
        private static void FileItemForFuture(MailItem item, DateTime date, FilingPriority priority) {
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"File order {priority}");
            string folderPath;
            Folder destination;
            switch (priority) {
                case FilingPriority.GENERAL:
                    folderPath = $"inbox/future orders/{FolderNameFromDate(_GetFirstSundayAfterDate(date))}";
                    try {
                        destination = GetFolderByPath(folderPath);
                    }
                    catch (DirectoryNotFoundException) {
                        destination = CreateFolderAtPath(folderPath);
                    }
                    SetupUserProperties(destination);
                    item.Move(destination);
                    break;
                case FilingPriority.COOKIE:
                    destination = GetFolderByPath($"inbox/cookie cakes");
                    item.Move(destination);
                    break;
                case FilingPriority.CUSTOM:
                    folderPath = $"inbox/custom orders/{FolderNameFromDate(date)}";
                    try {
                        destination = GetFolderByPath(folderPath);
                    }
                    catch (DirectoryNotFoundException) {
                        destination = CreateFolderAtPath(folderPath);
                    }
                    SetupUserProperties(destination);
                    item.Move(destination);
                    break;
                case FilingPriority.NONE:
                default:
                    break;
            }
        }
        private static string FolderNameFromDate(DateTime date) {
            if (date.Year > DateTime.Now.Year)
                return date.ToString("yyyy MMM dd").ToUpper();
            else
                return date.ToString("MMM dd").ToUpper();
        }

        internal static void SetupCategories() {
            foreach (KeyValuePair<String, ProductType> pair in XmlHandler.ProductTypesStandard) {
                if (pair.Value.Name == "Undefined" || !pair.Value.Categorise)
                    continue;

                if (CategoryExists(pair.Value.Name)) {
                    InternalCategories.Add(OutlookApp.Session.Categories[pair.Value.Name]);
                    continue;
                } else {
                    InternalCategories.Add(OutlookApp.Session.Categories.Add(pair.Value.Name));
                }
            }
        }
        internal static void ClearCategories() {
            throw new NotImplementedException();
        }
        private static bool CategoryExists(string name) {
            try {
                Category category =
                     OutlookApp.Session.Categories[name];
                if (category != null) {
                    return true;
                }
                else {
                    return false;
                }
            }
            catch { return false; }
        }

        internal static void SetupUserProperties(List<Folder> folders) {
            foreach (Folder folder in folders) {
                SetupUserProperties(folder);
                ; }
        }
        internal static void SetupUserProperties(Folder folder) {
            try {
                foreach ((string, OlUserPropertyType) entry in UserProperties) {
                    folder.UserDefinedProperties.Add(
                        entry.Item1,
                        entry.Item2);
                }
            }
            catch (System.Exception) { }
        }
    }

    internal static class Helper {
        private static string Desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private static string ExampleOrdersFolder =
            @"C:\Users\decroom\source\repos\OutlookInteropSolution\OutlookInterop\Example Orders";
        internal static void SaveSelected(string filename) {
            SaveHTML((MailItem)OutlookHandler.OutlookApp.ActiveExplorer().Selection[1], ExampleOrdersFolder, filename);
        }
        internal static void ExampleWufoo() {
            Items wufoo = OutlookHandler.DeletedItems.Items.Restrict("no-reply@wufoo.com");
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
    }
}

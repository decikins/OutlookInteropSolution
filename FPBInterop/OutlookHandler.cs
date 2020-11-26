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
        private static List<Category> InternalCategories = new List<Category>();

        internal const string MagentoFilter = "[SenderEmailAddress]=\"secureorders@fergusonplarre.com.au\"";

        internal static Application OutlookApp;
        internal static Folder RootFolder;
        internal static Folder Inbox;
        internal static Folder DeletedItems;

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
            Tracer.TraceEvent(TraceEventType.Information, 0, "Outlook instance closed, exiting...");
            TestHandler.StopTestEnv();
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

            List<MailItem> magentoOrdersUnformatted = GetFilteredMailList(folder, MagentoFilter);
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

        internal static void ProcessFolder(Folder folder, bool forceProcessAllItems = false) {
            ProcessItems(folder.Items, forceProcessAllItems);
        }
        internal static void ProcessItems(Items items, bool forceProcessAllItems) {
            ProgressBar pbar = new ProgressBar();
            pbar.Report(0);
            string query = "@SQL=(" + @"http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/AutoProcessed" + " IS NULL)";
            Debug.WriteLine(items.Count);
            if (!forceProcessAllItems)
               items = items.Restrict(query);
            Debug.WriteLine(items.Count);
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
                Tracer.TraceEvent(TraceEventType.Verbose, 0, item.Subject.Remove(0, 27));
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
                    foreach(MagentoProduct product in order.Products) {
                        item.AddCategory(InternalCategories.Where(c => c.Name == product.ProductType.Name).Single());
                    }
                    if (OrderToBeFiled(order))
                        FileItemForFuture(item, order.OrderPriority);
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
            DateTime sunday = date;
            while (sunday.DayOfWeek != 0) {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, sunday.ToString());
                sunday = sunday.AddDays(1);
            }
            Tracer.TraceEvent(TraceEventType.Verbose, 0, 
                $"First Sunday after date {date:dd/MM} is {sunday} ");
            return sunday;//.ToString("MMM dd").ToUpper(); ;
        }
        private static bool OrderToBeFiled(MagentoOrder order) {
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
        private static void FileItemForFuture(MailItem item, FilingPriority priority) {
            switch (priority) {
                case FilingPriority.GENERAL:
                    break;
                case FilingPriority.COOKIE:
                    break;
                case FilingPriority.CUSTOM:
                    break;
                case FilingPriority.NONE:
                default:
                    break;
            }
        }
        internal static void SetupCategories() {
            foreach(KeyValuePair<String,ProductType> pair in XmlHandler.ProductTypesStandard) {
                if (pair.Value.Name == "Undefined")
                    continue;

                if (CategoryExists(pair.Value.Name)) {
                    InternalCategories.Add(OutlookApp.Session.Categories[pair.Value.Name]);
                    continue;
                } else {
                    InternalCategories.Add(OutlookApp.Session.Categories.Add(pair.Value.Name));
                }
            }
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
        internal static void ClearCategories() {
            throw new NotImplementedException();
        }
        internal static void SetupUserProperties(List<Folder> folders) {
            foreach (Folder folder in folders) {
                try {
                    folder.UserDefinedProperties.Add(
                        "AutoProcessed",
                        OlUserPropertyType.olYesNo);
                }
                catch (System.Exception) { }
            }
        }
    }
}

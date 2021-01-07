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
    public static class OutlookHandler {
        internal static Application olApp = new Application();
        private static XmlHandler xmlHandle = new XmlHandler();
        private static readonly TraceSource Tracer = new TraceSource("FPBInterop.OutlookHandler");

        // Refactor to remove this internal category list and figure out
        // what to do with this UserProperties scheme.
        private static List<Category> InternalCategories = SetupCategories();

        private static readonly List<(string name, OlUserPropertyType type)> UserProperties =
            new List<(string name, OlUserPropertyType type)>() {
                        ("AutoProcessed",OlUserPropertyType.olYesNo)
        };

        public static void ProcessFolder(string folderPath, bool forceProcess, bool fileProcessed) {
            List<MailItem> items;
            Folder target;
            try {
                target = olApp.Session.GetFolderByPath(folderPath);
            }
            catch (DirectoryNotFoundException) {
                Console.WriteLine("Invalid directory - Not found");
                return;
            }
            items = new List<MailItem>(target.Items.Count);
            foreach (MailItem item in target.Items) {
                items.Add(item);
            }
            MailReader reader = new MailReader(forceProcess, fileProcessed);
            reader.ProcessItems(items);
        }
        public static void ProcessSelectedOrder(bool forceProcess, bool fileProcessed) {
            MailReader reader = new MailReader(forceProcess, fileProcessed);
            reader.ProcessItems(new List<MailItem>(1) { (MailItem)olApp.ActiveExplorer().Selection[1] });
        }

        internal class MailReader {
            public bool forceProcessAllItems { get; private set; }
            public bool fileToFolder { get; private set; }
            public MailReader(bool forceProcessAllItems = false, bool fileToFolder = false) {
                this.forceProcessAllItems = forceProcessAllItems;
                this.fileToFolder = fileToFolder;
            }

            public void ProcessItems(List<MailItem> items) {
                ProgressBar pbar = new ProgressBar();
                pbar.Report(0);
                if (!forceProcessAllItems)
                    items.RemoveAll(i => i.UserProperties.Find("AutoProcessed", true) != null);

                Tracer.TraceEvent(TraceEventType.Verbose, 0, $"{items.Count} items unprocessed");
                int totalItems = items.Count;
                for (int i = totalItems - 1; i >= 0; i--) {
                    _ProcessItem(items[i]);
                    pbar.Report((double)(totalItems - i) / (double)(totalItems));
                }
                pbar.Dispose();
                Console.WriteLine("Complete");
            }
            private void _ProcessItem(MailItem item) {
                if (item.SenderEmailAddress == "secureorders@fergusonplarre.com.au") {
                    MagentoOrder order = MagentoProcessor.Process(item);
                    if (order.Meta.HasFlag(OrderMetadata.DoNotProcess)) {
                        item.UnRead = false;
                        item.Move(olApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems));
                    }
                    else {
                        item.UnRead = true;
                        foreach (MagentoProduct product in order.Products) {
                            if (product.ProductType.Categorise)
                                item.AddCategory(InternalCategories.Where(
                                    c => c.Name == product.ProductType.Name).First());
                        }
                        item.Close(OlInspectorClose.olSave);
                    }
                    if (fileToFolder && _OrderShouldBeFiled(order))
                        _FileItemForFuture(item, order.DeliveryDate, order.OrderPriority);
                }
                if (item.SenderEmailAddress == "no-reply@wufoo.com") {

                }
            }

            internal static class MagentoProcessor {
                public static MagentoOrder Process(MailItem item) {
                    Tracer.TraceEvent(TraceEventType.Information, 0,
                         $"Magento Order: {item.Subject.Remove(0, 27)}");
                    try {
                        _ReformatDate(item);
                    }
                    catch (InvalidDateFormatException) {
                        Tracer.TraceEvent(TraceEventType.Information, 0,
                            "Date formatting failed, unrecognized date format");
                    }
                    UserProperty parsed =
                        item.UserProperties.Add("AutoProcessed", OlUserPropertyType.olYesNo, false);
                    parsed.Value = true;
                    DisableVisiblePrintUserProp(parsed);
                    MagentoOrder order;
                    try {
                        order = HtmlHandler.Magento.MagentoBuilder(item.HTMLBody);
                    }
                    catch (InvalidXPathException) {
                        Console.WriteLine("An error has occured, please check the log file");
                        return null;
                    }

                    try {
                        item.Close(OlInspectorClose.olSave);
                    }
                    catch (System.Runtime.InteropServices.COMException) {
                        Thread.Sleep(5000);
                        try {
                            item.Close(OlInspectorClose.olSave);
                        }
                        catch (System.Runtime.InteropServices.COMException) {
                            item.Close(OlInspectorClose.olPromptForSave);
                        }
                    }
                    return order;
                }
                private static void _ReformatDate(MailItem item) {
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
                        else throw new InvalidDateFormatException();
                    }

                    string newDateString = newDate.ToString("dddd, d MMMM yyyy");
                    item.HTMLBody = item.HTMLBody.Replace(dateMatch, newDateString);
                    item.Close(OlInspectorClose.olSave);
                }
            }

            internal static class WufooProcessor {

            }

            private bool _OrderShouldBeFiled(MagentoOrder order) {
                if (order.DeliveryDate < GetFirstSundayAfterDate(DateTime.Now).AddDays(1))
                    return false; // Order is for this week
                else {
                    // Order not for this week, but is custom
                    if (order.OrderPriority == FilingPriority.CUSTOM
                        && _FileCustomToday())
                        return true;
                    else
                        return false;
                }
            }

            private static bool _FileCustomToday() {
                List<DayOfWeek> CustomOrderDays = new List<DayOfWeek>
                {DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday};
                return CustomOrderDays.Contains(DateTime.Now.DayOfWeek);
            }
            
            private void _FileItemForFuture(MailItem item, DateTime date, FilingPriority priority) {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, $"File order {priority}");
                string folderPath;
                Folder destination;
                string destinationFolderName = FolderNameFromDate(GetFirstSundayAfterDate(date));
                switch (priority) {
                    case FilingPriority.GENERAL:
                        folderPath = 
                            FolderPaths.FutureGeneralOrders+$"/{destinationFolderName}";
                        break;
                    case FilingPriority.COOKIE:
                        folderPath = FolderPaths.CookieCakes;
                        break;
                    case FilingPriority.CUSTOM:
                        folderPath = FolderPaths.FutureCustomOrders+$"/{destinationFolderName}";
                        break;
                    case FilingPriority.NONE:
                    default:
                        return;
                }
                try {
                    destination = olApp.Session.GetFolderByPath(folderPath);
                }
                catch (DirectoryNotFoundException) {
                    destination = olApp.Session.CreateFolderAtPath(folderPath);
                }
                SetupUserProperties(destination);
                item.Move(destination);
            }

            private List<MailItem> _GetFilteredMailList(Folder folder, string restrictFilter) {
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
        }

        internal static DateTime GetFirstSundayAfterDate(DateTime date) {
            DateTime sunday = date;
            while (sunday.DayOfWeek != 0) {
                sunday = sunday.AddDays(1);
            }
            Tracer.TraceEvent(TraceEventType.Verbose, 0,
                $"First Sunday after date {date:dd/MM} is {sunday} ");
            return sunday;//.ToString("MMM dd").ToUpper(); ;
        }
        internal static string FolderNameFromDate(DateTime date) {
            if (date.Year > DateTime.Now.Year)
                return date.ToString("yyyy MMM dd").ToUpper();
            else
                return date.ToString("MMM dd").ToUpper();
        }

        internal static List<Category> SetupCategories() {
            List<Category> c = new List<Category>();
            foreach (KeyValuePair<String, ProductType> pair in xmlHandle.GetProductTypesStandard()) {
                if (pair.Value.Name == "Undefined" || !pair.Value.Categorise)
                    continue;

                if (olApp.Session.CategoryExists(pair.Value.Name)) {
                    c.Add(olApp.Session.Categories[pair.Value.Name]);
                    continue;
                }
                else {
                    c.Add(olApp.Session.Categories.Add(pair.Value.Name));
                }
            }
            return c;
        }
        internal static void ClearCategories() {
            throw new NotImplementedException();
        }

        internal static void SetupUserProperties(List<Folder> folders) {
            foreach (Folder folder in folders) {
                SetupUserProperties(folder);
            }
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

        public static void SaveSelectedItemHtml() {
            MailItem item = (MailItem)olApp.ActiveExplorer().Selection[1];
            File.WriteAllText("./test.html", item.HTMLBody);
        }
    }

    internal static class FolderPaths {
        internal const string FutureCustomOrders = @"inbox/custom orders";
        internal const string FutureGeneralOrders = @"inbox/future orders";
        internal const string CookieCakes = @"inbox/cookie cakes";
    }

    internal static class DASLQuery {
        internal const string UserPropertyQuery = @"http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";
    }
}

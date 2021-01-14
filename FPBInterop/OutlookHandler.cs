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
using System.Runtime.InteropServices;
using static olinteroplib.Methods;
using olinteroplib.ExtensionMethods;

namespace FPBInterop {
    public static class OutlookHandler {
        internal static Application olApp = new Application();
        internal static XmlHandler xmlHandle = new XmlHandler();

        public static Folder GetSelectedFolder { get { return (Folder)olApp.ActiveExplorer().CurrentFolder; } }

        // Refactor to remove this internal category list and figure out
        // what to do with this UserProperties scheme.
        private static List<Category> InternalCategories = SetupCategories();

        public static void ProcessFolder(string folderPath, bool forceProcess, bool moveToFolder) {
            Logger.TraceEvent(TraceEventType.Verbose, 
                $"Process folder {folderPath}; Force {forceProcess}; File {moveToFolder}");
            List<MailItem> items;
            Folder target;
            try {
                target = olApp.Session.GetFolderByPath(folderPath);
            }
            catch (DirectoryNotFoundException) {
                Console.WriteLine("Invalid directory - Not found");
                Logger.TraceEvent(TraceEventType.Error,
                    $"Folder {folderPath} not found");
                return;
            }
            items = new List<MailItem>(target.Items.Count);
            foreach (MailItem item in target.Items) {
                items.Add(item);
            }
            Marshal.ReleaseComObject(target);
            MailReader reader = new MailReader(forceProcess, moveToFolder);
            reader.ProcessItems(items);
            Console.WriteLine("Complete");
        }
        public static void ProcessSelectedOrder(bool forceProcess, bool fileProcessed) {
            Logger.TraceEvent(TraceEventType.Verbose,
                $"Process selected order: \"{((MailItem)olApp.ActiveExplorer().Selection[1]).Subject}\"");
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

               Logger.TraceEvent(TraceEventType.Verbose, $"{items.Count} items unprocessed");
                int totalItems = items.Count;
                for (int i = totalItems - 1; i >= 0; i--) {
                    _ProcessItem(items[i]);
                    pbar.Report((double)(totalItems - i) / (double)(totalItems));
                }
                Thread.Sleep(1000);
                pbar.Dispose();
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
                        _FileItemForFuture(item, order.DeliveryDate, order.Priority);
                }
                if (item.SenderEmailAddress == "no-reply@wufoo.com") {
                    switch (_GetWufooOrderType(item.Subject)) {
                        case WufooOrderType.Extras:
                            item.AddCategory()
                            break;
                    }
                }
            }

            internal static class MagentoProcessor {
                public static MagentoOrder Process(MailItem item) {
                    Logger.TraceEvent(TraceEventType.Information,
                         $"Magento Order: {item.Subject.Remove(0, 27)}");
                    try {
                        _ReformatDate(item);
                    }
                    catch (InvalidDateFormatException) {
                       Logger.TraceEvent(TraceEventType.Error,
                            "Date formatting failed, unrecognized date format");
                    }

                    MagentoOrder order;
                    try {
                        order = HtmlHandler.Magento.MagentoBuilder(item.HTMLBody);
                    }
                    catch (InvalidXPathException) {
                        Console.WriteLine("An error has occured, please check the log file");
                        return null;
                    }

                    UserProperty parsed =
                        item.UserProperties.Add("AutoProcessed", OlUserPropertyType.olYesNo, false);
                    parsed.Value = true;
                    DisableVisiblePrintUserProp(parsed);

                    try {
                        item.Close(OlInspectorClose.olSave);
                    }
                    catch (COMException) {
                        Thread.Sleep(3000);
                        try {
                            item.Close(OlInspectorClose.olSave);
                        }
                        catch (COMException) {
                            Logger.TraceEvent(TraceEventType.Error,
                                $"Unable to save categories, userproperties and/or date formatting to order");
                        }
                    }
                    return order;
                }
                private static void _ReformatDate(MailItem item) {
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    DateTime newDate = DateTime.MinValue;
                    string dateMatch;

                    Regex regex = new Regex(RegexStrings.FullDate);
                    if (regex.IsMatch(item.HTMLBody))
                        return;

                    if (regex.IsMatch(item.HTMLBody)) {
                        dateMatch = regex.Match(item.HTMLBody).Value;
                        newDate = DateTime.ParseExact(dateMatch, "dd/MM/yyyy", provider);
                    }
                    else {
                        regex = new Regex(RegexStrings.FullDateLeadingZero);
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

            private static WufooOrderType _GetWufooOrderType(string subject) {
                if (subject.Contains("Dec Room Extra"))
                    return WufooOrderType.Extras;

                if (subject.Contains("Design A Drip"))
                    return WufooOrderType.DesignADrip;

                if (subject.Contains("Vanilla Slice Cake Order Form"))
                    return WufooOrderType.VanillaSlice;

                if (subject.Contains("Flourless & Vegan Celebration Cake Order Form"))
                    return WufooOrderType.FlourlessAndVegan;

                if (subject.Contains("Custom Cake Order Form"))
                    return WufooOrderType.Custom;

                if (subject.Contains("Decorated Cake Order Form"))
                    return WufooOrderType.Decorated;

                return WufooOrderType.Misc;
            }

            private bool _OrderShouldBeFiled(BaseOrder order) {
                if (order.DeliveryDate < FolderNameHandler.GetFirstSundayAfterDate(DateTime.Now).AddDays(1))
                    return false; // Order is for this week
                else {
                    if (order.Priority == FilingPriority.CUSTOM) {
                        if (order.DeliveryDate > FolderNameHandler.GetFirstSundayAfterDate(DateTime.Now).AddDays(8))
                            return true;
                        else {
                            return _FileCustomToday();
                        }
                    }
                    else
                        return true;
                }
            }

            private static bool _FileGeneralToday() {
                List<DayOfWeek> GeneralOrderDays = new List<DayOfWeek>(2)
               {DayOfWeek.Saturday, DayOfWeek.Sunday};
                return GeneralOrderDays.Contains(DateTime.Now.DayOfWeek);
            }
            private static bool _FileCustomToday() {
                List<DayOfWeek> CustomOrderDays = new List<DayOfWeek>(4)
                {DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday};
                return !CustomOrderDays.Contains(DateTime.Now.DayOfWeek);
            }
            
            private void _FileItemForFuture(MailItem item, DateTime date, FilingPriority priority) {
               Logger.TraceEvent(TraceEventType.Verbose, $"File order {priority}");
                string folderPath;
                Folder destination;
                string destinationFolderName = 
                    FolderNameHandler.FolderNameFromDate(FolderNameHandler.GetFirstSundayAfterDate(date));
               Logger.TraceEvent(TraceEventType.Verbose,
                    $"First Sunday after {date} is {FolderNameHandler.GetFirstSundayAfterDate(date)}");
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
                item.Close(OlInspectorClose.olSave);
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
                   Logger.TraceEvent(TraceEventType.Information,
                        $"No items matching filter found in folder {folder.Name}");
                }
                catch (System.Exception) {
                   Logger.TraceEvent(TraceEventType.Error, "Getting order list failed");
                }

                return filteredOrders;
            }
        }

        internal static class FolderNameHandler {
            internal static DateTime GetFirstSundayAfterDate(DateTime date) {
                DateTime sunday = date;
                while (sunday.DayOfWeek != 0) {
                    sunday = sunday.AddDays(1);
                }
                return sunday;//.ToString("MMM dd").ToUpper(); ;
            }
            internal static string FolderNameFromDate(DateTime date) {
                if (date.Year > DateTime.Now.Year)
                    return date.ToString("yyyy MMM dd").ToUpper();
                else
                    return date.ToString("MMM dd").ToUpper();
            }
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

        public static void SetupUserProperties(Folder folder = null) {
            if (folder == null)
                folder = GetSelectedFolder;
            try {
                foreach ((string name, OlUserPropertyType type) entry in UserPropertyFields.UserProperties) {
                    if(folder.UserDefinedProperties.Find(entry.name)==null)
                        folder.UserDefinedProperties.Add(
                            entry.name,
                            entry.type);
                }
            }
            catch (System.Exception) { }
        }

        public static void SaveSelectedItemHtml() {
            MailItem item = (MailItem)olApp.ActiveExplorer().Selection[1];
            File.WriteAllText("./test.html", item.HTMLBody);
        }
    }

    internal enum WufooOrderType {
        None,
        Misc,
        Extras,
        Decorated,
        DesignADrip,
        FlourlessAndVegan,
        VanillaSlice,
        Custom
    }

    internal static class UserPropertyFields {
        internal const string CustomFolderViewName = "FPBInteropView";
        internal static readonly List<(string name, OlUserPropertyType type)> UserProperties =
            new List<(string name, OlUserPropertyType type)>() {
                        ("AutoProcessed",OlUserPropertyType.olYesNo)
        };
    }
    internal static class FolderPaths {
        internal const string FutureCustomOrders = @"inbox/custom orders";
        internal const string FutureGeneralOrders = @"inbox/future orders";
        internal const string CookieCakes = @"inbox/cookie cakes";
    }

    internal static class DASLQuery {
        internal const string UserPropertyQuery = @"http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";
    }

    internal static class RegexStrings {
        internal const string ShorthandDate = @"\d\d\/\d\d\/\d\d\d\d";
        internal const string FullDateLeadingZero = @"((\w){3,6}day), 0\d ((Jan|Febr)uary|Ma(rch|y)|A(pril|ugust)|Ju(ne|ly)|((Sept|Nov|Dec)em|Octo)ber) (\d){4}";
        internal const string FullDate = @"((\w){3,6}day),(( [1-3][0-9] )|( [1-9] ))((Jan|Febr)uary|Ma(rch|y)|A(pril|ugust)|Ju(ne|ly)|((Sept|Nov|Dec)em|Octo)ber) (\d){4}";
    }
}

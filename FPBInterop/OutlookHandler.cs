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
        private static XmlHandler xmlHandle = new XmlHandler();
        internal static Application olApp = new Application();
        private static readonly TraceSource Tracer = new TraceSource("FPBInterop.OutlookHandler");


        //Refactor to remove this internal category list
        private static List<Category> InternalCategories = SetupCategories();

        internal const string MagentoFilter = "[SenderEmailAddress]=\"secureorders@fergusonplarre.com.au\"";

        private static List<(string, OlUserPropertyType)> UserProperties =
            new List<(string, OlUserPropertyType)>() {
                ("AutoProcessed",OlUserPropertyType.olYesNo)
            };

        public static void ProcessFolder(string folderPath, bool forceProcess, bool fileProcessed) {
            MailProcessor mProc = new MailProcessor(forceProcess, fileProcessed);
            mProc.ProcessFolder(folderPath);
        }

        public static void ProcessSelectedOrder(bool forceProcess, bool fileProcessed) {
            MailProcessor mProc = new MailProcessor(forceProcess, fileProcessed);
            mProc.ProcessSelectedItem();
        }

        internal class MailProcessor {
            public bool forceProcessAllItems { get; private set; }
            public bool fileToFolder { get; private set; }
            public MailProcessor(bool forceProcessAllItems=false,bool fileToFolder=false) {
                this.forceProcessAllItems = forceProcessAllItems;
                this.fileToFolder = fileToFolder;
            }
                
            public void ProcessFolder(string folderPath) {
                _ProcessItems(olApp.Session.GetFolderByPath(folderPath).Items);
            }
            public void ProcessSelectedItem() {
                _ProcessItem((MailItem)olApp.ActiveExplorer().Selection[1]);
            }
            public void _ProcessItems(Items items) {
                ProgressBar pbar = new ProgressBar();
                pbar.Report(0);
                string query = "@SQL=(" + @"http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/AutoProcessed" + " IS NULL)";

                if (!forceProcessAllItems)
                    items = items.Restrict(query);

                Tracer.TraceEvent(TraceEventType.Verbose, 0, $"{items.Count} items unprocessed");
                int totalItems = items.Count;

                for (int i = totalItems; i > 0; i--) {
                    _ProcessItem((MailItem)items[i]);
                    pbar.Report((double)(totalItems - i) / (double)(totalItems + 1));
                }
                pbar.Dispose();
                Console.WriteLine("Complete");
            }
            public void _ProcessItem(MailItem item) {
                Tracer.TraceEvent(TraceEventType.Information, 0, "Test");
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
                        item.Save();
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
                    _ReformatDate(item);
                    MagentoOrder order = HtmlHandler.Magento.MagentoBuilder(item.HTMLBody);
                    UserProperty parsed = 
                        item.UserProperties.Add("AutoProcessed", OlUserPropertyType.olYesNo, false);
                    parsed.Value = true;
                    DisableVisiblePrintUserProp(parsed);
                    item.Save();
                    return order;
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
            }

            public bool _OrderShouldBeFiled(MagentoOrder order) {
                if (order.DeliveryDate < GetFirstSundayAfterDate(DateTime.Now).AddDays(1))
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
            public void _FileItemForFuture(MailItem item, DateTime date, FilingPriority priority) {
                Tracer.TraceEvent(TraceEventType.Verbose, 0, $"File order {priority}");
                string folderPath;
                Folder destination;
                switch (priority) {
                    case FilingPriority.GENERAL:
                        // if(date < DateTime.Now.AddDays(7))

                        folderPath = $"inbox/future orders/{FolderNameFromDate(GetFirstSundayAfterDate(date))}";
                        try {
                            destination = olApp.Session.GetFolderByPath(folderPath);
                        }
                        catch (DirectoryNotFoundException) {
                            destination = olApp.Session.CreateFolderAtPath(folderPath);
                        }
                        SetupUserProperties(destination);
                        item.Move(destination);
                        break;
                    case FilingPriority.COOKIE:
                        destination = olApp.Session.GetFolderByPath($"inbox/cookie cakes");
                        item.Move(destination);
                        break;
                    case FilingPriority.CUSTOM:
                        folderPath = $"inbox/custom orders/{FolderNameFromDate(date)}";
                        try {
                            destination = olApp.Session.GetFolderByPath(folderPath);
                        }
                        catch (DirectoryNotFoundException) {
                            destination = olApp.Session.CreateFolderAtPath(folderPath);
                        }
                        SetupUserProperties(destination);
                        item.Move(destination);
                        break;
                    case FilingPriority.NONE:
                    default:
                        break;
                }
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

        internal interface IProcessor {
            bool forceProcessAllItems { get; }
            bool fileToFolder { get; }
            void ProcessFolder(string folderPath);
            void ProcessSelectedItem();
            void _ProcessItems(Items items);
            void _ProcessItem(MailItem item);

            bool _OrderShouldBeFiled(BaseOrder Order);
            void _FileItemForFuture(MailItem item, DateTime date, FilingPriority priority);
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
                ;
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
    }
}

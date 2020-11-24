using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Threading;
using olinteroplib.ExtensionMethods;

namespace FPBInterop {
    internal static class TestHandler {
        private static readonly TraceSource Tracer = new TraceSource("FPBInterop.TestHandling");

        //TESTING SCENARIO 
        private static List<MailItem> _testItems = new List<MailItem>();
        private static bool _testSetup = false;
        private static Folder _testFolderParent;

        private static Folder TestFolder { get; set; }

        //TESTING SCENARIO
        internal static void StopTestEnv() {
            Tracer.TraceEvent(TraceEventType.Verbose, 0, "Begin test scenario cleanup");
            try {
                OutlookHandler.DeletedItems.Folders[TestFolder].Delete();
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
            TestFolder.MoveTo(OutlookHandler.DeletedItems);
            while (OutlookHandler.DeletedItems.Folders.Count == 0) { Thread.Sleep(100); }
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
    }
}

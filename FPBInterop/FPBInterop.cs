using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using static FPBInterop.OutlookHandler;

namespace FPBInterop {
    public static class FPBInterop {
        private static readonly TraceSource Tracer = new TraceSource("FPBInterop");
        private const string _testFolderName = "Test";

        public static void Init() {
            XmlHandler.LoadConfig();
            OutlookHandler.SetupAppRefs();
            OutlookHandler.SetupCategories();
        }

        public static bool SetupDefaultTestEnv(int maxItems, string sourceItemFilter = null) {
            if (maxItems < 1 | maxItems > 25) {
                Tracer.TraceEvent(TraceEventType.Information, 0, "Invalid number of items to copy");
                return false;
            }

            if (TestHandler.SetupTestEnv(DeletedItems, RootFolder, _testFolderName, maxItems, sourceItemFilter) == false) {
                Tracer.TraceEvent(TraceEventType.Error, 0, "Setting up test folder failed");
                return false;
            }
            else return true;
            
        }
        public static void StopTest() {
            TestHandler.StopTestEnv();
        }

        public static void ReformatMagentoDates(string folder) {
            OutlookHandler.ReformatMagentoDates(GetFolderByPath(folder));
        }

        public static void ProcessFolder(string folderPath, bool forceProcess, bool fileToFolder) {
            Tracer.TraceEvent(TraceEventType.Information, 0,
            $"##\tBEGIN PROCESSING FOLDER {folderPath.ToUpper()}\n" +
            $"##\tFORCE PROCESS ALL {forceProcess}\n" +
            $"##\tFILE ITEMS {fileToFolder}");
            OutlookHandler.ProcessItems(GetFolderByPath(folderPath).Items, forceProcess, fileToFolder);
            Tracer.TraceEvent(TraceEventType.Information, 0, $"## FOLDER PROCESSING COMPLETE");
        }
        public static void ProcessSelectedOrder() {
            OutlookHandler.ProcessSelectedItem();
        }

        public static void SaveSelected(string filename) {
            Helper.SaveSelected(filename);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static FPBInterop.OutlookHandler;

namespace FPBInterop {
    public static class FPBInterop {
        private const string _testFolderName = "Test";

        public static void Init() {
            OutlookHandler.SetupAppRefs();
            XmlHandler.LoadConfig();
        }

        public static bool SetupDefaultTestEnv(int maxItems, string sourceItemFilter = null) {
            return OutlookHandler.SetupTestEnv(DeletedItems, RootFolder, _testFolderName, maxItems, sourceItemFilter);
        }
        public static void StopTest() {
            OutlookHandler.StopTestEnv();
        }

        public static void ReformatMagentoDates(string folder) {
            OutlookHandler.ReformatMagentoDates(GetFolderByPath(folder));
        }

        public static void ProcessFolder(string folderPath, bool ignoreProcessed) {
            OutlookHandler.ProcessItems(GetFolderByPath(folderPath).Items, ignoreProcessed);
        }

        public static void ProcessSelectedOrder() {
            OutlookHandler.ProcessSelectedItem();
        }

        public static void SaveSelected(string filename) {
            OutlookHandler.SaveSelected(filename);
        }
    }
}

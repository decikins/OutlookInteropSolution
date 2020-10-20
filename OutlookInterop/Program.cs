using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Diagnostics;
using static FPBInterop.OutlookHandling.OutlookHandling;


namespace OutlookInterop {

    static class Program {

    /// PROPERTIES ///

        private static readonly bool DeleteTestFolderOnExit = true;
        private static readonly bool ClearCategoriesOnExit = true;
        private static bool TestingFolderScenario;
        private static bool TraceToLogFile;

        private static List<Franchise> StoreList = new List<Franchise>();

        private static bool CleanupCompleted = false;
        public static bool ApplicationIsExiting = false;

        private static readonly ConsoleTraceListener ConsoleTracer = new ConsoleTraceListener();

    /// METHODS ///

        static void Main(string[] args) {

            Trace.AutoFlush = true;
            TraceToLogFile = args.Contains("-l");
            Trace.WriteLine($"Write to log set to {TraceToLogFile.ToString().ToLower()}");
            if (TraceToLogFile) {
                string filePath = $"./log.txt {DateTime.Now.ToString("dd.MM.yy HH-mm-ss")}.txt";
                TextWriterTraceListener traceToLogFile =
                    new TextWriterTraceListener(filePath);
                Trace.Listeners.Add(traceToLogFile);
            }

            SetupOutlookRefs();

            UserInputLoop();
        }

        private static void UserInputLoop() {
            do {
                Console.Write("->");
                string input = Console.ReadLine();
                string stringArg = "";

                try {
                    stringArg = input.Contains('"') ?
                        input.Substring(input.IndexOf('"') + 1, input.LastIndexOf('"') - (input.IndexOf('"') + 1)) : null;
                }
                catch (ArgumentOutOfRangeException) {
                    Console.WriteLine("Invalid argument (check quote marks)");
                    continue;
                }

                string[] inputArgs = input.Split(' ');

                switch (inputArgs[0].ToLower()) {
                    case "-f":
                    case "formatdates":
                        ReformatMagentoDates(stringArg);
                        break;
                    case "-e":
                    case "enumfolders":
                        EnumerateFolders(inputArgs.Count() > 1 && inputArgs[1].ToLower() == "-h");
                        break;
                    case "-t":
                    case "setuptest":
                        int items;
                        if (inputArgs.Count() < 2) {
                            if (!int.TryParse(inputArgs[1], out items) || items < 1 || items > 25) {
                                Trace.WriteLine("Maxitems parameter must be non-negative integer greater than 1, less than 25");
                                break;
                            }
                        }
                        else
                            Trace.WriteLine("Please specify maximum number of items to duplicate: \"setuptest\" XX");

                        if (SetupDefaultTestEnv(int.Parse(inputArgs[1]), stringArg))
                            TestingFolderScenario = true;
                        break;
                    case "-s":
                    case "save":
                        SaveSelected(stringArg);
                        break;
                    case "-st":
                    case "stoptest":
                        StopTestEnv();
                        break;
                    case "-p":
                    case "parseorders":
                        if (String.IsNullOrEmpty(stringArg))
                            ParseSelectedOrder();
                        else
                            ParseOrdersInFolder(stringArg);
                        break;
                    /*case "help":
                        ShowHelp();
                        break;*/
                    case "":
                        Console.CursorTop--;
                        Console.WriteLine("");
                        Console.CursorTop--;
                        break;
                    case "x":
                    case "exit":
                    case "quit":
                    case "close":
                        ApplicationIsExiting = true;
                        break;
                    default:
                        Console.WriteLine("Invalid command; type 'help' for a list of valid commands");
                        break;

                }
            }
            while (!ApplicationIsExiting);
        }
        private static void ShowHelp() {
            Console.WriteLine("formatdates \"PATH\" OR -f \"PATH\":\n\tFormat dates for Magento orders in the specified\n\tfolder path");
            Console.WriteLine("");
        }
        private static void LoadShops(string filename, string path) {
            string storelist;
            try {
                storelist = File.ReadAllText(path + @"\" + filename);
            }
            catch (FileNotFoundException) {
                Trace.WriteLine("Filename not found at specified filepath!\n");
                return;
            }

            StringReader sr = new StringReader(storelist);

            while (sr.Peek() != -1) {
                string name = sr.ReadLine();
                Franchise store = new Franchise(name, String.Empty, true);
                StoreList.Add(store);
            };
        }
    }
}

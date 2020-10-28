using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using static FPBInterop.OutlookHandling;


namespace FPBInteropConsole {

    static class Program {

    /// PROPERTIES ///
        private static string filePath = $"./log.txt {DateTime.Now:dd.MM.yy HH-mm-ss}.txt";
        private static bool TraceToLogFile;

        private static List<Franchise> StoreList = new List<Franchise>();

        private static readonly ConsoleTraceListener ConsoleTracer = new ConsoleTraceListener();
        private readonly static TraceSource Tracer = new TraceSource("FPBInterop.Console");

    /// METHODS ///

        static void Main(string[] args) {
            Tracer.TraceEvent(TraceEventType.Verbose, 0, "Starting FPBInterop console app");
            TraceToLogFile = args.Contains("-l");
            Tracer.TraceEvent(TraceEventType.Verbose, 0, $"Write to log set to {TraceToLogFile.ToString().ToLower()}");
            if (TraceToLogFile) {
                TextWriterTraceListener traceToLogFile =
                    new TextWriterTraceListener(filePath);
                Tracer.Listeners.Add(traceToLogFile);
            }

            SetupOutlookRefs();
            UserInputLoop();
        }

        private static void UserInputLoop() {
            Regex rgx = new Regex(@"\b\d*\b");
            bool ApplicationIsExiting = false;
            string command;
            List<string> flags = new List<string>();
            string stringArg = null;
            int intArg = -1;

            bool hasStringArg = false;
            bool hasIntArg = false;
            bool hasFlags = false;

            do {
                Console.Write("->");
                string input = Console.ReadLine();

                if (!input.Contains(' '))
                    command = input;
                else {

                    command = input.Substring(0, input.IndexOf(" "));
                    input = input.Remove(command).Trim(' ');

                    if (input.Contains("\"")) {
                        try {
                            stringArg = input.Substring(input.IndexOf('"') + 1, 
                                input.LastIndexOf('"') - (input.IndexOf('"') + 1));
                            hasStringArg = true;
                            input = input.Remove(stringArg).Trim(' ');
                        }
                        catch (ArgumentOutOfRangeException) {
                            Console.WriteLine("Invalid argument (check quote marks)");
                            continue;
                        }
                    }

                    Match intMatch = rgx.Match(input);
                    if (intMatch.Success) {
                        intArg = int.Parse(intMatch.Value);
                        Tracer.TraceEvent(TraceEventType.Verbose, 0, "Has int");
                        hasIntArg = true;
                    }

                    if (input.Contains('-')) {
                        flags = GetFlags(input);
                        Tracer.TraceEvent(TraceEventType.Verbose, 0, "Has flags");
                        hasFlags = true;
                    }
                }

                switch (command) {
                    case "formatdates":
                        ReformatMagentoDates(stringArg);
                        break;
                    case "enumfolders":
                        EnumerateFolders(flags.Contains("-h"));
                        break;
                    case "setuptest":
                        if (!hasIntArg) {
                            if (intArg < 1 || intArg > 25) {
                                Tracer.TraceEvent(TraceEventType.Error, 0, 
                                    "Maxitems parameter must be non-negative integer greater than 1, less than 25");
                                break;
                            }
                            else
                                if(SetupDefaultTestEnv(intArg, stringArg) == false) {
                                Tracer.TraceEvent(TraceEventType.Error, 0, "Setting up test folder failed");
                            }
                        }
                        else
                            Tracer.TraceEvent(TraceEventType.Error, 0, 
                                "Please specify maximum number of items to duplicate: \"setuptest\" XX");
                        break;
                    case "save":
                        SaveSelected(stringArg);
                        break;
                    case "stoptest":
                        StopTestEnv();
                        break;
                    case "process":
                        if (String.IsNullOrEmpty(stringArg))
                            ProcessSelectedOrder();
                        else
                            ProcessFolder(stringArg);
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
                Tracer.TraceEvent(TraceEventType.Information, 0, "Filename not found at specified filepath!\n");
                return;
            }

            StringReader sr = new StringReader(storelist);

            while (sr.Peek() != -1) {
                string name = sr.ReadLine();
                Franchise store = new Franchise(name, String.Empty, true);
                StoreList.Add(store);
            };
        }
        private static List<string> GetFlags(string input) {
            List<string> flags = new List<string>();
            for(int i = input.IndexOf('-'); i >= 0; i = input.IndexOf('-', i + 1)) {
                flags.Add(String.Join(String.Empty, input[i], input[i + 1]));
            }
            return flags;
        }
    }

    public static class ExtensionMethods {
        public static string Remove(this String s, string substring) {
            if (!s.Contains(substring))
                throw new ArgumentException("String does not contain the provided text");

            return (s.Remove(s.IndexOf(substring), substring.Count()));
        }
    }

    public class NoHeaderTraceListener : TraceListener {
        public override void Write(string message) {
            Trace.Write(message);
        }

        public override void WriteLine(string message) {
            Trace.WriteLine(message);
        }
        public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string message) {
            Trace.WriteLine($"{message}");
        }
        public void TraceEvent(string message) {
            Trace.WriteLine(message);
        }
    }
}

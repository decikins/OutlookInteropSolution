using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using FPBInterop;

namespace FPBInteropConsole {

    static class Program {

    /// PROPERTIES ///
        private static string filePath = $"./log.txt {DateTime.Now:dd.MM.yy HH-mm-ss}.txt";

        private static readonly ConsoleTraceListener ConsoleTracer = new ConsoleTraceListener();
        private readonly static TraceSource Tracer = new TraceSource("FPBInterop.Console");

    /// METHODS ///

        static void Main(string[] args) {
            Tracer.TraceEvent(TraceEventType.Critical, 0, "Starting FPBInterop console app");
            UserInputLoop();
        }

        static void InitShutdown() {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        private static void UserInputLoop() {
            Regex rgx = new Regex(@"\b\d+\b");
            bool ApplicationIsExiting = false;
            string command;
            string input;
            List<string> flags = new List<string>();
            string stringArg = null;
            int intArg = -1;

            do {
                Console.Write("->");
                input = Console.ReadLine();

                if (!input.Contains(' '))
                    command = input;
                else {

                    command = input.Substring(0, input.IndexOf(" "));
                    input = input.Remove(command).Trim(' ');


                    if (input.Contains("\\"))
                        input.Replace('\\', '/');

                    if (input.Contains("/")) {
                        try {
                            stringArg = input.Substring(input.IndexOf('"') + 1, 
                                input.LastIndexOf('"') - (input.IndexOf('"') + 1));
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
                    }

                    if (input.Contains('-')) {
                        flags = GetFlags(input);
                        Tracer.TraceEvent(TraceEventType.Verbose, 0, "Has flags");
                    }
                }

                switch (command) {
                    case "processmagento":
                        OutlookHandler.ProcessFolder("inbox/online orders", flags.Contains("-f"), flags.Contains("-m"));
                        break;
                    case "processfolder":
                        OutlookHandler.ProcessFolder(stringArg, flags.Contains("-f"),flags.Contains("-m"));
                        break;
                    case "processitem":
                        OutlookHandler.ProcessSelectedOrder(flags.Contains("-f"), flags.Contains("-m"));
                        break;                     
                    case "":
                        Console.CursorTop--;
                        Console.WriteLine("");
                        Console.CursorTop--;
                        break;
                    case "saveselected":
                        OutlookHandler.SaveSelectedItemHtml();
                        break;
                    case "setupuserprops":
                        OutlookHandler.SetupUserProperties();
                        break;
                    case "x":
                    case "exit":
                    case "quit":
                    case "close":
                        ApplicationIsExiting = true;
                        break;
                    default:
                        Console.WriteLine("Invalid command");
                        break;

                }
            }
            while (!ApplicationIsExiting);
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

    public class NoHeaderLogListener : TextWriterTraceListener {
        public NoHeaderLogListener(string logFileName) : base(logFileName) {
            base.Writer = new StreamWriter(logFileName, false);
        }
        public override void Write(string message) {
            base.Writer.Write(message);
        }
        public override void WriteLine(string message) {
            base.Writer.WriteLine(message);
        }
        public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string message) {
            base.Writer.WriteLine(message);
        }
    }
}

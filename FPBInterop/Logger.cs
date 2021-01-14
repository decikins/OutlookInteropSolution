using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace FPBInterop {
    public static class Logger {
        public static readonly TraceSource Tracer = new TraceSource("FPBInterop");

        public static void TraceEvent(TraceEventType type, string message, int id = 0)  {
            if (type == TraceEventType.Critical)
                message = "#CRITICAL#: " + message;
            if (type == TraceEventType.Error)
                message = "!Error: " + message;
            Tracer.TraceEvent(type, id, message);
        }

        public static void Return() {
            Tracer.TraceEvent(TraceEventType.Critical,0, "");
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

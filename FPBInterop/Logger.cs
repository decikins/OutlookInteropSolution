using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace FPBInterop {
    public static class Logger {
        public static readonly TraceSource Tracer = new TraceSource("FPBInterop");

        public static void TraceEvent(TraceEventType type, int id, string message)  {
            Tracer.TraceEvent(type, id, message);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FPBInterop {
    [Serializable]
    public class InvalidXPathException : Exception {
        public InvalidXPathException() { }
        public InvalidXPathException(string message) : base(message) { }
        public InvalidXPathException(string message, Exception inner) : base(message, inner) { }
    }

    [Serializable]
    public class InvalidDateFormatException : Exception {
        public InvalidDateFormatException() { }
        public InvalidDateFormatException(string message) : base(message) { }
        public InvalidDateFormatException(string message, Exception inner) : base(message, inner) { }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookInterop {

    [Flags]
    enum OrderMetadata : short {
        IsMagento = 1,
        MustBeProcessed = 2,
        IncorrectForm = 4,
        DateMismatch = 8,

    }
}

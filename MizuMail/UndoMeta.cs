using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MizuMail
{
    public class UndoMeta
    {
        public string OldPath { get; set; }
        public string NewPath { get; set; }
        public string OldFolder { get; set; }
    }
}

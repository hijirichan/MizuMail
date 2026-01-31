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
        public FolderType OldFolder { get; set; }
        public string OldFolderPath { get; set; }
        public string MessageId { get; set; }
    }
}

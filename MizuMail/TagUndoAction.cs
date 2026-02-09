using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MizuMail
{
    public class TagUndoAction
    {
        public Mail Mail { get; set; }
        public List<string> OldLabels { get; set; }
        public List<string> NewLabels { get; set; }

        public void Undo()
        {
            Mail.Labels = new List<string>(OldLabels);
            TagStorage.SaveTags(Mail);   // ★ Message-ID ベースの保存
        }

        public void Redo()
        {
            Mail.Labels = new List<string>(NewLabels);
            TagStorage.SaveTags(Mail);   // ★ Message-ID ベースの保存
        }
    }
}

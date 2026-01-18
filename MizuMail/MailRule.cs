using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MizuMail
{
    public class MailRule
    {
        public string Contains { get; set; }   // 件名に含む
        public string From { get; set; }       // 差出人に含む
        public string MoveTo { get; set; }     // 移動先フォルダ名
    }
}

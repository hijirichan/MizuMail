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
        public bool UseRegex { get; set; }     // 正規表現を使用するかどうか
        public string Label { get; set; }      // 自動付与するラベル
    }
}

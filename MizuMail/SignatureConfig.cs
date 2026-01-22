using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MizuMail
{
    public class SignatureConfig
    {
        public bool Enabled { get; set; } = true;
        public string Signature { get; set; } = "";
    }
}

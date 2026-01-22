using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MizuMail
{
    public partial class FormSignature : Form
    {
        private SignatureConfig signature;

        public FormSignature()
        {
            InitializeComponent();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            SignatureConfig config = new SignatureConfig
            {
                Enabled = checkSignatureEnabled.Checked,
                Signature = textSignatureBody.Text
            };

            FormMain form = new FormMain();
            form.SaveSignature(config);
        }

        private void FormSignature_Load(object sender, EventArgs e)
        {
            FormMain form = new FormMain();
            signature = form.LoadSignature();

            if (signature == null)
            {
                signature = new SignatureConfig
                {
                    Enabled = false,
                    Signature = string.Empty
                };
            }
            checkSignatureEnabled.Checked = signature.Enabled;
            textSignatureBody.Text = signature.Signature;
        }
    }
}

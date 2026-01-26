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
    public partial class FormTagEditor : Form
    {
        public List<string> ResultTags { get; private set; }

        public FormTagEditor(List<string> currentTags)
        {
            InitializeComponent();
            textBoxTags.Text = string.Join(", ", currentTags);
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            ResultTags = textBoxTags.Text
                .Split(',')
                .Select(t => t.Trim())
                .Where(t => t.Length > 0)
                .ToList();

            DialogResult = DialogResult.OK;
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}

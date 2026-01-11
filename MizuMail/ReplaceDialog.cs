using System;
using System.Windows.Forms;

namespace MizuMail
{
    public partial class ReplaceDialog : Form
    {
        public event Action<string, bool, bool> FindNextRequested;
        public event Action<string, string, bool> ReplaceRequested;
        public event Action<string, string, bool> ReplaceAllRequested;

        public ReplaceDialog()
        {
            InitializeComponent();
        }

        private void btnFindNext_Click(object sender, EventArgs e)
        {
            FindNextRequested?.Invoke(
                txtFind.Text,
                chkCase.Checked,
                true
            );
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            ReplaceRequested?.Invoke(
                txtFind.Text,
                txtReplace.Text,
                chkCase.Checked
            );
        }

        private void btnReplaceAll_Click(object sender, EventArgs e)
        {
            ReplaceAllRequested?.Invoke(
                txtFind.Text,
                txtReplace.Text,
                chkCase.Checked
            );
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
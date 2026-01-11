using System;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;

namespace MizuMail
{
    public partial class FindDialog : Form
    {
        public event Action<string, bool, bool> FindNextRequested;

        public FindDialog()
        {
            InitializeComponent();
        }

        private void btnFindNext_Click(object sender, EventArgs e)
        {
            FindNextRequested?.Invoke(
                txtFind.Text,
                chkCase.Checked,
                rdoDown.Checked
            );
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
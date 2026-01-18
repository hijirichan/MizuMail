using System;
using System.Windows.Forms;

namespace MizuMail
{
    public partial class FormRuleEditDialog : Form
    {
        public MailRule Rule { get; private set; }
        private TreeNode rootNode;

        // 新規作成
        public FormRuleEditDialog(TreeNode root)
        {
            InitializeComponent();
            Rule = new MailRule();
            rootNode = root;
        }

        // 編集
        public FormRuleEditDialog(MailRule rule, TreeNode root)
        {
            InitializeComponent();
            Rule = rule;
            rootNode = root;

            textContains.Text = rule.Contains;
            textFrom.Text = rule.From;
            textMoveTo.Text = rule.MoveTo;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            Rule.Contains = textContains.Text.Trim();
            Rule.From = textFrom.Text.Trim();
            Rule.MoveTo = textMoveTo.Text.Trim();

            DialogResult = DialogResult.OK;
        }

        private void buttonSelectFolder_Click(object sender, EventArgs e)
        {
            var dlg = new FormFolderSelectDialog(rootNode);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textMoveTo.Text = dlg.SelectedPath;
            }
        }
    }
}

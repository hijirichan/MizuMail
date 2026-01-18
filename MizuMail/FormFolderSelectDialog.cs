using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MizuMail
{
    public partial class FormFolderSelectDialog : Form
    {
        public string SelectedPath { get; private set; }

        public FormFolderSelectDialog(TreeNode rootNode)
        {
            InitializeComponent();

            // TreeView にコピー
            treeViewFolders.Nodes.Add((TreeNode)rootNode.Clone());
            treeViewFolders.ExpandAll();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (treeViewFolders.SelectedNode == null)
                return;

            // MailFolder を取得
            var folder = treeViewFolders.SelectedNode.Tag as MailFolder;
            if (folder == null)
                return;

            // パスを作成（Inbox/Work/2024）
            SelectedPath = BuildFolderPath(treeViewFolders.SelectedNode);

            DialogResult = DialogResult.OK;
        }

        private string BuildFolderPath(TreeNode node)
        {
            List<string> parts = new List<string>();

            while (node != null)
            {
                var folder = node.Tag as MailFolder;

                if (folder != null)
                    parts.Add(folder.Name);  // 内部名だけを使う

                node = node.Parent;
            }

            parts.Reverse();
            return string.Join("/", parts);
        }
    }
}

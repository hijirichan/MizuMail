using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MizuMail
{
    public partial class FormRuleEditor : Form
    {
        public List<MailRule> Rules { get; private set; }
        private TreeNode rootNode;

        public FormRuleEditor(List<MailRule> rules, TreeNode root)
        {
            InitializeComponent();

            rootNode = root; // ← これを保持する

            Rules = rules.Select(r => new MailRule
            {
                UseRegex = r.UseRegex,
                Contains = r.Contains,
                From = r.From,
                MoveTo = r.MoveTo
            }).ToList();

            LoadList();
        }

        private void LoadList()
        {
            listViewRules.Items.Clear();

            foreach (var r in Rules)
            {
                var item = new ListViewItem();
                item.Checked = r.UseRegex;
                item.SubItems.Add(r.Contains);
                item.SubItems.Add(r.From);
                item.SubItems.Add(r.MoveTo);
                item.Tag = r;

                listViewRules.Items.Add(item);
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            var dlg = new FormRuleEditDialog(rootNode);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Rules.Add(dlg.Rule);
                LoadList();
            }
        }

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            if (listViewRules.SelectedItems.Count == 0)
                return;

            var rule = listViewRules.SelectedItems[0].Tag as MailRule;

            var dlg = new FormRuleEditDialog(rule, rootNode);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                LoadList();
            }
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listViewRules.SelectedItems.Count == 0)
                return;

            var rule = listViewRules.SelectedItems[0].Tag as MailRule;
            Rules.Remove(rule);

            LoadList();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            Rules.Clear();

            foreach (ListViewItem item in listViewRules.Items)
            {
                var rule = new MailRule();

                rule.UseRegex = item.Checked; // ★ ここで保存！

                rule.Contains = item.SubItems[1].Text;
                rule.From = item.SubItems[2].Text;
                rule.MoveTo = item.SubItems[3].Text;

                Rules.Add(rule);
            }

            SaveRules(); // ★ JSON などに保存する処理
            DialogResult = DialogResult.OK;
            Close();
        }

        private void SaveRules()
        {
            string json = JsonConvert.SerializeObject(Rules, Formatting.Indented);
            File.WriteAllText("rules.json", json);
        }
    }
}

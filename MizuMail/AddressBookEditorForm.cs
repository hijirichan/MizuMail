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
    public partial class AddressBookEditorForm : Form
    {
        private AddressBook addressBook;
        public AddressEntry SelectedEntry { get; private set; }
        public bool SelectMode { get; set; } = false;

        public AddressBookEditorForm(AddressBook book)
        {
            InitializeComponent();
            addressBook = book;

            LoadList();
        }

        // ★ ListView にアドレス帳を表示
        private void LoadList()
        {
            listView1.Items.Clear();

            foreach (var entry in addressBook.Entries)
            {
                var item = new ListViewItem(entry.DisplayName);
                item.SubItems.Add(entry.Email);
                item.SubItems.Add(entry.Note);
                item.Tag = entry;

                listView1.Items.Add(item);
            }
        }

        // ★ ListView 選択時に右側の編集欄へ反映
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                ClearEditFields();
                return;
            }

            var entry = (AddressEntry)listView1.SelectedItems[0].Tag;

            textDisplayName.Text = entry.DisplayName;
            textEmail.Text = entry.Email;
            textNote.Text = entry.Note;
        }

        private void ClearEditFields()
        {
            textDisplayName.Text = "";
            textEmail.Text = "";
            textNote.Text = "";
        }

        // ★ 新規追加
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            var entry = new AddressEntry()
            {
                DisplayName = "新しい連絡先",
                Email = "",
                Note = ""
            };

            addressBook.Entries.Add(entry);
            LoadList();
        }

        // ★ 削除
        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
                return;

            var entry = (AddressEntry)listView1.SelectedItems[0].Tag;

            addressBook.Entries.Remove(entry);
            LoadList();
            ClearEditFields();
        }

        // ★ 編集内容を反映
        private void buttonSave_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
                return;

            var entry = (AddressEntry)listView1.SelectedItems[0].Tag;

            entry.DisplayName = textDisplayName.Text.Trim();
            entry.Email = textEmail.Text.Trim();
            entry.Note = textNote.Text.Trim();

            LoadList();
        }

        // ★ OK（保存して閉じる）
        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (SelectMode && listView1.SelectedItems.Count > 0)
            {
                SelectedEntry = (AddressEntry)listView1.SelectedItems[0].Tag;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        // ★ キャンセル（変更破棄）
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

    }
}

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
    public partial class FormListEditor : Form
    {
        private IList<string> list;
        private Action saveCallback;

        public FormListEditor(string title, IList<string> list, Action saveCallback)
        {
            InitializeComponent();
            this.Text = title;
            if (this.Text.Contains("ブラック"))
            {
                labelMailList.Text = "ブロックされたメールアドレス一覧";
            }
            else
            {
                labelMailList.Text = "許可されたメールアドレス一覧";
            }
            this.list = list;
            this.saveCallback = saveCallback;

            RefreshList();
        }

        private void RefreshList()
        {
            listMailAddress.Items.Clear();
            foreach (var addr in list)
            {
                listMailAddress.Items.Add(addr);
            }
        }

        private void buttonAddListMailAddress_Click(object sender, EventArgs e)
        {
            string addr = textAddListMailAddress.Text.Trim();
            if (string.IsNullOrWhiteSpace(addr))
                return;

            if (!list.Contains(addr))
            {
                list.Add(addr);
                saveCallback?.Invoke();
                RefreshList();
            }

            textAddListMailAddress.Clear();
        }

        private void buttonDeleteListMailAddress_Click(object sender, EventArgs e)
        {
            if (listMailAddress.SelectedItem == null)
                return;

            string addr = listMailAddress.SelectedItem.ToString();
            list.Remove(addr);

            saveCallback?.Invoke();
            RefreshList();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

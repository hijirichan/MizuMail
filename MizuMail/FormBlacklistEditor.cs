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
    public partial class FormBlacklistEditor : Form
    {
        private BlockedList blockedList;
        private Action saveCallback;

        public FormBlacklistEditor(BlockedList list, Action saveCallback)
        {
            InitializeComponent();
            this.blockedList = list;
            this.saveCallback = saveCallback;

            RefreshList();
        }

        private void RefreshList()
        {
            listBlackListMail.Items.Clear();
            foreach (var addr in blockedList.blockedEmails)
            {
                listBlackListMail.Items.Add(addr);
            }
        }

        private void buttonAddBlackListMailAddress_Click(object sender, EventArgs e)
        {
            string addr = textAddBlackListMailAddress.Text.Trim();
            if (string.IsNullOrWhiteSpace(addr))
                return;

            if (!blockedList.blockedEmails.Contains(addr))
            {
                blockedList.blockedEmails.Add(addr);
                saveCallback?.Invoke();
                RefreshList();
            }

            textAddBlackListMailAddress.Clear();
        }

        private void buttonDeleteBlackListMailAddress_Click(object sender, EventArgs e)
        {
            if (listBlackListMail.SelectedItem == null)
                return;

            string addr = listBlackListMail.SelectedItem.ToString();
            blockedList.blockedEmails.Remove(addr);

            saveCallback?.Invoke();
            RefreshList();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

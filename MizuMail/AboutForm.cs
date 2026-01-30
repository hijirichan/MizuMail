using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MizuMail
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
        }

        private void linkUrl_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string url = "https://www.angel-teatime.com/"; // 開きたいURL
            try
            {
                Process.Start(url); // URLを指定すると既定のブラウザで開く
            }
            catch (Exception ex)
            {
                MessageBox.Show($"URLを開けませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AboutForm_Load(object sender, EventArgs e)
        {
            labelAppName.Text = Application.ProductName + " Version " + Application.ProductVersion;
            labelCopyright.Text = "Copyright (C) 2026 " + Application.CompanyName;
        }
    }
}

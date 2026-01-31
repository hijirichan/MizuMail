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
    public partial class FormMailAccount : Form
    {
        public FormMailAccount()
        {
            InitializeComponent();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            // アカウント情報を設定する
            Mail.fromName = textFromName.Text;
            Mail.userAddress = textUserAddress.Text;
            Mail.userName = textUserName.Text;
            Mail.smtpServerName = textSmtpServerName.Text;
            Mail.smtpPortNo = int.Parse(textSmtpPortNo.Text);
            Mail.password = textPassword.Text;
            Mail.deleteMail = checkDeleteMail.Checked;
            Mail.alertSound = checkAlertSound.Checked;
            Mail.alertSoundFile = textAlertSoundFile.Text;
            Mail.checkMail = checkReceiveInterval.Checked;
            Mail.checkInterval = (int)updownReceiveInterval.Value;
            if (radioPop3.Checked)
            {
                Mail.receiveMethod_Pop3 = true;
                Mail.popServerName = textPopServerName.Text;
                Mail.popPortNo = int.Parse(textPopServerPortNo.Text);
            }
            else
            {
                Mail.receiveMethod_Pop3 = false;
                Mail.imapServerName = textPopServerName.Text;
                Mail.imapPortNo = int.Parse(textPopServerPortNo.Text);
                Mail.deleteMail = false;
            }
            Mail.useSsl = checkUseSsl.Checked;

            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormMailAccount_Load(object sender, EventArgs e)
        {
            // アカウント情報を設定フォームに設定する
            textFromName.Text = Mail.fromName;
            textUserAddress.Text = Mail.userAddress;
            textUserName.Text = Mail.userName;
            textSmtpServerName.Text = Mail.smtpServerName;
            textSmtpPortNo.Text = Mail.smtpPortNo.ToString();
            if (Mail.receiveMethod_Pop3)
            {
                textPopServerPortNo.Text = Mail.popPortNo.ToString();
                textPopServerName.Text = Mail.popServerName;
            }
            else
            {
                textPopServerPortNo.Text = Mail.imapPortNo.ToString();
                textPopServerName.Text = Mail.imapServerName;
            }
            textPassword.Text = Mail.password;
            checkDeleteMail.Checked = Mail.deleteMail;
            checkAlertSound.Checked = Mail.alertSound;
            textAlertSoundFile.Text = Mail.alertSoundFile;
            checkReceiveInterval.Checked = Mail.checkMail;
            updownReceiveInterval.Value = Mail.checkInterval;
            if (Mail.receiveMethod_Pop3)
            {
                radioPop3.Checked = true;
                checkDeleteMail.Enabled = true;
            }
            else
            {
                radioImap4.Checked = true;
                checkDeleteMail.Checked = false;
                checkDeleteMail.Enabled = false;
            }
            checkUseSsl.Checked = Mail.useSsl;
        }

        private void checkAlertSound_CheckedChanged(object sender, EventArgs e)
        {
            textAlertSoundFile.Enabled = checkAlertSound.Checked;
            buttonAlertSoundFileBrowse.Enabled = checkAlertSound.Checked;
        }

        private void buttonAlertSoundFileBrowse_Click(object sender, EventArgs e)
        {
            if (openAlertSoundFile.ShowDialog() == DialogResult.OK)
            {
                textAlertSoundFile.Text = openAlertSoundFile.FileName;
            }
        }

        private void checkReceiveInterval_CheckedChanged(object sender, EventArgs e)
        {
            updownReceiveInterval.Enabled = checkReceiveInterval.Checked;
            labelReceiveInterval.Enabled = checkReceiveInterval.Checked;
        }

        private void radioImap4_Click(object sender, EventArgs e)
        {
            checkDeleteMail.Checked = false;
            checkDeleteMail.Enabled = false;
        }

        private void radioPop3_CheckedChanged(object sender, EventArgs e)
        {
            checkDeleteMail.Enabled = true;
        }
    }
}

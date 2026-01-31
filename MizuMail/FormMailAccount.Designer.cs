
namespace MizuMail
{
    partial class FormMailAccount
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textUserAddress = new System.Windows.Forms.TextBox();
            this.textUserName = new System.Windows.Forms.TextBox();
            this.textSmtpServerName = new System.Windows.Forms.TextBox();
            this.textSmtpPortNo = new System.Windows.Forms.TextBox();
            this.textPopServerName = new System.Windows.Forms.TextBox();
            this.textPopServerPortNo = new System.Windows.Forms.TextBox();
            this.textPassword = new System.Windows.Forms.TextBox();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.textFromName = new System.Windows.Forms.TextBox();
            this.checkDeleteMail = new System.Windows.Forms.CheckBox();
            this.checkAlertSound = new System.Windows.Forms.CheckBox();
            this.textAlertSoundFile = new System.Windows.Forms.TextBox();
            this.buttonAlertSoundFileBrowse = new System.Windows.Forms.Button();
            this.openAlertSoundFile = new System.Windows.Forms.OpenFileDialog();
            this.checkReceiveInterval = new System.Windows.Forms.CheckBox();
            this.updownReceiveInterval = new System.Windows.Forms.NumericUpDown();
            this.labelReceiveInterval = new System.Windows.Forms.Label();
            this.radioPop3 = new System.Windows.Forms.RadioButton();
            this.radioImap4 = new System.Windows.Forms.RadioButton();
            this.label9 = new System.Windows.Forms.Label();
            this.checkUseSsl = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.updownReceiveInterval)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(133, 15);
            this.label1.TabIndex = 3;
            this.label1.Text = "ユーザのメールアドレス";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 15);
            this.label2.TabIndex = 5;
            this.label2.Text = "ユーザ名";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 114);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 15);
            this.label3.TabIndex = 7;
            this.label3.Text = "送信サーバ名";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 142);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(151, 15);
            this.label4.TabIndex = 9;
            this.label4.Text = "送信サーバのポート番号";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 174);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 15);
            this.label5.TabIndex = 11;
            this.label5.Text = "受信サーバ名";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 204);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(151, 15);
            this.label6.TabIndex = 13;
            this.label6.Text = "受信サーバのポート番号";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 232);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 15);
            this.label7.TabIndex = 15;
            this.label7.Text = "パスワード";
            // 
            // textUserAddress
            // 
            this.textUserAddress.Location = new System.Drawing.Point(183, 48);
            this.textUserAddress.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textUserAddress.Name = "textUserAddress";
            this.textUserAddress.Size = new System.Drawing.Size(430, 22);
            this.textUserAddress.TabIndex = 4;
            // 
            // textUserName
            // 
            this.textUserName.Location = new System.Drawing.Point(183, 79);
            this.textUserName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textUserName.Name = "textUserName";
            this.textUserName.Size = new System.Drawing.Size(430, 22);
            this.textUserName.TabIndex = 6;
            // 
            // textSmtpServerName
            // 
            this.textSmtpServerName.Location = new System.Drawing.Point(183, 111);
            this.textSmtpServerName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textSmtpServerName.Name = "textSmtpServerName";
            this.textSmtpServerName.Size = new System.Drawing.Size(430, 22);
            this.textSmtpServerName.TabIndex = 8;
            // 
            // textSmtpPortNo
            // 
            this.textSmtpPortNo.Location = new System.Drawing.Point(183, 139);
            this.textSmtpPortNo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textSmtpPortNo.Name = "textSmtpPortNo";
            this.textSmtpPortNo.Size = new System.Drawing.Size(430, 22);
            this.textSmtpPortNo.TabIndex = 10;
            this.textSmtpPortNo.Text = "25";
            // 
            // textPopServerName
            // 
            this.textPopServerName.Location = new System.Drawing.Point(183, 170);
            this.textPopServerName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textPopServerName.Name = "textPopServerName";
            this.textPopServerName.Size = new System.Drawing.Size(430, 22);
            this.textPopServerName.TabIndex = 12;
            // 
            // textPopServerPortNo
            // 
            this.textPopServerPortNo.Location = new System.Drawing.Point(183, 199);
            this.textPopServerPortNo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textPopServerPortNo.Name = "textPopServerPortNo";
            this.textPopServerPortNo.Size = new System.Drawing.Size(430, 22);
            this.textPopServerPortNo.TabIndex = 14;
            this.textPopServerPortNo.Text = "110";
            // 
            // textPassword
            // 
            this.textPassword.Location = new System.Drawing.Point(183, 229);
            this.textPassword.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textPassword.Name = "textPassword";
            this.textPassword.PasswordChar = '●';
            this.textPassword.Size = new System.Drawing.Size(430, 22);
            this.textPassword.TabIndex = 16;
            // 
            // buttonOK
            // 
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.Location = new System.Drawing.Point(403, 414);
            this.buttonOK.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(97, 34);
            this.buttonOK.TabIndex = 29;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(506, 414);
            this.buttonCancel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(109, 34);
            this.buttonCancel.TabIndex = 30;
            this.buttonCancel.Text = "キャンセル";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(16, 22);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(94, 15);
            this.label8.TabIndex = 1;
            this.label8.Text = "差出人の名前";
            // 
            // textFromName
            // 
            this.textFromName.Location = new System.Drawing.Point(183, 14);
            this.textFromName.Margin = new System.Windows.Forms.Padding(4);
            this.textFromName.Name = "textFromName";
            this.textFromName.Size = new System.Drawing.Size(430, 22);
            this.textFromName.TabIndex = 2;
            // 
            // checkDeleteMail
            // 
            this.checkDeleteMail.AutoSize = true;
            this.checkDeleteMail.Location = new System.Drawing.Point(183, 319);
            this.checkDeleteMail.Margin = new System.Windows.Forms.Padding(4);
            this.checkDeleteMail.Name = "checkDeleteMail";
            this.checkDeleteMail.Size = new System.Drawing.Size(219, 19);
            this.checkDeleteMail.TabIndex = 21;
            this.checkDeleteMail.Text = "メール受信時にメールを削除する";
            this.checkDeleteMail.UseVisualStyleBackColor = true;
            // 
            // checkAlertSound
            // 
            this.checkAlertSound.AutoSize = true;
            this.checkAlertSound.Location = new System.Drawing.Point(183, 346);
            this.checkAlertSound.Margin = new System.Windows.Forms.Padding(4);
            this.checkAlertSound.Name = "checkAlertSound";
            this.checkAlertSound.Size = new System.Drawing.Size(139, 19);
            this.checkAlertSound.TabIndex = 22;
            this.checkAlertSound.Text = "通知音を再生する";
            this.checkAlertSound.UseVisualStyleBackColor = true;
            this.checkAlertSound.CheckedChanged += new System.EventHandler(this.checkAlertSound_CheckedChanged);
            // 
            // textAlertSoundFile
            // 
            this.textAlertSoundFile.Enabled = false;
            this.textAlertSoundFile.Location = new System.Drawing.Point(329, 344);
            this.textAlertSoundFile.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textAlertSoundFile.Name = "textAlertSoundFile";
            this.textAlertSoundFile.Size = new System.Drawing.Size(244, 22);
            this.textAlertSoundFile.TabIndex = 23;
            // 
            // buttonAlertSoundFileBrowse
            // 
            this.buttonAlertSoundFileBrowse.Enabled = false;
            this.buttonAlertSoundFileBrowse.Location = new System.Drawing.Point(579, 341);
            this.buttonAlertSoundFileBrowse.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonAlertSoundFileBrowse.Name = "buttonAlertSoundFileBrowse";
            this.buttonAlertSoundFileBrowse.Size = new System.Drawing.Size(34, 27);
            this.buttonAlertSoundFileBrowse.TabIndex = 24;
            this.buttonAlertSoundFileBrowse.Text = "...";
            this.buttonAlertSoundFileBrowse.UseVisualStyleBackColor = true;
            this.buttonAlertSoundFileBrowse.Click += new System.EventHandler(this.buttonAlertSoundFileBrowse_Click);
            // 
            // openAlertSoundFile
            // 
            this.openAlertSoundFile.DefaultExt = "*.wav";
            this.openAlertSoundFile.Filter = "WAVファイル(*.wav)|*wav|すべてのファイル|*.*";
            this.openAlertSoundFile.Title = "通知音の選択";
            // 
            // checkReceiveInterval
            // 
            this.checkReceiveInterval.AutoSize = true;
            this.checkReceiveInterval.Location = new System.Drawing.Point(183, 373);
            this.checkReceiveInterval.Margin = new System.Windows.Forms.Padding(4);
            this.checkReceiveInterval.Name = "checkReceiveInterval";
            this.checkReceiveInterval.Size = new System.Drawing.Size(166, 19);
            this.checkReceiveInterval.TabIndex = 25;
            this.checkReceiveInterval.Text = "自動受信を有効にする";
            this.checkReceiveInterval.UseVisualStyleBackColor = true;
            this.checkReceiveInterval.CheckedChanged += new System.EventHandler(this.checkReceiveInterval_CheckedChanged);
            // 
            // updownReceiveInterval
            // 
            this.updownReceiveInterval.Enabled = false;
            this.updownReceiveInterval.Location = new System.Drawing.Point(356, 371);
            this.updownReceiveInterval.Maximum = new decimal(new int[] {
            120,
            0,
            0,
            0});
            this.updownReceiveInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.updownReceiveInterval.Name = "updownReceiveInterval";
            this.updownReceiveInterval.Size = new System.Drawing.Size(59, 22);
            this.updownReceiveInterval.TabIndex = 26;
            this.updownReceiveInterval.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // labelReceiveInterval
            // 
            this.labelReceiveInterval.AutoSize = true;
            this.labelReceiveInterval.Enabled = false;
            this.labelReceiveInterval.Location = new System.Drawing.Point(421, 375);
            this.labelReceiveInterval.Name = "labelReceiveInterval";
            this.labelReceiveInterval.Size = new System.Drawing.Size(118, 15);
            this.labelReceiveInterval.TabIndex = 27;
            this.labelReceiveInterval.Text = "分間隔に受信する";
            // 
            // radioPop3
            // 
            this.radioPop3.AutoSize = true;
            this.radioPop3.Checked = true;
            this.radioPop3.Location = new System.Drawing.Point(183, 266);
            this.radioPop3.Name = "radioPop3";
            this.radioPop3.Size = new System.Drawing.Size(65, 19);
            this.radioPop3.TabIndex = 18;
            this.radioPop3.TabStop = true;
            this.radioPop3.Text = "POP3";
            this.radioPop3.UseVisualStyleBackColor = true;
            this.radioPop3.CheckedChanged += new System.EventHandler(this.radioPop3_CheckedChanged);
            // 
            // radioImap4
            // 
            this.radioImap4.AutoSize = true;
            this.radioImap4.Location = new System.Drawing.Point(254, 266);
            this.radioImap4.Name = "radioImap4";
            this.radioImap4.Size = new System.Drawing.Size(69, 19);
            this.radioImap4.TabIndex = 19;
            this.radioImap4.Text = "IMAP4";
            this.radioImap4.UseVisualStyleBackColor = true;
            this.radioImap4.Click += new System.EventHandler(this.radioImap4_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(16, 266);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(67, 15);
            this.label9.TabIndex = 17;
            this.label9.Text = "受信方式";
            // 
            // checkUseSsl
            // 
            this.checkUseSsl.AutoSize = true;
            this.checkUseSsl.Location = new System.Drawing.Point(183, 292);
            this.checkUseSsl.Margin = new System.Windows.Forms.Padding(4);
            this.checkUseSsl.Name = "checkUseSsl";
            this.checkUseSsl.Size = new System.Drawing.Size(162, 19);
            this.checkUseSsl.TabIndex = 20;
            this.checkUseSsl.Text = "接続にSSLを使用する";
            this.checkUseSsl.UseVisualStyleBackColor = true;
            // 
            // FormMailAccount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(627, 461);
            this.Controls.Add(this.checkUseSsl);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.radioImap4);
            this.Controls.Add(this.radioPop3);
            this.Controls.Add(this.labelReceiveInterval);
            this.Controls.Add(this.updownReceiveInterval);
            this.Controls.Add(this.checkReceiveInterval);
            this.Controls.Add(this.buttonAlertSoundFileBrowse);
            this.Controls.Add(this.textAlertSoundFile);
            this.Controls.Add(this.checkAlertSound);
            this.Controls.Add(this.checkDeleteMail);
            this.Controls.Add(this.textFromName);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.textPassword);
            this.Controls.Add(this.textPopServerPortNo);
            this.Controls.Add(this.textPopServerName);
            this.Controls.Add(this.textSmtpPortNo);
            this.Controls.Add(this.textSmtpServerName);
            this.Controls.Add(this.textUserName);
            this.Controls.Add(this.textUserAddress);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormMailAccount";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "アカウント設定";
            this.Load += new System.EventHandler(this.FormMailAccount_Load);
            ((System.ComponentModel.ISupportInitialize)(this.updownReceiveInterval)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textUserAddress;
        private System.Windows.Forms.TextBox textUserName;
        private System.Windows.Forms.TextBox textSmtpServerName;
        private System.Windows.Forms.TextBox textSmtpPortNo;
        private System.Windows.Forms.TextBox textPopServerName;
        private System.Windows.Forms.TextBox textPopServerPortNo;
        private System.Windows.Forms.TextBox textPassword;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textFromName;
        private System.Windows.Forms.CheckBox checkDeleteMail;
        private System.Windows.Forms.CheckBox checkAlertSound;
        private System.Windows.Forms.TextBox textAlertSoundFile;
        private System.Windows.Forms.Button buttonAlertSoundFileBrowse;
        private System.Windows.Forms.OpenFileDialog openAlertSoundFile;
        private System.Windows.Forms.CheckBox checkReceiveInterval;
        private System.Windows.Forms.NumericUpDown updownReceiveInterval;
        private System.Windows.Forms.Label labelReceiveInterval;
        private System.Windows.Forms.RadioButton radioPop3;
        private System.Windows.Forms.RadioButton radioImap4;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox checkUseSsl;
    }
}
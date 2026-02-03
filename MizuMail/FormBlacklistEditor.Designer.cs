namespace MizuMail
{
    partial class FormBlacklistEditor
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
            this.listBlackListMail = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textAddBlackListMailAddress = new System.Windows.Forms.TextBox();
            this.buttonAddBlackListMailAddress = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonDeleteBlackListMailAddress = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(214, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "ブロックされているメールアドレス一覧";
            // 
            // listBlackListMail
            // 
            this.listBlackListMail.FormattingEnabled = true;
            this.listBlackListMail.ItemHeight = 15;
            this.listBlackListMail.Location = new System.Drawing.Point(15, 27);
            this.listBlackListMail.Name = "listBlackListMail";
            this.listBlackListMail.Size = new System.Drawing.Size(534, 199);
            this.listBlackListMail.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 239);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "追加:";
            // 
            // textAddBlackListMailAddress
            // 
            this.textAddBlackListMailAddress.Location = new System.Drawing.Point(58, 236);
            this.textAddBlackListMailAddress.Name = "textAddBlackListMailAddress";
            this.textAddBlackListMailAddress.Size = new System.Drawing.Size(393, 22);
            this.textAddBlackListMailAddress.TabIndex = 3;
            // 
            // buttonAddBlackListMailAddress
            // 
            this.buttonAddBlackListMailAddress.Location = new System.Drawing.Point(457, 232);
            this.buttonAddBlackListMailAddress.Name = "buttonAddBlackListMailAddress";
            this.buttonAddBlackListMailAddress.Size = new System.Drawing.Size(92, 33);
            this.buttonAddBlackListMailAddress.TabIndex = 4;
            this.buttonAddBlackListMailAddress.Text = "追加";
            this.buttonAddBlackListMailAddress.UseVisualStyleBackColor = true;
            this.buttonAddBlackListMailAddress.Click += new System.EventHandler(this.buttonAddBlackListMailAddress_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 280);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 15);
            this.label3.TabIndex = 5;
            this.label3.Text = "選択項目を削除";
            // 
            // buttonDeleteBlackListMailAddress
            // 
            this.buttonDeleteBlackListMailAddress.Location = new System.Drawing.Point(126, 272);
            this.buttonDeleteBlackListMailAddress.Name = "buttonDeleteBlackListMailAddress";
            this.buttonDeleteBlackListMailAddress.Size = new System.Drawing.Size(92, 30);
            this.buttonDeleteBlackListMailAddress.TabIndex = 6;
            this.buttonDeleteBlackListMailAddress.Text = "削除";
            this.buttonDeleteBlackListMailAddress.UseVisualStyleBackColor = true;
            this.buttonDeleteBlackListMailAddress.Click += new System.EventHandler(this.buttonDeleteBlackListMailAddress_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(459, 272);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(90, 31);
            this.buttonClose.TabIndex = 7;
            this.buttonClose.Text = "閉じる";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // FormBlacklistEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(561, 315);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonDeleteBlackListMailAddress);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.buttonAddBlackListMailAddress);
            this.Controls.Add(this.textAddBlackListMailAddress);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.listBlackListMail);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormBlacklistEditor";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "ブラックリストメールアドレスの編集";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textAddBlackListMailAddress;
        private System.Windows.Forms.Button buttonAddBlackListMailAddress;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonDeleteBlackListMailAddress;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.ListBox listBlackListMail;
    }
}
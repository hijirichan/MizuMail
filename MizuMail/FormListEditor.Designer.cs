namespace MizuMail
{
    partial class FormListEditor
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
            this.labelMailList = new System.Windows.Forms.Label();
            this.listMailAddress = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textAddListMailAddress = new System.Windows.Forms.TextBox();
            this.buttonAddListMailAddress = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonDeleteBlackListMailAddress = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelMailList
            // 
            this.labelMailList.AutoSize = true;
            this.labelMailList.Location = new System.Drawing.Point(12, 9);
            this.labelMailList.Name = "labelMailList";
            this.labelMailList.Size = new System.Drawing.Size(115, 15);
            this.labelMailList.TabIndex = 0;
            this.labelMailList.Text = "メールアドレス一覧";
            // 
            // listMailAddress
            // 
            this.listMailAddress.FormattingEnabled = true;
            this.listMailAddress.ItemHeight = 15;
            this.listMailAddress.Location = new System.Drawing.Point(15, 27);
            this.listMailAddress.Name = "listMailAddress";
            this.listMailAddress.Size = new System.Drawing.Size(534, 199);
            this.listMailAddress.TabIndex = 1;
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
            // textAddListMailAddress
            // 
            this.textAddListMailAddress.Location = new System.Drawing.Point(58, 236);
            this.textAddListMailAddress.Name = "textAddListMailAddress";
            this.textAddListMailAddress.Size = new System.Drawing.Size(393, 22);
            this.textAddListMailAddress.TabIndex = 3;
            // 
            // buttonAddListMailAddress
            // 
            this.buttonAddListMailAddress.Location = new System.Drawing.Point(457, 232);
            this.buttonAddListMailAddress.Name = "buttonAddListMailAddress";
            this.buttonAddListMailAddress.Size = new System.Drawing.Size(92, 33);
            this.buttonAddListMailAddress.TabIndex = 4;
            this.buttonAddListMailAddress.Text = "追加";
            this.buttonAddListMailAddress.UseVisualStyleBackColor = true;
            this.buttonAddListMailAddress.Click += new System.EventHandler(this.buttonAddListMailAddress_Click);
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
            this.buttonDeleteBlackListMailAddress.Click += new System.EventHandler(this.buttonDeleteListMailAddress_Click);
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
            // FormListEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(561, 315);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonDeleteBlackListMailAddress);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.buttonAddListMailAddress);
            this.Controls.Add(this.textAddListMailAddress);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.listMailAddress);
            this.Controls.Add(this.labelMailList);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormListEditor";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "メールリスト編集";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelMailList;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textAddListMailAddress;
        private System.Windows.Forms.Button buttonAddListMailAddress;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonDeleteBlackListMailAddress;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.ListBox listMailAddress;
    }
}
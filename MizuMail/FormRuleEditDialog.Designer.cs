namespace MizuMail
{
    partial class FormRuleEditDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textContains;
        private System.Windows.Forms.TextBox textFrom;
        private System.Windows.Forms.TextBox textMoveTo;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonSelectFolder;

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
            this.textContains = new System.Windows.Forms.TextBox();
            this.textFrom = new System.Windows.Forms.TextBox();
            this.textMoveTo = new System.Windows.Forms.TextBox();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonSelectFolder = new System.Windows.Forms.Button();
            this.checkUseRegex = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 15);
            this.label1.TabIndex = 7;
            this.label1.Text = "件名に含む";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 15);
            this.label2.TabIndex = 5;
            this.label2.Text = "差出人に含む";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(95, 15);
            this.label3.TabIndex = 3;
            this.label3.Text = "移動先フォルダ";
            // 
            // textContains
            // 
            this.textContains.Location = new System.Drawing.Point(137, 12);
            this.textContains.Name = "textContains";
            this.textContains.Size = new System.Drawing.Size(366, 22);
            this.textContains.TabIndex = 6;
            // 
            // textFrom
            // 
            this.textFrom.Location = new System.Drawing.Point(137, 47);
            this.textFrom.Name = "textFrom";
            this.textFrom.Size = new System.Drawing.Size(366, 22);
            this.textFrom.TabIndex = 4;
            // 
            // textMoveTo
            // 
            this.textMoveTo.Location = new System.Drawing.Point(137, 82);
            this.textMoveTo.Name = "textMoveTo";
            this.textMoveTo.Size = new System.Drawing.Size(285, 22);
            this.textMoveTo.TabIndex = 2;
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(325, 145);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(86, 30);
            this.buttonOK.TabIndex = 1;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(417, 145);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(86, 30);
            this.buttonCancel.TabIndex = 0;
            this.buttonCancel.Text = "キャンセル";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // buttonSelectFolder
            // 
            this.buttonSelectFolder.Location = new System.Drawing.Point(428, 82);
            this.buttonSelectFolder.Name = "buttonSelectFolder";
            this.buttonSelectFolder.Size = new System.Drawing.Size(75, 23);
            this.buttonSelectFolder.TabIndex = 0;
            this.buttonSelectFolder.Text = "選択...";
            this.buttonSelectFolder.Click += new System.EventHandler(this.buttonSelectFolder_Click);
            // 
            // checkUseRegex
            // 
            this.checkUseRegex.AutoSize = true;
            this.checkUseRegex.Location = new System.Drawing.Point(137, 120);
            this.checkUseRegex.Name = "checkUseRegex";
            this.checkUseRegex.Size = new System.Drawing.Size(154, 19);
            this.checkUseRegex.TabIndex = 8;
            this.checkUseRegex.Text = "正規表現を使用する";
            this.checkUseRegex.UseVisualStyleBackColor = true;
            // 
            // FormRuleEditDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 184);
            this.Controls.Add(this.checkUseRegex);
            this.Controls.Add(this.buttonSelectFolder);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.textMoveTo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textFrom);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textContains);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormRuleEditDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "ルール編集";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkUseRegex;
    }
}
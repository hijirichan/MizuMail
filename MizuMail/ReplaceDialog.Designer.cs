using System.Windows.Forms;

namespace MizuMail
{
    partial class ReplaceDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private Label labelFind;
        private Label labelReplace;
        private TextBox txtFind;
        private TextBox txtReplace;
        private CheckBox chkCase;
        private Button btnFindNext;
        private Button btnReplace;
        private Button btnReplaceAll;
        private Button btnCancel;

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
            this.labelFind = new System.Windows.Forms.Label();
            this.labelReplace = new System.Windows.Forms.Label();
            this.txtFind = new System.Windows.Forms.TextBox();
            this.txtReplace = new System.Windows.Forms.TextBox();
            this.chkCase = new System.Windows.Forms.CheckBox();
            this.btnFindNext = new System.Windows.Forms.Button();
            this.btnReplace = new System.Windows.Forms.Button();
            this.btnReplaceAll = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelFind
            // 
            this.labelFind.AutoSize = true;
            this.labelFind.Location = new System.Drawing.Point(12, 15);
            this.labelFind.Name = "labelFind";
            this.labelFind.Size = new System.Drawing.Size(109, 15);
            this.labelFind.TabIndex = 0;
            this.labelFind.Text = "検索する文字列:";
            // 
            // labelReplace
            // 
            this.labelReplace.AutoSize = true;
            this.labelReplace.Location = new System.Drawing.Point(12, 50);
            this.labelReplace.Name = "labelReplace";
            this.labelReplace.Size = new System.Drawing.Size(112, 15);
            this.labelReplace.TabIndex = 2;
            this.labelReplace.Text = "置換後の文字列:";
            // 
            // txtFind
            // 
            this.txtFind.Location = new System.Drawing.Point(140, 12);
            this.txtFind.Name = "txtFind";
            this.txtFind.Size = new System.Drawing.Size(200, 22);
            this.txtFind.TabIndex = 1;
            // 
            // txtReplace
            // 
            this.txtReplace.Location = new System.Drawing.Point(140, 47);
            this.txtReplace.Name = "txtReplace";
            this.txtReplace.Size = new System.Drawing.Size(200, 22);
            this.txtReplace.TabIndex = 3;
            // 
            // chkCase
            // 
            this.chkCase.Location = new System.Drawing.Point(15, 85);
            this.chkCase.Name = "chkCase";
            this.chkCase.Size = new System.Drawing.Size(255, 24);
            this.chkCase.TabIndex = 4;
            this.chkCase.Text = "大文字と小文字を区別する";
            // 
            // btnFindNext
            // 
            this.btnFindNext.Location = new System.Drawing.Point(360, 10);
            this.btnFindNext.Name = "btnFindNext";
            this.btnFindNext.Size = new System.Drawing.Size(92, 23);
            this.btnFindNext.TabIndex = 5;
            this.btnFindNext.Text = "次を検索";
            this.btnFindNext.Click += new System.EventHandler(this.btnFindNext_Click);
            // 
            // btnReplace
            // 
            this.btnReplace.Location = new System.Drawing.Point(360, 45);
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Size = new System.Drawing.Size(92, 23);
            this.btnReplace.TabIndex = 6;
            this.btnReplace.Text = "置換";
            this.btnReplace.Click += new System.EventHandler(this.btnReplace_Click);
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.Location = new System.Drawing.Point(360, 80);
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.Size = new System.Drawing.Size(92, 23);
            this.btnReplaceAll.TabIndex = 7;
            this.btnReplaceAll.Text = "すべて置換";
            this.btnReplaceAll.Click += new System.EventHandler(this.btnReplaceAll_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(360, 115);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(92, 23);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // ReplaceDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(469, 160);
            this.Controls.Add(this.labelFind);
            this.Controls.Add(this.txtFind);
            this.Controls.Add(this.labelReplace);
            this.Controls.Add(this.txtReplace);
            this.Controls.Add(this.chkCase);
            this.Controls.Add(this.btnFindNext);
            this.Controls.Add(this.btnReplace);
            this.Controls.Add(this.btnReplaceAll);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReplaceDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "置換";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}
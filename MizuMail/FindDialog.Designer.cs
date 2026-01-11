using System.Windows.Forms;

namespace MizuMail
{
    partial class FindDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private Label label1;
        private TextBox txtFind;
        private CheckBox chkCase;
        private GroupBox groupDirection;
        private RadioButton rdoDown;
        private RadioButton rdoUp;
        private Button btnFindNext;
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtFind = new System.Windows.Forms.TextBox();
            this.chkCase = new System.Windows.Forms.CheckBox();
            this.groupDirection = new System.Windows.Forms.GroupBox();
            this.rdoDown = new System.Windows.Forms.RadioButton();
            this.rdoUp = new System.Windows.Forms.RadioButton();
            this.btnFindNext = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupDirection.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "検索する文字列:";
            // 
            // txtFind
            // 
            this.txtFind.Location = new System.Drawing.Point(127, 11);
            this.txtFind.Name = "txtFind";
            this.txtFind.Size = new System.Drawing.Size(200, 22);
            this.txtFind.TabIndex = 1;
            // 
            // chkCase
            // 
            this.chkCase.Location = new System.Drawing.Point(15, 50);
            this.chkCase.Name = "chkCase";
            this.chkCase.Size = new System.Drawing.Size(276, 24);
            this.chkCase.TabIndex = 2;
            this.chkCase.Text = "大文字と小文字を区別する";
            // 
            // groupDirection
            // 
            this.groupDirection.Controls.Add(this.rdoDown);
            this.groupDirection.Controls.Add(this.rdoUp);
            this.groupDirection.Location = new System.Drawing.Point(15, 80);
            this.groupDirection.Name = "groupDirection";
            this.groupDirection.Size = new System.Drawing.Size(200, 70);
            this.groupDirection.TabIndex = 3;
            this.groupDirection.TabStop = false;
            this.groupDirection.Text = "検索方向";
            // 
            // rdoDown
            // 
            this.rdoDown.Checked = true;
            this.rdoDown.Location = new System.Drawing.Point(10, 20);
            this.rdoDown.Name = "rdoDown";
            this.rdoDown.Size = new System.Drawing.Size(104, 24);
            this.rdoDown.TabIndex = 0;
            this.rdoDown.TabStop = true;
            this.rdoDown.Text = "下へ";
            // 
            // rdoUp
            // 
            this.rdoUp.Location = new System.Drawing.Point(10, 40);
            this.rdoUp.Name = "rdoUp";
            this.rdoUp.Size = new System.Drawing.Size(104, 24);
            this.rdoUp.TabIndex = 1;
            this.rdoUp.Text = "上へ";
            // 
            // btnFindNext
            // 
            this.btnFindNext.Location = new System.Drawing.Point(339, 9);
            this.btnFindNext.Name = "btnFindNext";
            this.btnFindNext.Size = new System.Drawing.Size(91, 24);
            this.btnFindNext.TabIndex = 4;
            this.btnFindNext.Text = "次を検索";
            this.btnFindNext.Click += new System.EventHandler(this.btnFindNext_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(339, 38);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(91, 25);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // FindDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(448, 160);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtFind);
            this.Controls.Add(this.chkCase);
            this.Controls.Add(this.groupDirection);
            this.Controls.Add(this.btnFindNext);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FindDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "検索";
            this.groupDirection.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}
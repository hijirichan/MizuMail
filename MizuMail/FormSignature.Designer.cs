namespace MizuMail
{
    partial class FormSignature
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
            this.checkSignatureEnabled = new System.Windows.Forms.CheckBox();
            this.textSignatureBody = new System.Windows.Forms.TextBox();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCencel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // checkSignatureEnabled
            // 
            this.checkSignatureEnabled.AutoSize = true;
            this.checkSignatureEnabled.Location = new System.Drawing.Point(12, 12);
            this.checkSignatureEnabled.Name = "checkSignatureEnabled";
            this.checkSignatureEnabled.Size = new System.Drawing.Size(124, 19);
            this.checkSignatureEnabled.TabIndex = 0;
            this.checkSignatureEnabled.Text = "署名を使用する";
            this.checkSignatureEnabled.UseVisualStyleBackColor = true;
            // 
            // textSignatureBody
            // 
            this.textSignatureBody.Font = new System.Drawing.Font("MS UI Gothic", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textSignatureBody.Location = new System.Drawing.Point(12, 37);
            this.textSignatureBody.Multiline = true;
            this.textSignatureBody.Name = "textSignatureBody";
            this.textSignatureBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textSignatureBody.Size = new System.Drawing.Size(548, 305);
            this.textSignatureBody.TabIndex = 1;
            // 
            // buttonOK
            // 
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.Location = new System.Drawing.Point(377, 348);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 28);
            this.buttonOK.TabIndex = 2;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCencel
            // 
            this.buttonCencel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCencel.Location = new System.Drawing.Point(458, 348);
            this.buttonCencel.Name = "buttonCencel";
            this.buttonCencel.Size = new System.Drawing.Size(100, 28);
            this.buttonCencel.TabIndex = 3;
            this.buttonCencel.Text = "キャンセル";
            this.buttonCencel.UseVisualStyleBackColor = true;
            // 
            // FormSignature
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 382);
            this.Controls.Add(this.buttonCencel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.textSignatureBody);
            this.Controls.Add(this.checkSignatureEnabled);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormSignature";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "署名の設定";
            this.Load += new System.EventHandler(this.FormSignature_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkSignatureEnabled;
        private System.Windows.Forms.TextBox textSignatureBody;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCencel;
    }
}
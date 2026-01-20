namespace MizuMail
{
    partial class FormMailCreate
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMailCreate));
            this.label1 = new System.Windows.Forms.Label();
            this.textMailTo = new System.Windows.Forms.TextBox();
            this.contextAddress = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuSelectFromAddressBook = new System.Windows.Forms.ToolStripMenuItem();
            this.textMailSubject = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonSend = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.label4 = new System.Windows.Forms.Label();
            this.textMailBcc = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textMailCc = new System.Windows.Forms.TextBox();
            this.textMailBody = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSend = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.menuClose = new System.Windows.Forms.ToolStripMenuItem();
            this.menuEdit = new System.Windows.Forms.ToolStripMenuItem();
            this.menuUndo = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.menuCut = new System.Windows.Forms.ToolStripMenuItem();
            this.menuCopy = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPaste = new System.Windows.Forms.ToolStripMenuItem();
            this.menuDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripSeparator();
            this.menuFind = new System.Windows.Forms.ToolStripMenuItem();
            this.menuReplace = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripSeparator();
            this.menuAllSelect = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButtonSend = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolAddAttachment = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButtonCut = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonCopy = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonPaste = new System.Windows.Forms.ToolStripButton();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.labelMessage = new System.Windows.Forms.ToolStripStatusLabel();
            this.buttonAttachList = new System.Windows.Forms.ToolStripDropDownButton();
            this.contextAddress.SuspendLayout();
            this.toolStripContainer1.ContentPanel.SuspendLayout();
            this.toolStripContainer1.TopToolStripPanel.SuspendLayout();
            this.toolStripContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "宛先";
            // 
            // textMailTo
            // 
            this.textMailTo.ContextMenuStrip = this.contextAddress;
            this.textMailTo.Location = new System.Drawing.Point(56, 11);
            this.textMailTo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textMailTo.Name = "textMailTo";
            this.textMailTo.Size = new System.Drawing.Size(1104, 22);
            this.textMailTo.TabIndex = 1;
            this.textMailTo.Leave += new System.EventHandler(this.textMailTo_Leave);
            // 
            // contextAddress
            // 
            this.contextAddress.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextAddress.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuSelectFromAddressBook});
            this.contextAddress.Name = "contextAddress";
            this.contextAddress.Size = new System.Drawing.Size(212, 28);
            // 
            // menuSelectFromAddressBook
            // 
            this.menuSelectFromAddressBook.Name = "menuSelectFromAddressBook";
            this.menuSelectFromAddressBook.Size = new System.Drawing.Size(211, 24);
            this.menuSelectFromAddressBook.Text = "アドレス帳から選択(&A)";
            this.menuSelectFromAddressBook.Click += new System.EventHandler(this.menuSelectFromAddressBook_Click);
            // 
            // textMailSubject
            // 
            this.textMailSubject.Location = new System.Drawing.Point(55, 92);
            this.textMailSubject.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textMailSubject.Name = "textMailSubject";
            this.textMailSubject.Size = new System.Drawing.Size(1105, 22);
            this.textMailSubject.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 95);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 15);
            this.label2.TabIndex = 6;
            this.label2.Text = "件名";
            // 
            // buttonSend
            // 
            this.buttonSend.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonSend.Location = new System.Drawing.Point(869, 345);
            this.buttonSend.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonSend.Name = "buttonSend";
            this.buttonSend.Size = new System.Drawing.Size(144, 39);
            this.buttonSend.TabIndex = 6;
            this.buttonSend.Text = "送信(&S)";
            this.buttonSend.UseVisualStyleBackColor = true;
            this.buttonSend.Visible = false;
            this.buttonSend.Click += new System.EventHandler(this.buttonSend_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(1019, 345);
            this.buttonCancel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(144, 39);
            this.buttonCancel.TabIndex = 7;
            this.buttonCancel.Text = "キャンセル";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Visible = false;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // toolStripContainer1
            // 
            // 
            // toolStripContainer1.ContentPanel
            // 
            this.toolStripContainer1.ContentPanel.Controls.Add(this.splitContainer1);
            this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(1163, 535);
            this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer1.Name = "toolStripContainer1";
            this.toolStripContainer1.Size = new System.Drawing.Size(1163, 590);
            this.toolStripContainer1.TabIndex = 0;
            // 
            // toolStripContainer1.TopToolStripPanel
            // 
            this.toolStripContainer1.TopToolStripPanel.Controls.Add(this.menuStrip1);
            this.toolStripContainer1.TopToolStripPanel.Controls.Add(this.toolStrip1);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.label4);
            this.splitContainer1.Panel1.Controls.Add(this.textMailBcc);
            this.splitContainer1.Panel1.Controls.Add(this.label3);
            this.splitContainer1.Panel1.Controls.Add(this.textMailCc);
            this.splitContainer1.Panel1.Controls.Add(this.label2);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Controls.Add(this.textMailTo);
            this.splitContainer1.Panel1.Controls.Add(this.textMailSubject);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.buttonCancel);
            this.splitContainer1.Panel2.Controls.Add(this.buttonSend);
            this.splitContainer1.Panel2.Controls.Add(this.textMailBody);
            this.splitContainer1.Size = new System.Drawing.Size(1163, 535);
            this.splitContainer1.SplitterDistance = 123;
            this.splitContainer1.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 69);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 15);
            this.label4.TabIndex = 4;
            this.label4.Text = "BCC";
            // 
            // textMailBcc
            // 
            this.textMailBcc.ContextMenuStrip = this.contextAddress;
            this.textMailBcc.Location = new System.Drawing.Point(55, 66);
            this.textMailBcc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textMailBcc.Name = "textMailBcc";
            this.textMailBcc.Size = new System.Drawing.Size(1104, 22);
            this.textMailBcc.TabIndex = 5;
            this.textMailBcc.Leave += new System.EventHandler(this.textMailBcc_Leave);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 40);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(27, 15);
            this.label3.TabIndex = 2;
            this.label3.Text = "CC";
            // 
            // textMailCc
            // 
            this.textMailCc.ContextMenuStrip = this.contextAddress;
            this.textMailCc.Location = new System.Drawing.Point(56, 37);
            this.textMailCc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textMailCc.Name = "textMailCc";
            this.textMailCc.Size = new System.Drawing.Size(1104, 22);
            this.textMailCc.TabIndex = 3;
            this.textMailCc.Leave += new System.EventHandler(this.textMailCc_Leave);
            // 
            // textMailBody
            // 
            this.textMailBody.AllowDrop = true;
            this.textMailBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textMailBody.Font = new System.Drawing.Font("Yu Gothic UI", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textMailBody.Location = new System.Drawing.Point(0, 0);
            this.textMailBody.Multiline = true;
            this.textMailBody.Name = "textMailBody";
            this.textMailBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textMailBody.Size = new System.Drawing.Size(1163, 408);
            this.textMailBody.TabIndex = 8;
            this.textMailBody.WordWrap = false;
            this.textMailBody.DragDrop += new System.Windows.Forms.DragEventHandler(this.textMailBody_DragDrop);
            this.textMailBody.DragEnter += new System.Windows.Forms.DragEventHandler(this.textMailBody_DragEnter);
            // 
            // menuStrip1
            // 
            this.menuStrip1.AllowDrop = true;
            this.menuStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.menuEdit});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1163, 28);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuSend,
            this.toolStripMenuItem2,
            this.menuClose});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(82, 24);
            this.toolStripMenuItem1.Text = "ファイル(&F)";
            // 
            // menuSend
            // 
            this.menuSend.Name = "menuSend";
            this.menuSend.Size = new System.Drawing.Size(189, 26);
            this.menuSend.Text = "送信(&S)";
            this.menuSend.Click += new System.EventHandler(this.menuSend_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(186, 6);
            // 
            // menuClose
            // 
            this.menuClose.Name = "menuClose";
            this.menuClose.Size = new System.Drawing.Size(189, 26);
            this.menuClose.Text = "画面を閉じる(&C)";
            this.menuClose.Click += new System.EventHandler(this.menuClose_Click);
            // 
            // menuEdit
            // 
            this.menuEdit.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuUndo,
            this.toolStripMenuItem3,
            this.menuCut,
            this.menuCopy,
            this.menuPaste,
            this.menuDelete,
            this.toolStripMenuItem6,
            this.menuFind,
            this.menuReplace,
            this.toolStripMenuItem4,
            this.menuAllSelect});
            this.menuEdit.Name = "menuEdit";
            this.menuEdit.Size = new System.Drawing.Size(71, 24);
            this.menuEdit.Text = "編集(&E)";
            this.menuEdit.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // menuUndo
            // 
            this.menuUndo.Name = "menuUndo";
            this.menuUndo.Size = new System.Drawing.Size(177, 26);
            this.menuUndo.Text = "元に戻す(&U)";
            this.menuUndo.Click += new System.EventHandler(this.menuUndo_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(174, 6);
            // 
            // menuCut
            // 
            this.menuCut.Name = "menuCut";
            this.menuCut.Size = new System.Drawing.Size(177, 26);
            this.menuCut.Text = "切り取り(&T)";
            this.menuCut.Click += new System.EventHandler(this.menuCut_Click);
            // 
            // menuCopy
            // 
            this.menuCopy.Name = "menuCopy";
            this.menuCopy.Size = new System.Drawing.Size(177, 26);
            this.menuCopy.Text = "コピー(&C)";
            this.menuCopy.Click += new System.EventHandler(this.menuCopy_Click);
            // 
            // menuPaste
            // 
            this.menuPaste.Name = "menuPaste";
            this.menuPaste.Size = new System.Drawing.Size(177, 26);
            this.menuPaste.Text = "貼り付け(&P)";
            this.menuPaste.Click += new System.EventHandler(this.menuPaste_Click);
            // 
            // menuDelete
            // 
            this.menuDelete.Name = "menuDelete";
            this.menuDelete.Size = new System.Drawing.Size(177, 26);
            this.menuDelete.Text = "削除(&D)";
            this.menuDelete.Click += new System.EventHandler(this.menuDelete_Click);
            // 
            // toolStripMenuItem6
            // 
            this.toolStripMenuItem6.Name = "toolStripMenuItem6";
            this.toolStripMenuItem6.Size = new System.Drawing.Size(174, 6);
            // 
            // menuFind
            // 
            this.menuFind.Name = "menuFind";
            this.menuFind.Size = new System.Drawing.Size(177, 26);
            this.menuFind.Text = "検索(&F)";
            this.menuFind.Click += new System.EventHandler(this.menuFind_Click);
            // 
            // menuReplace
            // 
            this.menuReplace.Name = "menuReplace";
            this.menuReplace.Size = new System.Drawing.Size(177, 26);
            this.menuReplace.Text = "置換(&R)";
            this.menuReplace.Click += new System.EventHandler(this.menuReplace_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(174, 6);
            // 
            // menuAllSelect
            // 
            this.menuAllSelect.Name = "menuAllSelect";
            this.menuAllSelect.Size = new System.Drawing.Size(177, 26);
            this.menuAllSelect.Text = "すべて選択(&A)";
            this.menuAllSelect.Click += new System.EventHandler(this.menuAllSelect_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.CanOverflow = false;
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButtonSend,
            this.toolStripSeparator1,
            this.toolAddAttachment,
            this.toolStripSeparator2,
            this.toolStripButtonCut,
            this.toolStripButtonCopy,
            this.toolStripButtonPaste});
            this.toolStrip1.Location = new System.Drawing.Point(4, 28);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(322, 27);
            this.toolStrip1.TabIndex = 3;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonSend
            // 
            this.toolStripButtonSend.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonSend.Image")));
            this.toolStripButtonSend.Name = "toolStripButtonSend";
            this.toolStripButtonSend.Size = new System.Drawing.Size(63, 24);
            this.toolStripButtonSend.Text = "送信";
            this.toolStripButtonSend.Click += new System.EventHandler(this.menuSend_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 27);
            // 
            // toolAddAttachment
            // 
            this.toolAddAttachment.Image = ((System.Drawing.Image)(resources.GetObject("toolAddAttachment.Image")));
            this.toolAddAttachment.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolAddAttachment.Name = "toolAddAttachment";
            this.toolAddAttachment.Size = new System.Drawing.Size(147, 24);
            this.toolAddAttachment.Text = "添付ファイルの追加";
            this.toolAddAttachment.Click += new System.EventHandler(this.toolAddAttachment_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 27);
            // 
            // toolStripButtonCut
            // 
            this.toolStripButtonCut.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonCut.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonCut.Image")));
            this.toolStripButtonCut.Name = "toolStripButtonCut";
            this.toolStripButtonCut.Size = new System.Drawing.Size(29, 24);
            this.toolStripButtonCut.Text = "切り取り";
            this.toolStripButtonCut.Click += new System.EventHandler(this.menuCut_Click);
            // 
            // toolStripButtonCopy
            // 
            this.toolStripButtonCopy.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonCopy.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonCopy.Image")));
            this.toolStripButtonCopy.Name = "toolStripButtonCopy";
            this.toolStripButtonCopy.Size = new System.Drawing.Size(29, 24);
            this.toolStripButtonCopy.Text = "コピー";
            this.toolStripButtonCopy.Click += new System.EventHandler(this.menuCopy_Click);
            // 
            // toolStripButtonPaste
            // 
            this.toolStripButtonPaste.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonPaste.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonPaste.Image")));
            this.toolStripButtonPaste.Name = "toolStripButtonPaste";
            this.toolStripButtonPaste.Size = new System.Drawing.Size(29, 24);
            this.toolStripButtonPaste.Text = "貼り付け";
            this.toolStripButtonPaste.Click += new System.EventHandler(this.menuPaste_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.labelMessage,
            this.buttonAttachList});
            this.statusStrip1.Location = new System.Drawing.Point(0, 564);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1163, 26);
            this.statusStrip1.TabIndex = 8;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // labelMessage
            // 
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(1148, 20);
            this.labelMessage.Spring = true;
            this.labelMessage.Text = "メールを新規作成します";
            this.labelMessage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // buttonAttachList
            // 
            this.buttonAttachList.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.buttonAttachList.Image = ((System.Drawing.Image)(resources.GetObject("buttonAttachList.Image")));
            this.buttonAttachList.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonAttachList.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonAttachList.Name = "buttonAttachList";
            this.buttonAttachList.Size = new System.Drawing.Size(73, 24);
            this.buttonAttachList.Text = "添付";
            this.buttonAttachList.Visible = false;
            this.buttonAttachList.DropDownItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.buttonAttachList_DropDownItemClicked);
            // 
            // FormMailCreate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1163, 590);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.toolStripContainer1);
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(600, 400);
            this.Name = "FormMailCreate";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "FormMailCreate";
            this.Load += new System.EventHandler(this.FormMailCreate_Load);
            this.SizeChanged += new System.EventHandler(this.FormMailCreate_SizeChanged);
            this.contextAddress.ResumeLayout(false);
            this.toolStripContainer1.ContentPanel.ResumeLayout(false);
            this.toolStripContainer1.TopToolStripPanel.ResumeLayout(false);
            this.toolStripContainer1.TopToolStripPanel.PerformLayout();
            this.toolStripContainer1.ResumeLayout(false);
            this.toolStripContainer1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.TextBox textMailTo;
        public System.Windows.Forms.TextBox textMailSubject;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Button buttonSend;
        public System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.ToolStripContainer toolStripContainer1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem menuSend;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem menuClose;
        private System.Windows.Forms.ToolStripMenuItem menuEdit;
        private System.Windows.Forms.ToolStripMenuItem menuUndo;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem menuCut;
        private System.Windows.Forms.ToolStripMenuItem menuCopy;
        private System.Windows.Forms.ToolStripMenuItem menuPaste;
        private System.Windows.Forms.ToolStripMenuItem menuDelete;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem menuAllSelect;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButtonSend;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton toolStripButtonCut;
        private System.Windows.Forms.ToolStripButton toolStripButtonCopy;
        private System.Windows.Forms.ToolStripButton toolStripButtonPaste;
        public System.Windows.Forms.TextBox textMailBody;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox textMailCc;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.TextBox textMailBcc;
        private System.Windows.Forms.ToolStripButton toolAddAttachment;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel labelMessage;
        public System.Windows.Forms.ToolStripDropDownButton buttonAttachList;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem6;
        private System.Windows.Forms.ToolStripMenuItem menuFind;
        private System.Windows.Forms.ToolStripMenuItem menuReplace;
        private System.Windows.Forms.ContextMenuStrip contextAddress;
        private System.Windows.Forms.ToolStripMenuItem menuSelectFromAddressBook;
    }
}
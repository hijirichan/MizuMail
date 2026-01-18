
namespace MizuMail
{
    partial class FormMain
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("受信メール(0)", 1, 1);
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("送信メール(0)", 2, 2);
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("下書き");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("ごみ箱(0)");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("メール", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4});
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.ファイルFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSaveAs = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripSeparator();
            this.menuFileClearTrash = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.menuAppExit = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMail = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMailSend = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMailReceive = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.menuMailNewItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuMailReply = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.menuMailDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.設定SToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuAccountSetting = new System.Windows.Forms.ToolStripMenuItem();
            this.menuReleEdit = new System.Windows.Forms.ToolStripMenuItem();
            this.ヘルプHToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuHelpView = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem8 = new System.Windows.Forms.ToolStripSeparator();
            this.menuHelpVersionCheck = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripSeparator();
            this.menuHelpAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.labelMessage = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolMailProgress = new System.Windows.Forms.ToolStripProgressBar();
            this.labelDate = new System.Windows.Forms.ToolStripStatusLabel();
            this.buttonAtachMenu = new System.Windows.Forms.ToolStripDropDownButton();
            this.toolMain = new System.Windows.Forms.ToolStrip();
            this.toolSendButton = new System.Windows.Forms.ToolStripButton();
            this.toolReceiveButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolNewButton = new System.Windows.Forms.ToolStripButton();
            this.toolReplyButton = new System.Windows.Forms.ToolStripButton();
            this.toolDeleteButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolHelpButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.toolFilterCombo = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.textSearch = new System.Windows.Forms.ToolStripTextBox();
            this.toolSearchButton = new System.Windows.Forms.ToolStripButton();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.treeMain = new System.Windows.Forms.TreeView();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuCreateFolder = new System.Windows.Forms.ToolStripMenuItem();
            this.menuRenameFolder = new System.Windows.Forms.ToolStripMenuItem();
            this.menuDeleteFolder = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem7 = new System.Windows.Forms.ToolStripSeparator();
            this.menuClearTrash = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.listMain = new System.Windows.Forms.ListView();
            this.columnFromOrTo = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnSubject = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnSize = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnMailName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.menuUndoMail = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripSeparator();
            this.menuNotReadYet = new System.Windows.Forms.ToolStripMenuItem();
            this.menuRead = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.menuSpeechMail = new System.Windows.Forms.ToolStripMenuItem();
            this.richTextBody = new System.Windows.Forms.RichTextBox();
            this.timerMain = new System.Windows.Forms.Timer(this.components);
            this.timerAutoReceive = new System.Windows.Forms.Timer(this.components);
            this.browserMail = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.menuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.toolMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.contextMenuStrip2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.browserMail)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ファイルFToolStripMenuItem,
            this.menuMail,
            this.設定SToolStripMenuItem,
            this.ヘルプHToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(5, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1165, 28);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // ファイルFToolStripMenuItem
            // 
            this.ファイルFToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuSaveAs,
            this.toolStripMenuItem4,
            this.menuFileClearTrash,
            this.toolStripMenuItem3,
            this.menuAppExit});
            this.ファイルFToolStripMenuItem.Name = "ファイルFToolStripMenuItem";
            this.ファイルFToolStripMenuItem.Size = new System.Drawing.Size(82, 24);
            this.ファイルFToolStripMenuItem.Text = "ファイル(&F)";
            // 
            // menuSaveAs
            // 
            this.menuSaveAs.Name = "menuSaveAs";
            this.menuSaveAs.Size = new System.Drawing.Size(292, 26);
            this.menuSaveAs.Text = "名前をつけて保存(&A)";
            this.menuSaveAs.ToolTipText = "選択したメールに名前をつけて保存します。";
            this.menuSaveAs.Click += new System.EventHandler(this.menuSaveAs_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(289, 6);
            // 
            // menuFileClearTrash
            // 
            this.menuFileClearTrash.Name = "menuFileClearTrash";
            this.menuFileClearTrash.Size = new System.Drawing.Size(292, 26);
            this.menuFileClearTrash.Text = "ごみ箱を空にする(&Y)";
            this.menuFileClearTrash.ToolTipText = "ごみ箱フォルダのメールを削除します。";
            this.menuFileClearTrash.Click += new System.EventHandler(this.menuClearTrash_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(289, 6);
            // 
            // menuAppExit
            // 
            this.menuAppExit.Name = "menuAppExit";
            this.menuAppExit.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.menuAppExit.Size = new System.Drawing.Size(292, 26);
            this.menuAppExit.Text = "アプリケーションの終了(&X)";
            this.menuAppExit.ToolTipText = "アプリケーションを終了します。";
            this.menuAppExit.Click += new System.EventHandler(this.menuAppExit_Click);
            // 
            // menuMail
            // 
            this.menuMail.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuMailSend,
            this.menuMailReceive,
            this.toolStripMenuItem1,
            this.menuMailNewItem,
            this.menuMailReply,
            this.toolStripMenuItem2,
            this.menuMailDelete});
            this.menuMail.Name = "menuMail";
            this.menuMail.Size = new System.Drawing.Size(79, 24);
            this.menuMail.Text = "メール(&M)";
            // 
            // menuMailSend
            // 
            this.menuMailSend.Name = "menuMailSend";
            this.menuMailSend.Size = new System.Drawing.Size(226, 26);
            this.menuMailSend.Text = "送信(&S)";
            this.menuMailSend.ToolTipText = "メールを送信します。";
            this.menuMailSend.Click += new System.EventHandler(this.toolSendButton_Click);
            // 
            // menuMailReceive
            // 
            this.menuMailReceive.Name = "menuMailReceive";
            this.menuMailReceive.Size = new System.Drawing.Size(226, 26);
            this.menuMailReceive.Text = "受信(&M)";
            this.menuMailReceive.ToolTipText = "メールを受信します。";
            this.menuMailReceive.Click += new System.EventHandler(this.toolReceiveButton_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(223, 6);
            // 
            // menuMailNewItem
            // 
            this.menuMailNewItem.Name = "menuMailNewItem";
            this.menuMailNewItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.menuMailNewItem.Size = new System.Drawing.Size(226, 26);
            this.menuMailNewItem.Text = "新規作成(&N)";
            this.menuMailNewItem.ToolTipText = "メールを新規作成します。";
            this.menuMailNewItem.Click += new System.EventHandler(this.toolNewButton_Click);
            // 
            // menuMailReply
            // 
            this.menuMailReply.Name = "menuMailReply";
            this.menuMailReply.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.R)));
            this.menuMailReply.Size = new System.Drawing.Size(226, 26);
            this.menuMailReply.Text = "返信(&R)";
            this.menuMailReply.ToolTipText = "選択したメールを返信します。";
            this.menuMailReply.Click += new System.EventHandler(this.toolReplyButton_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(223, 6);
            // 
            // menuMailDelete
            // 
            this.menuMailDelete.Name = "menuMailDelete";
            this.menuMailDelete.ShortcutKeys = System.Windows.Forms.Keys.Delete;
            this.menuMailDelete.Size = new System.Drawing.Size(226, 26);
            this.menuMailDelete.Text = "削除(&D)";
            this.menuMailDelete.ToolTipText = "選択したメールを削除します。";
            this.menuMailDelete.Click += new System.EventHandler(this.toolDeleteButton_Click);
            // 
            // 設定SToolStripMenuItem
            // 
            this.設定SToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuAccountSetting,
            this.menuReleEdit});
            this.設定SToolStripMenuItem.Name = "設定SToolStripMenuItem";
            this.設定SToolStripMenuItem.Size = new System.Drawing.Size(71, 24);
            this.設定SToolStripMenuItem.Text = "設定(&S)";
            // 
            // menuAccountSetting
            // 
            this.menuAccountSetting.Name = "menuAccountSetting";
            this.menuAccountSetting.Size = new System.Drawing.Size(224, 26);
            this.menuAccountSetting.Text = "アカウントの設定(&M)";
            this.menuAccountSetting.ToolTipText = "メールを送受信する情報を設定します。";
            this.menuAccountSetting.Click += new System.EventHandler(this.menuAccountSetting_Click);
            // 
            // menuReleEdit
            // 
            this.menuReleEdit.Name = "menuReleEdit";
            this.menuReleEdit.Size = new System.Drawing.Size(224, 26);
            this.menuReleEdit.Text = "振り分け設定(&R)";
            this.menuReleEdit.Click += new System.EventHandler(this.menuReleEdit_Click);
            // 
            // ヘルプHToolStripMenuItem
            // 
            this.ヘルプHToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuHelpView,
            this.toolStripMenuItem8,
            this.menuHelpVersionCheck,
            this.toolStripMenuItem5,
            this.menuHelpAbout});
            this.ヘルプHToolStripMenuItem.Name = "ヘルプHToolStripMenuItem";
            this.ヘルプHToolStripMenuItem.Size = new System.Drawing.Size(79, 24);
            this.ヘルプHToolStripMenuItem.Text = "ヘルプ(&H)";
            // 
            // menuHelpView
            // 
            this.menuHelpView.Name = "menuHelpView";
            this.menuHelpView.ShortcutKeys = System.Windows.Forms.Keys.F1;
            this.menuHelpView.Size = new System.Drawing.Size(248, 26);
            this.menuHelpView.Text = "ヘルプの表示(&V)";
            this.menuHelpView.Click += new System.EventHandler(this.menuHelpView_Click);
            // 
            // toolStripMenuItem8
            // 
            this.toolStripMenuItem8.Name = "toolStripMenuItem8";
            this.toolStripMenuItem8.Size = new System.Drawing.Size(245, 6);
            // 
            // menuHelpVersionCheck
            // 
            this.menuHelpVersionCheck.Name = "menuHelpVersionCheck";
            this.menuHelpVersionCheck.Size = new System.Drawing.Size(248, 26);
            this.menuHelpVersionCheck.Text = "最新バージョンのチェック(&K)";
            this.menuHelpVersionCheck.ToolTipText = "最新バージョンがあるかを確認します。";
            this.menuHelpVersionCheck.Click += new System.EventHandler(this.menuHelpVersionCheck_Click);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(245, 6);
            // 
            // menuHelpAbout
            // 
            this.menuHelpAbout.Name = "menuHelpAbout";
            this.menuHelpAbout.Size = new System.Drawing.Size(248, 26);
            this.menuHelpAbout.Text = "MizuMailについて(&A)";
            this.menuHelpAbout.ToolTipText = "アプリケーションのバージョン情報を表示します。";
            this.menuHelpAbout.Click += new System.EventHandler(this.menuHelpAbout_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.statusStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible;
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.labelMessage,
            this.toolMailProgress,
            this.labelDate,
            this.buttonAtachMenu});
            this.statusStrip1.Location = new System.Drawing.Point(0, 575);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 19, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1165, 30);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusMain";
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = false;
            this.labelMessage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(1072, 24);
            this.labelMessage.Spring = true;
            this.labelMessage.Text = "現在の状況";
            this.labelMessage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // toolMailProgress
            // 
            this.toolMailProgress.Name = "toolMailProgress";
            this.toolMailProgress.Size = new System.Drawing.Size(100, 22);
            this.toolMailProgress.Visible = false;
            // 
            // labelDate
            // 
            this.labelDate.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.labelDate.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left;
            this.labelDate.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust;
            this.labelDate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.labelDate.Name = "labelDate";
            this.labelDate.Size = new System.Drawing.Size(73, 24);
            this.labelDate.Text = "現在日時";
            this.labelDate.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            // 
            // buttonAtachMenu
            // 
            this.buttonAtachMenu.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonAtachMenu.Image = ((System.Drawing.Image)(resources.GetObject("buttonAtachMenu.Image")));
            this.buttonAtachMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonAtachMenu.MergeAction = System.Windows.Forms.MergeAction.Insert;
            this.buttonAtachMenu.Name = "buttonAtachMenu";
            this.buttonAtachMenu.Size = new System.Drawing.Size(34, 28);
            this.buttonAtachMenu.Text = "toolStripDropDownButton1";
            this.buttonAtachMenu.ToolTipText = "添付ファイルが存在するときのメニュー";
            this.buttonAtachMenu.Visible = false;
            this.buttonAtachMenu.DropDownItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.buttonAtachMenu_DropDownItemClicked);
            // 
            // toolMain
            // 
            this.toolMain.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolSendButton,
            this.toolReceiveButton,
            this.toolStripSeparator1,
            this.toolNewButton,
            this.toolReplyButton,
            this.toolDeleteButton,
            this.toolStripSeparator2,
            this.toolHelpButton,
            this.toolStripSeparator3,
            this.toolFilterCombo,
            this.toolStripSeparator4,
            this.textSearch,
            this.toolSearchButton});
            this.toolMain.Location = new System.Drawing.Point(0, 28);
            this.toolMain.Name = "toolMain";
            this.toolMain.Size = new System.Drawing.Size(1165, 28);
            this.toolMain.TabIndex = 2;
            this.toolMain.Text = "toolStrip1";
            // 
            // toolSendButton
            // 
            this.toolSendButton.Image = ((System.Drawing.Image)(resources.GetObject("toolSendButton.Image")));
            this.toolSendButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolSendButton.Name = "toolSendButton";
            this.toolSendButton.Size = new System.Drawing.Size(63, 25);
            this.toolSendButton.Text = "送信";
            this.toolSendButton.Click += new System.EventHandler(this.toolSendButton_Click);
            // 
            // toolReceiveButton
            // 
            this.toolReceiveButton.Image = ((System.Drawing.Image)(resources.GetObject("toolReceiveButton.Image")));
            this.toolReceiveButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolReceiveButton.Name = "toolReceiveButton";
            this.toolReceiveButton.Size = new System.Drawing.Size(63, 25);
            this.toolReceiveButton.Text = "受信";
            this.toolReceiveButton.Click += new System.EventHandler(this.toolReceiveButton_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 28);
            // 
            // toolNewButton
            // 
            this.toolNewButton.Image = ((System.Drawing.Image)(resources.GetObject("toolNewButton.Image")));
            this.toolNewButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolNewButton.Name = "toolNewButton";
            this.toolNewButton.Size = new System.Drawing.Size(93, 25);
            this.toolNewButton.Text = "新規作成";
            this.toolNewButton.Click += new System.EventHandler(this.toolNewButton_Click);
            // 
            // toolReplyButton
            // 
            this.toolReplyButton.Enabled = false;
            this.toolReplyButton.Image = ((System.Drawing.Image)(resources.GetObject("toolReplyButton.Image")));
            this.toolReplyButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolReplyButton.Name = "toolReplyButton";
            this.toolReplyButton.Size = new System.Drawing.Size(63, 25);
            this.toolReplyButton.Text = "返信";
            this.toolReplyButton.Click += new System.EventHandler(this.toolReplyButton_Click);
            // 
            // toolDeleteButton
            // 
            this.toolDeleteButton.Enabled = false;
            this.toolDeleteButton.Image = ((System.Drawing.Image)(resources.GetObject("toolDeleteButton.Image")));
            this.toolDeleteButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolDeleteButton.Name = "toolDeleteButton";
            this.toolDeleteButton.Size = new System.Drawing.Size(63, 25);
            this.toolDeleteButton.Text = "削除";
            this.toolDeleteButton.Click += new System.EventHandler(this.toolDeleteButton_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 28);
            // 
            // toolHelpButton
            // 
            this.toolHelpButton.Image = ((System.Drawing.Image)(resources.GetObject("toolHelpButton.Image")));
            this.toolHelpButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolHelpButton.Name = "toolHelpButton";
            this.toolHelpButton.Size = new System.Drawing.Size(68, 25);
            this.toolHelpButton.Text = "ヘルプ";
            this.toolHelpButton.Click += new System.EventHandler(this.menuHelpView_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 28);
            // 
            // toolFilterCombo
            // 
            this.toolFilterCombo.Items.AddRange(new object[] {
            "すべて",
            "未読",
            "添付あり",
            "今日"});
            this.toolFilterCombo.Name = "toolFilterCombo";
            this.toolFilterCombo.Size = new System.Drawing.Size(121, 28);
            this.toolFilterCombo.SelectedIndexChanged += new System.EventHandler(this.toolFilterCombo_SelectedIndexChanged);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 28);
            // 
            // textSearch
            // 
            this.textSearch.Font = new System.Drawing.Font("Yu Gothic UI", 9F);
            this.textSearch.Name = "textSearch";
            this.textSearch.Size = new System.Drawing.Size(200, 28);
            // 
            // toolSearchButton
            // 
            this.toolSearchButton.Image = ((System.Drawing.Image)(resources.GetObject("toolSearchButton.Image")));
            this.toolSearchButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolSearchButton.Name = "toolSearchButton";
            this.toolSearchButton.Size = new System.Drawing.Size(63, 25);
            this.toolSearchButton.Text = "検索";
            this.toolSearchButton.Click += new System.EventHandler(this.toolSearchButton_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 56);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.treeMain);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(1165, 519);
            this.splitContainer1.SplitterDistance = 262;
            this.splitContainer1.TabIndex = 3;
            // 
            // treeMain
            // 
            this.treeMain.AllowDrop = true;
            this.treeMain.ContextMenuStrip = this.contextMenuStrip2;
            this.treeMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeMain.ImageIndex = 0;
            this.treeMain.ImageList = this.imageList1;
            this.treeMain.LabelEdit = true;
            this.treeMain.Location = new System.Drawing.Point(0, 0);
            this.treeMain.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.treeMain.Name = "treeMain";
            treeNode1.ImageIndex = 1;
            treeNode1.Name = "nodeReceive";
            treeNode1.SelectedImageIndex = 1;
            treeNode1.Text = "受信メール(0)";
            treeNode2.ImageIndex = 2;
            treeNode2.Name = "nodeSend";
            treeNode2.SelectedImageIndex = 2;
            treeNode2.Text = "送信メール(0)";
            treeNode3.ImageIndex = 3;
            treeNode3.Name = "nodeDraft";
            treeNode3.Text = "下書き";
            treeNode4.ImageIndex = 4;
            treeNode4.Name = "nodeDelete";
            treeNode4.Text = "ごみ箱(0)";
            treeNode5.ImageIndex = 0;
            treeNode5.Name = "rootMail";
            treeNode5.Text = "メール";
            this.treeMain.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode5});
            this.treeMain.SelectedImageIndex = 0;
            this.treeMain.Size = new System.Drawing.Size(262, 519);
            this.treeMain.TabIndex = 0;
            this.treeMain.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.treeMain_AfterLabelEdit);
            this.treeMain.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeMain_AfterSelect);
            this.treeMain.DragDrop += new System.Windows.Forms.DragEventHandler(this.treeMain_DragDrop);
            this.treeMain.DragEnter += new System.Windows.Forms.DragEventHandler(this.treeMain_DragEnter);
            this.treeMain.DragOver += new System.Windows.Forms.DragEventHandler(this.treeMain_DragOver);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuCreateFolder,
            this.menuRenameFolder,
            this.menuDeleteFolder,
            this.toolStripMenuItem7,
            this.menuClearTrash});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(196, 106);
            // 
            // menuCreateFolder
            // 
            this.menuCreateFolder.Name = "menuCreateFolder";
            this.menuCreateFolder.Size = new System.Drawing.Size(195, 24);
            this.menuCreateFolder.Text = "フォルダの作成(&C)";
            this.menuCreateFolder.Click += new System.EventHandler(this.menuCreateFolder_Click);
            // 
            // menuRenameFolder
            // 
            this.menuRenameFolder.Name = "menuRenameFolder";
            this.menuRenameFolder.Size = new System.Drawing.Size(195, 24);
            this.menuRenameFolder.Text = "名前の変更(&R)";
            this.menuRenameFolder.Click += new System.EventHandler(this.menuRenameFolder_Click);
            // 
            // menuDeleteFolder
            // 
            this.menuDeleteFolder.Name = "menuDeleteFolder";
            this.menuDeleteFolder.Size = new System.Drawing.Size(195, 24);
            this.menuDeleteFolder.Text = "フォルダの削除(&D)";
            this.menuDeleteFolder.Click += new System.EventHandler(this.menuDeleteFolder_Click);
            // 
            // toolStripMenuItem7
            // 
            this.toolStripMenuItem7.Name = "toolStripMenuItem7";
            this.toolStripMenuItem7.Size = new System.Drawing.Size(192, 6);
            // 
            // menuClearTrash
            // 
            this.menuClearTrash.Name = "menuClearTrash";
            this.menuClearTrash.Size = new System.Drawing.Size(195, 24);
            this.menuClearTrash.Text = "ごみ箱を空にする(&Y)";
            this.menuClearTrash.Click += new System.EventHandler(this.menuClearTrash_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "foler");
            this.imageList1.Images.SetKeyName(1, "inbox");
            this.imageList1.Images.SetKeyName(2, "send");
            this.imageList1.Images.SetKeyName(3, "draft");
            this.imageList1.Images.SetKeyName(4, "trash");
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.listMain);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.browserMail);
            this.splitContainer2.Panel2.Controls.Add(this.richTextBody);
            this.splitContainer2.Size = new System.Drawing.Size(899, 519);
            this.splitContainer2.SplitterDistance = 174;
            this.splitContainer2.TabIndex = 0;
            // 
            // listMain
            // 
            this.listMain.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.listMain.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnFromOrTo,
            this.columnSubject,
            this.columnDate,
            this.columnSize,
            this.columnMailName});
            this.listMain.ContextMenuStrip = this.contextMenuStrip1;
            this.listMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listMain.FullRowSelect = true;
            this.listMain.HideSelection = false;
            this.listMain.Location = new System.Drawing.Point(0, 0);
            this.listMain.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.listMain.Name = "listMain";
            this.listMain.ShowItemToolTips = true;
            this.listMain.Size = new System.Drawing.Size(899, 174);
            this.listMain.TabIndex = 0;
            this.listMain.UseCompatibleStateImageBehavior = false;
            this.listMain.View = System.Windows.Forms.View.Details;
            this.listMain.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listMain_ColumnClick);
            this.listMain.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.listMain_ItemDrag);
            this.listMain.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.listMain_ItemSelectionChanged);
            this.listMain.DoubleClick += new System.EventHandler(this.listMain_DoubleClick);
            // 
            // columnFromOrTo
            // 
            this.columnFromOrTo.Text = "メールボックス名";
            this.columnFromOrTo.Width = 142;
            // 
            // columnSubject
            // 
            this.columnSubject.Text = "件名";
            this.columnSubject.Width = 210;
            // 
            // columnDate
            // 
            this.columnDate.Text = "受信日時";
            this.columnDate.Width = 137;
            // 
            // columnSize
            // 
            this.columnSize.Text = "サイズ";
            // 
            // columnMailName
            // 
            this.columnMailName.Text = "メールファイル名";
            this.columnMailName.Width = 0;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuDelete,
            this.menuUndoMail,
            this.toolStripMenuItem6,
            this.menuNotReadYet,
            this.menuRead,
            this.toolStripSeparator5,
            this.menuSpeechMail});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(178, 136);
            // 
            // menuDelete
            // 
            this.menuDelete.Name = "menuDelete";
            this.menuDelete.Size = new System.Drawing.Size(177, 24);
            this.menuDelete.Text = "削除(&D)";
            this.menuDelete.Click += new System.EventHandler(this.menuDelete_Click);
            // 
            // menuUndoMail
            // 
            this.menuUndoMail.Name = "menuUndoMail";
            this.menuUndoMail.Size = new System.Drawing.Size(177, 24);
            this.menuUndoMail.Text = "元に戻す(&U)";
            this.menuUndoMail.Click += new System.EventHandler(this.menuUndoMail_Click);
            // 
            // toolStripMenuItem6
            // 
            this.toolStripMenuItem6.Name = "toolStripMenuItem6";
            this.toolStripMenuItem6.Size = new System.Drawing.Size(174, 6);
            // 
            // menuNotReadYet
            // 
            this.menuNotReadYet.Name = "menuNotReadYet";
            this.menuNotReadYet.Size = new System.Drawing.Size(177, 24);
            this.menuNotReadYet.Text = "未読みにする(&N)";
            this.menuNotReadYet.Click += new System.EventHandler(this.menuNotReadYet_Click);
            // 
            // menuRead
            // 
            this.menuRead.Name = "menuRead";
            this.menuRead.Size = new System.Drawing.Size(177, 24);
            this.menuRead.Text = "既読にする(&Y)";
            this.menuRead.Click += new System.EventHandler(this.menuRead_Click);
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(174, 6);
            // 
            // menuSpeechMail
            // 
            this.menuSpeechMail.Name = "menuSpeechMail";
            this.menuSpeechMail.Size = new System.Drawing.Size(177, 24);
            this.menuSpeechMail.Text = "読み上げ(&R)";
            this.menuSpeechMail.Click += new System.EventHandler(this.menuSpeechMail_Click);
            // 
            // richTextBody
            // 
            this.richTextBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBody.Font = new System.Drawing.Font("Yu Gothic UI", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.richTextBody.Location = new System.Drawing.Point(0, 0);
            this.richTextBody.Margin = new System.Windows.Forms.Padding(4);
            this.richTextBody.Name = "richTextBody";
            this.richTextBody.Size = new System.Drawing.Size(899, 341);
            this.richTextBody.TabIndex = 1;
            this.richTextBody.Text = "";
            this.richTextBody.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.richTextBody_LinkClicked);
            // 
            // timerMain
            // 
            this.timerMain.Enabled = true;
            this.timerMain.Interval = 500;
            this.timerMain.Tick += new System.EventHandler(this.timerMain_Tick);
            // 
            // timerAutoReceive
            // 
            this.timerAutoReceive.Tick += new System.EventHandler(this.timerAutoReceive_Tick);
            // 
            // browserMail
            // 
            this.browserMail.AllowExternalDrop = true;
            this.browserMail.CreationProperties = null;
            this.browserMail.DefaultBackgroundColor = System.Drawing.Color.White;
            this.browserMail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.browserMail.Location = new System.Drawing.Point(0, 0);
            this.browserMail.Name = "browserMail";
            this.browserMail.Size = new System.Drawing.Size(899, 341);
            this.browserMail.TabIndex = 3;
            this.browserMail.ZoomFactor = 1D;
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1165, 605);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.toolMain);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "FormMain";
            this.Text = "MizuMail";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMain_FormClosing);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.toolMain.ResumeLayout(false);
            this.toolMain.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.contextMenuStrip2.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.browserMail)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ファイルFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem menuAppExit;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStrip toolMain;
        private System.Windows.Forms.ToolStripButton toolNewButton;
        private System.Windows.Forms.ToolStripButton toolSendButton;
        private System.Windows.Forms.ToolStripButton toolReceiveButton;
        private System.Windows.Forms.ToolStripButton toolHelpButton;
        private System.Windows.Forms.ToolStripStatusLabel labelMessage;
        private System.Windows.Forms.SplitContainer splitContainer1;
        public System.Windows.Forms.TreeView treeMain;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.ListView listMain;
        private System.Windows.Forms.ColumnHeader columnFromOrTo;
        private System.Windows.Forms.ColumnHeader columnSubject;
        private System.Windows.Forms.ColumnHeader columnDate;
        private System.Windows.Forms.ToolStripMenuItem 設定SToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem menuAccountSetting;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuDelete;
        private System.Windows.Forms.ToolStripMenuItem menuNotReadYet;
        private System.Windows.Forms.ToolStripStatusLabel labelDate;
        private System.Windows.Forms.Timer timerMain;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ColumnHeader columnMailName;
        private System.Windows.Forms.RichTextBox richTextBody;
        private System.Windows.Forms.ToolStripDropDownButton buttonAtachMenu;
        private System.Windows.Forms.ColumnHeader columnSize;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton toolReplyButton;
        private System.Windows.Forms.ToolStripButton toolDeleteButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.ToolStripMenuItem menuClearTrash;
        private System.Windows.Forms.ToolStripMenuItem menuMail;
        private System.Windows.Forms.ToolStripMenuItem menuMailSend;
        private System.Windows.Forms.ToolStripMenuItem menuMailReceive;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem menuMailNewItem;
        private System.Windows.Forms.ToolStripMenuItem menuMailReply;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem menuMailDelete;
        private System.Windows.Forms.ToolStripMenuItem ヘルプHToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem menuHelpAbout;
        private System.Windows.Forms.Timer timerAutoReceive;
        private System.Windows.Forms.ToolStripMenuItem menuSaveAs;
        private System.Windows.Forms.ToolStripMenuItem menuFileClearTrash;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem menuHelpVersionCheck;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem5;
        private System.Windows.Forms.ToolStripMenuItem menuUndoMail;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem6;
        private System.Windows.Forms.ToolStripProgressBar toolMailProgress;
        private System.Windows.Forms.ToolStripMenuItem menuSpeechMail;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripTextBox textSearch;
        private System.Windows.Forms.ToolStripButton toolSearchButton;
        private System.Windows.Forms.ToolStripComboBox toolFilterCombo;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripMenuItem menuHelpView;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem8;
        private System.Windows.Forms.ToolStripMenuItem menuCreateFolder;
        private System.Windows.Forms.ToolStripMenuItem menuDeleteFolder;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem7;
        private System.Windows.Forms.ToolStripMenuItem menuRenameFolder;
        private System.Windows.Forms.ToolStripMenuItem menuRead;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripMenuItem menuReleEdit;
        private Microsoft.Web.WebView2.WinForms.WebView2 browserMail;
    }
}


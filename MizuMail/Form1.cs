using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Pop3;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Win32;
using MimeKit;
using MimeKit.Text;
using Newtonsoft.Json;
using NLog;
using NLog.Targets;
using Org.BouncyCastle.Asn1.X509;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Org.BouncyCastle.Tls.Certificate;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MizuMail
{
    public partial class FormMain : Form
    {
        // UIDL格納用の配列
        public List<string> localUidls = new List<string>();

        // メールボックス情報を表示しているときのフラグ
        public bool mailBoxViewFlag = false;

        // ListViewItemSorterに指定するフィールド
        public ListViewItemComparer listViewItemSorter;

        // 現在の検索キーワードを格納するフィールド
        private string currentKeyword = "";

        // 選択された行を格納するフィールド
        private Mail currentMail;
        private FolderManager folderManager;
        bool isBuildingTree = false;

        // ★ メールルール
        private List<MailRule> rules = new List<MailRule>();

        // ★ フォルダごとのキャッシュ
        private Dictionary<string, Mail> mailCache = new Dictionary<string, Mail>();

        // ロガーの取得
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// ListViewの項目の並び替えに使用するクラス
        /// </summary>
        public class ListViewItemComparer : System.Collections.IComparer
        {

            public static ListViewItemComparer Default
            {
                get { return _default; }
            }

            static ListViewItemComparer()
            {
                _default = new ListViewItemComparer
                {
                    Column = 2,
                    Order = SortOrder.Descending,
                    ColumnModes = new[] { ComparerMode.String, ComparerMode.String, ComparerMode.DateTime, ComparerMode.String, ComparerMode.String }
                };
            }

            /// <summary>
            /// 比較する方法
            /// </summary>
            public enum ComparerMode
            {
                String,
                Integer,
                DateTime
            };

            private int _column;
            private static ListViewItemComparer _default;

            /// <summary>
            /// 並び替えるListView列の番号
            /// </summary>
            public int Column
            {
                set
                {
                    if (_column == value)
                    {
                        if (Order == SortOrder.Ascending)
                            Order = SortOrder.Descending;
                        else if (Order == SortOrder.Descending)
                            Order = SortOrder.Ascending;
                    }
                    _column = value;
                }
                get
                {
                    return _column;
                }
            }

            /// <summary>
            /// 昇順か降順か
            /// </summary>
            public SortOrder Order { get; set; }

            /// <summary>
            /// 並び替えの方法
            /// </summary>
            public ComparerMode Mode { get; private set; }

            /// <summary>
            /// 列ごとの並び替えの方法
            /// </summary>
            public ComparerMode[] ColumnModes { get; set; }

            /// <summary>
            /// ListViewItemComparerクラスのコンストラクタ
            /// </summary>
            /// <param name="col">並び替える列番号</param>
            /// <param name="ord">昇順か降順か</param>
            /// <param name="cmod">並び替えの方法</param>
            public ListViewItemComparer(int col, SortOrder ord, ComparerMode cmod)
            {
                _column = col;
                Order = ord;
                Mode = cmod;
            }

            public ListViewItemComparer()
            {
                _column = 0;
                Order = SortOrder.Ascending;
                Mode = ComparerMode.String;
            }

            // xがyより小さいときはマイナスの数、大きいときはプラスの数、
            // 同じときは0を返す
            public int Compare(object x, object y)
            {
                int result = 0;

                // ListViewItemの取得
                ListViewItem itemx = (ListViewItem)x;
                ListViewItem itemy = (ListViewItem)y;

                //並べ替えの方法を決定
                if (ColumnModes != null && ColumnModes.Length > _column)
                    Mode = ColumnModes[_column];

                // 並び替えの方法別に、xとyを比較する
                switch (Mode)
                {
                    case ComparerMode.String:
                        result = string.Compare(itemx.SubItems[_column].Text,
                            itemy.SubItems[_column].Text);
                        break;
                    case ComparerMode.Integer:
                        result = int.Parse(itemx.SubItems[_column].Text) -
                            int.Parse(itemy.SubItems[_column].Text);
                        break;
                    case ComparerMode.DateTime:
                        DateTime dx, dy;
                        bool okx = TryParseListViewDate(itemx.SubItems[_column].Text, out dx);
                        bool oky = TryParseListViewDate(itemy.SubItems[_column].Text, out dy);

                        if (!okx && !oky)
                        {
                            // どちらもパース不能 → 文字列比較にフォールバック
                            result = string.Compare(itemx.SubItems[_column].Text, itemy.SubItems[_column].Text);
                        }
                        else if (!okx)
                        {
                            // x だけ無効 → 後ろに回す（好みで逆にしてもOK）
                            result = 1;
                        }
                        else if (!oky)
                        {
                            // y だけ無効
                            result = -1;
                        }
                        else
                        {
                            result = DateTime.Compare(dx, dy);
                        }
                        break;
                }

                // 降順の時は結果を+-逆にする
                if (Order == SortOrder.Descending)
                    result = -result;
                else if (Order == SortOrder.None)
                    result = 0;

                // 結果を返す
                return result;
            }
        }

        public FormMain()
        {
            InitializeComponent();

            // 初期化
            currentMail = null;

            System.Windows.Forms.Application.Idle += Application_Idle;

            listMain.ColumnClick += listMain_ColumnClick;
            listMain.SmallImageList = new ImageList { ImageSize = new Size(1, 20) };
            listViewItemSorter = ListViewItemComparer.Default;
            listMain.ListViewItemSorter = listViewItemSorter;
            listMain.Columns.Add("プレビュー", 100);
        }

        /// <summary>
        /// 日付用パースメソッド
        /// </summary>
        /// <param name="text"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        private static bool TryParseListViewDate(string text, out DateTime dt)
        {
            dt = DateTime.MinValue;

            if (string.IsNullOrEmpty(text) || text == "未送信")
                return false;

            // FormatReceivedDate で "yyyy/MM/dd HH:mm:ss" にしている前提
            if (DateTime.TryParseExact(
                    text,
                    "yyyy/MM/dd HH:mm:ss",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out dt))
            {
                return true;
            }

            // 念のための保険: 通常の TryParse も試す
            return DateTime.TryParse(text, out dt);
        }

        /// <summary>
        /// メール送受信後のTreeView、ListViewの更新
        /// </summary>
        private void UpdateView()
        {
            listMain.ListViewItemSorter = null;

            UpdateTreeView();
            UpdateListView();

            listMain.ListViewItemSorter = listViewItemSorter;

            UpdateUndoState();
        }

        private void UpdateTreeView()
        {
            // 受信
            int inboxCount = Directory.GetFiles(folderManager.Inbox.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[0].Text = $"受信メール({inboxCount})";

            // 送信
            int sendCount = Directory.GetFiles(folderManager.Send.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[1].Text = $"送信メール({sendCount})";

            // 下書き
            int draftCount = Directory.GetFiles(folderManager.Draft.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[2].Text = $"下書き({draftCount})";

            // ごみ箱
            int trashCount = Directory.GetFiles(folderManager.Trash.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[3].Text = $"ごみ箱({trashCount})";
        }

        private long GetMailFileSize(Mail mail)
        {
            string path = ResolveMailPath(mail);
            if (string.IsNullOrEmpty(path))
                return 0;

            if (!File.Exists(path))
                return 0;

            FileInfo fi = new FileInfo(path);
            return fi.Length;
        }

        private string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            double kb = bytes / 1024.0;
            if (kb < 1024) return kb.ToString("F1") + " KB";
            double mb = kb / 1024.0;
            return mb.ToString("F2") + " MB";
        }

        private void UpdateListView()
        {
            listMain.SelectedItems.Clear();
            listMain.Items.Clear();

            var prevSorter = listMain.ListViewItemSorter;
            listMain.ListViewItemSorter = null;

            listMain.BeginUpdate();
            try
            {
                listMain.Items.Clear();

                TreeNode selected = treeMain.SelectedNode;
                if (selected == null)
                    return;

                MailFolder folder = selected.Tag as MailFolder;
                if (folder == null)
                {
                    ShowMailboxInfo();
                    return;
                }

                // ★ すべてのフォルダを LoadEmlFolder に統一
                IEnumerable<Mail> displayList = LoadEmlFolder(folder);

                // ★ 検索フィルタ
                if (!string.IsNullOrEmpty(currentKeyword))
                {
                    displayList = displayList.Where(m =>
                        (m.subject?.Contains(currentKeyword) ?? false) ||
                        (m.body?.Contains(currentKeyword) ?? false) ||
                        (m.address?.Contains(currentKeyword) ?? false)
                    );
                }

                // ★ フィルタコンボ
                string filter = toolFilterCombo.SelectedItem?.ToString();

                if (filter == "未読")
                    displayList = displayList.Where(m => m.notReadYet);

                else if (filter == "添付あり")
                    displayList = displayList.Where(m => m.hasAtach).ToList();

                else if (filter == "今日")
                    displayList = displayList.Where(m =>
                    {
                        if (DateTime.TryParse(m.date, out DateTime dt))
                            return dt.Date == DateTime.Now.Date;
                        return false;
                    });

                mailBoxViewFlag = false;

                // ★ ListView に追加
                var baseFont = listMain.Font;

                foreach (Mail mail in displayList)
                {
                    string col0;

                    if (mail.Folder.Type == FolderType.Send || mail.Folder.Type == FolderType.Draft)
                    {
                        // 送信メール・下書き → 宛先を表示
                        col0 = mail.address;
                    }
                    else
                    {
                        // 受信メール・その他 → 差出人を表示
                        col0 = mail.from;
                    }

                    string displayDate = FormatReceivedDate(mail.date);

                    ListViewItem item = new ListViewItem(col0);

                    item.SubItems.Add(mail.subject);
                    item.SubItems.Add(displayDate);

                    long sizeBytes = GetMailFileSize(mail);
                    item.SubItems.Add(FormatSize(sizeBytes));

                    item.SubItems.Add(mail.mailName);

                    string preview = mail.body?.Replace("\r", "").Replace("\n", " ");
                    if (!string.IsNullOrEmpty(preview) && preview.Length > 30)
                        preview = preview.Substring(0, 30) + "…";
                    item.SubItems.Add(preview);

                    item.Tag = mail;

                    bool isDraftMail = (mail.Folder.Type == FolderType.Draft && mail.isDraft);

                    if (mail.notReadYet || isDraftMail)
                    {
                        item.BackColor = Color.FromArgb(0xE8, 0xF4, 0xFF);
                        item.Font = new Font(baseFont, FontStyle.Bold);
                    }
                    else
                    {
                        item.Font = new Font(baseFont, FontStyle.Regular);
                    }

                    listMain.Items.Add(item);
                }
            }
            finally
            {
                listMain.EndUpdate();
                listMain.ListViewItemSorter = prevSorter;
            }
            listMain.SelectedItems.Clear();
        }

        private void treeMain_AfterSelect(object sender, TreeViewEventArgs e)
        {
            richTextBody.Clear();
            currentKeyword = "";

            buttonAtachMenu.DropDownItems.Clear();
            buttonAtachMenu.Visible = false;

            currentMail = null;
            listMain.SelectedItems.Clear();

            var folder = e.Node.Tag as MailFolder;

            if (folder == null)
            {
                // ルートノード（メール）
                listMain.Columns[0].Text = "メールボックス名";
                listMain.Columns[1].Text = "メールアドレス";
                listMain.Columns[2].Text = "更新日時";
            }
            else
            {
                switch (folder.Type)
                {
                    case FolderType.Inbox:
                    case FolderType.InboxSub:
                        listMain.Columns[0].Text = "差出人";
                        listMain.Columns[1].Text = "件名";
                        listMain.Columns[2].Text = "受信日時";
                        break;

                    case FolderType.Send:
                        listMain.Columns[0].Text = "宛先";
                        listMain.Columns[1].Text = "件名";
                        listMain.Columns[2].Text = "送信日時";
                        break;

                    case FolderType.Draft:
                        listMain.Columns[0].Text = "宛先(下書き)";
                        listMain.Columns[1].Text = "件名";
                        listMain.Columns[2].Text = "作成日時";
                        break;

                    case FolderType.Trash:
                        listMain.Columns[0].Text = "差出人または宛先";
                        listMain.Columns[1].Text = "件名";
                        listMain.Columns[2].Text = "受信日時または送信日時";
                        break;
                }
            }

            UpdateListView();
        }

        private void menuAccountSetting_Click(object sender, EventArgs e)
        {
            FormMailAccount form = new FormMailAccount();
            DialogResult ret = form.ShowDialog();

            if (ret == DialogResult.OK)
            {
                // 設定が変更された場合の処理（必要なら追加）
                SetTimer(Mail.checkMail, Mail.checkInterval);
            }
        }

        private async void toolSendButton_Click(object sender, EventArgs e)
        {
            toolMailProgress.Minimum = 0;
            toolMailProgress.Maximum = 100;
            toolMailProgress.Value = 0;

            try
            {
                labelMessage.Text = "メール送信中...";
                statusStrip1.Refresh();

                using (var client = new SmtpClient())
                {
                    await client.ConnectAsync(Mail.smtpServerName, Mail.smtpPortNo, SecureSocketOptions.Auto);
                    await client.AuthenticateAsync(Mail.userName, Mail.password);

                    // ★ 送信対象は Draft フォルダの isDraft=true のメール
                    var sendList = LoadEmlFolder(folderManager.Draft)
                        .Where(m => m.isDraft)
                        .ToList();

                    int total = sendList.Count;
                    int sentCount = 0;
                    toolMailProgress.Visible = true;

                    foreach (Mail mail in sendList)
                    {
                        // ★ 送信前のパスを取得
                        string oldPath = ResolveMailPath(mail);

                        // ★ .eml をそのまま読み込む（添付も含めて完全）
                        var message = MimeMessage.Load(mail.mailPath);

                        // ★ 送信
                        await client.SendAsync(message);

                        // ★ 進捗更新
                        sentCount++;
                        toolMailProgress.Value = (int)(sentCount * 100.0 / total);
                        statusStrip1.Refresh();

                        // ★ メール状態更新
                        mail.date = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                        mail.notReadYet = false;
                        mail.isDraft = false;

                        // ★ 送信済みフォルダへ移動（唯一の移動）
                        MoveMailWithUndo(mail, folderManager.Send);

                        // ★ mailCache 更新
                        string newPath = mail.mailPath; // MoveMailWithUndo が更新済み
                        if (mailCache.ContainsKey(oldPath))
                            mailCache.Remove(oldPath);
                        mailCache[newPath] = mail;

                        // ★ Draft フォルダの一覧を更新
                        LoadEmlFolder(folderManager.Draft);

                        labelMessage.Text = "送信: " + mail.address;
                        statusStrip1.Refresh();
                    }

                    await client.DisconnectAsync(true);
                }

                labelMessage.Text = "メール送信完了";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                labelMessage.Text = "メール送信エラー : " + ex.Message;
                statusStrip1.Refresh();
            }

            toolMailProgress.Value = 100;
            await Task.Delay(300);
            toolMailProgress.Value = 0;
            toolMailProgress.Visible = false;

            BuildTree();
            UpdateView();
        }

        private async void toolReceiveButton_Click(object sender, EventArgs e)
        {
            // ステータスバーの状況を初期化する
            toolMailProgress.Minimum = 0;
            toolMailProgress.Maximum = 100;
            toolMailProgress.Value = 0;

            // メール受信方式がPOP3かIMAP4かで処理を分岐
            if (Mail.receiveMethod_Pop3)
            {
                // POP3メール受信処理を呼び出す
                await Pop3Receive();
            }
            else
            {
                // IMAP4メール受信処理を呼び出す
                await Imap4Receive();
            }
        }

        private void listMain_DoubleClick(object sender, EventArgs e)
        {
            if (listMain.SelectedItems.Count == 0)
                return;

            Mail mail = listMain.SelectedItems[0].Tag as Mail;
            if (mail == null)
                return;

            // ★ Trash 以外は既読・未読トグル
            if (mail.Folder.Type != FolderType.Trash)
            {
                ToggleReadState(mail);
            }

            // ★ 送信メールだけは編集画面を開く
            if (mail.Folder.Type == FolderType.Send || mail.Folder.Type == FolderType.Draft)
            {
                OpenSendMailEditor(mail);
                return;
            }

            UpdateListView();
        }

        private void listMain_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (!e.IsSelected)
                return;

            Mail mail = e.Item.Tag as Mail;
            if (mail == null)
                return;

            currentMail = mail;

            // ================================
            // ★ フルロード（ここで初めて本文を読む）
            // ================================
            var full = mail.message;

            // ================================
            // ★ HTML 抽出（再帰で確実に拾う）
            // ================================
            string html = GetHtmlBody(full.Body);

            // ★ HTMLがDOCTYPEから始まっていたら、整形する
            if (!string.IsNullOrEmpty(html))
            {
                // IEモードではDOCTYPEがあると描画されないことがあるので除去
                if (html.StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase))
                {
                    int index = html.IndexOf("<html", StringComparison.OrdinalIgnoreCase);
                    if (index > 0)
                        html = html.Substring(index);
                }

                // baseタグが無ければ挿入
                if (html.IndexOf("<base", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    // headタグを柔軟に検出（大文字・空白対応）
                    var headRegex = new Regex("<head[^>]*>", RegexOptions.IgnoreCase);
                    if (headRegex.IsMatch(html))
                    {
                        html = headRegex.Replace(html, m => m.Value + "<base href='file:///' />", 1);
                    }
                    else
                    {
                        // headタグが無い場合は bodyの前に挿入
                        var bodyRegex = new Regex("<body[^>]*>", RegexOptions.IgnoreCase);
                        if (bodyRegex.IsMatch(html))
                        {
                            html = bodyRegex.Replace(html,
                                "<head><base href='file:///' /></head>\r\n<body>", 1);
                        }
                        else
                        {
                            // 最後の手段：htmlタグの後に挿入
                            var htmlRegex = new Regex("<html[^>]*>", RegexOptions.IgnoreCase);
                            if (htmlRegex.IsMatch(html))
                            {
                                html = htmlRegex.Replace(html,
                                    m => m.Value + "<head><base href='file:///' /></head>", 1);
                            }
                        }
                    }
                }

                browserMail.Visible = true;
                ShowHtml(html);
                richTextBody.Text = "";
            }
            else
            {
                // fallback: plain text
                string text = full.TextBody ?? full.GetTextBody(MimeKit.Text.TextFormat.Plain);
                text = FixBrokenHtml(text);
                if(text.ToUpper().Contains("<HTML"))
                {
                    browserMail.Visible = true;
                    ShowHtml(text);
                    richTextBody.Text = "";
                }
                else
                {
                    browserMail.Visible = false;
                    richTextBody.Text = text ?? "";
                }
            }

            // ================================
            // ★ 添付ファイル処理（B方式）
            // ================================
            buttonAtachMenu.DropDownItems.Clear();

            foreach (var part in FindAttachments(full.Body))
            {
                string name = GetAttachmentName(part);
                string ext = Path.GetExtension(name);

                Icon icon = null;
                try
                {
                    icon = GetIconFromExtension(ext);
                }
                catch
                {
                    icon = SystemIcons.WinLogo;
                }

                var item = new ToolStripMenuItem(name, icon.ToBitmap());
                item.Tag = name;
                buttonAtachMenu.DropDownItems.Add(item);
            }

            buttonAtachMenu.Visible = buttonAtachMenu.DropDownItems.Count > 0;

            UpdateUndoState();
        }

        private string GetHtmlBody(MimeEntity entity)
        {
            // text/html
            var text = entity as TextPart;
            if (text != null && text.IsHtml)
                return text.Text;

            // multipart
            var multipart = entity as Multipart;
            if (multipart != null)
            {
                foreach (var part in multipart)
                {
                    string html = GetHtmlBody(part);
                    if (!string.IsNullOrEmpty(html))
                        return html;
                }
            }

            return null;
        }

        private void AddAttachmentMenuItems(MimeMessage full)
        {
            buttonAtachMenu.DropDownItems.Clear();

            foreach (var part in full.BodyParts)
            {
                var mp = part as MimePart;
                if (mp != null && mp.IsAttachment)
                {
                    string fileName = mp.FileName;

                    // ★ 一時フォルダに保存
                    string tempPath = Path.Combine(Path.GetTempPath(), fileName);
                    using (var stream = File.Create(tempPath))
                    {
                        mp.Content.DecodeTo(stream);
                    }

                    // ★ アイコン取得（ここで初めて成功する）
                    Icon icon = Icon.ExtractAssociatedIcon(tempPath);

                    // ★ メニューに追加
                    var item = new ToolStripMenuItem(fileName, icon.ToBitmap());
                    item.Tag = tempPath; // 保存先パスを保持
                    buttonAtachMenu.DropDownItems.Add(item);
                }
            }

            buttonAtachMenu.Visible = buttonAtachMenu.DropDownItems.Count > 0;
        }

        private void SaveAttachment(MimeMessage full, string fileName, Mail mail)
        {
            foreach (var part in full.BodyParts)
            {
                var mp = part as MimePart;
                if (mp != null && mp.IsAttachment && mp.FileName == fileName)
                {
                    using (var dialog = new SaveFileDialog())
                    {
                        dialog.FileName = fileName;

                        if (dialog.ShowDialog() == DialogResult.OK)
                        {
                            using (var stream = File.Create(dialog.FileName))
                            {
                                mp.Content.DecodeTo(stream);
                            }
                        }
                    }
                    return;
                }
            }
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            // listMainのカラムサイズを保存する
            Properties.Settings.Default.ColWidth0 = listMain.Columns[0].Width;
            Properties.Settings.Default.ColWidth1 = listMain.Columns[1].Width;
            Properties.Settings.Default.ColWidth2 = listMain.Columns[2].Width;
            Properties.Settings.Default.ColWidth3 = listMain.Columns[3].Width;
            Properties.Settings.Default.ColWidth4 = listMain.Columns[4].Width;
            Properties.Settings.Default.ColWidth5 = listMain.Columns[5].Width;
            Properties.Settings.Default.Save();

            SaveUidls();

            // 設定を保存する
            SaveRules();
            SaveSettings();
        }

        private async void FormMain_Load(object sender, EventArgs e)
        {
            mailCache.Clear();

            // ① FolderManager 初期化
            folderManager = new FolderManager();
            await browserMail.EnsureCoreWebView2Async(null);

            // ② 設定読み込み
            LoadSettings();
            LoadRules();

            // ③ mbox フォルダ構造保証
            Directory.CreateDirectory(folderManager.Inbox.FullPath);
            Directory.CreateDirectory(folderManager.Send.FullPath);
            Directory.CreateDirectory(folderManager.Draft.FullPath);
            Directory.CreateDirectory(folderManager.Trash.FullPath);

            BuildTree();

            // ④ TreeView の Tag を MailFolder に統一
            TreeNode root = treeMain.Nodes[0];
            TreeNode inboxNode = root.Nodes[0];
            MailFolder inboxFolder = folderManager.Inbox;
            root.Nodes[0].Tag = folderManager.Inbox;
            root.Nodes[1].Tag = folderManager.Send;
            root.Nodes[2].Tag = folderManager.Draft;
            root.Nodes[3].Tag = folderManager.Trash;

            // ⑤ inbox サブフォルダ読み込み（MailFolder 再帰）
            LoadInboxFolders(inboxNode, inboxFolder);

            LoadUidls();

            // ⑥ カラム幅復元
            RestoreColumnWidths();

            // ⑦ 表示更新（LoadEmlFolder を使う）
            UpdateView();

            // ⑧ タイマー開始
            SetTimer(Mail.checkMail, Mail.checkInterval);

            // ⑨ 展開
            treeMain.ExpandAll();
        }

        private void RestoreColumnWidths()
        {
            int[] defaults = { 150, 200, 150, 120, 0, 200 }; // 好みで調整可能
            int[] widths = new int[6];

            object[] settings =
            {
                Properties.Settings.Default.ColWidth0,
                Properties.Settings.Default.ColWidth1,
                Properties.Settings.Default.ColWidth2,
                Properties.Settings.Default.ColWidth3,
                Properties.Settings.Default.ColWidth4,
                Properties.Settings.Default.ColWidth5
            };

            for (int i = 0; i < 6; i++)
            {
                int w;

                // null → default
                if (settings[i] == null)
                {
                    widths[i] = defaults[i];
                    continue;
                }

                // 数値としてパースできるか？
                if (int.TryParse(settings[i].ToString(), out w))
                {
                    // 0 やマイナスは異常値 → default
                    widths[i] = (w > 0) ? w : defaults[i];
                }
                else
                {
                    // パース不能 → default
                    widths[i] = defaults[i];
                }
            }

            // ListView に反映
            for (int i = 0; i < 6; i++)
            {
                if (i < listMain.Columns.Count)
                    listMain.Columns[i].Width = widths[i];
            }
        }

        private void menuDelete_Click(object sender, EventArgs e)
        {
            var selected = listMain.SelectedItems.Cast<ListViewItem>()
                                                 .Select(i => i.Tag as Mail)
                                                 .Where(m => m != null)
                                                 .ToList();
            if (selected.Count == 0)
                return;

            var trashTargets = new List<Mail>();
            var permanentTargets = new List<Mail>();

            foreach (var mail in selected)
            {
                bool isTrash = (mail.Folder == folderManager.Trash);

                if (!isTrash)
                    trashTargets.Add(mail);
                else
                    permanentTargets.Add(mail);
            }

            // ごみ箱へ移動
            if (trashTargets.Count > 0)
            {
                string msg = $"選択したメール {trashTargets.Count} 件をごみ箱に移動します。よろしいですか？";
                if (MessageBox.Show(msg, "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    foreach (var m in trashTargets)
                        MoveMailWithUndo(m, folderManager.Trash);
                }
            }

            // ごみ箱内の完全削除
            if (permanentTargets.Count > 0)
            {
                string msg = $"ごみ箱内のメール {permanentTargets.Count} 件を完全に削除します。元に戻せません。";
                if (MessageBox.Show(msg, "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    foreach (var m in permanentTargets)
                        DeletePermanently(m);
                }
            }

            // ★ UpdateView() は危険なので使わない
            BuildTree();
            UpdateListView();

            currentMail = null;
            UpdateUndoState();
        }

        private void menuNotReadYet_Click(object sender, EventArgs e)
        {
            if (currentMail == null)
                return;

            if (!currentMail.notReadYet)
                ToggleReadState(currentMail);

            UpdateView();
        }

        private void menuAppExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void timerMain_Tick(object sender, EventArgs e)
        {
            labelDate.Text = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString();
            statusStrip1.Refresh();
        }

        private void toolNewButton_Click(object sender, EventArgs e)
        {
            // メール作成ウィンドウを作成する
            FormMailCreate form = new FormMailCreate();

            // ウィンドウのタイトルを設定する
            form.Text = "新規作成";

            // ウィンドウを表示する
            DialogResult ret = form.ShowDialog();

            if (ret == DialogResult.OK)
            {
                string to = form.textMailTo.Text;
                string subject = form.textMailSubject.Text;
                string body = form.textMailBody.Text;
                string cc = form.textMailCc.Text;
                string bcc = form.textMailBcc.Text;
                string atach = string.Join(";", form.buttonAttachList.DropDownItems.Cast<ToolStripItem>().Select(item => item.Text));

                // 宛先、本文がある場合
                if (to != "" | body != "")
                {
                    // 件名がない場合は無題
                    if (subject == "")
                    {
                        subject = "無題";
                    }
                    // コレクションに追加する
                    Mail mail = new Mail(to, cc, bcc, subject, body, atach, "未送信", "", "", true);
                    mail.isDraft = true;
                    mail.notReadYet = true;
                    mail.Folder = folderManager.Draft;
                    SaveMail(mail);
                }

                // ツリービューとリストビューの表示を更新する
                UpdateView();
            }
        }

        /// <summary>
        /// 設定ファイルからアプリケーション設定を読み出す
        /// </summary>
        public void LoadSettings()
        {
            // 環境設定保存クラスを作成する
            MailSettings MailSetting = new MailSettings();

            // アカウント情報(初期値)を設定する
            Mail.fromName = "";
            Mail.userAddress = "";
            Mail.userName = "";
            Mail.password = "";
            Mail.smtpServerName = "";
            Mail.popServerName = "";
            Mail.imapServerName = "";
            Mail.popPortNo = 110;
            Mail.imapPortNo = 143;
            Mail.smtpPortNo = 25;
            Mail.deleteMail = false;
            Mail.alertSound = false;
            Mail.alertSoundFile = "";
            Mail.checkInterval = 5;
            Mail.checkMail = false;
            Mail.useSsl = false;
            Mail.receiveMethod_Pop3 = true;

            // 環境設定ファイルが存在する場合は環境設定情報を読み込んでアカウント情報に設定する
            if (File.Exists(System.Windows.Forms.Application.StartupPath + @"\MizuMail.xml"))
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(MailSettings));
                using (var fs = new FileStream(System.Windows.Forms.Application.StartupPath + @"\MizuMail.xml", FileMode.Open))
                {
                    MailSetting = (MailSettings)serializer.Deserialize(fs);
                }

                // アカウント情報
                Mail.fromName = MailSetting.m_fromName;
                Mail.userAddress = MailSetting.m_mailAddress;
                Mail.userName = MailSetting.m_userName;
                Mail.password = Decrypt(MailSetting.m_passWord);

                // 接続情報
                Mail.smtpServerName = MailSetting.m_smtpServer;
                Mail.popServerName = MailSetting.m_popServer;
                Mail.imapServerName = MailSetting.m_imapServer;
                Mail.imapPortNo = MailSetting.m_imapPortNo;
                Mail.popPortNo = MailSetting.m_popPortNo;
                Mail.smtpPortNo = MailSetting.m_smtpPortNo;
                Mail.deleteMail = MailSetting.m_deleteMail;
                Mail.useSsl = MailSetting.m_useSsl;
                Mail.receiveMethod_Pop3 = MailSetting.m_ReceiveMethod_Pop3;

                // 通知音設定
                Mail.alertSound = MailSetting.m_alertSound;
                Mail.alertSoundFile = MailSetting.m_alertSoundFile;

                // メールチェック設定
                Mail.checkInterval = MailSetting.m_checkInterval;
                Mail.checkMail = MailSetting.m_checkMail;

                // 画面の表示が通常のとき 
                if (MailSetting.m_windowStat == FormWindowState.Normal)
                {
                    // 過去のバージョンから環境設定ファイルを流用した初期起動以外はこの中に入る
                    if (MailSetting.m_windowLeft != 0 && MailSetting.m_windowTop != 0 && MailSetting.m_windowWidth != 0 && MailSetting.m_windowHeight != 0)
                    {
                        this.Left = MailSetting.m_windowLeft;
                        this.Top = MailSetting.m_windowTop;
                        this.Width = MailSetting.m_windowWidth;
                        this.Height = MailSetting.m_windowHeight;
                    }
                }
                this.WindowState = MailSetting.m_windowStat;
            }
        }

        /// <summary>
        /// アプリケーション設定を設定ファイルに書き出す
        /// </summary>
        public void SaveSettings()
        {
            MailSettings MailSetting = new MailSettings()
            {
                // アカウント情報
                m_fromName = Mail.fromName,
                m_mailAddress = Mail.userAddress,
                m_userName = Mail.userName,
                m_passWord = Encrypt(Mail.password),

                // 接続情報
                m_smtpServer = Mail.smtpServerName,
                m_popServer = Mail.popServerName,
                m_smtpPortNo = Mail.smtpPortNo,
                m_popPortNo = Mail.popPortNo,
                m_imapServer = Mail.imapServerName,
                m_imapPortNo = Mail.imapPortNo,
                m_deleteMail = Mail.deleteMail,
                m_useSsl = Mail.useSsl,
                m_ReceiveMethod_Pop3 = Mail.receiveMethod_Pop3,

                // 通知音設定
                m_alertSound = Mail.alertSound,
                m_alertSoundFile = Mail.alertSoundFile,

                // メールチェック設定
                m_checkInterval = Mail.checkInterval,
                m_checkMail = Mail.checkMail,

                // ウィンドウ設定
                m_windowLeft = this.Left,
                m_windowTop = this.Top,
                m_windowWidth = this.Width,
                m_windowHeight = this.Height,
                m_windowStat = this.WindowState
            };

            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(MailSettings));

            using (var fs = new FileStream(System.Windows.Forms.Application.StartupPath + @"\MizuMail.xml", FileMode.Create))
            {
                serializer.Serialize(fs, MailSetting);
            }
        }

        private void richTextBody_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            string url = e.LinkText;

            // URL の安全性チェック
            var result = CheckUrlSafety(url);

            if (!result.IsSafe)
            {
                MessageBox.Show( $"このリンクは安全ではない可能性があります。\n理由: {result.Reason}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 開く前に確認
            if (MessageBox.Show($"リンクを開きますか？\n{url}", "リンクを開く", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(url);
            }
        }

        private void buttonAtachMenu_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (listMain.SelectedItems.Count == 0)
                return;

            Mail mail = listMain.SelectedItems[0].Tag as Mail;
            if (mail == null)
                return;

            string fileName = e.ClickedItem.Tag as string;

            var message = MimeMessage.Load(mail.mailPath);

            // 添付パートを検索
            var part = FindAttachments(message.Body)
                .FirstOrDefault(p =>
                    string.Equals(GetAttachmentName(p), fileName, StringComparison.OrdinalIgnoreCase));

            if (part == null)
            {
                MessageBox.Show("添付ファイルが見つかりませんでした。");
                return;
            }

            // Temp に展開
            string tempDir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "tmp");
            Directory.CreateDirectory(tempDir);

            string tempPath = Path.Combine(tempDir, fileName);

            using (var stream = File.Create(tempPath))
                part.Content.DecodeTo(stream);

            Process.Start(tempPath);
        }

        // 受信日時をローカル時刻に変換して表示する（日本語曜日対応）
        private string FormatReceivedDate(string dateText)
        {
            if (string.IsNullOrEmpty(dateText))
                return "";

            if (dateText == "未送信")
                return dateText;

            var jp = new System.Globalization.CultureInfo("ja-JP");

            // DateTimeOffset（オフセット付き）
            if (DateTimeOffset.TryParse(dateText, out DateTimeOffset dto))
            {
                var local = dto.ToLocalTime();
                return local.ToString("yyyy/MM/dd (ddd) HH:mm:ss", jp);
            }

            // DateTime（オフセットなし）
            if (DateTime.TryParse(dateText, out DateTime dt))
            {
                if (dt.Kind == DateTimeKind.Utc)
                    return dt.ToLocalTime().ToString("yyyy/MM/dd (ddd) HH:mm:ss", jp);

                if (dt.Kind == DateTimeKind.Unspecified)
                {
                    var localSpecified = DateTime.SpecifyKind(dt, DateTimeKind.Local);
                    return localSpecified.ToString("yyyy/MM/dd (ddd) HH:mm:ss", jp);
                }
                return dt.ToString("yyyy/MM/dd (ddd) HH:mm:ss", jp);
            }

            return dateText;
        }

        private void Application_Idle(object sender, EventArgs e)
        {
            toolReplyButton.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            toolDeleteButton.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuRead.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuNotReadYet.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuDelete.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuMailDelete.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuMailReply.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            var trash = folderManager.Trash.FullPath;
            menuFileClearTrash.Enabled = Directory.GetFiles(trash, "*.eml").Length > 0 || Directory.GetFiles(trash, "*.meta").Length > 0;
            menuClearTrash.Enabled = Directory.GetFiles(trash, "*.eml").Length > 0 || Directory.GetFiles(trash, "*.meta").Length > 0;
            menuSaveAs.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            menuSpeechMail.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            UpdateUndoState();
        }

        private void toolDeleteButton_Click(object sender, EventArgs e)
        {
            // 削除メニューと同じ処理を呼び出す
            menuDelete_Click(sender, e);
        }

        private void toolReplyButton_Click(object sender, EventArgs e)
        {
            // 返信対象のメールが選択されていない場合は何もしない
            if (currentMail == null)
                return;

            // メール作成ウィンドウを作成する
            FormMailCreate form = new FormMailCreate();

            // ウィンドウのタイトルを設定する
            form.Text = "返信";

            // 返信するメールを設定する
            form.textMailSubject.Text = "Re: " + currentMail.subject;
            form.textMailTo.Text = currentMail.from;
            if (currentMail.body.Trim() != string.Empty)
            {
                form.textMailBody.Text = "\r\n\r\n------------------------------\r\n" + currentMail.body.TrimEnd('\r', '\n');
            }

            // ウィンドウを表示する
            DialogResult ret = form.ShowDialog();

            if (ret == DialogResult.OK)
            {
                string to = form.textMailTo.Text;
                string subject = form.textMailSubject.Text;
                string body = form.textMailBody.Text;
                string cc = form.textMailCc.Text;
                string bcc = form.textMailBcc.Text;
                string atach = string.Join(";", form.buttonAttachList.DropDownItems.Cast<ToolStripItem>().Select(item => item.Text));

                // 宛先、本文がある場合
                if (to != "" | body != "")
                {
                    // 件名がない場合は無題
                    if (subject == "")
                    {
                        subject = "無題";
                    }
                    // コレクションに追加する
                    Mail mail = new Mail(to, cc, bcc, subject, body, atach, "未送信", "", "", true);
                    mail.isDraft = true;
                    mail.notReadYet = true;
                    mail.Folder = folderManager.Draft;
                    SaveMail(mail);
                }

                // ツリービューとリストビューの表示を更新する
                UpdateView();
            }
        }

        private void menuClearTrash_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("ごみ箱内のメールをすべて完全に削除します。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
                return;

            var trash = folderManager.Trash;

            // ★ ごみ箱内の全ファイルを削除（.eml と .meta）
            foreach (var file in Directory.GetFiles(trash.FullPath))
            {
                try
                {
                    File.Delete(file);

                    // ★ mailCache からも削除（フルパスキー）
                    mailCache.Remove(file);
                }
                catch { }
            }

            // ★ ごみ箱を空にした後は Undo を無効化
            currentMail = null;
            menuUndoMail.Enabled = false;

            UpdateView();
        }

        private void menuHelpAbout_Click(object sender, EventArgs e)
        {
            var aboutForm = new AboutForm();
            aboutForm.ShowDialog();
        }

        private void SetTimer(bool isEnabled, int intervalMinutes)
        {
            // 60,000(msec)
            timerAutoReceive.Interval = intervalMinutes * 60000;
            timerAutoReceive.Enabled = isEnabled;
        }

        private void timerAutoReceive_Tick(object sender, EventArgs e)
        {
            toolReceiveButton_Click(sender, e);
        }

        private void menuSaveAs_Click(object sender, EventArgs e)
        {
            if (currentMail == null)
                return;

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.FileName = (currentMail.subject ?? "mail") + ".eml";
                sfd.Filter = "EMLファイル (*.eml)|*.eml|すべてのファイル (*.*)|*.*";
                sfd.FilterIndex = 0;
                sfd.RestoreDirectory = true;
                sfd.OverwritePrompt = true;

                if (sfd.ShowDialog() != DialogResult.OK)
                    return;

                // ★ 送信メールは再構築
                if (currentMail.Folder.Type == FolderType.Send)
                {
                    MimeMessage msg = BuildMimeMessageFromMail(currentMail);

                    using (var stream = File.Create(sfd.FileName))
                        msg.WriteTo(stream);

                    return;
                }

                // ★ 受信・ごみ箱・サブフォルダ → 既存の .eml をコピー
                string src = ResolveMailPath(currentMail);
                if (!string.IsNullOrEmpty(src) && File.Exists(src))
                {
                    File.Copy(src, sfd.FileName, true);
                }
                else
                {
                    MessageBox.Show("メールファイルが見つかりません。", "エラー");
                }
            }
        }

        private MimeMessage BuildMimeMessageFromMail(Mail mail)
        {
            var message = new MimeMessage();

            // From
            if (!string.IsNullOrWhiteSpace(Mail.userAddress))
                message.From.Add(MailboxAddress.Parse(Mail.userAddress));

            // To
            if (!string.IsNullOrWhiteSpace(mail.address))
                foreach (var addr in mail.address.Split(';'))
                    if (!string.IsNullOrWhiteSpace(addr))
                        message.To.Add(MailboxAddress.Parse(addr.Trim()));

            // Cc
            if (!string.IsNullOrWhiteSpace(mail.ccaddress))
                foreach (var addr in mail.ccaddress.Split(';'))
                    if (!string.IsNullOrWhiteSpace(addr))
                        message.Cc.Add(MailboxAddress.Parse(addr.Trim()));

            // Bcc
            if (!string.IsNullOrWhiteSpace(mail.bccaddress))
                foreach (var addr in mail.bccaddress.Split(';'))
                    if (!string.IsNullOrWhiteSpace(addr))
                        message.Bcc.Add(MailboxAddress.Parse(addr.Trim()));

            message.Subject = mail.subject ?? "";

            var textPart = new TextPart(TextFormat.Text)
            {
                Text = mail.body ?? ""
            };

            var files = (mail.atach ?? "")
                .Split(';')
                .Select(f => f.Trim())
                .Where(f => File.Exists(f))
                .ToList();

            if (files.Any())
            {
                var multipart = new Multipart("mixed");
                multipart.Add(textPart);

                foreach (var file in files)
                {
                    multipart.Add(new MimePart()
                    {
                        Content = new MimeContent(File.OpenRead(file)),
                        ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                        ContentTransferEncoding = ContentEncoding.Base64,
                        FileName = Path.GetFileName(file)
                    });
                }

                message.Body = multipart;
            }
            else
            {
                message.Body = textPart;
            }

            return message;
        }

        private void GetDirectorySize(string targetDirectory, ref long size)
        {
            // ディレクトリ内のファイルサイズを取得する
            string[] files = Directory.GetFiles(targetDirectory);
            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                size += fi.Length;
            }
            // サブディレクトリ内のファイルサイズを取得する（再帰処理）
            string[] directories = Directory.GetDirectories(targetDirectory);
            foreach (string directory in directories)
            {
                GetDirectorySize(directory, ref size);
            }
        }

        private void menuHelpVersionCheck_Click(object sender, EventArgs e)
        {
            // バージョンチェックを行う
            if (IsNewVersionAvailable())
            {
                MessageBox.Show("新しいバージョンが利用可能です。\nhttps://www.angel-teatime.com/ からダウンロードしてください。", "バージョンチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("現在お使いのバージョンは最新です。", "バージョンチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static bool IsNewVersionAvailable()
        {
            try
            {
                string url = "https://www.angel-teatime.com/files/mizumail/mizumail_version.txt";

                using (var client = new WebClient())
                {
                    string versionText = client.DownloadString(url).Trim();
                    Version serverVersion = new Version(versionText);
                    Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                    return serverVersion > currentVersion;
                }
            }
            catch
            {
                // 通信エラーなどは false 扱い
                return false;
            }
        }

        public static string Encrypt(string plainText)
        {
            byte[] data = Encoding.UTF8.GetBytes(plainText);
            byte[] encrypted = ProtectedData.Protect(
                data,
                null,
                DataProtectionScope.CurrentUser
            );
            return Convert.ToBase64String(encrypted);
        }

        public static string Decrypt(string encryptedText)
        {
            byte[] decrypted = null;
            try
            {
                byte[] data = Convert.FromBase64String(encryptedText);
                decrypted = ProtectedData.Unprotect(
                    data,
                    null,
                    DataProtectionScope.CurrentUser
                );
            }
            catch(Exception)
            {
                // 復号に失敗した場合は空文字を返す
                return string.Empty;
            }

            return Encoding.UTF8.GetString(decrypted);
        }

        private async Task Pop3Receive()
        {
            // メール受信件数
            int mailCount = 0;

            try
            {
                labelMessage.Text = "メール受信中...";
                statusStrip1.Refresh();

                using (var client = new Pop3Client())
                {
                    await client.ConnectAsync(Mail.popServerName, Mail.popPortNo, Mail.useSsl);
                    await client.AuthenticateAsync(Mail.userName, Mail.password);

                    int total = client.Count;
                    labelMessage.Text = $"{total}件のメッセージがあります";
                    statusStrip1.Refresh();

                    toolMailProgress.Visible = true;

                    for (int i = 0; i < total; i++)
                    {
                        toolMailProgress.Value = (int)((i + 1) * 100.0 / total);
                        statusStrip1.Refresh();

                        string uidl = client.GetMessageUid(i);

                        if (localUidls.Contains(uidl))
                            continue;

                        labelMessage.Text = $"{i + 1}件目のメール受信中";
                        statusStrip1.Refresh();

                        var message = await client.GetMessageAsync(i);

                        // ★ MessageId が無い場合の fallback
                        string baseName = message.MessageId;
                        if (string.IsNullOrWhiteSpace(baseName))
                            baseName = Guid.NewGuid().ToString();

                        // ★ ファイル名サニタイズ
                        foreach (char c in Path.GetInvalidFileNameChars())
                        {
                            baseName = baseName.Replace(c, '_');
                        }

                        string mailName = baseName + "_unread.eml";

                        string inboxPath = Path.Combine(folderManager.Inbox.FullPath, mailName);
                        Directory.CreateDirectory(folderManager.Inbox.FullPath);

                        await Task.Run(() => message.WriteTo(inboxPath));

                        // ★ Mail オブジェクト生成（folder は不要）
                        Mail mail = Mail.FromMimeMessage(message);
                        mail.mailName = mailName;
                        mail.notReadYet = true;

                        // ★ パス設定（必須）
                        mail.mailPath = inboxPath;

                        // ★ mailCache に登録（必須）
                        mailCache[inboxPath] = mail;

                        // 振り分け処理
                        ApplyRules(mail);

                        // ★ UIDL 保存
                        localUidls.Add(uidl);
                        mailCount++;

                        if (Mail.deleteMail)
                            await client.DeleteMessageAsync(i);
                    }

                    await client.DisconnectAsync(true);
                }

                // ★ UIDL 永続化（最後に1回だけ）
                SaveUidls();

                labelMessage.Text = "メール受信完了";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                labelMessage.Text = "メール受信エラー : " + ex.Message;
                statusStrip1.Refresh();
            }

            // 通知音など（既存の安全対策はそのまま）
            if (mailCount > 0)
            {
                if (Mail.alertSound && !string.IsNullOrWhiteSpace(Mail.alertSoundFile))
                {
                    try
                    {
                        if (File.Exists(Mail.alertSoundFile))
                        {
                            using (var p = new SoundPlayer(Mail.alertSoundFile))
                            {
                                p.Play();
                            }
                        }
                        else
                        {
                            logger.Error($"Alert sound file not found: {Mail.alertSoundFile}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Alert sound error: " + ex.Message);
                    }
                }

                labelMessage.Text = $"{mailCount}件の新着メールがあります";
                statusStrip1.Refresh();
            }

            toolMailProgress.Value = 100;
            await Task.Delay(300);
            toolMailProgress.Value = 0;
            toolMailProgress.Visible = false;

            // ツリービューとリストビューの表示を更新する
            BuildTree();
            UpdateView();
        }

        private async Task Imap4Receive()
        {
            int mailCount = 0;

            try
            {
                labelMessage.Text = "メール受信中...";
                statusStrip1.Refresh();

                using (var client = new ImapClient())
                {
                    await client.ConnectAsync(Mail.imapServerName, Mail.imapPortNo, Mail.useSsl);
                    await client.AuthenticateAsync(Mail.userName, Mail.password);

                    var inbox = client.Inbox;
                    await inbox.OpenAsync(FolderAccess.ReadOnly);

                    labelMessage.Text = $"{inbox.Count}件のメッセージがあります";
                    statusStrip1.Refresh();

                    var summaries = await inbox.FetchAsync(0, -1, MessageSummaryItems.UniqueId);

                    toolMailProgress.Visible = true;
                    int total = summaries.Count;

                    for (int i = 0; i < total; i++)
                    {
                        toolMailProgress.Value = (int)((i + 1) * 100.0 / total);
                        statusStrip1.Refresh();

                        var summary = summaries[i];
                        var uid = summary.UniqueId.Id.ToString();

                        if (localUidls.Contains(uid))
                            continue;

                        labelMessage.Text = $"{i + 1}件目のメール受信中";
                        statusStrip1.Refresh();

                        var message = await inbox.GetMessageAsync(summary.UniqueId);

                        // ★ UID をファイル名にする（IMAP の正しい方法）
                        string mailName = uid + "_unread.eml";

                        string inboxPath = Path.Combine(folderManager.Inbox.FullPath, mailName);
                        Directory.CreateDirectory(folderManager.Inbox.FullPath);

                        await Task.Run(() => message.WriteTo(inboxPath));

                        // ★ Mail オブジェクト生成
                        Mail mail = new Mail(
                            message.From.ToString(),
                            message.Cc.ToString(),
                            message.Bcc.ToString(),
                            message.Subject,
                            message.TextBody ?? "",
                            null,
                            message.Date.ToString(),
                            mailName,
                            uid,
                            true
                        );

                        // ★ パス設定（必須）
                        mail.mailPath = inboxPath;

                        // ★ mailCache に登録（必須）
                        mailCache[inboxPath] = mail;

                        // 振り分け処理
                        ApplyRules(mail);

                        localUidls.Add(uid);
                        mailCount++;
                    }

                    await client.DisconnectAsync(true);
                }

                // ★ UIDL 永続化（最後に1回だけ）
                SaveUidls();

                labelMessage.Text = "メール受信完了";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                labelMessage.Text = "メール受信エラー : " + ex.Message;
                statusStrip1.Refresh();
            }

            // 通知音など
            if (mailCount > 0)
            {
                if (Mail.alertSound && !string.IsNullOrWhiteSpace(Mail.alertSoundFile))
                {
                    try
                    {
                        if (File.Exists(Mail.alertSoundFile))
                        {
                            using (var p = new SoundPlayer(Mail.alertSoundFile))
                            {
                                p.Play();
                            }
                        }
                        else
                        {
                            logger.Error($"Alert sound file not found: {Mail.alertSoundFile}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Alert sound error: " + ex.Message);
                    }
                }
                labelMessage.Text = $"{mailCount}件の新着メールがあります";
                statusStrip1.Refresh();
            }

            toolMailProgress.Value = 100;
            await Task.Delay(300);
            toolMailProgress.Value = 0;
            toolMailProgress.Visible = false;

            BuildTree();
            UpdateView();
        }

        private void menuUndoMail_Click(object sender, EventArgs e)
        {
            // ★ Trash フォルダを走査して .meta を探す
            var metaFiles = Directory.GetFiles(folderManager.Trash.FullPath, "*.meta");
            if (metaFiles.Length == 0)
                return;

            foreach (var metaFile in metaFiles)
            {
                var json = File.ReadAllText(metaFile);
                var meta = JsonConvert.DeserializeObject<UndoMeta>(json);

                // ★ 対象メールのパス
                string newPath = meta.NewPath;
                string oldPath = meta.OldPath;

                if (!File.Exists(newPath))
                    continue;

                // ★ メールを元の場所へ戻す
                Directory.CreateDirectory(Path.GetDirectoryName(oldPath));
                File.Move(newPath, oldPath);

                // ★ mailCache 更新
                if (mailCache.ContainsKey(newPath))
                    mailCache.Remove(newPath);

                Mail mail = LoadSingleMail(oldPath);
                mailCache[oldPath] = mail;

                // ★ フォルダ復元
                mail.Folder = folderManager.GetFolderByType(meta.OldFolder);

                // ★ 保存
                SaveMail(mail);

                // ★ .meta 削除
                File.Delete(metaFile);
            }

            UpdateView();
        }

        // 音声読み上げの準備
        private SpeechSynthesizer synth = new SpeechSynthesizer();

        private async void menuSpeechMail_Click(object sender, EventArgs e)
        {
            // 選択中のメールを音声で読み上げる
            if (currentMail != null)
            {
                string toSpeak = $"件名: {currentMail.subject}。差出人: {currentMail.address}。本文: {currentMail.body}";
                try
                {
                    if (await IsVoiceVoxRunning())
                    {
                        await SpeakWithVoiceVox(toSpeak, 2);
                    }
                    else
                    {
                        synth.SpeakAsync(toSpeak);
                    }
                }
                catch (Exception ex)
                {
                    logger.Debug($"出力エラー:{ex.Message}");
                }
            }
        }

        private static readonly HttpClient client = new HttpClient();

        public async Task SpeakWithVoiceVox(string text, int speakerId = 2)
        {
            // 1. audio_query
            var query = await client.PostAsync($"http://127.0.0.1:50021/audio_query?text={Uri.EscapeDataString(text)}&speaker={speakerId}", null);
            var queryJson = await query.Content.ReadAsStringAsync();

            // speedScale を上げて高速化
            dynamic queryObj = Newtonsoft.Json.JsonConvert.DeserializeObject(queryJson);
            queryObj.speedScale = 1.3;
            queryJson = Newtonsoft.Json.JsonConvert.SerializeObject(queryObj);

            // 2. synthesis
            var audio = await client.PostAsync($"http://127.0.0.1:50021/synthesis?speaker={speakerId}", new StringContent(queryJson, Encoding.UTF8, "application/json"));

            // 3. WAV 再生（MemoryStream で高速化）
            var stream = await audio.Content.ReadAsStreamAsync();
            var mem = new MemoryStream();
            await stream.CopyToAsync(mem);
            mem.Position = 0;

            var player = new SoundPlayer(mem);
            player.Play();
        }

        private async Task<bool> IsVoiceVoxRunning()
        {
            try
            {
                var res = await client.GetAsync("http://127.0.0.1:50021/version");
                return res.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        private void toolSearchButton_Click(object sender, EventArgs e)
        {
            string keyword = textSearch.Text.Trim();
            currentKeyword = keyword;

            if (string.IsNullOrEmpty(keyword))
            {
                UpdateListView();
                return;
            }

            // ★ Tag は MailFolder なので string ではない
            MailFolder folder = treeMain.SelectedNode?.Tag as MailFolder;
            if (folder == null)
            {
                UpdateListView();
                return;
            }

            // ★ 選択フォルダのメールをロード
            List<Mail> sourceList = LoadEmlFolder(folder);

            // ★ キーワードでフィルタ
            var filtered = sourceList
                .Where(m =>
                    (!string.IsNullOrEmpty(m.address) && m.address.Contains(keyword)) ||
                    (!string.IsNullOrEmpty(m.subject) && m.subject.Contains(keyword)) ||
                    (!string.IsNullOrEmpty(m.body) && m.body.Contains(keyword))
                )
                .ToList();

            // ★ 検索結果表示
            ShowSearchResult(filtered);
        }

        private void ShowSearchResult(List<Mail> list)
        {
            listMain.BeginUpdate();
            listMain.Items.Clear();

            var baseFont = listMain.Font;

            foreach (var mail in list)
            {
                // ★ 常に差出人を表示
                ListViewItem item = new ListViewItem(mail.from);

                item.SubItems.Add(mail.subject);
                item.SubItems.Add(FormatReceivedDate(mail.date));

                long sizeBytes = GetMailFileSize(mail);
                item.SubItems.Add(FormatSize(sizeBytes));

                item.SubItems.Add(mail.mailName);

                if (HasAttachment(mail))
                    item.ImageKey = "attach";

                item.Tag = mail;

                // 未読は太字＋背景色
                if (mail.notReadYet)
                {
                    item.BackColor = Color.FromArgb(0xE8, 0xF4, 0xFF);
                    item.Font = new Font(baseFont, FontStyle.Bold);
                }
                else
                {
                    item.Font = new Font(baseFont, FontStyle.Regular);
                }

                listMain.Items.Add(item);
            }

            listMain.EndUpdate();
        }

        private void listMain_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // 既存の sorter を使う
            if (listViewItemSorter == null)
                listViewItemSorter = ListViewItemComparer.Default;

            listViewItemSorter.Column = e.Column;
            listMain.ListViewItemSorter = listViewItemSorter;

            listMain.Sort();
        }

        private void toolFilterCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateListView();
        }

        private bool HasAttachment(Mail mail)
        {
            string path = ResolveMailPath(mail);
            if (string.IsNullOrEmpty(path))
                return false;

            // 送信メール（.mail）
            if (path.EndsWith(".eml") && string.IsNullOrEmpty(mail.uidl))
                return !string.IsNullOrEmpty(mail.atach);

            // 受信メール（.eml）
            if (path.EndsWith(".eml") && File.Exists(path))
            {
                var msg = MimeKit.MimeMessage.Load(path);
                return msg.Attachments.Any();
            }

            return false;
        }

        public string ResolveMailPath(Mail mail)
        {
            if (mail == null)
                return null;

            // ★ 新方式（MailFolder）だけで十分
            if (mail.Folder != null)
                return Path.Combine(mail.Folder.FullPath, mail.mailName);

            // ★ fallback（Undo直後などの一時的な null 対策）
            return Path.Combine(folderManager.Inbox.FullPath, mail.mailName);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);

            // richTextBody がフォーカスされていて Ctrl が押されているときだけズーム
            if (richTextBody.Focused && ModifierKeys == Keys.Control)
            {
                float size = richTextBody.Font.Size + (e.Delta > 0 ? 1 : -1);
                size = Math.Max(8, Math.Min(48, size)); // 最小8pt 最大48pt

                richTextBody.Font = new Font(richTextBody.Font.FontFamily, size);
            }
        }

        private void ColorizeQuoteLines()
        {
            // カーソル位置を保存
            int selStart = richTextBody.SelectionStart;
            int selLength = richTextBody.SelectionLength;

            richTextBody.SuspendLayout();

            // 全体を標準色に戻す
            richTextBody.SelectAll();
            richTextBody.SelectionColor = Color.Black;

            string[] lines = richTextBody.Lines;
            int pos = 0;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                // 先頭の '>' の数を数える
                int depth = 0;
                foreach (char c in line)
                {
                    if (c == '>') depth++;
                    else if (c == ' ') continue; // "> > >" のようなパターンにも対応
                    else break;
                }

                if (depth > 0)
                {
                    // 深さに応じた色を決める
                    Color color = GetQuoteColor(depth);

                    // 色付け
                    richTextBody.Select(pos, line.Length);
                    richTextBody.SelectionColor = color;
                }

                pos += line.Length + 1; // 改行分 +1
            }

            // カーソル位置を復元
            richTextBody.SelectionStart = selStart;
            richTextBody.SelectionLength = selLength;

            richTextBody.ResumeLayout();
        }

        private Color GetQuoteColor(int depth)
        {
            switch (depth)
            {
                case 1: return Color.FromArgb(0, 90, 200);      // 青
                case 2: return Color.FromArgb(0, 130, 160);     // 青緑
                case 3: return Color.FromArgb(0, 160, 120);     // 緑寄り
                case 4: return Color.FromArgb(120, 120, 120);   // グレー
                default: return Color.FromArgb(160, 160, 160);  // それ以上は薄いグレー
            }
        }

        private void menuHelpView_Click(object sender, EventArgs e)
        {
            // ヘルプを表示する
            OpenHelp();
        }

        private void OpenHelp()
        {
            try
            {
                string helpPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "help", "MizuMail.html");

                if (File.Exists(helpPath))
                {
                    System.Diagnostics.Process.Start(helpPath);
                }
                else
                {
                    MessageBox.Show("ヘルプファイルが見つかりません。\n" + helpPath, "ヘルプ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ヘルプを開けませんでした。\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SaveMail(Mail mail)
        {
            if (mail == null)
                return;

            // ★ 保存先フォルダが未設定なら Inbox
            if (mail.Folder == null)
                mail.Folder = folderManager.Inbox;

            // ★ ファイル名が未設定なら生成
            if (string.IsNullOrEmpty(mail.mailName))
                mail.mailName = $"{DateTime.Now.Ticks}.eml";

            // ★ 未読/既読のファイル名調整
            if (mail.notReadYet)
            {
                if (!mail.mailName.EndsWith("_unread.eml", StringComparison.OrdinalIgnoreCase))
                    mail.mailName = Path.GetFileNameWithoutExtension(mail.mailName) + "_unread.eml";
            }
            else
            {
                if (mail.mailName.EndsWith("_unread.eml", StringComparison.OrdinalIgnoreCase))
                    mail.mailName = mail.mailName.Replace("_unread.eml", ".eml");
                else if (!mail.mailName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase))
                    mail.mailName = Path.GetFileNameWithoutExtension(mail.mailName) + ".eml";
            }

            string savePath = Path.Combine(mail.Folder.FullPath, mail.mailName);

            try
            {
                var message = new MimeMessage();

                // ★ 差出人（From）を上書きしない
                if (!string.IsNullOrEmpty(mail.from))
                {
                    // 受信メールの From をそのまま使う
                    message.From.Add(MailboxAddress.Parse(mail.from));
                }
                else
                {
                    // 新規作成・下書き・送信メール
                    message.From.Add(new MailboxAddress(Mail.fromName, Mail.userAddress));
                }

                // To
                if (!string.IsNullOrWhiteSpace(mail.address))
                {
                    foreach (var addr in mail.address.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
                        message.To.Add(MailboxAddress.Parse(addr.Trim()));
                }

                // Cc
                if (!string.IsNullOrWhiteSpace(mail.ccaddress))
                {
                    foreach (var addr in mail.ccaddress.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
                        message.Cc.Add(MailboxAddress.Parse(addr.Trim()));
                }

                // Bcc
                if (!string.IsNullOrWhiteSpace(mail.bccaddress))
                {
                    foreach (var addr in mail.bccaddress.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
                        message.Bcc.Add(MailboxAddress.Parse(addr.Trim()));
                }

                // 件名
                message.Subject = mail.subject ?? "";

                // 本文＋添付
                var builder = new BodyBuilder
                {
                    TextBody = mail.body ?? ""
                };

                if (!string.IsNullOrWhiteSpace(mail.atach))
                {
                    foreach (var file in mail.atach.Split(';'))
                    {
                        var trimmed = file.Trim();
                        if (!string.IsNullOrEmpty(trimmed) && File.Exists(trimmed))
                            builder.Attachments.Add(trimmed);
                    }
                }

                message.Body = builder.ToMessageBody();

                // 日付
                if (DateTime.TryParse(mail.date, out var dt))
                    message.Date = new DateTimeOffset(dt);
                else
                    message.Date = DateTimeOffset.Now;

                // X-MizuMail-Draft
                message.Headers["X-MizuMail-Draft"] = mail.isDraft ? "1" : "0";

                // X-Mailer
                message.Headers.Add("X-Mailer", "MizuMail " + System.Windows.Forms.Application.ProductVersion);

                // 保存
                Directory.CreateDirectory(mail.Folder.FullPath);
                message.WriteTo(savePath);
            }
            catch (Exception ex)
            {
                logger.Error($"SaveMail error: {savePath} : {ex.Message}");
            }
        }

        /// <summary>
        /// 送信メールの読み出しメソッド
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public Mail LoadSendMail(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            try
            {
                string json = File.ReadAllText(filePath, Encoding.UTF8);
                Mail mail = JsonConvert.DeserializeObject<Mail>(json);

                // 読み込んだ Mail に mailName を付与（ファイル名から）
                if (mail != null)
                {
                    mail.mailName = Path.GetFileName(filePath);
                }

                return mail;
            }
            catch
            {
                return null;
            }
        }

        public class UrlCheckResult
        {
            public bool IsSafe { get; set; }
            public string Reason { get; set; }
        }

        private UrlCheckResult CheckUrlSafety(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                return new UrlCheckResult { IsSafe = false, Reason = "URL が空です" };

            // 1. http/https 以外は拒否
            if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                return new UrlCheckResult { IsSafe = false, Reason = "http/https 以外のプロトコル" };

            // 2. javascript: や file: を拒否
            if (url.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase))
                return new UrlCheckResult { IsSafe = false, Reason = "JavaScript スキーム" };

            if (url.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
                return new UrlCheckResult { IsSafe = false, Reason = "ローカルファイルアクセス" };

            // 3. URL が長すぎる（フィッシングの典型）
            if (url.Length > 2048)
                return new UrlCheckResult { IsSafe = false, Reason = "URL が異常に長い" };

            // 4. ドメイン部分を抽出
            Uri uri;
            if (!Uri.TryCreate(url, UriKind.Absolute, out uri))
                return new UrlCheckResult { IsSafe = false, Reason = "URL が不正" };

            string host = uri.Host;

            // 5. 国際化ドメインを punycode に変換
            try
            {
                var idn = new System.Globalization.IdnMapping();
                host = idn.GetAscii(host);
            }
            catch
            {
                return new UrlCheckResult { IsSafe = false, Reason = "国際化ドメインが不正" };
            }

            // 6. 似た文字の混在（フィッシング対策）
            if (ContainsMixedScripts(host))
                return new UrlCheckResult { IsSafe = false, Reason = "ドメインに混在文字（フィッシングの可能性）" };

            return new UrlCheckResult { IsSafe = true };
        }

        private bool ContainsMixedScripts(string text)
        {
            bool hasLatin = text.Any(c => c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z');
            bool hasNonLatin = text.Any(c => c > 127);

            return hasLatin && hasNonLatin;
        }

        private void menuCreateFolder_Click(object sender, EventArgs e)
        {
            TreeNode node = treeMain.SelectedNode;
            if (node == null)
                return;

            // ★ Tag は MailFolder
            MailFolder parent = node.Tag as MailFolder;
            if (parent == null)
                return;

            // inbox / send / trash 以外はサブフォルダ作成OK
            if (parent.Type == FolderType.Send || parent.Type == FolderType.Draft || parent.Type == FolderType.Trash)
            {
                MessageBox.Show("このフォルダにはサブフォルダを作成できません。");
                return;
            }

            string name = Prompt.ShowDialog("フォルダ名を入力してください", "新規フォルダ");
            if (string.IsNullOrWhiteSpace(name))
                return;

            if (!IsValidFolderName(name))
            {
                MessageBox.Show("フォルダ名に使用できない文字が含まれています。");
                return;
            }

            // ★ 新しいフォルダのパス
            string newDir = Path.Combine(parent.FullPath, name);

            if (Directory.Exists(newDir))
            {
                MessageBox.Show("同じ名前のフォルダが既に存在します。");
                return;
            }

            Directory.CreateDirectory(newDir);

            // ★ 新しい MailFolder を作成
            var newFolder = new MailFolder(name, newDir, FolderType.InboxSub);

            // ★ TreeView に追加
            TreeNode newNode = new TreeNode(name);
            newNode.Tag = newFolder;
            node.Nodes.Add(newNode);
            node.Expand();

            // ★ FolderManager に登録（InboxSubFolders は第一階層だけ管理）
            if (parent.Type == FolderType.Inbox)
                folderManager.InboxSubFolders.Add(newFolder);

            BuildTree();
            UpdateListView();
        }

        private void treeMain_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label == null)
                return;

            string newName = e.Label.Trim();

            if (!IsValidFolderName(newName))
            {
                MessageBox.Show("フォルダ名に使用できない文字が含まれています。");
                e.CancelEdit = true;
                return;
            }

            TreeNode node = e.Node;
            MailFolder folder = node.Tag as MailFolder;

            if (folder == null)
            {
                e.CancelEdit = true;
                return;
            }

            if (folder.Type == FolderType.Inbox ||
                folder.Type == FolderType.Send ||
                folder.Type == FolderType.Draft ||
                folder.Type == FolderType.Trash)
            {
                MessageBox.Show("このフォルダ名は変更できません。");
                e.CancelEdit = true;
                return;
            }

            if (string.IsNullOrEmpty(folder.FullPath))
            {
                e.CancelEdit = true;
                return;
            }

            DirectoryInfo parentInfo = Directory.GetParent(folder.FullPath);
            if (parentInfo == null)
            {
                e.CancelEdit = true;
                return;
            }

            string parentDir = parentInfo.FullName;
            string newDir = Path.Combine(parentDir, newName);

            if (Directory.Exists(newDir))
            {
                MessageBox.Show("同じ名前のフォルダが既に存在します。");
                e.CancelEdit = true;
                return;
            }

            try
            {
                Directory.Move(folder.FullPath, newDir);
            }
            catch
            {
                MessageBox.Show("フォルダ名を変更できませんでした。");
                e.CancelEdit = true;
                return;
            }

            // ★ FullPath は触らない（不変）
            folder.Name = newName;
            folder.DisplayName = newName;

            // ★ フォルダ一覧をディスクから再構築
            TreeNode root = treeMain.Nodes[0];
            TreeNode inboxNode = root.Nodes[0];
            MailFolder inboxFolder = folderManager.Inbox;

            LoadInboxFolders(inboxNode, inboxFolder);
            UpdateTreeView();
            UpdateListView();
        }

        private void listMain_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (e.Item is ListViewItem item && item.Tag is Mail mail)
            {
                DoDragDrop(mail, DragDropEffects.Move);
            }
        }

        private void treeMain_DragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(Mail)))
                return;

            Mail mail = e.Data.GetData(typeof(Mail)) as Mail;
            if (mail == null)
                return;

            Point pt = treeMain.PointToClient(new Point(e.X, e.Y));
            TreeNode node = treeMain.GetNodeAt(pt);
            if (node == null)
                return;

            MailFolder targetFolder = node.Tag as MailFolder;
            if (targetFolder == null)
                return;

            // ★ メール移動だけ行う（UI更新は絶対にしない）
            if (targetFolder.Type == FolderType.Trash)
            {
                MoveMailWithUndo(mail, folderManager.Trash);
            }
            else if (mail.Folder != targetFolder)
            {
                MoveMailWithUndo(mail, targetFolder);
            }

            // ★ UI更新は DragDrop 完了後に行う（ここが重要）
            this.BeginInvoke(new Action(delegate
            {
                // TreeView がまだ構築されていない可能性がある
                if (treeMain.Nodes.Count == 0)
                    return;

                BuildTree();

                // ★ SelectedNode が null なら Inbox を選択
                if (treeMain.SelectedNode == null)
                {
                    if (treeMain.Nodes.Count > 0 && treeMain.Nodes[0].Nodes.Count > 0)
                        treeMain.SelectedNode = treeMain.Nodes[0].Nodes[0];
                    else
                        return;
                }

                // ★ Tag が MailFolder でない場合も防御
                MailFolder folder = treeMain.SelectedNode.Tag as MailFolder;
                if (folder == null)
                    return;

                UpdateListView();
            }));
        }

        private TreeNode FindNodeByFolder(MailFolder folder)
        {
            foreach (TreeNode node in treeMain.Nodes[0].Nodes)
            {
                if (node.Tag == folder)
                    return node;
            }
            return null;
        }

        private void MoveMailToFolder(Mail mail, MailFolder targetFolder)
        {
            string oldPath = ResolveMailPath(mail);
            string newPath = Path.Combine(targetFolder.FullPath, mail.mailName);

            // 物理移動
            File.Move(oldPath, newPath);

            // Mail の状態更新
            mail.Folder = targetFolder;

            // mailCache のキー更新
            mailCache.Remove(oldPath);
            mailCache[newPath] = mail;
        }

        private void MoveMailToTrash(Mail mail)
        {
            string oldPath = ResolveMailPath(mail);
            string trashPath = Path.Combine(folderManager.Trash.FullPath, mail.mailName);

            // Undo 情報を保存
            string metaPath = trashPath + ".meta";
            File.WriteAllLines(metaPath, new[]
            {
                "OriginalFolder=" + mail.Folder.FullPath,
                "OriginalName=" + mail.mailName
            });

            // 物理移動
            MoveMailWithUndo(mail, folderManager.Trash);
            BuildTree();

            // Mail の状態更新
            mail.Folder = folderManager.Trash;

            // mailCache のキー更新
            mailCache.Remove(oldPath);
            mailCache[trashPath] = mail;
        }

        public List<Mail> LoadEmlFolder(MailFolder folder)
        {
            var list = new List<Mail>();

            if (!Directory.Exists(folder.FullPath))
                return list;

            // ★ フォルダ読み込み前に mailCache を整理
            //   このフォルダ配下のキャッシュだけ削除する
            var removeKeys = mailCache.Keys
                .Where(k => k.StartsWith(folder.FullPath, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var key in removeKeys)
                mailCache.Remove(key);

            string[] files = Directory.GetFiles(folder.FullPath, "*.eml");

            foreach (string file in files)
            {
                try
                {
                    var message = MimeMessage.Load(file);

                    var mail = new Mail();
                    mail.mailName = Path.GetFileName(file);
                    mail.Folder = folder;

                    // ★ これが絶対必要
                    mail.mailPath = file;

                    mail.notReadYet = mail.mailName.EndsWith("_unread.eml", StringComparison.OrdinalIgnoreCase);

                    var fromMailbox = message.From.Mailboxes.FirstOrDefault();
                    mail.from = fromMailbox != null
                        ? (!string.IsNullOrEmpty(fromMailbox.Name)
                            ? fromMailbox.Name + " <" + fromMailbox.Address + ">"
                            : fromMailbox.Address)
                        : "(差出人なし)";

                    if (message.To != null && message.To.Mailboxes.Any())
                    {
                        var listTo = new List<string>();
                        foreach (var mb in message.To.Mailboxes)
                        {
                            listTo.Add(!string.IsNullOrEmpty(mb.Name)
                                ? mb.Name + " <" + mb.Address + ">"
                                : mb.Address);
                        }
                        mail.address = string.Join("; ", listTo.ToArray());
                    }
                    else
                    {
                        mail.address = "(宛先なし)";
                    }

                    mail.subject = message.Subject ?? "";

                    if (message.Date != DateTimeOffset.MinValue)
                        mail.date = message.Date.LocalDateTime.ToString("yyyy/MM/dd HH:mm:ss");
                    else
                        mail.date = "";

                    mail.hasAtach = message.Attachments != null && message.Attachments.Any();
                    mail.body = "";
                    mail.message = message;
                    mail.isDraft = message.Headers["X-MizuMail-Draft"] == "1";

                    // ★ 本文読み込み（TextPart / Multipart 両対応）
                    string bodyText = "";

                    if (message.Body is TextPart)
                    {
                        bodyText = ((TextPart)message.Body).Text;
                    }
                    else if (message.Body is Multipart)
                    {
                        var mp = (Multipart)message.Body;
                        for (int i = 0; i < mp.Count; i++)
                        {
                            if (mp[i] is TextPart)
                            {
                                bodyText = ((TextPart)mp[i]).Text;
                                break;
                            }
                        }
                    }

                    mail.body = bodyText ?? "";

                    list.Add(mail);

                    // ★ mailCache に登録（これが CountUnread と同期の鍵）
                    mailCache[file] = mail;
                }
                catch (Exception ex)
                {
                    logger.Error("LoadEmlFolder error: " + file + " : " + ex.Message);
                }
            }

            return list;
        }

        public MimeMessage LoadFullMail(Mail mail)
        {
            string file = Path.Combine(mail.Folder.FullPath, mail.mailName);
            return MimeMessage.Load(file);
        }

        private Mail LoadMailHeaderOnly(string path)
        {
            var mail = new Mail();
            mail.mailName = Path.GetFileName(path);

            // ★ メールのエンコーディングを自動判定
            var raw = File.ReadAllBytes(path);
            var text = DecodeMimeMessage(raw);

            using (var reader = new StringReader(text))
            {
                string line;
                while (!string.IsNullOrEmpty(line = reader.ReadLine()))
                {
                    if (line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
                        mail.subject = DecodeMimeHeader(line.Substring(8).Trim());

                    else if (line.StartsWith("From:", StringComparison.OrdinalIgnoreCase))
                        mail.address = DecodeMimeHeader(line.Substring(5).Trim());

                    else if (line.StartsWith("Date:", StringComparison.OrdinalIgnoreCase))
                        mail.date = line.Substring(5).Trim();
                }
            }

            return mail;
        }

        private string DecodeMimeHeader(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            var regex = new System.Text.RegularExpressions.Regex(
                @"=\?(?<charset>[^?]+)\?(?<encoding>[bBqQ])\?(?<data>[^?]+)\?=");

            return regex.Replace(value, m =>
            {
                string charset = m.Groups["charset"].Value;
                string encoding = m.Groups["encoding"].Value.ToUpper();
                string data = m.Groups["data"].Value;

                try
                {
                    byte[] bytes;

                    if (encoding == "B")
                    {
                        bytes = Convert.FromBase64String(data);
                    }
                    else
                    {
                        bytes = DecodeQuotedPrintableBytes(data);
                    }

                    var enc = Encoding.GetEncoding(charset);
                    return enc.GetString(bytes);
                }
                catch
                {
                    return m.Value;
                }
            });
        }

        private byte[] DecodeQuotedPrintableBytes(string input)
        {
            input = input.Replace("=\r\n", ""); // ソフト改行を削除

            var ms = new MemoryStream();
            for (int i = 0; i < input.Length; i++)
            {
                if (input[i] == '=' && i + 2 < input.Length)
                {
                    string hex = input.Substring(i + 1, 2);
                    ms.WriteByte(Convert.ToByte(hex, 16));
                    i += 2;
                }
                else
                {
                    ms.WriteByte((byte)input[i]);
                }
            }
            return ms.ToArray();
        }

        public void EnsureMailBodyLoaded(Mail mail)
        {
            if (mail.body != null)
                return;

            string path = ResolveMailPath(mail);
            var raw = File.ReadAllBytes(path);
            mail.body = DecodeMimeMessage(raw);
        }

        private string DecodeMimeMessage(byte[] raw)
        {
            // ★ まずは UTF-8
            try { return Encoding.UTF8.GetString(raw); } catch { }

            // ★ ISO-2022-JP
            try { return Encoding.GetEncoding("iso-2022-jp").GetString(raw); } catch { }

            // ★ Shift_JIS
            try { return Encoding.GetEncoding("shift_jis").GetString(raw); } catch { }

            // ★ EUC-JP
            try { return Encoding.GetEncoding("euc-jp").GetString(raw); } catch { }

            // ★ 最後の fallback
            return Encoding.Default.GetString(raw);
        }

        private void treeMain_DragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(Mail)))
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            Point pt = treeMain.PointToClient(new Point(e.X, e.Y));
            TreeNode node = treeMain.GetNodeAt(pt);

            if (node == null)
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            MailFolder folder = node.Tag as MailFolder;
            if (folder == null)
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            // ★ Trash だけは Drop 可能（MoveMailToTrash で処理）
            // ★ それ以外のフォルダ（Inbox, Send, サブフォルダ）は全部 Move 可能
            e.Effect = DragDropEffects.Move;
        }

        private void DeleteInboxSubFolder(string folderName)
        {
            string folderPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", folderName);

            if (!Directory.Exists(folderPath))
                return;

            // ★ ごみ箱フォルダ
            string trashDir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trash");
            Directory.CreateDirectory(trashDir);

            // ★ フォルダ内の .eml をすべてごみ箱へ移動
            foreach (var file in Directory.GetFiles(folderPath, "*.eml"))
            {
                string fileName = Path.GetFileName(file);
                string dest = Path.Combine(trashDir, fileName);

                // 同名ファイルがあれば上書きしないようにリネーム
                if (File.Exists(dest))
                {
                    string newName = $"{Path.GetFileNameWithoutExtension(fileName)}_{Guid.NewGuid()}.eml";
                    dest = Path.Combine(trashDir, newName);
                }

                File.Move(file, dest);
            }

            // ★ 実フォルダ削除
            Directory.Delete(folderPath, true);
        }

        private void DeleteFolderRecursive(string folder)
        {
            string folderPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", folder.Replace("/", "\\"));

            if (!Directory.Exists(folderPath))
                return;

            // ごみ箱フォルダ
            string trashDir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trash");
            Directory.CreateDirectory(trashDir);

            // ★ .eml をすべてごみ箱へ移動
            foreach (var file in Directory.GetFiles(folderPath, "*.eml", SearchOption.AllDirectories))
            {
                string fileName = Path.GetFileName(file);
                string dest = Path.Combine(trashDir, fileName);

                if (File.Exists(dest))
                {
                    string newName = $"{Path.GetFileNameWithoutExtension(fileName)}_{Guid.NewGuid()}.eml";
                    dest = Path.Combine(trashDir, newName);
                }

                File.Move(file, dest);
            }

            // ★ フォルダごと削除
            Directory.Delete(folderPath, true);
        }

        private void menuRenameFolder_Click(object sender, EventArgs e)
        {
            TreeNode node = treeMain.SelectedNode;
            if (node == null)
                return;

            TreeNode inboxNode = treeMain.Nodes[0].Nodes[0];

            // ★ 受信メール直下のフォルダ以外は名前変更禁止
            if (node.Parent != inboxNode)
            {
                MessageBox.Show("フォルダ名は「受信メール」の下にあるフォルダだけ変更できます。", "フォルダ名変更", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ★ ラベル編集開始
            node.BeginEdit();
        }

        private void OpenSendMailEditor(Mail mail)
        {
            FormMailCreate form = new FormMailCreate();
            form.Text = "編集";

            form.textMailSubject.Text = mail.subject;
            form.textMailTo.Text = mail.address;
            form.textMailCc.Text = mail.ccaddress;
            form.textMailBcc.Text = mail.bccaddress;
            form.textMailBody.Text = mail.body.TrimEnd('\r', '\n');

            DialogResult ret = form.ShowDialog();

            if (ret == DialogResult.OK)
            {
                string to = form.textMailTo.Text;
                string subject = form.textMailSubject.Text;
                string body = form.textMailBody.Text;
                string cc = form.textMailCc.Text;
                string bcc = form.textMailBcc.Text;
                string atach = string.Join(";", form.buttonAttachList.DropDownItems.Cast<ToolStripItem>().Select(i => i.Text));

                if (!string.IsNullOrEmpty(to) || !string.IsNullOrEmpty(body))
                {
                    if (string.IsNullOrEmpty(subject))
                        subject = "無題";

                    mail.address = to;
                    mail.ccaddress = cc;
                    mail.bccaddress = bcc;
                    mail.subject = subject;
                    mail.body = body;
                    mail.atach = atach;
                    mail.isDraft = true;
                    mail.notReadYet = true;
                    mail.Folder = folderManager.Send;

                    SaveMail(mail);
                }
            }

            UpdateListView();
        }

        private void CreateSubFolder(TreeNode parentNode)
        {
            // 親フォルダの実パス（例：inbox）
            string parentFolder = parentNode.Tag as string;

            // 新しいフォルダ名を入力
            string newName = Prompt.ShowDialog("フォルダ名を入力してください", "新規フォルダ");
            if (string.IsNullOrWhiteSpace(newName))
                return;

            newName = newName.Trim();

            // ★ バリデーション
            if (!IsValidFolderName(newName))
            {
                MessageBox.Show("フォルダ名に使用できない文字が含まれています。");
                return;
            }

            // ★ 新しいフォルダの実パス
            string newFolder = parentFolder + "/" + newName;

            // ★ ディレクトリ作成
            string dir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", newFolder.Replace("/", "\\"));
            if (Directory.Exists(dir))
            {
                MessageBox.Show("同じ名前のフォルダが既に存在します。");
                return;
            }

            Directory.CreateDirectory(dir);

            // ★ TreeNode 作成
            TreeNode node = new TreeNode(newName);
            node.Tag = newFolder;   // ← これが最重要

            parentNode.Nodes.Add(node);
            parentNode.Expand();
        }

        private bool IsValidFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            // 前後の空白禁止
            if (name != name.Trim())
                return false;

            // 禁止文字
            char[] invalid = Path.GetInvalidFileNameChars();
            if (name.Any(c => invalid.Contains(c)))
                return false;

            // スラッシュ禁止（階層を勝手に作らせない）
            if (name.Contains("/") || name.Contains("\\"))
                return false;

            return true;
        }

        private void ShowMailboxInfo()
        {
            // メールフォルダ一覧の場合はメールボックスの情報を出して終了
            ListViewItem item = new ListViewItem(Mail.fromName);
            item.SubItems.Add(Mail.userAddress);
            // メールフォルダの最終更新日を表示
            item.SubItems.Add(System.IO.Directory.GetLastWriteTime(System.Windows.Forms.Application.StartupPath + "\\mbox\\inbox\\").ToString("yyyy/MM/dd HH:mm:ss"));
            // メールフォルダのサイズを表示
            var directorySize = 0L;
            GetDirectorySize(System.Windows.Forms.Application.StartupPath + "\\mbox\\", ref directorySize);
            item.SubItems.Add(FormatSize(directorySize));
            listMain.Items.Add(item);
            mailBoxViewFlag = true;
        }

        private void menuDeleteFolder_Click(object sender, EventArgs e)
        {
            TreeNode node = treeMain.SelectedNode;
            if (node == null)
                return;

            MailFolder folder = node.Tag as MailFolder;
            if (folder == null)
                return;

            // システムフォルダ禁止
            if (folder.Type == FolderType.Inbox ||
                folder.Type == FolderType.Send ||
                folder.Type == FolderType.Draft ||
                folder.Type == FolderType.Trash)
            {
                MessageBox.Show("このフォルダは削除できません。");
                return;
            }

            // 親フォルダを取得
            MailFolder parent = GetParentFolder(folder);
            if (parent != null)
            {
                parent.SubFolders.Remove(folder);   // ★ これが必須
            }

            folderManager.InboxSubFolders.Remove(folder);

            // 実フォルダ削除
            try
            {
                Directory.Delete(folder.FullPath, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("フォルダを削除できませんでした。\n" + ex.Message);
                return;
            }

            // ★ UI更新はここでまとめて行う
            BuildTree();
            UpdateListView();
        }

        private MailFolder GetParentFolder(MailFolder folder)
        {
            // Inbox の直下
            foreach (var sub in folderManager.Inbox.SubFolders)
            {
                if (sub == folder)
                    return folderManager.Inbox;
            }

            // 再帰で探す
            return FindParentRecursive(folderManager.Inbox, folder);
        }

        private MailFolder FindParentRecursive(MailFolder parent, MailFolder target)
        {
            foreach (var sub in parent.SubFolders)
            {
                if (sub == target)
                    return parent;

                var found = FindParentRecursive(sub, target);
                if (found != null)
                    return found;
            }
            return null;
        }

        private void MoveFolderMailsToTrashRecursive(MailFolder folder)
        {
            // ★ このフォルダ内のメールを移動
            var mails = LoadEmlFolder(folder);

            foreach (var mail in mails)
            {
                string oldPath = Path.Combine(folder.FullPath, mail.mailName);
                string newPath = Path.Combine(folderManager.Trash.FullPath, mail.mailName);

                try
                {
                    File.Move(oldPath, newPath);
                }
                catch { }
            }

            // ★ サブフォルダも再帰的に処理
            foreach (var subDir in Directory.GetDirectories(folder.FullPath))
            {
                string name = Path.GetFileName(subDir);
                var subFolder = new MailFolder(name, subDir, FolderType.InboxSub);

                MoveFolderMailsToTrashRecursive(subFolder);
            }
        }

        private void MoveFolderMailsToTrash(MailFolder folder)
        {
            var mails = LoadEmlFolder(folder);

            foreach (var mail in mails)
            {
                string oldPath = Path.Combine(folder.FullPath, mail.mailName);
                string newPath = Path.Combine(folderManager.Trash.FullPath, mail.mailName);

                try
                {
                    File.Move(oldPath, newPath);
                }
                catch { }
            }
        }

        private void LoadInboxFolders(TreeNode inboxNode, MailFolder inboxFolder)
        {
            inboxNode.Nodes.Clear();
            inboxFolder.SubFolders.Clear();
            folderManager.InboxSubFolders.Clear();

            foreach (var dir in Directory.GetDirectories(inboxFolder.FullPath))
            {
                string name = Path.GetFileName(dir);
                var folder = new MailFolder(name, dir, FolderType.InboxSub);

                inboxFolder.SubFolders.Add(folder);
                folderManager.InboxSubFolders.Add(folder);

                TreeNode node = new TreeNode(name);
                node.Tag = folder;
                inboxNode.Nodes.Add(node);

                LoadSubFoldersRecursive(node, folder);
            }
        }

        private void LoadSubFoldersRecursive(TreeNode parentNode, MailFolder parentFolder)
        {
            foreach (var dir in Directory.GetDirectories(parentFolder.FullPath))
            {
                string name = Path.GetFileName(dir);
                var sub = new MailFolder(name, dir, FolderType.InboxSub);

                // ★ MailFolder 側に追加（重要）
                parentFolder.SubFolders.Add(sub);

                // ★ TreeView ノード
                TreeNode node = new TreeNode(name);
                node.Tag = sub;
                parentNode.Nodes.Add(node);

                // ★ 再帰
                LoadSubFoldersRecursive(node, sub);
            }
        }

        private void LoadFolderRecursive(TreeNode parentNode, string parentFolder, string parentPath)
        {
            foreach (var dir in Directory.GetDirectories(parentPath))
            {
                string folderName = Path.GetFileName(dir);          // 例: "仕事"
                string folderPath = parentFolder + "/" + folderName; // 例: "inbox/仕事"

                TreeNode node = new TreeNode(folderName);
                node.Tag = folderPath;  // ← 最重要

                parentNode.Nodes.Add(node);

                // 再帰的にサブフォルダを読み込む
                LoadFolderRecursive(node, folderPath, dir);
            }
        }

        private void treeMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Mail)))
                e.Effect = DragDropEffects.Move;
            else
                e.Effect = DragDropEffects.None;
        }

        private Dictionary<string, string> ParseHeaders(string headerText)
        {
            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string currentKey = null;

            using (var reader = new StringReader(headerText))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        break;

                    if (line.StartsWith(" ") || line.StartsWith("\t"))
                    {
                        if (currentKey != null)
                            headers[currentKey] += " " + line.Trim();
                    }
                    else
                    {
                        int idx = line.IndexOf(':');
                        if (idx > 0)
                        {
                            currentKey = line.Substring(0, idx);
                            string value = line.Substring(idx + 1).Trim();
                            headers[currentKey] = value;
                        }
                    }
                }
            }

            return headers;
        }

        public Mail ParseMail(string path)
        {
            var rawBytes = File.ReadAllBytes(path);
            var rawText = DecodeRawBytesFallback(rawBytes); // とりあえず読める形に

            // 1. ヘッダと本文を分離
            string headerText;
            string bodyText;
            SplitHeaderAndBody(rawText, out headerText, out bodyText);

            // 2. ヘッダを解析（folding 対応）
            var headers = ParseHeaders(headerText);

            var mail = new Mail();
            mail.mailName = Path.GetFileName(path);

            // 3. ヘッダから各種情報を取り出し
            if (headers.TryGetValue("Subject", out var subjectRaw))
                mail.subject = DecodeMimeHeader(subjectRaw);

            if (headers.TryGetValue("From", out var fromRaw))
                mail.address = DecodeMimeHeader(fromRaw);

            if (headers.TryGetValue("Date", out var dateRaw))
                mail.date = dateRaw; // 表示時に FormatReceivedDate で整形

            // 4. 本文デコード（シンプル版：text/plain 優先）
            mail.body = DecodeBody(headers, bodyText);

            return mail;
        }

        private Mail ParseMailWithMimeKit(string path)
        {
            var message = MimeMessage.Load(path);

            var mail = new Mail();
            mail.mailName = Path.GetFileName(path);

            // 件名（MIME デコード済み）
            mail.subject = message.Subject;

            // 差出人（MIME デコード済み）
            var from = message.From.Mailboxes.FirstOrDefault();
            mail.address = !string.IsNullOrEmpty(from?.Name)
                ? from.ToString()
                : from?.Address;

            // 日付（folding も MIME も自動処理）
            mail.date = message.Date.LocalDateTime.ToString("yyyy/MM/dd HH:mm:ss");

            // 本文（text/plain を優先）
            mail.body = message.TextBody ?? message.HtmlBody ?? "";

            return mail;
        }

        private void SplitHeaderAndBody(string rawText, out string headerText, out string bodyText)
        {
            var sbHeader = new StringBuilder();
            var sbBody = new StringBuilder();

            using (var reader = new StringReader(rawText))
            {
                string line;
                bool inHeader = true;

                while ((line = reader.ReadLine()) != null)
                {
                    if (inHeader)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                        {
                            inHeader = false;
                            continue;
                        }
                        sbHeader.AppendLine(line);
                    }
                    else
                    {
                        sbBody.AppendLine(line);
                    }
                }
            }

            headerText = sbHeader.ToString();
            bodyText = sbBody.ToString();
        }

        private string DecodeBody(Dictionary<string, string> headers, string bodyText)
        {
            // Content-Type → charset 抽出
            headers.TryGetValue("Content-Type", out var contentType);
            headers.TryGetValue("Content-Transfer-Encoding", out var transferEncoding);

            string charset = null;

            if (!string.IsNullOrEmpty(contentType))
            {
                var parts = contentType.Split(';');
                foreach (var part in parts)
                {
                    var p = part.Trim();
                    if (p.StartsWith("charset=", StringComparison.OrdinalIgnoreCase))
                    {
                        charset = p.Substring("charset=".Length).Trim(' ', '"', '\'');
                    }
                }
            }

            byte[] rawBody;

            // ★ Base64
            if (!string.IsNullOrEmpty(transferEncoding) &&
                transferEncoding.Equals("base64", StringComparison.OrdinalIgnoreCase))
            {
                // Base64 は空白・改行を除去
                var normalized = string.Concat(
                    bodyText
                        .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(l => l.Trim())
                );
                rawBody = Convert.FromBase64String(normalized);


                try
                {
                    rawBody = Convert.FromBase64String(normalized);
                }
                catch
                {
                    // 壊れている場合は ASCII として扱う
                    rawBody = Encoding.ASCII.GetBytes(bodyText);
                }
            }
            // ★ Quoted-Printable
            else if (!string.IsNullOrEmpty(transferEncoding) &&
                     transferEncoding.Equals("quoted-printable", StringComparison.OrdinalIgnoreCase))
            {
                rawBody = DecodeQuotedPrintableBytes(bodyText);
            }
            // ★ それ以外（7bit / 8bit / binary）
            else
            {
                rawBody = Encoding.ASCII.GetBytes(bodyText);
            }

            // ★ charset に基づいてデコード
            Encoding enc;

            try
            {
                enc = !string.IsNullOrEmpty(charset)
                    ? Encoding.GetEncoding(charset)
                    : DetectJapaneseEncoding(rawBody); // 自動判定
            }
            catch
            {
                enc = Encoding.UTF8;
            }

            return enc.GetString(rawBody);
        }

        private Encoding DetectJapaneseEncoding(byte[] raw)
        {
            Encoding[] candidates =
            {
                Encoding.UTF8,
                Encoding.GetEncoding("iso-2022-jp"),
                Encoding.GetEncoding("shift_jis"),
                Encoding.GetEncoding("euc-jp")
            };

            foreach (var enc in candidates)
            {
                try
                {
                    enc.GetString(raw);
                    return enc;
                }
                catch { }
            }

            return Encoding.UTF8;
        }

        private string DecodeRawBytesFallback(byte[] raw)
        {
            // 順番は好みだけど、日本語メールならこのあたり
            Encoding[] candidates =
            {
                Encoding.UTF8,
                Encoding.GetEncoding("iso-2022-jp"),
                Encoding.GetEncoding("shift_jis"),
                Encoding.GetEncoding("euc-jp")
            };

            foreach (var enc in candidates)
            {
                try
                {
                    return enc.GetString(raw);
                }
                catch { }
            }

            return Encoding.Default.GetString(raw);
        }

        private void LoadUidls()
        {
            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "uidl.txt");
            if (File.Exists(path))
                localUidls = File.ReadAllLines(path).ToList();
        }

        private void SaveUidls()
        {
            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "uidl.txt");
            File.WriteAllLines(path, localUidls);
        }

        private void MoveToTrash(Mail mail)
        {
            if (mail == null)
                return;

            MoveMailWithUndo(mail, folderManager.Trash);
        }

        private void DeletePermanently(Mail mail)
        {
            if (mail == null) return;

            string path = ResolveMailPath(mail);
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
                File.Delete(path);

            // Mail オブジェクト自体は ListView 再描画で消えるので特に保持しない
        }

        private void UndoDelete(Mail mail)
        {
            if (mail == null) return;

            string currentPath = ResolveMailPath(mail);
            if (!File.Exists(currentPath)) return;

            string metaPath = currentPath + ".meta";
            if (!File.Exists(metaPath)) return;

            // ★ .meta を読む
            var lines = File.ReadAllLines(metaPath);
            string originalFolderPath = lines.First(l => l.StartsWith("OriginalFolder="))
                                             .Substring("OriginalFolder=".Length);
            string originalName = lines.First(l => l.StartsWith("OriginalName="))
                                       .Substring("OriginalName=".Length);

            // 元フォルダ
            MailFolder originalFolder = folderManager.FindByPath(originalFolderPath);
            if (originalFolder == null) return;

            string newPath = Path.Combine(originalFolder.FullPath, originalName);

            if (File.Exists(newPath))
            {
                string unique = Guid.NewGuid().ToString() + "_" + originalName;
                newPath = Path.Combine(originalFolder.FullPath, unique);
                mail.mailName = unique;
            }
            else
            {
                mail.mailName = originalName;
            }

            File.Move(currentPath, newPath);

            // ★ .meta 削除
            File.Delete(metaPath);

            // mail の状態更新
            mail.Folder = originalFolder;

            // ★ mailCache のキー更新
            string oldKey = currentPath;
            string newKey = newPath;
            mailCache.Remove(oldKey);
            mailCache[newKey] = mail;

            // メモリ上の Undo 情報はクリア
            mail.LastFolder = null;
            mail.LastMailName = null;
        }

        private void UpdateUndoState()
        {
            // ★ Trash フォルダに .meta があれば Undo 可能
            bool hasUndo = Directory.GetFiles(folderManager.Trash.FullPath, "*.meta").Any();
            menuUndoMail.Enabled = hasUndo;
        }

        private void menuRead_Click(object sender, EventArgs e)
        {
            if (currentMail == null)
                return;

            if (currentMail.notReadYet)
                ToggleReadState(currentMail);

            UpdateView();
        }

        private void ToggleReadState(Mail mail)
        {
            if (mail.Folder.Type == FolderType.Trash)
                return;

            string oldPath = ResolveMailPath(mail);
            string fileName = Path.GetFileName(oldPath);

            bool isUnreadFile = fileName.EndsWith("_unread.eml", StringComparison.OrdinalIgnoreCase);
            bool isReadFile = fileName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase) && !isUnreadFile;

            string newPath;

            // 未読 → 既読
            if (mail.notReadYet)
            {
                oldPath = ResolveMailPath(mail);
                string folder = Path.GetDirectoryName(oldPath);

                string newName = mail.mailName.Replace("_unread.eml", ".eml");
                newPath = Path.Combine(folder, newName);

                File.Move(oldPath, newPath);

                mail.mailName = newName;
                mail.notReadYet = false;

                mailCache.Remove(oldPath);
                mailCache[newPath] = mail;
            }
            // ★ 既読 → 未読
            else if (!mail.notReadYet)
            {
                oldPath = ResolveMailPath(mail);
                string folder = Path.GetDirectoryName(oldPath);

                string baseName = Path.GetFileNameWithoutExtension(mail.mailName);
                string newName = baseName + "_unread.eml";
                newPath = Path.Combine(folder, newName);

                File.Move(oldPath, newPath);

                mail.mailName = newName;
                mail.notReadYet = true;

                mailCache.Remove(oldPath);
                mailCache[newPath] = mail;
            }
            else
            {
                return;
            }

            // ★ mailCache 更新
            mailCache.Remove(oldPath);
            mailCache[newPath] = mail;

            mail.mailName = Path.GetFileName(newPath);
        }

        private void SyncNotReadFlagFromFileName(Mail mail)
        {
            string path = ResolveMailPath(mail);
            string fileName = Path.GetFileName(path);

            // 末尾が _unread.eml なら未読
            if (fileName.EndsWith("_unread.eml", StringComparison.OrdinalIgnoreCase))
            {
                mail.notReadYet = true;
            }
            else
            {
                mail.notReadYet = false;
            }

            // 念のため mailName も同期
            mail.mailName = fileName;
        }

        private void MoveMailWithUndo(Mail mail, MailFolder newFolder)
        {
            string oldPath = ResolveMailPath(mail);
            string oldName = mail.mailName;

            // ★ Draft に移動する場合は必ず未送信扱いにする
            if (newFolder.Type == FolderType.Draft)
            {
                mail.notReadYet = true;

                // ファイル名を _unread に統一
                if (!oldName.EndsWith("_unread.eml"))
                {
                    string baseName = Path.GetFileNameWithoutExtension(oldName);
                    mail.mailName = baseName + "_unread.eml";
                }
            }
            else
            {
                // Draft 以外は元の未読状態を維持
                if (mail.notReadYet)
                {
                    if (!oldName.EndsWith("_unread.eml"))
                        mail.mailName = Path.GetFileNameWithoutExtension(oldName) + "_unread.eml";
                }
                else
                {
                    mail.mailName = oldName.Replace("_unread.eml", ".eml");
                }
            }

            string newPath = Path.Combine(newFolder.FullPath, mail.mailName);

            Directory.CreateDirectory(newFolder.FullPath);
            if (File.Exists(oldPath))
                File.Move(oldPath, newPath);

            if (mailCache.ContainsKey(oldPath))
                mailCache.Remove(oldPath);

            mail.mailPath = newPath;
            mail.Folder = newFolder;
            mailCache[newPath] = mail;
        }

        private Mail LoadSingleMail(string path)
        {
            if (!File.Exists(path))
                return null;

            var message = MimeMessage.Load(path);

            Mail mail = new Mail();
            mail.mailName = Path.GetFileName(path);

            // ★ From
            var fromMailbox = message.From.Mailboxes.FirstOrDefault();
            mail.from = fromMailbox != null ? fromMailbox.ToString() : "";

            // ★ To
            mail.address = string.Join("; ", message.To.Mailboxes.Select(m => m.ToString()));

            // ★ Cc
            mail.ccaddress = string.Join("; ", message.Cc.Mailboxes.Select(m => m.ToString()));

            // ★ Bcc
            mail.bccaddress = string.Join("; ", message.Bcc.Mailboxes.Select(m => m.ToString()));

            // ★ Subject
            mail.subject = message.Subject ?? "";

            // ★ Body
            mail.body = message.TextBody ?? "";

            // ★ Date
            mail.date = message.Date.LocalDateTime.ToString("yyyy/MM/dd HH:mm:ss");

            // ★ 未読判定
            mail.notReadYet = path.EndsWith("_unread.eml");

            return mail;
        }

        private void SaveRules()
        {
            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "rules.json");
            File.WriteAllText(path, JsonConvert.SerializeObject(rules, Formatting.Indented));
        }

        private void LoadRules()
        {
            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "rules.json");
            if (File.Exists(path))
            {
                rules = JsonConvert.DeserializeObject<List<MailRule>>(File.ReadAllText(path));
            }
        }

        private void ApplyRules(Mail mail)
        {
            if (rules == null || rules.Count == 0)
                return;

            string subject = mail.subject ?? "";
            string from = mail.from ?? "";

            // ★ メールアドレス部分だけ抽出
            string fromAddress = from;
            int lt = fromAddress.IndexOf('<');
            int gt = fromAddress.IndexOf('>');
            if (lt >= 0 && gt > lt)
                fromAddress = fromAddress.Substring(lt + 1, gt - lt - 1);

            foreach (var rule in rules)
            {
                bool match = false;

                // 件名に含む
                if (!string.IsNullOrWhiteSpace(rule.Contains))
                {
                    if (subject.IndexOf(rule.Contains, StringComparison.OrdinalIgnoreCase) >= 0)
                        match = true;
                }

                // 差出人に含む（メールアドレスで比較）
                if (!string.IsNullOrWhiteSpace(rule.From))
                {
                    if (fromAddress.IndexOf(rule.From, StringComparison.OrdinalIgnoreCase) >= 0)
                        match = true;
                }

                if (!match)
                    continue;

                // ★ 移動先フォルダを取得（必要なら作成）
                MailFolder target = folderManager.GetOrCreateFolderByPath(rule.MoveTo);
                if (target == null)
                    return;

                // ★ 正式な移動処理に統一
                MoveMailWithUndo(mail, target);

                break; // 最初に一致したルールだけ適用
            }
        }

        private void AddFolderToTreeView(MailFolder folder)
        {
            // Inbox ノードを取得
            TreeNode inboxNode = treeMain.Nodes[0].Nodes[0];

            TreeNode node = new TreeNode(folder.Name);
            node.Tag = folder;

            inboxNode.Nodes.Add(node);
        }

        private void menuReleEdit_Click(object sender, EventArgs e)
        {
            var dlg = new FormRuleEditor(rules, treeMain.Nodes[0]);
            dlg.Owner = this;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                rules = dlg.Rules;   // 更新されたルールを受け取る
                SaveRules();
            }
        }

        private void BuildTree()
        {
            if (isBuildingTree)
                return;

            isBuildingTree = true;

            try
            {
                treeMain.BeginUpdate();
                treeMain.Nodes.Clear();

                TreeNode root = new TreeNode("メール");
                root.Tag = null;
                root.ImageKey = "folder";
                root.SelectedImageKey = "folder";
                treeMain.Nodes.Add(root);

                AddSystemFolder(root, folderManager.Inbox);
                AddSystemFolder(root, folderManager.Send);
                AddSystemFolder(root, folderManager.Draft);
                AddSystemFolder(root, folderManager.Trash);

                treeMain.ExpandAll();
                treeMain.EndUpdate();
            }
            finally
            {
                isBuildingTree = false;
            }
        }

        private void AddSystemFolder(TreeNode parent, MailFolder folder)
        {
            // ★ フォルダが壊れていたらスキップ
            if (folder == null || string.IsNullOrEmpty(folder.FullPath))
                return;

            if (!Directory.Exists(folder.FullPath))
                return;

            int unread = CountUnread(folder);

            string text = unread >= 0
                ? $"{folder.DisplayName}({unread})"
                : folder.DisplayName;

            TreeNode node = new TreeNode(text);
            node.Tag = folder;

            string icon = GetIconKey(folder, unread);
            node.ImageKey = icon;
            node.SelectedImageKey = icon;

            parent.Nodes.Add(node);

            // ★ Inbox のときだけサブフォルダを読み込む
            if (folder.Type == FolderType.Inbox)
            {
                LoadInboxFolders(node, folder);
            }
        }

        private void AddFolderToTreeView(TreeNode parent, MailFolder folder)
        {
            int unread = CountUnread(folder);

            string text = unread > 0
                ? $"{folder.DisplayName ?? folder.Name} ({unread})"
                : folder.DisplayName ?? folder.Name;

            TreeNode node = new TreeNode(text);
            node.Tag = folder;

            string icon = GetIconKey(folder, unread);
            node.ImageKey = icon;
            node.SelectedImageKey = icon;

            parent.Nodes.Add(node);

            foreach (var sub in folder.SubFolders)
            {
                AddFolderToTreeView(node, sub);
            }
        }

        private int CountUnread(MailFolder folder)
        {
            if (folder == null || string.IsNullOrEmpty(folder.FullPath))
                return 0;

            if (!Directory.Exists(folder.FullPath))
                return 0;

            int count = 0;
            try
            {
                string[] files = Directory.GetFiles(folder.FullPath, "*.eml", SearchOption.TopDirectoryOnly);
                foreach (string file in files)
                {
                    string name = Path.GetFileName(file);
                    if (name.EndsWith("_unread.eml", StringComparison.OrdinalIgnoreCase))
                        count++;
                }
            }
            catch (Exception ex)
            {
                logger.Warn("CountUnread error: " + folder.FullPath + " : " + ex.Message);
            }

            return count;
        }

        private string GetIconKey(MailFolder folder, int unread)
        {
            switch (folder.Type)
            {
                case FolderType.Inbox:
                    return "inbox";
                case FolderType.Send:
                    return "send";
                case FolderType.Trash:
                    return "trash";
                case FolderType.Draft:
                    return "draft";
                default:
                    return unread > 0 ? "folder_unread" : "folder";
            }
        }

        public static void SetBrowserFeatureControl()
        {
            var fileName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION",
                    RegistryKeyPermissionCheck.ReadWriteSubTree))
                {
                    // 11001 = IE11 Edge モード
                    key.SetValue(fileName, 11001, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void ShowHtml(string html)
        {
            if (browserMail.CoreWebView2 != null)
            {
                browserMail.NavigateToString(html);
            }
            else
            {
                // 初期化完了後に実行する
                browserMail.CoreWebView2InitializationCompleted += (s, e) =>
                {
                    browserMail.NavigateToString(html);
                };
            }
        }

        string FixBrokenHtml(string html)
        {
            // 不正な属性を削除
            html = Regex.Replace(html, @"\bsvg=""[^""]*""", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"\bXMLNS:\[default\]", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"\b2000 www\.w3\.org http:", "", RegexOptions.IgnoreCase);

            // <html> タグを正常化
            html = Regex.Replace(html, @"<html[^>]*>", "<html lang=\"ja\">", RegexOptions.IgnoreCase);

            return html;
        }

        private Mail LoadMail(string path, MailFolder folder)
        {
            var message = MimeMessage.Load(path);

            Mail mail = new Mail();
            mail.mailPath = path;
            mail.mailName = Path.GetFileName(path);
            mail.Folder = folder;

            // 件名
            mail.subject = message.Subject ?? "";

            // 差出人
            if (message.From != null && message.From.Count > 0)
                mail.from = message.From.ToString();
            else
                mail.from = "";

            // 宛先
            if (message.To != null && message.To.Count > 0)
                mail.address = message.To.ToString();
            else
                mail.address = "";

            // Cc
            if (message.Cc != null && message.Cc.Count > 0)
                mail.ccaddress = message.Cc.ToString();
            else
                mail.ccaddress = "";

            // Bcc（通常は空）
            if (message.Bcc != null && message.Bcc.Count > 0)
                mail.bccaddress = message.Bcc.ToString();
            else
                mail.bccaddress = "";

            // 日付
            if (message.Date != DateTimeOffset.MinValue)
                mail.date = message.Date.LocalDateTime.ToString("yyyy/MM/dd HH:mm:ss");
            else
                mail.date = File.GetLastWriteTime(path).ToString("yyyy/MM/dd HH:mm:ss");

            // ★ 既読フラグ（X-MizuMail-Read）
            mail.notReadYet = message.Headers["X-MizuMail-Read"] != "1";

            // ★ 下書きフラグ（X-MizuMail-Draft）
            mail.isDraft = message.Headers["X-MizuMail-Draft"] == "1";

            // ★ 本文（TextPart / Multipart 両対応）
            string bodyText = "";

            if (message.Body is TextPart)
            {
                bodyText = ((TextPart)message.Body).Text;
            }
            else if (message.Body is Multipart)
            {
                var mp = (Multipart)message.Body;
                for (int i = 0; i < mp.Count; i++)
                {
                    if (mp[i] is TextPart)
                    {
                        bodyText = ((TextPart)mp[i]).Text;
                        break;
                    }
                }
            }

            mail.body = bodyText ?? "";

            return mail;
        }

        private void UpdateAttachmentMenu(Mail mail)
        {
            buttonAtachMenu.DropDownItems.Clear();

            if (mail == null || mail.attachList.Count == 0)
                return;

            foreach (var name in mail.attachList)
            {
                var item = new ToolStripMenuItem(name);
                item.Tag = name; // ファイル名だけ
                buttonAtachMenu.DropDownItems.Add(item);
            }
        }

        private IEnumerable<MimePart> FindAttachments(MimeEntity entity)
        {
            if (entity is MimePart part)
            {
                // 添付判定（FileName が null の場合も拾う）
                if (part.IsAttachment ||
                    part.ContentDisposition?.Disposition == ContentDisposition.Attachment ||
                    !string.IsNullOrEmpty(part.FileName) ||
                    !string.IsNullOrEmpty(part.ContentDisposition?.FileName) ||
                    !string.IsNullOrEmpty(part.ContentType?.Name))
                {
                    yield return part;
                }
            }

            if (entity is Multipart multipart)
            {
                foreach (var child in multipart)
                {
                    foreach (var found in FindAttachments(child))
                        yield return found;
                }
            }
        }

        private string GetAttachmentName(MimePart part)
        {
            return part.FileName
                ?? part.ContentDisposition?.FileName
                ?? part.ContentType?.Name
                ?? "attachment.bin";
        }

        [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SHGetFileInfo(
            string pszPath,
            uint dwFileAttributes,
            out SHFILEINFO psfi,
            uint cbFileInfo,
            uint uFlags
        );

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        private Icon GetIconFromExtension(string ext)
        {
            SHFILEINFO shinfo = new SHFILEINFO();
            SHGetFileInfo(ext,
                0,
                out shinfo,
                (uint)Marshal.SizeOf(shinfo),
                0x100 | 0x1); // SHGFI_ICON | SHGFI_USEFILEATTRIBUTES

            return Icon.FromHandle(shinfo.hIcon);
        }

        private void MoveDraftToSent(Mail mail)
        {
            // 現在の .eml のパス
            string src = mail.mailPath;

            // FolderManager から送信フォルダを取得
            MailFolder sendFolder = folderManager.Send;

            // 移動先パス
            string dst = Path.Combine(sendFolder.FullPath, Path.GetFileName(src));

            // ファイル移動
            File.Move(src, dst);

            // Mail オブジェクト更新
            mail.mailPath = dst;
            mail.Folder = sendFolder;
        }

        private void menuLocalFiltter_Click(object sender, EventArgs e)
        {
            ApplyLocalFilters();
            BuildTree();
            UpdateView();
        }

        private void ApplyLocalFilters()
        {
            // ★ Inbox のメールを全部読み込む
            var inbox = folderManager.Inbox;
            var mails = LoadEmlFolder(inbox);

            foreach (var mail in mails)
            {
                ApplyRules(mail);  // ← 受信時と同じルールを適用
            }
        }

    }
}
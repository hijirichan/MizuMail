using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Pop3;
using MailKit.Security;
using Microsoft.Web.WebView2.Core;
using MimeKit;
using MimeKit.Text;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MizuMail
{
    public partial class FormMain : Form
    {
        // UIDL格納用の配列
        public List<string> localUidls = new List<string>();

        // メールボックス情報を表示しているときのフラグ
        public bool mailBoxViewFlag = false;

        // 現在の検索キーワードを格納するフィールド
        private string currentKeyword = "";

        // 選択された行を格納するフィールド
        private Mail currentMail;
        private FolderManager folderManager;
        bool isBuildingTree = false;
        private bool showHeader = false;
        private int charWidth;
        private bool isUpdatingList = false;
        private bool suppressSelect = false;

        private System.Windows.Forms.Timer resizeTimer;
        private int updateListViewCount = 0;
        private DateTime updateListViewStart = DateTime.MinValue;
        HashSet<MailFolder> updatedFolders = new HashSet<MailFolder>();
        private List<Mail> _virtualList = new List<Mail>();
        private Dictionary<MailFolder, TreeNode> folderNodeMap = new Dictionary<MailFolder, TreeNode>();
        private Stack<TagUndoAction> tagUndoStack = new Stack<TagUndoAction>();
        private Stack<TagUndoAction> tagRedoStack = new Stack<TagUndoAction>();
        // ★ 事前に1回だけ作る（FormMain のフィールド）
        private readonly Font boldFont;
        private readonly Font normalFont;
        private int sortColumn = 0;
        private bool sortAscending = true;

        // ★ メールルール
        private List<MailRule> rules = new List<MailRule>();

        // ★ フォルダごとのキャッシュ
        private Dictionary<string, Mail> mailCache = new Dictionary<string, Mail>();

        // ロガーの取得
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public FormMain()
        {
            // ★ WebView2 Runtime チェック
            if (!CheckWebView2Runtime())
            {
                MessageBox.Show(
                    "WebView2 Runtimeがインストールされていないため、MizuMailを起動できません。\n" +
                    "Microsoft Edge WebView2 Runtime をインストールしてください。",
                    "WebView2 Runtime が必要です",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                Environment.Exit(1);
                return;
            }

            InitializeComponent();

            normalFont = listMain.Font;
            boldFont = new Font(listMain.Font, FontStyle.Bold);

            // 初期化
            currentMail = null;

            Application.Idle += Application_Idle;
            listMain.SmallImageList = new ImageList { ImageSize = new Size(1, 20) };
            listMain.VirtualMode = true;
            listMain.RetrieveVirtualItem += listMain_RetrieveVirtualItem;
        }

        private bool CheckWebView2Runtime()
        {
            try
            {
                string version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                return !string.IsNullOrEmpty(version);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 日付用パースメソッド
        /// </summary>
        /// <param name="text"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        private static bool TryParseListViewDate(string text, out DateTime dt)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                dt = default;
                return false;
            }

            // 日本語ロケール（曜日：月, 火, 水, 木, 金, 土, 日）
            var jp = new System.Globalization.CultureInfo("ja-JP");

            string[] formats = new[]
            {
                "yyyy/MM/dd (ddd) HH:mm:ss",
                "yyyy/MM/dd (ddd) HH:mm",
                "yyyy/MM/dd HH:mm:ss",
                "yyyy/MM/dd HH:mm",
                "yyyy/MM/dd"
            };

            return DateTime.TryParseExact(
                text.Trim(),
                formats,
                jp,   // ← ここが超重要
                System.Globalization.DateTimeStyles.None,
                out dt
            );
        }

        /// <summary>
        /// メール送受信後のTreeView、ListViewの更新
        /// </summary>
        private void UpdateView()
        {
            logger.Debug($"[UpdateView] 呼び出し元: {Environment.StackTrace}");
            logger.Debug($"UpdateListView: mailCache.Count={mailCache.Count}");

            listMain.ListViewItemSorter = null;

            UpdateTreeView();
            UpdateListView();

            UpdateUndoState();
        }

        private void UpdateTreeView()
        {
            // 受信
            int inboxCount = Directory.GetFiles(folderManager.Inbox.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[0].Text = $"受信メール({inboxCount})";

            // 受信
            int spamCount = Directory.GetFiles(folderManager.Spam.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[1].Text = $"迷惑メール({spamCount})";

            // 送信
            int sendCount = Directory.GetFiles(folderManager.Send.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[2].Text = $"送信メール({sendCount})";

            // 下書き
            int draftCount = Directory.GetFiles(folderManager.Draft.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[3].Text = $"下書き({draftCount})";

            // ごみ箱
            int trashCount = Directory.GetFiles(folderManager.Trash.FullPath, "*.eml").Length;
            treeMain.Nodes[0].Nodes[4].Text = $"ごみ箱({trashCount})";
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
            if (isUpdatingList)
                return;

            isUpdatingList = true;
            updateListViewCount++;
            logger.Debug($"[UpdateListView] 呼び出し {updateListViewCount} 回目  時刻: {DateTime.Now:HH:mm:ss.fff}");

            try
            {
                TreeNode node = treeMain.SelectedNode;
                MailFolder folder = node?.Tag as MailFolder;

                //
                // ★★★ メールボックス一覧モード（folder == null）
                //
                if (folder == null)
                {
                    logger.Debug("[UpdateListView] folder == null → MailboxView");

                    // ★ VirtualMode の内部データを完全リセット
                    _virtualList.Clear();

                    // ★ 選択をクリア（古い選択イベント対策）
                    listMain.SelectedIndices.Clear();

                    // ★ VirtualListSize を 0 にして内部状態をリセット
                    if (listMain.VirtualMode)
                        listMain.VirtualListSize = 0;

                    // ★ VirtualMode を OFF に戻す
                    if (listMain.VirtualMode)
                    {
                        listMain.VirtualMode = false;
                        listMain.RetrieveVirtualItem -= listMain_RetrieveVirtualItem;
                    }

                    listMain.BeginUpdate();
                    try
                    {
                        listMain.Items.Clear();
                        ShowMailboxInfo();   // Items.Add ベースで OK
                    }
                    finally
                    {
                        try { listMain.EndUpdate(); } catch { }
                    }

                    labelMessage.Text = "メールボックス一覧";

                    // ★★★ Visible を true に戻す（消失対策）
                    listMain.Visible = true;

                    return;
                }

                //
                // ★★★ メール一覧モード（VirtualMode ON）
                //
                mailBoxViewFlag = false;

                if (!listMain.VirtualMode)
                {
                    listMain.Items.Clear();
                    listMain.VirtualMode = true;
                    listMain.RetrieveVirtualItem += listMain_RetrieveVirtualItem;
                }

                // ★ displayList 抽出
                var displayList = mailCache.Values
                    .Where(m => m != null &&
                                m.Folder != null &&
                                m.Folder.FullPath == folder.FullPath)
                    .ToList();

                // ★ デフォルトは日付降順
                displayList = displayList
                    .OrderByDescending(m =>
                    {
                        if (DateTime.TryParse(m.date, out DateTime dt))
                            return dt;
                        return DateTime.MinValue;
                    })
                    .ToList();

                // ★ キーワードフィルタ
                if (!string.IsNullOrEmpty(currentKeyword))
                {
                    string kw = currentKeyword;
                    displayList = displayList
                        .Where(m =>
                            (m.subject?.Contains(kw) ?? false) ||
                            (m.body?.Contains(kw) ?? false) ||
                            (m.address?.Contains(kw) ?? false))
                        .ToList();
                }

                // ★ フィルタコンボ
                string filter = toolFilterCombo.SelectedItem?.ToString();
                if (filter == "未読")
                    displayList = displayList.Where(m => m.notReadYet).ToList();
                else if (filter == "添付あり")
                    displayList = displayList.Where(m => m.hasAtach).ToList();
                else if (filter == "今日")
                    displayList = displayList.Where(m =>
                    {
                        if (DateTime.TryParse(m.date, out DateTime dt))
                            return dt.Date == DateTime.Now.Date;
                        return false;
                    }).ToList();

                // ★ VirtualMode 用データソースにセット
                _virtualList = displayList;

                // ★ VirtualListSize を更新（描画トリガー）
                listMain.VirtualListSize = _virtualList.Count;

                listMain.Visible = true;
                labelMessage.Text = $"{_virtualList.Count}件読み込みました。";

                logger.Debug($"VirtualListSize={listMain.VirtualListSize}, _virtualList.Count={_virtualList.Count}");

                //
                // ★★★ 選択の安定化（最重要）
                //
                listMain.SelectedIndices.Clear();

                if (_virtualList.Count > 0)
                {
                    listMain.SelectedIndices.Add(0);
                }
            }
            finally
            {
                isUpdatingList = false;
            }

            logger.Debug($"[UpdateListView] 完了 {updateListViewCount} 回目");
        }

        // フィールドに追加
        private bool _firstSelectHandled = false;

        private async void treeMain_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (suppressSelect)
                return;

            if (isBuildingTree)
                return;

            // ★ フォーム起動直後の「勝手に当たるフォーカス」は無視する
            if (!_firstSelectHandled)
            {
                _firstSelectHandled = true;
                return;
            }

            richTextBody.Clear();
            currentKeyword = "";
            buttonAtachMenu.DropDownItems.Clear();
            buttonAtachMenu.Visible = false;
            currentMail = null;

            var node = e.Node;
            listMain.Visible = false;
            await LoadAndShowFolderAsync(node);

            // ★★★ これが無かったせいで画面が更新されていなかった
            UpdateView();
        }

        private async Task LoadAndShowFolderAsync(TreeNode node)
        {
            var folder = node.Tag as MailFolder;

            // カラムだけ先に更新（UIスレッドでOK）
            UpdateColumnHeaders(folder);

            if (folder != null)
            {
                // すでにロード済みか？
                bool folderLoaded = mailCache.Values.Any(m =>
                    m != null &&
                    m.Folder != null &&
                    m.Folder.FullPath == folder.FullPath
                );

                if (!folderLoaded)
                {
                    // ★★★ 重い処理はバックグラウンドで実行（UIを止めない）
                    await Task.Run(() => LoadEmlFolder(folder));

                    // ★ null の削除（軽いのでUIスレッドでOK）
                    foreach (var key in mailCache.Where(kv => kv.Value == null)
                                                 .Select(kv => kv.Key)
                                                 .ToList())
                    {
                        mailCache.Remove(key);
                    }
                }
            }

            // ★★★ ロード完了後に UI 更新（UIスレッドで実行）
            UpdateView();
        }

        private void UpdateColumnHeaders(MailFolder folder)
        {
            if (folder == null)
            {
                listMain.Columns[1].Text = "メールボックス名";
                listMain.Columns[2].Text = "メールアドレス";
                listMain.Columns[3].Text = "更新日時";

                // メールボックス一覧ではプレビュー・タグ・ファイル名は非表示扱い
                listMain.Columns[5].Text = "";
                listMain.Columns[6].Text = "";
                listMain.Columns[7].Text = "";
                return;
            }

            switch (folder.Type)
            {
                case FolderType.Inbox:
                case FolderType.InboxSub:
                case FolderType.Spam:
                    listMain.Columns[1].Text = "差出人";
                    listMain.Columns[2].Text = "件名";
                    listMain.Columns[3].Text = "受信日時";
                    break;

                case FolderType.Send:
                    listMain.Columns[1].Text = "宛先";
                    listMain.Columns[2].Text = "件名";
                    listMain.Columns[3].Text = "送信日時";
                    break;

                case FolderType.Draft:
                    listMain.Columns[1].Text = "宛先(下書き)";
                    listMain.Columns[2].Text = "件名";
                    listMain.Columns[3].Text = "作成日時";
                    break;

                case FolderType.Trash:
                    listMain.Columns[1].Text = "差出人または宛先";
                    listMain.Columns[2].Text = "件名";
                    listMain.Columns[3].Text = "受信日時または送信日時";
                    break;
            }

            listMain.Columns[5].Text = "プレビュー";
            listMain.Columns[6].Text = "タグ";
            listMain.Columns[7].Text = "メールファイル名";
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

                using (var client = new MailKit.Net.Smtp.SmtpClient())
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

                        // ★ 送信メールフォルダの一覧を更新
                        LoadEmlFolder(folderManager.Send);

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

            UpdateTreeView();   // 件数だけ更新
            UpdateListView();   // 選択中フォルダの一覧を更新
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
            if (listMain.SelectedIndices.Count == 0)
                return;

            int index = listMain.SelectedIndices[0];

            // ★ VirtualMode 安全対策：範囲外なら無視
            if (index < 0 || index >= _virtualList.Count)
                return;

            Mail mail = _virtualList[index];
            if (mail == null)
                return;

            // ★ 送信メール・下書きメールは編集画面を開く（最優先）
            if (mail.Folder.Type == FolderType.Send || mail.Folder.Type == FolderType.Draft)
            {
                OpenSendMailEditor(mail);
                return;
            }

            // ★ Trash 以外は既読化
            if (mail.Folder.Type != FolderType.Trash)
            {
                ToggleReadState(mail);

                // ★ VirtualMode ではこの行だけ再描画すれば十分
                listMain.RedrawItems(index, index, true);
            }

            // ★ 本文表示
            currentMail = mail;
            UpdateMailView();
        }

        private void listMain_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (listMain.SelectedIndices.Count != 1)
            {
                currentMail = null;
                UpdateMailView();
                return;
            }

            int index = listMain.SelectedIndices[0];

            // ★ VirtualMode 安全対策：範囲外なら無視
            if (index < 0 || index >= _virtualList.Count)
            {
                currentMail = null;
                UpdateMailView();
                return;
            }

            currentMail = _virtualList[index];
            UpdateMailView();
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
                if (part is MimePart mp && mp.IsAttachment)
                {
                    string fileName = mp.FileName;

                    // ★ 一時ファイル名をユニークにする
                    string tempPath = Path.Combine(
                        Path.GetTempPath(),
                        Guid.NewGuid().ToString("N") + "_" + fileName
                    );

                    using (var stream = File.Create(tempPath))
                    {
                        mp.Content.DecodeTo(stream);
                    }

                    // ★ アイコン取得（例外対策）
                    Icon icon;
                    try
                    {
                        icon = Icon.ExtractAssociatedIcon(tempPath);
                    }
                    catch
                    {
                        icon = SystemIcons.Application;
                    }

                    var item = new ToolStripMenuItem(fileName, icon.ToBitmap());
                    item.Tag = tempPath;

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
            Properties.Settings.Default.ColWidth6 = listMain.Columns[6].Width;
            Properties.Settings.Default.ColWidth7 = listMain.Columns[7].Width;
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
            var env = await CoreWebView2Environment.CreateAsync(null, null, new CoreWebView2EnvironmentOptions("--allow-file-access-from-files"));
            await browserMail.EnsureCoreWebView2Async(env);
            charWidth = TextRenderer.MeasureText("あ", listMain.Font).Width;
            listMain.Sorting = SortOrder.None;

            // ② 設定読み込み
            LoadSettings();
            LoadRules();

            // ③ mbox フォルダ構造保証
            Directory.CreateDirectory(folderManager.Inbox.FullPath);
            Directory.CreateDirectory(folderManager.Spam.FullPath);
            Directory.CreateDirectory(folderManager.Send.FullPath);
            Directory.CreateDirectory(folderManager.Draft.FullPath);
            Directory.CreateDirectory(folderManager.Trash.FullPath);

            BuildTree();

            await Task.Delay(50); // ★ TreeView 初期化完了を待つ

            // ④ TreeView の Tag を MailFolder に統一
            TreeNode root = treeMain.Nodes[0];
            TreeNode inboxNode = root.Nodes[0];
            MailFolder inboxFolder = folderManager.Inbox;
            root.Nodes[0].Tag = folderManager.Inbox;
            root.Nodes[1].Tag = folderManager.Spam;
            root.Nodes[2].Tag = folderManager.Send;
            root.Nodes[3].Tag = folderManager.Draft;
            root.Nodes[4].Tag = folderManager.Trash;

            var img = new ImageList();
            img.ImageSize = new Size(16, 16);
            img.Images.Add(Properties.Resources.unread);
            img.Images.Add(Properties.Resources.read);
            img.Images.Add(Properties.Resources.attach);
            img.Images.Add(Properties.Resources.spam);
            listMain.SmallImageList = img;

            // ⑤ inbox サブフォルダ読み込み（MailFolder 再帰）
            LoadInboxFolders(inboxNode, inboxFolder);

            LoadUidls();

            // ⑥ カラム幅復元
            RestoreColumnWidths();

            // ⑦ 表示更新
            suppressSelect = true;
            treeMain.SelectedNode = inboxNode;
            suppressSelect = false;

            await LoadAndShowFolderAsync(inboxNode);
            UpdateView();

            // ⑧ タイマー開始
            SetTimer(Mail.checkMail, Mail.checkInterval);

            // ⑨ 展開
            treeMain.ExpandAll();

            // AfterSelect の「初回スキップ」があるなら、ここで済んだことにしておく
            _firstSelectHandled = true;
        }

        private void RestoreColumnWidths()
        {
            int[] defaults = { 24, 150, 200, 150, 120, 200, 120, 0 }; // 好みで調整可能
            int[] widths = new int[8];

            object[] settings =
            {
                Properties.Settings.Default.ColWidth0,
                Properties.Settings.Default.ColWidth1,
                Properties.Settings.Default.ColWidth2,
                Properties.Settings.Default.ColWidth3,
                Properties.Settings.Default.ColWidth4,
                Properties.Settings.Default.ColWidth5,
                Properties.Settings.Default.ColWidth6,
                Properties.Settings.Default.ColWidth7
            };

            for (int i = 0; i < 8; i++)
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
            for (int i = 0; i < 8; i++)
            {
                if (i < listMain.Columns.Count)
                    listMain.Columns[i].Width = widths[i];
            }
        }

        private void menuDelete_Click(object sender, EventArgs e)
        {
            if (listMain.SelectedIndices.Count == 0)
                return;

            // ★ VirtualMode では SelectedItems を使わない
            //    まず選択された Mail を取得
            var selected = listMain.SelectedIndices
                                   .Cast<int>()
                                   .Select(i => _virtualList[i])
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

            //
            // ★ ごみ箱へ移動（確認ダイアログあり）
            //
            if (trashTargets.Count > 0)
            {
                string msg = $"選択したメール {trashTargets.Count} 件をごみ箱に移動します。よろしいですか？";
                if (MessageBox.Show(msg, "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    foreach (var m in trashTargets)
                        MoveMailWithUndo(m, folderManager.Trash);
                }
            }

            //
            // ★ ごみ箱内の完全削除（確認ダイアログあり）
            //
            if (permanentTargets.Count > 0)
            {
                string msg = $"ごみ箱内のメール {permanentTargets.Count} 件を完全に削除します。元に戻せません。";
                if (MessageBox.Show(msg, "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    foreach (var m in permanentTargets)
                        DeletePermanently(m);
                }
            }

            //
            // ★ VirtualMode では _virtualList から削除されたメールを除外する必要がある
            //
            _virtualList = _virtualList
                .Where(m => m != null && !trashTargets.Contains(m) && !permanentTargets.Contains(m))
                .ToList();

            // ★ VirtualListSize を更新（最重要）
            listMain.VirtualListSize = _virtualList.Count;

            // ★ 再描画
            listMain.Invalidate();

            // ★ TreeView の件数更新
            UpdateTreeView();

            // ★ 本文クリア
            currentMail = null;
            UpdateMailView();

            UpdateUndoState();
        }

        private void menuNotReadYet_Click(object sender, EventArgs e)
        {
            if (listMain.SelectedIndices.Count == 0)
                return;

            foreach (int index in listMain.SelectedIndices)
            {
                Mail mail = _virtualList[index];
                if (mail != null && !mail.notReadYet)
                {
                    ToggleReadState(mail); // ★ 既読 → 未読 に戻す
                }
            }

            // ★ VirtualMode では部分再描画が最速
            int first = listMain.SelectedIndices[0];
            int last = listMain.SelectedIndices[listMain.SelectedIndices.Count - 1];
            listMain.RedrawItems(first, last, true);

            // ★ 本文表示を更新（未読に戻した場合でも自然）
            UpdateMailView();
        }

        private void menuAppExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
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
                string atach = string.Join(";", form.buttonAttachList.DropDownItems.Cast<ToolStripItem>().Select(item => item.Tag?.ToString() ?? ""));
                string date = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                // 宛先、本文がある場合
                if (to != "" | body != "")
                {
                    // 件名がない場合は無題
                    if (subject == "")
                    {
                        subject = "無題";
                    }
                    // コレクションに追加する
                    Mail mail = new Mail(to, cc, bcc, subject, body, atach, date, "", "", true);
                    mail.isDraft = true;
                    mail.notReadYet = true;
                    mail.Folder = folderManager.Draft;
                    SaveMail(mail);

                    // ★ Draft フォルダの一覧を更新
                    LoadEmlFolder(folderManager.Draft);
                    treeMain.SelectedNode = folderNodeMap[folderManager.Draft];
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
            if (File.Exists(Application.StartupPath + @"\MizuMail.xml"))
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(MailSettings));
                using (var fs = new FileStream(Application.StartupPath + @"\MizuMail.xml", FileMode.Open))
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

            using (var fs = new FileStream(Application.StartupPath + @"\MizuMail.xml", FileMode.Create))
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
            if (listMain.SelectedIndices.Count == 0)
                return;

            int index = listMain.SelectedIndices[0];
            Mail mail = _virtualList[index];
            if (mail == null)
                return;

            string fileName = e.ClickedItem.Tag as string;
            if (fileName == null)
                return;

            var message = MimeMessage.Load(mail.mailPath);

            var part = FindAttachments(message.Body)
                .FirstOrDefault(p =>
                    string.Equals(GetAttachmentName(p), fileName, StringComparison.OrdinalIgnoreCase));

            if (part == null)
            {
                MessageBox.Show("添付ファイルが見つかりませんでした。");
                return;
            }

            string tempDir = Path.Combine(Application.StartupPath, "mbox", "tmp");
            Directory.CreateDirectory(tempDir);

            // ★ ユニークな一時ファイル名
            string tempPath = Path.Combine(
                tempDir,
                Guid.NewGuid().ToString("N") + "_" + fileName
            );

            using (var stream = File.Create(tempPath))
                part.Content.DecodeTo(stream);

            if (MessageBox.Show(
                $"選択したファイル {tempPath} を開きますか？\n" +
                "ファイルによってはウィルスの可能性もありますので注意してください。",
                "確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    Process.Start(tempPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ファイルを開けませんでした。\n" + ex.Message);
                }
            }
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
            // ★ メールボックス一覧モードでは UI を触らない
            if (mailBoxViewFlag)
            {
                toolReplyButton.Enabled = false;
                toolDeleteButton.Enabled = false;
                menuRead.Enabled = false;
                menuNotReadYet.Enabled = false;
                menuDelete.Enabled = false;
                menuMailDelete.Enabled = false;
                menuMailReply.Enabled = false;
                toolShowHeader.Enabled = false;
                menuEditTags.Enabled = false;
                menuSaveAs.Enabled = false;
                menuSpeechMail.Enabled = false;
                menuAddToAddressBook.Enabled = false;
                menuUndoTags.Enabled = false;
                menuRedoTags.Enabled = false;
                menuAttachmentFileAllSave.Enabled = false;
                return;
            }

            int selCount = listMain.SelectedIndices.Count;
            bool hasOne = selCount == 1;
            bool hasAny = selCount > 0;

            toolReplyButton.Enabled = hasOne;
            toolDeleteButton.Enabled = hasAny;
            menuRead.Enabled = hasAny;
            menuNotReadYet.Enabled = hasAny;
            menuDelete.Enabled = hasAny;
            menuMailDelete.Enabled = hasAny;
            menuMailReply.Enabled = hasOne;
            toolShowHeader.Enabled = hasOne;
            menuEditTags.Enabled = hasOne;

            // ★ Trash フォルダのチェック（ここは今のままでOK）
            if (folderManager?.Trash == null || folderManager.Trash.FullPath == null)
            {
                menuFileClearTrash.Enabled = false;
                menuClearTrash.Enabled = false;
                return;
            }

            string fullPath = folderManager.Trash.FullPath;

            if (!Directory.Exists(fullPath))
            {
                menuFileClearTrash.Enabled = false;
                menuClearTrash.Enabled = false;
                return;
            }

            bool hasEml = Directory.GetFiles(fullPath, "*.eml").Length > 0;
            bool hasMeta = Directory.GetFiles(fullPath, "*.meta").Length > 0;
            menuFileClearTrash.Enabled = hasEml || hasMeta;
            menuClearTrash.Enabled = hasEml || hasMeta;

            menuSaveAs.Enabled = hasOne;
            menuSpeechMail.Enabled = hasOne;
            menuAddToAddressBook.Enabled = hasOne;
            menuUndoTags.Enabled = hasOne && tagUndoStack.Count > 0;
            menuRedoTags.Enabled = hasOne && tagRedoStack.Count > 0;
            menuAttachmentFileAllSave.Enabled = hasOne && buttonAtachMenu.DropDownItems.Count > 0;

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
                string atach = string.Join(";", form.buttonAttachList.DropDownItems.Cast<ToolStripItem>().Select(item => item.Tag?.ToString() ?? ""));
                string date = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                // 宛先、本文がある場合
                if (to != "" | body != "")
                {
                    // 件名がない場合は無題
                    if (subject == "")
                    {
                        subject = "無題";
                    }
                    // コレクションに追加する
                    Mail mail = new Mail(to, cc, bcc, subject, body, atach, date, "", "", true);
                    mail.isDraft = true;
                    mail.notReadYet = true;
                    mail.Folder = folderManager.Draft;
                    SaveMail(mail);
                    // ★ Draft フォルダの一覧を更新
                    LoadEmlFolder(folderManager.Draft);
                    treeMain.SelectedNode = folderNodeMap[folderManager.Draft];
                }

                // ツリービューとリストビューの表示を更新する
                UpdateView();
            }
        }

        private void menuClearTrash_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("ごみ箱内のメールをすべて完全に削除します。",
                                "確認",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Warning) != DialogResult.OK)
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

            // ★ VirtualMode の内部リストも空にする
            _virtualList.Clear();
            listMain.VirtualListSize = 0;
            listMain.Invalidate();

            // ★ 現在メールをクリア
            currentMail = null;

            // ★ TreeView の件数更新
            UpdateTreeView();

            // ★ 画面全体を今の状態に合わせて更新
            UpdateView();

            // ★ Undo 状態更新
            UpdateUndoState();
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

        private async void menuHelpVersionCheck_Click(object sender, EventArgs e)
        {
            if (await IsNewVersionAvailableAsync())
            {
                MessageBox.Show("新しいバージョンが利用可能です。\nhttps://www.angel-teatime.com/からダウンロードしてください。", "バージョンチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("現在お使いのバージョンは最新です。", "バージョンチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static readonly HttpClient http = new HttpClient();

        public static async Task<bool> IsNewVersionAvailableAsync()
        {
            try
            {
                string url = "https://www.angel-teatime.com/files/mizumail/mizumail_version.txt";
                string versionText = (await http.GetStringAsync(url)).Trim();
                Version serverVersion = new Version(versionText);
                Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                return serverVersion > currentVersion;
            }
            catch
            {
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
            int mailCount = 0;
            string prevFolderPath = (treeMain.SelectedNode?.Tag as MailFolder)?.FullPath;

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

                        // ★ Message-ID を取得（空なら no_message_id）
                        string baseName = message.MessageId;
                        if (string.IsNullOrWhiteSpace(baseName))
                            baseName = "no_message_id";

                        // ★ ファイル名サニタイズ
                        foreach (char c in Path.GetInvalidFileNameChars())
                            baseName = baseName.Replace(c, '_');

                        // ★ GUID を付けて絶対衝突しないファイル名にする
                        string unique = Guid.NewGuid().ToString("N");
                        string mailName = $"{baseName}_{unique}_unread.eml";

                        string inboxPath = Path.Combine(folderManager.Inbox.FullPath, mailName);
                        Directory.CreateDirectory(folderManager.Inbox.FullPath);

                        // ★ EML 保存
                        // ★ FileStream + Flush(true) で保存完了を保証
                        using (var fs = new FileStream(inboxPath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            message.WriteTo(fs);
                            fs.Flush(true); // ディスクに確実に書き込む
                        }

                        // ★ Mail オブジェクト生成
                        Mail mail = Mail.FromMimeMessage(message);
                        mail.mailName = mailName;
                        mail.notReadYet = true;
                        mail.mailPath = inboxPath;
                        mail.Folder = folderManager.Inbox;

                        // ★ 振り分け処理
                        ApplyRulesForNewMail(mail);

                        // ★ mailCache に登録
                        mailCache[inboxPath] = mail;

                        updatedFolders.Add(mail.Folder);

                        // ★ UIDL 保存
                        localUidls.Add(uidl);
                        mailCount++;

                        if (Mail.deleteMail)
                            await client.DeleteMessageAsync(i);
                    }

                    await client.DisconnectAsync(true);
                }

                SaveUidls();

                labelMessage.Text = "メール受信完了";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                labelMessage.Text = "メール受信エラー : " + ex.Message;
                logger.Error(ex);
                statusStrip1.Refresh();
            }

            if (mailCount > 0)
            {
                if (Mail.alertSound && !string.IsNullOrWhiteSpace(Mail.alertSoundFile))
                {
                    try
                    {
                        if (File.Exists(Mail.alertSoundFile))
                        {
                            using (var p = new SoundPlayer(Mail.alertSoundFile))
                                p.Play();
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

            UpdateTreeView();

            if (!string.IsNullOrEmpty(prevFolderPath))
            {
                TreeNode found = FindNodeByFolderPath(treeMain.Nodes[0], prevFolderPath);
                if (found != null)
                    treeMain.SelectedNode = found;
            }

            foreach (var f in updatedFolders)
                LoadEmlFolder(f);

            UpdateListView();
        }

        private async Task Imap4Receive()
        {
            int mailCount = 0;
            string prevFolderPath = (treeMain.SelectedNode?.Tag as MailFolder)?.FullPath;

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

                        // ★ UID をベースにする（IMAP の正しい方法）
                        string baseName = uid;

                        // ★ ファイル名サニタイズ
                        foreach (char c in Path.GetInvalidFileNameChars())
                            baseName = baseName.Replace(c, '_');

                        // ★ GUID を付けて絶対衝突しないファイル名にする
                        string unique = Guid.NewGuid().ToString("N");
                        string mailName = $"{baseName}_{unique}_unread.eml";

                        string inboxPath = Path.Combine(folderManager.Inbox.FullPath, mailName);
                        Directory.CreateDirectory(folderManager.Inbox.FullPath);

                        // ★ FileStream + Flush(true) で保存完了を保証
                        using (var fs = new FileStream(inboxPath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            message.WriteTo(fs);
                            fs.Flush(true);
                        }

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

                        // ★ ここが絶対に必要（今回の NullReference の原因）
                        mail.Folder = folderManager.Inbox;

                        // 振り分け処理
                        ApplyRulesForNewMail(mail);

                        // ★ mailCache に登録（必須）
                        mailCache[inboxPath] = mail;

                        // 変更のあったフォルダを登録
                        updatedFolders.Add(mail.Folder);

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

            UpdateTreeView();

            if (!string.IsNullOrEmpty(prevFolderPath))
            {
                TreeNode found = FindNodeByFolderPath(treeMain.Nodes[0], prevFolderPath);
                if (found != null)
                    treeMain.SelectedNode = found;
            }

            // ★ 受信フォルダを再読み込みして mailCache を更新
            foreach (var f in updatedFolders)
                LoadEmlFolder(f);

            UpdateListView();
        }

        private TreeNode FindNodeByFolderPath(TreeNode node, string path)
        {
            if (node.Tag is MailFolder mf && mf.FullPath == path)
                return node;

            foreach (TreeNode child in node.Nodes)
            {
                TreeNode result = FindNodeByFolderPath(child, path);
                if (result != null)
                    return result;
            }

            return null;
        }

        private void menuUndoMail_Click(object sender, EventArgs e)
        {
            var metaFiles = Directory.GetFiles(folderManager.Trash.FullPath, "*.meta");
            if (metaFiles.Length == 0)
                return;

            foreach (var metaFile in metaFiles)
            {
                var json = File.ReadAllText(metaFile);
                var meta = JsonConvert.DeserializeObject<UndoMeta>(json);

                string newPath = meta.NewPath;
                string oldPath = meta.OldPath;

                if (!File.Exists(newPath))
                    continue;

                // ★ ごみ箱側のキャッシュを必ず削除
                if (mailCache.ContainsKey(newPath))
                    mailCache.Remove(newPath);

                // ★ oldPath に既存ファイルがあれば削除
                if (File.Exists(oldPath))
                    File.Delete(oldPath);

                // ★ メールを元の場所へ戻す
                Directory.CreateDirectory(Path.GetDirectoryName(oldPath));
                File.Move(newPath, oldPath);

                // ★ .meta 削除
                File.Delete(metaFile);
            }

            // ★ 各フォルダを再読み込み（ここで mailCache が最新になる）
            LoadEmlFolder(folderManager.Inbox);
            LoadEmlFolder(folderManager.Trash);
            LoadEmlFolder(folderManager.Send);
            LoadEmlFolder(folderManager.Draft);

            // ★ 画面全体を「今の選択フォルダ」に合わせて更新
            UpdateView();

            // ★ Undo 後に最初のメールを選択し直す（VirtualMode 対応）
            if (_virtualList.Count > 0)
            {
                listMain.SelectedIndices.Clear();
                listMain.SelectedIndices.Add(0);
            }

            // ★ Undo 状態更新（もしあれば）
            UpdateUndoState();
        }

        // 音声読み上げの準備
        private SpeechSynthesizer synth = new SpeechSynthesizer();

        // HttpClient は使い回す（現代的）
        private static readonly HttpClient client = new HttpClient();

        private async void menuSpeechMail_Click(object sender, EventArgs e)
        {
            if (currentMail == null)
                return;

            string toSpeak = $"件名: {currentMail.subject}。差出人: {currentMail.from}。本文: {currentMail.body}";

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
                logger.Debug($"読み上げエラー: {ex.Message}");
            }
        }

        public async Task SpeakWithVoiceVox(string text, int speakerId = 2)
        {
            // 1. audio_query
            var query = await client.PostAsync(
                $"http://127.0.0.1:50021/audio_query?text={Uri.EscapeDataString(text)}&speaker={speakerId}",
                null);

            query.EnsureSuccessStatusCode();

            var queryJson = await query.Content.ReadAsStringAsync();

            // JObject を使う（dynamic より現代的で安全）
            var obj = Newtonsoft.Json.Linq.JObject.Parse(queryJson);
            obj["speedScale"] = 1.3;   // 読み上げ速度調整
            queryJson = obj.ToString();

            // 2. synthesis
            var audio = await client.PostAsync(
                $"http://127.0.0.1:50021/synthesis?speaker={speakerId}",
                new StringContent(queryJson, Encoding.UTF8, "application/json"));

            audio.EnsureSuccessStatusCode();

            // 3. WAV 再生（MemoryStream）
            using (var stream = await audio.Content.ReadAsStreamAsync())
            using (var mem = new MemoryStream())
            {
                await stream.CopyToAsync(mem);
                mem.Position = 0;

                var player = new SoundPlayer(mem);
                player.Play();
            }
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
            // ★ VirtualMode の内部リストを検索結果に置き換える
            _virtualList = list;

            // ★ 件数を更新
            listMain.VirtualListSize = _virtualList.Count;

            // ★ 再描画
            listMain.Invalidate();

            // ★ 本文をクリア
            currentMail = null;
            UpdateMailView();
        }

        private void listMain_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (sortColumn == e.Column)
                sortAscending = !sortAscending;
            else
            {
                sortColumn = e.Column;
                sortAscending = true;
            }

            SortVirtualList();
        }

        private void SortVirtualList()
        {
            if (_virtualList == null || _virtualList.Count == 0)
                return;

            switch (sortColumn)
            {
                case 0: // アイコン（既読/未読）
                    _virtualList = sortAscending
                        ? _virtualList.OrderBy(m => m.notReadYet).ToList()
                        : _virtualList.OrderByDescending(m => m.notReadYet).ToList();
                    break;

                case 1: // From
                    _virtualList = sortAscending
                        ? _virtualList.OrderBy(m => m.from).ToList()
                        : _virtualList.OrderByDescending(m => m.from).ToList();
                    break;

                case 2: // Subject
                    _virtualList = sortAscending
                        ? _virtualList.OrderBy(m => m.subject).ToList()
                        : _virtualList.OrderByDescending(m => m.subject).ToList();
                    break;

                case 3: // Date
                    _virtualList = sortAscending
                        ? _virtualList.OrderBy(m => ParseDate(m.date)).ToList()
                        : _virtualList.OrderByDescending(m => ParseDate(m.date)).ToList();
                    break;

                case 4: // Size
                    _virtualList = sortAscending
                        ? _virtualList.OrderBy(m => m.sizeBytes).ToList()
                        : _virtualList.OrderByDescending(m => m.sizeBytes).ToList();
                    break;

                case 5: // プレビュー
                    _virtualList = sortAscending
                        ? _virtualList.OrderBy(m => m.preview).ToList()
                        : _virtualList.OrderByDescending(m => m.preview).ToList();
                    break;

                case 6: // タグ（ラベル）
                    _virtualList = sortAscending
                        ? _virtualList.OrderBy(m => string.Join(", ", m.Labels)).ToList()
                        : _virtualList.OrderByDescending(m => string.Join(", ", m.Labels)).ToList();
                    break;

                default:
                    break;
            }

            // ★ VirtualMode の高速再描画
            listMain.RedrawItems(0, _virtualList.Count - 1, true);
        }

        private DateTime ParseDate(string s)
        {
            if (DateTime.TryParse(s, out DateTime dt))
                return dt;
            return DateTime.MinValue;
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
                string helpPath = Path.Combine(Application.StartupPath, "help", "MizuMail.html");

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
            logger.Debug($"SaveMail called: {mail.mailPath}");

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
                message.Headers.Add("X-Mailer", "MizuMail " + Application.ProductVersion);

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

            // spam / send / draft / trash 以外はサブフォルダ作成OK
            if (parent.Type == FolderType.Spam || parent.Type == FolderType.Send || parent.Type == FolderType.Draft || parent.Type == FolderType.Trash)
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
                folder.Type == FolderType.Spam ||
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
            // VirtualMode では SelectedItems は使えない
            if (listMain.SelectedIndices.Count == 0)
                return;

            // ★ 選択されたメールを取得
            var mails = listMain.SelectedIndices
                                .Cast<int>()
                                .Select(i => _virtualList[i])
                                .Where(m => m != null)
                                .ToList();

            if (mails.Count == 0)
                return;

            // ★ ドラッグ開始（従来の処理）
            DoDragDrop(mails, DragDropEffects.Move);
        }

        private void treeMain_DragDrop(object sender, DragEventArgs e)
        {
            // ★ 複数メール対応：List<Mail> を受け取る
            if (!e.Data.GetDataPresent(typeof(List<Mail>)))
                return;

            var mails = e.Data.GetData(typeof(List<Mail>)) as List<Mail>;
            if (mails == null || mails.Count == 0)
                return;

            // ★ ドロップ先フォルダを取得
            Point pt = treeMain.PointToClient(new Point(e.X, e.Y));
            TreeNode node = treeMain.GetNodeAt(pt);
            if (node == null)
                return;

            MailFolder targetFolder = node.Tag as MailFolder;
            if (targetFolder == null)
                return;

            // ★ 複数メールをまとめて移動
            foreach (var mail in mails)
            {
                bool fromTrash = (mail.Folder.Type == FolderType.Trash);

                if (targetFolder.Type == FolderType.Trash)
                {
                    // ごみ箱へ移動 → Undo 情報を作る
                    MoveMailWithUndo(mail, folderManager.Trash);
                }
                else if (mail.Folder != targetFolder)
                {
                    // ごみ箱から出す場合は .meta を削除
                    if (fromTrash)
                    {
                        string meta = mail.mailPath + ".meta";
                        if (File.Exists(meta))
                            File.Delete(meta);
                    }

                    // 通常の移動
                    MoveMailWithUndo(mail, targetFolder);
                }
            }

            // ★ D&D 完了後にフォルダを再読み込み（Undo と同じ）
            LoadEmlFolder(folderManager.Inbox);
            LoadEmlFolder(folderManager.Trash);
            LoadEmlFolder(folderManager.Send);
            LoadEmlFolder(folderManager.Draft);

            // ★ UI更新は DragDrop 完了後に行う（元のコードをそのまま）
            this.BeginInvoke(new Action(delegate
            {
                if (treeMain.Nodes.Count == 0)
                    return;

                BuildTree();

                if (treeMain.SelectedNode == null)
                {
                    if (treeMain.Nodes.Count > 0 && treeMain.Nodes[0].Nodes.Count > 0)
                        treeMain.SelectedNode = treeMain.Nodes[0].Nodes[0];
                    else
                        return;
                }

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
            logger.Debug($"LoadEmlFolder START: {folder.FullPath}");

            var list = new List<Mail>();

            if (!Directory.Exists(folder.FullPath))
                return list;

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
                    if (fromMailbox != null)
                    {
                        // 既存の文字列版（UI 表示用）
                        mail.from = !string.IsNullOrEmpty(fromMailbox.Name)
                            ? fromMailbox.Name + " <" + fromMailbox.Address + ">"
                            : fromMailbox.Address;

                        // ★ 新しい MailAddress 版（なりすまし検知用）
                        mail.From = new MailAddress(fromMailbox.Address, fromMailbox.Name);
                    }
                    else
                    {
                        mail.from = "(差出人なし)";
                        mail.From = new MailAddress("unknown@example.com", "(差出人なし)");
                    }

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
                    string bodyText = message.GetTextBody(MimeKit.Text.TextFormat.Plain) 
                        ?? message.GetTextBody(MimeKit.Text.TextFormat.Html)
                        ?? "";

                    mail.body = bodyText;
                    mail.isHtml = message.GetTextBody(MimeKit.Text.TextFormat.Html) != null;
                    mail.preview = BuildPreview(mail.body, mail.isHtml);

                    // ★ タグ読み込み
                    var labels = message.Headers["X-MizuMail-Labels"];
                    if (!string.IsNullOrEmpty(labels))
                    {
                        mail.Labels = labels
                            .Split(',')
                            .Select(t => t.Trim())
                            .Where(t => t.Length > 0)
                            .ToList();
                    }
                    else
                    {
                        mail.Labels = new List<string>();
                    }

                    // ★ ここで ApplyRules を呼ぶ（最適）
                    ApplyLabelRules(mail);

                    LoadMailCache(mail);

                    list.Add(mail);

                    // ★ mailCache に登録（これが CountUnread と同期の鍵）
                    if (mail != null)
                        mailCache[file] = mail;
                }
                catch (Exception ex)
                {
                    logger.Error("LoadEmlFolder error: " + file + " : " + ex.Message);
                    if (mailCache.ContainsKey(file))
                        mailCache.Remove(file);
                }
            }

            logger.Debug($"LoadEmlFolder END: {folder.FullPath}, count={list.Count}");

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
            // ★ 複数メール対応：List<Mail> を受け入れる
            if (!e.Data.GetDataPresent(typeof(List<Mail>)))
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
            string folderPath = Path.Combine(Application.StartupPath, "mbox", "inbox", folderName);

            if (!Directory.Exists(folderPath))
                return;

            // ★ ごみ箱フォルダ
            string trashDir = Path.Combine(Application.StartupPath, "mbox", "trash");
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
            string folderPath = Path.Combine(Application.StartupPath, "mbox", folder.Replace("/", "\\"));

            if (!Directory.Exists(folderPath))
                return;

            // ごみ箱フォルダ
            string trashDir = Path.Combine(Application.StartupPath, "mbox", "trash");
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
                string atach = string.Join(";", form.buttonAttachList.DropDownItems.Cast<ToolStripItem>().Select(i => i.Tag?.ToString() ?? ""));

                if (!string.IsNullOrEmpty(to) || !string.IsNullOrEmpty(body))
                {
                    if (mail.Folder.Type == FolderType.Send)
                    {
                        var draft = new Mail
                        {
                            address = to,
                            ccaddress = cc,
                            bccaddress = bcc,
                            subject = string.IsNullOrEmpty(subject) ? "無題" : subject,
                            body = body,
                            atach = atach,
                            isDraft = true,
                            notReadYet = true,
                            Folder = folderManager.Draft
                        };
                        SaveMail(draft);
                        // ★ Draft フォルダの一覧を更新
                        LoadEmlFolder(folderManager.Draft);
                    }
                    else
                    {
                        mail.address = to;
                        mail.ccaddress = cc;
                        mail.bccaddress = bcc;
                        mail.subject = string.IsNullOrEmpty(subject) ? "無題" : subject;
                        mail.body = body;
                        mail.atach = atach;
                        mail.isDraft = true;
                        mail.notReadYet = true;
                        mail.Folder = folderManager.Draft;
                        SaveMail(mail);
                        // ★ Draft フォルダの一覧を更新
                        LoadEmlFolder(folderManager.Draft);
                        treeMain.SelectedNode = folderNodeMap[folderManager.Draft];
                    }
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
            string dir = Path.Combine(Application.StartupPath, "mbox", newFolder.Replace("/", "\\"));
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
            mailBoxViewFlag = true;

            listMain.BeginUpdate();
            try
            {
                // ★ ルートフォルダ用のカラム名に切り替え
                listMain.Columns[1].Text = "メールボックス名";
                listMain.Columns[2].Text = "メールアドレス";
                listMain.Columns[3].Text = "更新日時";
                listMain.Columns[4].Text = "サイズ";
                listMain.Columns[5].Text = "";   // プレビュー列は使わない
                listMain.Columns[6].Text = "";   // タグ列も使わない
                listMain.Columns[7].Text = "";   // ファイル名列も使わない

                // ★ 表示する情報を作成
                var item = new ListViewItem(" ", 0); // アイコン列

                item.SubItems.Add(Mail.fromName);      // メールボックス名
                item.SubItems.Add(Mail.userAddress);   // メールアドレス

                // 更新日時
                string inboxPath = Path.Combine(Application.StartupPath, "mbox", "inbox");
                item.SubItems.Add(Directory.GetLastWriteTime(inboxPath).ToString("yyyy/MM/dd HH:mm:ss"));

                // サイズ
                long directorySize = 0;
                GetDirectorySize(Path.Combine(Application.StartupPath, "mbox"), ref directorySize);
                item.SubItems.Add(FormatSize(directorySize));

                // 使わない列は空文字で埋める
                item.SubItems.Add(""); // preview
                item.SubItems.Add(""); // preview
                item.SubItems.Add(""); // mailName

                listMain.Items.Add(item);
            }
            finally
            {
                listMain.EndUpdate();
            }
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
                folder.Type == FolderType.Spam ||
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
            inboxFolder.SubFolders.Clear(); // ★ ここが唯一のサブフォルダリスト

            foreach (var dir in Directory.GetDirectories(inboxFolder.FullPath))
            {
                string name = Path.GetFileName(dir);

                // ★ 既存のサブフォルダがあるか探す
                var folder = inboxFolder.SubFolders
                    .FirstOrDefault(f => f.FullPath.Equals(dir, StringComparison.OrdinalIgnoreCase));

                // ★ なければ作る（これが唯一の new）
                if (folder == null)
                {
                    folder = new MailFolder(name, dir, FolderType.InboxSub);
                    inboxFolder.SubFolders.Add(folder);
                }

                // ★ TreeView に folderManager のインスタンスを入れる
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
            if (e.Data.GetDataPresent(typeof(List<Mail>)))
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
            string path = Path.Combine(Application.StartupPath, "uidl.txt");
            if (File.Exists(path))
                localUidls = File.ReadAllLines(path).ToList();
        }

        private void SaveUidls()
        {
            string path = Path.Combine(Application.StartupPath, "uidl.txt");
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

            // ★ これが必要
            mailCache.Remove(path);
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
            if (listMain.SelectedIndices.Count == 0)
                return;

            foreach (int index in listMain.SelectedIndices)
            {
                Mail mail = _virtualList[index];
                if (mail != null && mail.notReadYet)
                {
                    ToggleReadState(mail);
                }
            }

            // ★ VirtualMode では全体再描画ではなく、部分再描画が高速
            listMain.RedrawItems(
                listMain.SelectedIndices[0],
                listMain.SelectedIndices[listMain.SelectedIndices.Count - 1],
                true
            );

            UpdateMailView(); // 本文更新
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
            if (newFolder == null)
            {
                logger.Error("MoveMailWithUndo: newFolder is null");
                return;
            }

            string oldPath = ResolveMailPath(mail);

            // ★ 残像対策（1回でOK）
            if (mailCache.ContainsKey(oldPath))
                mailCache.Remove(oldPath);

            FolderType oldFolderType = mail.Folder.Type;
            string oldFolderPath = mail.Folder.FullPath;
            string oldName = mail.mailName;

            // ★ 新しいファイル名を決める
            if (newFolder.Type == FolderType.Draft)
            {
                mail.notReadYet = true;

                if (!oldName.EndsWith("_unread.eml"))
                {
                    string baseName = Path.GetFileNameWithoutExtension(oldName);
                    mail.mailName = baseName + "_unread.eml";
                }
            }
            else
            {
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

            // ★ 先にファイルを移動する
            Directory.CreateDirectory(newFolder.FullPath);
            if (File.Exists(oldPath))
                File.Move(oldPath, newPath);

            // ★ mailCache 更新（ここでは oldPath はもう存在しないので Remove 不要）
            mail.mailPath = newPath;
            mail.Folder = newFolder;
            mailCache[newPath] = mail;

            // ★ ごみ箱へ移動する場合だけ Undo 情報を JSON で保存
            if (newFolder.Type == FolderType.Trash)
            {
                var meta = new UndoMeta
                {
                    OldPath = oldPath,
                    NewPath = newPath,
                    OldFolder = oldFolderType,
                    OldFolderPath = oldFolderPath
                };

                string metaJson = JsonConvert.SerializeObject(meta, Formatting.Indented);
                File.WriteAllText(newPath + ".meta", metaJson);
            }
        }

        private Mail LoadSingleMail(string path)
        {
            if (!File.Exists(path))
                return null;

            var message = MimeMessage.Load(path);

            Mail mail = new Mail();
            mail.mailName = Path.GetFileName(path);

            // From
            var fromMailbox = message.From.Mailboxes.FirstOrDefault();
            mail.from = fromMailbox != null ? fromMailbox.ToString() : "";

            // To
            mail.address = string.Join("; ", message.To.Mailboxes.Select(m => m.ToString()));

            // Cc
            mail.ccaddress = string.Join("; ", message.Cc.Mailboxes.Select(m => m.ToString()));

            // Bcc
            mail.bccaddress = string.Join("; ", message.Bcc.Mailboxes.Select(m => m.ToString()));

            // Subject
            mail.subject = message.Subject ?? "";

            // ★ HTML メール対応（最重要）
            mail.body = message.HtmlBody ?? message.TextBody ?? "";

            // Date
            mail.date = message.Date.LocalDateTime.ToString("yyyy/MM/dd HH:mm:ss");

            // 未読判定
            mail.notReadYet = path.EndsWith("_unread.eml");

            return mail;
        }

        public static void SaveMailLabels(Mail mail)
        {
            if (mail == null || string.IsNullOrEmpty(mail.mailPath))
                return;

            try
            {
                var msg = MimeMessage.Load(mail.mailPath);

                // 既存のタグヘッダを削除
                msg.Headers.Remove("X-MizuMail-Labels");

                // タグが残っている場合だけ書き込む
                if (mail.Labels != null && mail.Labels.Count > 0)
                {
                    msg.Headers.Add("X-MizuMail-Labels", string.Join(", ", mail.Labels));
                }

                // 上書き保存
                using (var stream = File.Create(mail.mailPath))
                {
                    msg.WriteTo(stream);
                }
            }
            catch (Exception ex)
            {
                logger.Error($"タグ保存中にエラー: {ex}");
            }
        }

        private void SaveRules()
        {
            string path = Path.Combine(Application.StartupPath, "rules.json");
            File.WriteAllText(path, JsonConvert.SerializeObject(rules, Formatting.Indented));
        }

        private void LoadRules()
        {
            string path = Path.Combine(Application.StartupPath, "rules.json");
            if (File.Exists(path))
            {
                rules = JsonConvert.DeserializeObject<List<MailRule>>(File.ReadAllText(path));
            }
        }

        private void ApplyRulesForNewMail(Mail mail)
        {
            if (rules == null || rules.Count == 0)
                return;

            string subject = mail.subject ?? "";
            string from = mail.from ?? "";

            // メールアドレス抽出
            string fromAddress = ExtractAddress(from);

            string oldPath = mail.mailPath;

            foreach (var rule in rules)
            {
                bool match = false;

                // 件名
                if (!string.IsNullOrWhiteSpace(rule.Contains))
                {
                    if (rule.UseRegex)
                    {
                        if (Regex.IsMatch(subject, rule.Contains, RegexOptions.IgnoreCase))
                            match = true;
                    }
                    else
                    {
                        if (subject.IndexOf(rule.Contains, StringComparison.OrdinalIgnoreCase) >= 0)
                            match = true;
                    }
                }

                // 差出人（表示名＋アドレス両方を対象にする）
                string fromFull = mail.from ?? "";

                if (!string.IsNullOrWhiteSpace(rule.From))
                {
                    if (rule.UseRegex)
                    {
                        if (Regex.IsMatch(fromFull, rule.From, RegexOptions.IgnoreCase) ||
                            Regex.IsMatch(fromAddress, rule.From, RegexOptions.IgnoreCase))
                            match = true;
                    }
                    else
                    {
                        if (fromFull.IndexOf(rule.From, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            fromAddress.IndexOf(rule.From, StringComparison.OrdinalIgnoreCase) >= 0)
                            match = true;
                    }
                }

                // ★ なりすまし防止（新着メールのみ）
                if (IsSuspiciousSender(mail))
                {
                    MoveMailWithUndo(mail, folderManager.Spam);
                    UpdateMailCacheAfterMove(oldPath, mail);
                    return;
                }

                if (!match)
                    continue;

                // ★ ラベル付与
                if (!string.IsNullOrWhiteSpace(rule.Label))
                {
                    if (!mail.Labels.Contains(rule.Label))
                        mail.Labels.Add(rule.Label);
                }

                // ★ 移動先フォルダ
                if (!string.IsNullOrWhiteSpace(rule.MoveTo))
                {
                    MailFolder target = folderManager.GetOrCreateFolderByPath(rule.MoveTo);
                    if (target != null)
                    {
                        MoveMailWithUndo(mail, target);
                        UpdateMailCacheAfterMove(oldPath, mail);
                    }
                }

                break; // 最初の一致だけ適用
            }
        }

        private void ApplyLabelRules(Mail mail)
        {
            if (rules == null || rules.Count == 0)
                return;

            string subject = mail.subject ?? "";
            string from = mail.from ?? "";
            string fromAddress = ExtractAddress(from);

            foreach (var rule in rules)
            {
                bool match = false;

                // 件名
                if (!string.IsNullOrWhiteSpace(rule.Contains))
                {
                    if (rule.UseRegex)
                    {
                        if (Regex.IsMatch(subject, rule.Contains, RegexOptions.IgnoreCase | RegexOptions.Compiled))
                            match = true;
                    }
                    else
                    {
                        if (subject.IndexOf(rule.Contains, StringComparison.OrdinalIgnoreCase) >= 0)
                            match = true;
                    }
                }

                // 差出人（表示名＋アドレス両方を対象にする）
                string fromFull = mail.from ?? "";

                if (!string.IsNullOrWhiteSpace(rule.From))
                {
                    if (rule.UseRegex)
                    {
                        if (Regex.IsMatch(fromFull, rule.From, RegexOptions.IgnoreCase | RegexOptions.Compiled) ||
                            Regex.IsMatch(fromAddress, rule.From, RegexOptions.IgnoreCase | RegexOptions.Compiled))
                            match = true;
                    }
                    else
                    {
                        if (fromFull.IndexOf(rule.From, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            fromAddress.IndexOf(rule.From, StringComparison.OrdinalIgnoreCase) >= 0)
                            match = true;
                    }
                }

                if (!match)
                    continue;

                // ★ ラベル付与のみ
                if (!string.IsNullOrWhiteSpace(rule.Label))
                {
                    if (!mail.Labels.Contains(rule.Label))
                        mail.Labels.Add(rule.Label);
                }
            }
        }

        private string ExtractAddress(string from)
        {
            string addr = from ?? "";
            int lt = addr.IndexOf('<');
            int gt = addr.IndexOf('>');
            if (lt >= 0 && gt > lt)
                return addr.Substring(lt + 1, gt - lt - 1);
            return addr;
        }

        private void UpdateMailCacheAfterMove(string oldPath, Mail mail)
        {
            string newPath = mail.mailPath;

            if (mailCache.ContainsKey(oldPath))
                mailCache.Remove(oldPath);

            mailCache[newPath] = mail;
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
            isBuildingTree = true;
            try
            {
                treeMain.BeginUpdate();

                treeMain.Nodes.Clear();
                folderNodeMap.Clear();

                // ルートノード
                var root = new TreeNode("メール");
                root.Tag = null; // ルートは MailFolder ではない
                treeMain.Nodes.Add(root);

                // 受信
                var inboxNode = CreateFolderNode("受信メール", folderManager.Inbox);
                root.Nodes.Add(inboxNode);

                // 迷惑
                var spamNode = CreateFolderNode("迷惑メール", folderManager.Spam);
                root.Nodes.Add(spamNode);

                // 送信
                var sendNode = CreateFolderNode("送信メール", folderManager.Send);
                root.Nodes.Add(sendNode);

                // 下書き
                var draftNode = CreateFolderNode("下書き", folderManager.Draft);
                root.Nodes.Add(draftNode);

                // ごみ箱
                var trashNode = CreateFolderNode("ごみ箱", folderManager.Trash);
                root.Nodes.Add(trashNode);

                // Inbox のサブフォルダ（既存の再帰ロジックをそのまま利用）
                LoadInboxFolders(inboxNode, folderManager.Inbox);

                treeMain.ExpandAll();
            }
            finally
            {
                treeMain.EndUpdate();
                isBuildingTree = false;
            }

            // 件数表示は既存の UpdateTreeView を使う
            UpdateTreeView();
        }

        private TreeNode CreateFolderNode(string caption, MailFolder folder)
        {
            int count = Directory.GetFiles(folder.FullPath, "*.eml").Length;
            var node = new TreeNode($"{caption}({count})");
            node.Tag = folder;

            // ★ フォルダ種別ごとにアイコンを設定
            switch (folder.Type)
            {
                case FolderType.Inbox:
                    node.ImageIndex = 1;            // 受信
                    node.SelectedImageIndex = 1;
                    break;

                case FolderType.Spam:
                    node.ImageIndex = 0;            // 迷惑
                    node.SelectedImageIndex = 0;
                    break;

                case FolderType.Send:
                    node.ImageIndex = 2;            // 送信
                    node.SelectedImageIndex = 2;
                    break;

                case FolderType.Draft:
                    node.ImageIndex = 3;            // 下書き
                    node.SelectedImageIndex = 3;
                    break;

                case FolderType.Trash:
                    node.ImageIndex = 4;            // ごみ箱
                    node.SelectedImageIndex = 4;
                    break;

                default:
                    node.ImageIndex = 0;            // 通常フォルダ
                    node.SelectedImageIndex = 0;
                    break;
            }

            folderNodeMap[folder] = node;
            return node;
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
                case FolderType.Spam:
                    return "spam";
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

        private void ShowHtml(string html)
        {
            if (browserMail.CoreWebView2 != null)
            {
                browserMail.CoreWebView2.NavigateToString(html);
                return;
            }

            // 初期化前なら初期化後に一度だけ実行
            browserMail.CoreWebView2InitializationCompleted += (s, e) =>
            {
                browserMail.CoreWebView2.NavigateToString(html);
            };
        }

        private async void ShowHtmlWithInlineImages(MimeMessage message, Mail mail)
        {
            if (message == null)
            {
                ShowHtml("");
                return;
            }

            // ================================
            // ★ HTML を安全に取得（最重要）
            // ================================
            string html =
                message.GetTextBody(MimeKit.Text.TextFormat.Html)
                ?? message.GetTextBody(MimeKit.Text.TextFormat.Plain)
                ?? "";

            // HTML が空なら空表示
            if (string.IsNullOrWhiteSpace(html))
            {
                ShowHtml("");
                return;
            }

            // ================================
            // ★ inline 画像の抽出
            // ================================
            IEnumerable<MimePart> GetAllParts(MimeEntity entity)
            {
                if (entity is MimePart p)
                    yield return p;

                if (entity is Multipart mp)
                {
                    foreach (var child in mp)
                        foreach (var c in GetAllParts(child))
                            yield return c;
                }
            }

            var inlineImages = new Dictionary<string, string>();

            foreach (var part in GetAllParts(message.Body))
            {
                if (part.ContentId != null &&
                    part.ContentType.MediaType.Equals("image", StringComparison.OrdinalIgnoreCase))
                {
                    string cid = part.ContentId.Trim('<', '>');

                    using (var ms = new MemoryStream())
                    {
                        part.Content.DecodeTo(ms);
                        var bytes = ms.ToArray();
                        string base64 = Convert.ToBase64String(bytes);

                        string mime = part.ContentType.MimeType; // 例: image/jpeg
                        string dataUri = $"data:{mime};base64,{base64}";

                        inlineImages[cid] = dataUri;
                    }
                }
            }

            // ================================
            // ★ cid:xxx → data URI に置換
            // ================================
            foreach (var kv in inlineImages)
            {
                html = html.Replace("cid:" + kv.Key, kv.Value);
            }

            // ================================
            // ★ UI を固めないための遅延（重要）
            // ================================
            await Task.Yield(); // UI に Visible 切り替えを反映させる

            ShowHtml(html);
        }

        private string GetExtension(MimePart part)
        {
            // 1. FileName があればそこから
            if (!string.IsNullOrEmpty(part.FileName))
                return Path.GetExtension(part.FileName);

            // 2. Content-Type から推測
            string mime = part.ContentType.MimeType.ToLower();

            return mime switch
            {
                "image/jpeg" => ".jpg",
                "image/jpg" => ".jpg",
                "image/png" => ".png",
                "image/gif" => ".gif",
                "image/bmp" => ".bmp",
                "image/webp" => ".webp",
                "image/svg+xml" => ".svg",
                _ => ".bin"
            };
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

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SHGetFileInfo(
            string pszPath,
            uint dwFileAttributes,
            out SHFILEINFO psfi,
            uint cbFileInfo,
            uint uFlags);

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

        public static Icon GetIconFromExtension(string ext)
        {
            SHFILEINFO shinfo = new SHFILEINFO();

            IntPtr hImg = SHGetFileInfo(
                ext,
                0x80, // FILE_ATTRIBUTE_NORMAL
                out shinfo,
                (uint)Marshal.SizeOf(shinfo),
                0x100 | 0x10 | 0x4000);
            // SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES

            if (shinfo.hIcon != IntPtr.Zero)
                return Icon.FromHandle(shinfo.hIcon);

            return SystemIcons.WinLogo; // fallback
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
                ApplyLabelRules(mail);  // ← ラベル付与のルールを適用
            }
        }

        private void menuAddressBook_Click(object sender, EventArgs e)
        {
            // アドレス帳を読み込む
            var book = AddressBook.LoadAddressBook();

            using (var dlg = new AddressBookEditorForm(book))
            {
                dlg.SelectMode = false; // 編集モード

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // 保存
                    AddressBook.SaveAddressBook(book);
                }
            }
        }

        private (string name, string email) ParseFrom(string from)
        {
            string name = from;
            string email = from;

            int lt = from.IndexOf('<');
            int gt = from.IndexOf('>');

            if (lt >= 0 && gt > lt)
            {
                name = from.Substring(0, lt).Trim();
                email = from.Substring(lt + 1, gt - lt - 1).Trim();
            }

            return (name, email);
        }

        private void menuAddToAddressBook_Click(object sender, EventArgs e)
        {
            if (currentMail == null)
                return;

            // From から抽出
            var (name, email) = ParseFrom(currentMail.from);

            if (string.IsNullOrWhiteSpace(email))
                return;

            // アドレス帳読み込み
            var book = AddressBook.LoadAddressBook();

            // 既に登録済みかチェック
            if (book.Entries.Any(a =>a.Email.Equals(email, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("このアドレスは既に登録されています。");
                return;
            }

            // 新規追加
            var entry = new AddressEntry()
            {
                DisplayName = name,
                Email = email,
                Note = ""
            };

            book.Entries.Add(entry);

            // 保存
            AddressBook.SaveAddressBook(book);

            MessageBox.Show("アドレス帳に追加しました。");
        }

        public static int Levenshtein(string s, string t)
        {
            if (string.IsNullOrEmpty(s)) return t?.Length ?? 0;
            if (string.IsNullOrEmpty(t)) return s.Length;

            int[,] d = new int[s.Length + 1, t.Length + 1];

            for (int i = 0; i <= s.Length; i++)
                d[i, 0] = i;

            for (int j = 0; j <= t.Length; j++)
                d[0, j] = j;

            for (int i = 1; i <= s.Length; i++)
            {
                for (int j = 1; j <= t.Length; j++)
                {
                    int cost = (s[i - 1] == t[j - 1]) ? 0 : 1;

                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost
                    );
                }
            }

            return d[s.Length, t.Length];
        }

        public static string ExtractBrandName(string displayName)
        {
            if (string.IsNullOrEmpty(displayName))
                return "";

            // 英字だけ抽出
            var alpha = new string(displayName.Where(c => char.IsLetter(c) && c <= 127).ToArray());
            if (!string.IsNullOrEmpty(alpha))
                return alpha.ToLowerInvariant();

            // 日本語だけ抽出（漢字・ひらがな・カタカナ）
            var jp = new string(displayName.Where(c =>
                (c >= 0x4E00 && c <= 0x9FFF) ||     // 漢字
                (c >= 0x3040 && c <= 0x309F) ||     // ひらがな
                (c >= 0x30A0 && c <= 0x30FF)        // カタカナ
            ).ToArray());

            return jp;
        }

        public static string ExtractDomain(string address)
        {
            if (string.IsNullOrEmpty(address))
                return "";

            int at = address.IndexOf('@');
            if (at < 0) return "";

            return address.Substring(at + 1).ToLowerInvariant();
        }

        public static string ExtractDomainMain(string domain)
        {
            if (string.IsNullOrEmpty(domain))
                return "";

            var parts = domain.Split('.');

            // 例: e.resonabank.co.jp → ["e","resonabank","co","jp"]
            if (parts.Length >= 3)
            {
                // 最後が jp / kr / in など → 企業名は最後から3番目
                return parts[parts.Length - 3].ToLowerInvariant();
            }

            // それ以外は最後から2番目
            if (parts.Length >= 2)
                return parts[parts.Length - 2].ToLowerInvariant();

            return domain.ToLowerInvariant();
        }

        public static bool IsJapanese(string s)
        {
            return s.Any(c =>
                (c >= 0x3040 && c <= 0x309F) || // ひらがな
                (c >= 0x30A0 && c <= 0x30FF) || // カタカナ
                (c >= 0x4E00 && c <= 0x9FFF));  // 漢字
        }

        public static bool IsAscii(string s)
        {
            return s.All(c => c < 128);
        }

        static readonly HashSet<string> DomainIgnoreList = new HashSet<string>
        {
            "info", "mail", "news", "mailing", "support", "contact",
            "noreply", "no-reply", "service", "update", "notification"
        };

        static readonly Dictionary<string, HashSet<string>> BrandDomainWhitelist = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase)
        {
            // Amazon
            ["amazon"] = new HashSet<string>
            {
                "amazon", "amazonaws", "amazonses", "amazonservices", "amazon-adsystem", "amazonpay"
            },
            
            // Apple
            ["apple"] = new HashSet<string>
            {
                "apple", "insideapple", "email", "developer", "id", "updates", "news"
            },
            // 楽天
            ["rakuten"] = new HashSet<string>
            {
                "rakuten", "rakuten-card", "rakuten-bank", "rakuten-sec", "rakuten-pay"
            },
            // 三井住友カード
            ["smbc"] = new HashSet<string> { "smbc", "vpass" },
            ["三井住友"] = new HashSet<string> { "smbc", "vpass" },
            
            // ゆうちょ銀行
            ["ゆうちょ"] = new HashSet<string> { "jp-bank", "japanpost", "yucho" },
            
            // PayPay銀行
            ["paypay-bank"] = new HashSet<string> { "paypay-bank", "jnb", "japannetbank" },
            
            // PayPay（決済）
            ["paypay"] = new HashSet<string> { "paypay", "paypay-bank", "paypay-card", "yahoo" },

            // Yahoo!
            ["yahoo"] = new HashSet<string> { "yahoo", "ybb" },

            // au / au PAY
            ["au"] = new HashSet<string> { "au", "kddi", "auone", "aupay" },
            ["kddi"] = new HashSet<string> { "au", "kddi", "auone", "aupay" },

            // docomo / d払い
            ["docomo"] = new HashSet<string> { "docomo", "nttdocomo", "dpoint" },
            ["ドコモ"] = new HashSet<string> { "docomo", "nttdocomo", "dpoint" },

            // SoftBank
            ["softbank"] = new HashSet<string> { "softbank", "ymobile" },
            
            // LINE
            ["line"] = new HashSet<string> { "line", "linecorp", "linepay" },
            
            // メルカリ
            ["mercari"] = new HashSet<string> { "mercari" },

            // JTB
            ["jtb"] = new HashSet<string>
            {
                "jtb", "jtbcorp", "jtbtravel", "jtb-global"
            },

            ["ジェイティービー"] = new HashSet<string>
            {
                "jtb", "jtbcorp", "jtbtravel", "jtb-global"
            }
        };

        public static bool IsSuspiciousSender(Mail mail)
        {
            // ★ 基本的な null チェック
            if (mail == null || mail.message == null || mail.From == null)
                return false;

            if (mail.From.DisplayName == null || mail.From.Address == null)
                return false;

            // ★ 差出人名からブランド名を抽出（Leet 正規化込み）
            string brand = ExtractBrandName(mail.From.DisplayName);
            brand = NormalizeLeet(brand); // ← 追加

            // ★ メールアドレスからドメイン抽出
            string domain = ExtractDomain(mail.From.Address);
            string domainMain = ExtractDomainMain(domain);

            if (string.IsNullOrEmpty(domain) || string.IsNullOrEmpty(domainMain))
                return false;

            // ★ TLD（トップレベルドメイン）が怪しい場合は即アウト
            if (IsSuspiciousTld(domain))
                return true;

            // ★ 無視すべき一般語ドメイン（info, mail, support など）
            if (DomainIgnoreList.Contains(domainMain))
                return false;

            // ★ ブランド名が空 → 判定不可（普通の個人メールなど）
            if (string.IsNullOrEmpty(brand))
                return false;

            // ★ ブランド名とホワイトリストの照合
            foreach (var kv in BrandDomainWhitelist)
            {
                // ブランド名にキーが含まれる、または同義語が含まれる
                bool brandMatchesKey =
                    brand.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0;

                bool brandMatchesAlias =
                    kv.Value.Any(v => brand.IndexOf(v, StringComparison.OrdinalIgnoreCase) >= 0);

                if (brandMatchesKey || brandMatchesAlias)
                {
                    // 許可ドメインに含まれていなければ偽装
                    if (!kv.Value.Contains(domainMain))
                        return true;

                    // 許可ドメインなら安全
                    return false;
                }
            }

            // ★ 日本語ブランド名は距離判定しない
            if (!IsAscii(brand))
                return false;

            // ★ 長さ差が大きい場合は距離判定しない（Apple Developer 対策）
            if (Math.Abs(brand.Length - domainMain.Length) >= 5)
                return false;

            // ★ Levenshtein 距離判定（動的閾値）
            int distance = Levenshtein(brand, domainMain);

            // 閾値：ブランド名の長さに応じて変動
            int threshold = Math.Max(4, brand.Length / 2 + 2);

            logger.Debug($"[SpoofCheck] Brand='{brand}', Domain='{domainMain}', Distance={distance}, Threshold={threshold}, From='{mail.from}'");

            return distance >= threshold;
        }

        static readonly Dictionary<char, char> LeetMap = new Dictionary<char, char>
        {
            ['0'] = 'o',
            ['1'] = 'l',
            ['3'] = 'e',
            ['4'] = 'a',
            ['5'] = 's',
            ['7'] = 't'
        };

        public static string NormalizeLeet(string s)
        {
            if (string.IsNullOrEmpty(s))
                return s;

            return new string(s.Select(c =>
                LeetMap.TryGetValue(c, out var mapped) ? mapped : c
            ).ToArray());
        }

        static readonly HashSet<string> SuspiciousTlds = new HashSet<string>
        {
            "xyz", "top", "shop", "work", "loan", "click", "info"
        };

        public static bool IsSuspiciousTld(string domain)
        {
            var parts = domain.Split('.');
            string tld = parts.Last();
            return SuspiciousTlds.Contains(tld);
        }

        public static MailAddress ParseMailAddress(string raw)
        {
            try
            {
                return new MailAddress(raw);
            }
            catch
            {
                // 失敗したら DisplayName だけでも返す
                return new MailAddress("unknown@example.com", raw);
            }
        }

        private void menuAttachmentFileAllSave_Click(object sender, EventArgs e)
        {
            if (listMain.SelectedIndices.Count == 0)
                return;

            // ★ VirtualMode では SelectedItems を使わない
            int index = listMain.SelectedIndices[0];
            Mail mail = _virtualList[index];

            if (mail == null || mail.message == null)
                return;

            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "添付ファイルの保存先フォルダを選択してください";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                string saveDir = dialog.SelectedPath;

                foreach (var attachment in mail.message.Attachments)
                {
                    if (attachment is MimePart part)
                    {
                        string fileName = part.FileName;
                        if (string.IsNullOrEmpty(fileName))
                            fileName = "attachment.bin";

                        string savePath = Path.Combine(saveDir, fileName);

                        using (var stream = File.Create(savePath))
                        {
                            part.Content.DecodeTo(stream);
                        }
                    }
                }

                MessageBox.Show("添付ファイルを保存しました。", "完了",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void SaveSignature(SignatureConfig config)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "signature.json");
            string json = JsonConvert.SerializeObject(config, Formatting.Indented);
            File.WriteAllText(path, json, Encoding.UTF8);
        }

        public static SignatureConfig LoadSignature()
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "signature.json");

            if (!File.Exists(path))
                return new SignatureConfig(); // 初期値

            string json = File.ReadAllText(path, Encoding.UTF8);
            return JsonConvert.DeserializeObject<SignatureConfig>(json);
        }

        private void menuSignatureSetting_Click(object sender, EventArgs e)
        {
            FormSignature form = new FormSignature();
            form.ShowDialog(this);
            if (form.DialogResult == DialogResult.OK)
            {
                // 保存
                MessageBox.Show("署名を保存しました。", "署名", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void toolShowHeader_Click(object sender, EventArgs e)
        {
            showHeader = !showHeader;
            toolShowHeader.Checked = showHeader;
            UpdateMailView();
        }

        private void UpdateMailView()
        {
            if (currentMail == null)
            {
                richTextBody.Clear();
                browserMail.NavigateToString("<html></html>");
                return;
            }

            if (showHeader)
                ShowHeaderView(currentMail);
            else
                ShowNormalView(currentMail);
        }

        private void ShowHeaderView(Mail mail)
        {
            if (mail == null || mail.message == null)
            {
                richTextBody.Clear();
                browserMail.NavigateToString("<html></html>");
                return;
            }

            richTextBody.Visible = true;
            richTextBody.BackColor = SystemColors.Window;   // ★ これが決定打
            richTextBody.ForeColor = SystemColors.WindowText;

            var msg = mail.message;

            StringBuilder sb = new StringBuilder();

            // ★ MIME ヘッダをそのまま全部表示
            foreach (var h in msg.Headers)
            {
                sb.AppendLine($"{h.Field}: {h.Value}");
            }

            browserMail.Visible = false;
            richTextBody.Text = sb.ToString();
            richTextBody.BringToFront();
        }

        private void ShowNormalView(Mail mail)
        {
            browserMail.Visible = true;
            richTextBody.Clear();
            browserMail.BringToFront();

            if (mail == null || mail.message == null)
            {
                richTextBody.Clear();
                browserMail.NavigateToString("<html></html>");
                return;
            }

            var message = mail.message;

            // ================================
            // ★ HTML メールの場合（isHtml で判定）
            // ================================
            if (mail.isHtml)
            {
                string html = message.GetTextBody(MimeKit.Text.TextFormat.Html) ?? "";

                // HTML が空なら fallback
                if (string.IsNullOrEmpty(html))
                    html = message.GetTextBody(MimeKit.Text.TextFormat.Plain) ?? "";

                browserMail.Visible = true;
                richTextBody.Visible = false;
                browserMail.BringToFront();

                ShowHtmlWithInlineImages(mail.message, mail);
            }
            else
            {
                // ================================
                // ★ テキストメールの場合
                // ================================
                string text = message.GetTextBody(MimeKit.Text.TextFormat.Plain)
                            ?? message.TextBody
                            ?? "";

                // HTMLタグ混入の簡易修正
                text = FixBrokenHtml(text);

                browserMail.Visible = false;
                richTextBody.Visible = true;
                richTextBody.BringToFront();

                richTextBody.Text = text;
            }

            // ================================
            // ★ 添付ファイルメニューの構築
            // ================================
            buttonAtachMenu.DropDownItems.Clear();

            foreach (var part in FindAttachments(message.Body))
            {
                string name = GetAttachmentName(part);
                string ext = Path.GetExtension(name);

                Icon icon = null;
                try { icon = GetIconFromExtension(ext); }
                catch { icon = SystemIcons.WinLogo; }

                var item = new ToolStripMenuItem(name, icon.ToBitmap());
                item.Tag = name;
                buttonAtachMenu.DropDownItems.Add(item);
            }

            buttonAtachMenu.Visible = buttonAtachMenu.DropDownItems.Count > 0;

            UpdateUndoState();
        }

        private int GetMailIconIndex(Mail mail)
        {
            if (mail.Folder.Type == FolderType.Spam)
                return 3; // 迷惑メール

            if (mail.hasAtach)
                return 2; // 添付あり

            if (mail.notReadYet)
                return 0; // 未読

            return 1; // 既読
        }

        /// <summary>
        /// カラム幅に応じて文字列を省略する（末尾に … を付ける）
        /// </summary>
        /// <param name="text">元の文字列</param>
        /// <param name="columnWidth">カラム幅（px）</param>
        /// <param name="font">使用フォント</param>
        /// <returns>省略後の文字列</returns>
        private string FitTextToColumn(string text, int columnWidth, Font font)
        {
            if (string.IsNullOrEmpty(text))
                return "";

            int maxChars = Math.Max(5, columnWidth / charWidth); // charWidth は FormMain_Load で測定
            if (text.Length <= maxChars)
                return text;

            return text.Substring(0, maxChars) + "…";
        }

        private void listMain_ColumnWidthChanged(object sender, ColumnWidthChangedEventArgs e)
        {
            if (resizeTimer == null)
            {
                resizeTimer = new System.Windows.Forms.Timer();
                resizeTimer.Interval = 50; // ← これが重要
                resizeTimer.Tick += (s, ev) =>
                {
                    resizeTimer.Stop();
                };
            }

            resizeTimer.Stop();
            resizeTimer.Start();
        }

        private void LoadMailCache(Mail mail)
        {
            // ★ ファイルサイズをキャッシュ
            if (File.Exists(mail.mailPath))
            {
                mail.sizeBytes = new FileInfo(mail.mailPath).Length;
            }
            else
            {
                mail.sizeBytes = 0;
            }

            // ★ プレビューをキャッシュ（改行除去）
            if (!string.IsNullOrEmpty(mail.body))
            {
                mail.preview = mail.body.Replace("\r", "").Replace("\n", " ");
            }
            else
            {
                mail.preview = "";
            }
        }

        private string BuildPreview(string htmlOrText, bool isHtml)
        {
            string text = htmlOrText ?? "";

            if (isHtml)
            {
                // 1. 改行タグを先に置換
                text = Regex.Replace(text, @"<(br|BR)\s*/?>", "\n");
                text = Regex.Replace(text, @"</p>", "\n");

                // 2. タグ除去
                text = Regex.Replace(text, "<.*?>", "");

                // 3. HTML エンティティの簡易デコード
                text = WebUtility.HtmlDecode(text);
            }

            // 4. 改行・空白整形
            text = text.Replace("\r", "")
                       .Replace("\n", " ")
                       .Replace("\t", " ");

            // 5. 連続スペースを 1 個に
            text = Regex.Replace(text, @"\s{2,}", " ");

            // 6. トリム
            return text.Trim();
        }

        private void menuEditTags_Click(object sender, EventArgs e)
        {
            if (currentMail == null)
                return;

            using (var dlg = new FormTagEditor(currentMail.Labels))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    var action = new TagUndoAction
                    {
                        Mail = currentMail,
                        OldLabels = new List<string>(currentMail.Labels),
                        NewLabels = new List<string>(dlg.ResultTags)
                    };

                    tagUndoStack.Push(action);
                    tagRedoStack.Clear(); // Redo は無効化

                    // ★ タグを更新
                    currentMail.Labels = dlg.ResultTags;
                    SaveMailLabels(currentMail);

                    // ★ VirtualMode の部分再描画（これだけで十分）
                    int index = _virtualList.IndexOf(currentMail);
                    if (index >= 0)
                        listMain.RedrawItems(index, index, true);

                    UpdateUndoState();
                }
            }
        }

        private void menuUndoTags_Click(object sender, EventArgs e)
        {
            if (tagUndoStack.Count == 0)
                return;

            var action = tagUndoStack.Pop();

            // ★ Undo 実行（Mail.Labels を元に戻す）
            action.Undo();

            // ★ Redo スタックへ積む
            tagRedoStack.Push(action);

            // ★ VirtualMode の再描画
            int index = _virtualList.IndexOf(action.Mail);
            if (index >= 0)
                listMain.RedrawItems(index, index, true);

            UpdateUndoState();
        }

        private void menuRedoTags_Click(object sender, EventArgs e)
        {
            if (tagRedoStack.Count == 0)
                return;

            var action = tagRedoStack.Pop();

            // ★ Redo 実行（Mail.Labels を新しい状態に戻す）
            action.Redo();

            // ★ Undo スタックへ戻す
            tagUndoStack.Push(action);

            // ★ VirtualMode の再描画
            int index = _virtualList.IndexOf(action.Mail);
            if (index >= 0)
                listMain.RedrawItems(index, index, true);

            UpdateUndoState();
        }


        private void listMain_RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
            try
            {
                if (_virtualList == null || e.ItemIndex < 0 || e.ItemIndex >= _virtualList.Count)
                {
                    e.Item = new ListViewItem(" ");
                    return;
                }

                Mail mail = _virtualList[e.ItemIndex];
                if (mail == null)
                {
                    e.Item = new ListViewItem(" ");
                    return;
                }

                int iconIndex = GetMailIconIndex(mail);
                var item = new ListViewItem(" ", iconIndex);

                item.SubItems.Add(mail.from ?? mail.address ?? "");
                item.SubItems.Add(mail.subject ?? "");
                item.SubItems.Add(FormatReceivedDate(mail.date));
                item.SubItems.Add(FormatSize(mail.sizeBytes));
                item.SubItems.Add(mail.preview ?? "");
                item.SubItems.Add(mail.Labels != null ? string.Join(", ", mail.Labels) : "");
                item.SubItems.Add(mail.mailName ?? "");

                //
                // ★ 未読メールのスタイルを適用（VirtualMode ではここでしかできない）
                //

                item.Font = mail.notReadYet ? boldFont : listMain.Font;

                if (mail.notReadYet)
                {
                    item.BackColor = Color.FromArgb(230, 245, 255);
                }
                else
                {
                    item.BackColor = listMain.BackColor;
                }

                e.Item = item;
            }
            catch (Exception ex)
            {
                logger.Error("RetrieveVirtualItem error: " + ex);
                e.Item = new ListViewItem("ERR");
            }
        }
    }
}
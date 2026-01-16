using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Pop3;
using MailKit.Net.Smtp;
using MimeKit;
using MimeKit.Text;
using Newtonsoft.Json;
using NLog;
using NLog.Targets;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MizuMail
{
    public partial class FormMain : Form
    {
        // メールを格納するコレクションの配列
        public List<Mail>[] collectionMail = new List<Mail>[3];

        // UIDL格納用の配列
        public List<string> localUidls = new List<string>();

        // メールの種類を識別する定数
        public const int RECEIVE = 0;   // 受信メール
        public const int SEND = 1;      // 送信メール
        public const int DELETE = 2;    // ごみ箱メール

        // メールボックス情報を表示しているときのフラグ
        public bool mailBoxViewFlag = false;

        // ListViewItemSorterに指定するフィールド
        public ListViewItemComparer listViewItemSorter;

        // 現在の検索キーワードを格納するフィールド
        private string currentKeyword = "";

        // 選択された行を格納するフィールド
        private int currentRow;
        private Mail currentMail;
        private FolderManager folderManager;

        // ★ フォルダごとのキャッシュ
        private Dictionary<string, List<Mail>> mailCache = new Dictionary<string, List<Mail>>();

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

            // コレクションを作成する
            collectionMail[RECEIVE] = new List<Mail>();
            collectionMail[SEND] = new List<Mail>();
            collectionMail[DELETE] = new List<Mail>();

            // 初期化
            currentRow = -1;
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
            // ListViewItemSorterを解除する
            listMain.ListViewItemSorter = null;

            // ツリービューとリストビューの表示を更新する
            UpdateTreeView();
            UpdateListView();

            // ListViewItemSorterを指定する
            listMain.ListViewItemSorter = listViewItemSorter;
        }

        private void UpdateTreeView()
        {
            // メールの件数を設定する
            treeMain.Nodes[0].Nodes[0].Text = "受信メール(" + collectionMail[RECEIVE].Count + ")";
            treeMain.Nodes[0].Nodes[1].Text = "送信メール(" + collectionMail[SEND].Count + ")";
            treeMain.Nodes[0].Nodes[2].Text = "ごみ箱(" + collectionMail[DELETE].Count + ")";
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
                    displayList = displayList.Where(m => HasAttachment(m));

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
                    ListViewItem item = new ListViewItem(mail.address);
                    item.SubItems.Add(mail.subject);

                    string displayDate = FormatReceivedDate(mail.date);
                    item.SubItems.Add(displayDate);

                    long sizeBytes = GetMailFileSize(mail);
                    item.SubItems.Add(FormatSize(sizeBytes));

                    item.SubItems.Add(mail.mailName);

                    string preview = mail.body?.Replace("\r", "").Replace("\n", " ");
                    if (!string.IsNullOrEmpty(preview) && preview.Length > 30)
                        preview = preview.Substring(0, 30) + "…";
                    item.SubItems.Add(preview);

                    item.Tag = mail;

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
            }
            finally
            {
                listMain.EndUpdate();
                listMain.ListViewItemSorter = prevSorter;
            }
        }

        private void treeMain_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // 下のペインをクリアする
            richTextBody.Clear();
            browserMail.DocumentText = "";
            currentKeyword = "";

            // 添付ファイルメニューに登録されている要素を破棄する
            buttonAtachMenu.DropDownItems.Clear();
            buttonAtachMenu.Visible = false;

            // フォルダ切替時に stale な選択をクリアして誤参照を防ぐ
            currentRow = -1;
            currentMail = null;
            listMain.SelectedItems.Clear();

            if (e.Node.Index == 0 && e.Node.Text == "メール")
            {
                listMain.Columns[0].Text = "メールボックス名";
                listMain.Columns[1].Text = "メールアドレス";
                listMain.Columns[2].Text = "更新日時";
            }
            else if (e.Node.Index == 0)
            {
                // 受信メールが選択された場合
                listMain.Columns[0].Text = "差出人";
                listMain.Columns[1].Text = "件名";
                listMain.Columns[2].Text = "受信日時";
            }
            else if (e.Node.Index == 1)
            {
                // 送信メールが選択された場合
                listMain.Columns[0].Text = "宛先";
                listMain.Columns[1].Text = "件名";
                listMain.Columns[2].Text = "送信日時";
            }
            else if (e.Node.Index == 2)
            {
                // ごみ箱が選択された場合
                listMain.Columns[0].Text = "差出人または宛先";
                listMain.Columns[1].Text = "件名";
                listMain.Columns[2].Text = "受信日時または送信日時";
            }

            // リストビューを更新する
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
            // ステータスバーの状況を初期化する
            toolMailProgress.Minimum = 0;
            toolMailProgress.Maximum = 100;
            toolMailProgress.Value = 0;

            try
            {
                // ステータスバーに状況を表示する
                labelMessage.Text = "メール送信中...";
                statusStrip1.Refresh();

                using (var client = new SmtpClient())
                {
                    // 非同期で送信サーバに接続・認証する
                    await client.ConnectAsync(Mail.smtpServerName, Mail.smtpPortNo, MailKit.Security.SecureSocketOptions.Auto);
                    await client.AuthenticateAsync(Mail.userName, Mail.password);

                    int total = collectionMail[SEND].Count;
                    int sentCount = 0;
                    toolMailProgress.Visible = true;

                    foreach (Mail mail in collectionMail[SEND])
                    {
                        // 未送信かどうかチェックする
                        if (mail.notReadYet == true)
                        {
                            // MimeMessageを作成する
                            var message = new MimeMessage();

                            // ヘッダーの追加
                            message.Headers.Add("X-Mailer", "MizuMail version " + System.Windows.Forms.Application.ProductVersion);

                            // 送信元情報
                            message.From.Add(MailboxAddress.Parse(Mail.fromName + " <" + Mail.userAddress + ">"));
                            // 宛先
                            message.To.Add(MailboxAddress.Parse(mail.address));
                            // CC
                            if (!string.IsNullOrWhiteSpace(mail.ccaddress))
                            {
                                message.Cc.Add(MailboxAddress.Parse(mail.ccaddress));
                            }
                            // BCC
                            if (!string.IsNullOrWhiteSpace(mail.bccaddress))
                            {
                                message.Bcc.Add(MailboxAddress.Parse(mail.bccaddress));
                            }
                            // 件名
                            message.Subject = mail.subject;
                            // 本文
                            var textPart = new TextPart(TextFormat.Text)
                            {
                                Text = mail.body,
                            };

                            // 添付ファイルの確認
                            var files = mail.atach.Split(';').Select(f => f.Trim()).Where(f => File.Exists(f));

                            // 添付ファイルが複数ある場合
                            if (!string.IsNullOrWhiteSpace(mail.atach) && files.Count() > 0)
                            {
                                // マルチパートを作成して添付ファイルを添付する
                                var multipart = new Multipart("mixed");
                                multipart.Add(textPart);
                                foreach (var file in files)
                                {
                                    if (!File.Exists(file))
                                        continue;

                                    var attachment = new MimePart()
                                    {
                                        Content = new MimeContent(File.OpenRead(file)),
                                        ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                                        ContentTransferEncoding = ContentEncoding.Base64,
                                        FileName = Path.GetFileName(file)
                                    };
                                    multipart.Add(attachment);
                                }
                                message.Body = multipart;
                            }
                            else
                            {
                                message.Body = textPart;
                            }

                            // 非同期で送信する
                            await client.SendAsync(message);

                            // 送信済み件数をカウント
                            sentCount++;

                            // 進捗更新
                            int percent = (int)(sentCount * 100.0 / total);
                            toolMailProgress.Value = percent;
                            statusStrip1.Refresh();

                            // 送信日時を設定する
                            mail.date = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString();
                            mail.notReadYet = false;
                            mail.folder = "send";

                            // ★ 送信済みとしてファイルに上書き保存
                            SaveMail(mail);

                            // ステータス更新（任意）
                            labelMessage.Text = "送信: " + mail.address;
                            statusStrip1.Refresh();
                        }
                    }

                    // 切断（非同期）
                    await client.DisconnectAsync(true);
                }

                // ステータスバーに状況を表示する
                labelMessage.Text = "メール送信完了";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                // ステータスバーに状況を表示する
                labelMessage.Text = "メール送信エラー : " + ex.Message;
                statusStrip1.Refresh();
            }

            toolMailProgress.Value = 100;
            await Task.Delay(300);
            toolMailProgress.Value = 0;
            toolMailProgress.Visible = false;

            // ツリービューとリストビューの表示を更新する
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

            if (mail.notReadYet)
            {
                string oldPath = folderManager.ResolveMailPath(mail);
                string newPath = oldPath.Replace("_unread.eml", ".eml");

                File.Move(oldPath, newPath);

                mail.notReadYet = false;
                mail.mailName = Path.GetFileName(newPath);
            }

            // ★ 送信メールだけは編集画面を開く
            if (mail.folder == "send")
            {
                OpenSendMailEditor(mail);
                return;
            }

            // ★ それ以外はすべて ShowMailPreview に任せる
            ShowMailPreview(mail);
        }

        private void listMain_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            // 選択された行を取得する
            currentRow = e.ItemIndex;

            // 選択中の ListViewItem から Mail を取得して保持する（null 安全）
            if (e.Item != null && e.Item.Tag is Mail)
            {
                currentMail = (Mail)e.Item.Tag;
            }
            else
            {
                currentMail = null;
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

            // 設定を保存する
            SaveSettings();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            // ① FolderManager 初期化
            folderManager = new FolderManager();

            // ② 設定読み込み
            LoadSettings();

            // ③ mbox フォルダ構造保証
            Directory.CreateDirectory(folderManager.Inbox.FullPath);
            Directory.CreateDirectory(folderManager.Send.FullPath);
            Directory.CreateDirectory(folderManager.Trash.FullPath);

            // ④ TreeView の Tag を MailFolder に統一
            TreeNode root = treeMain.Nodes[0];
            root.Nodes[0].Tag = folderManager.Inbox;
            root.Nodes[1].Tag = folderManager.Send;
            root.Nodes[2].Tag = folderManager.Trash;

            // ⑤ inbox サブフォルダ読み込み（MailFolder 再帰）
            LoadInboxFolders();

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
            var selectedItems = listMain.SelectedItems.Cast<ListViewItem>().ToList();
            if (selectedItems.Count == 0)
                return;

            var trashMails = new List<Mail>();

            foreach (var item in selectedItems)
            {
                if (!(item.Tag is Mail mail))
                    continue;

                string folder = (mail.folder ?? "inbox").Replace("\\", "/");

                // ★ 単一選択時の確認ダイアログ（send も含めて統一）
                if (selectedItems.Count == 1 && folder != "trash")
                {
                    string msg = $"選択したメール「{mail.subject}」を削除しますか？";
                    if (MessageBox.Show(msg, "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                        continue;
                }

                // ★ ごみ箱内 → 完全削除へ回す
                if (folder == "trash")
                {
                    trashMails.Add(mail);
                    continue;
                }

                // ★ 元のパス
                string oldPath = ResolveMailPath(mail);

                // ★ ごみ箱のパス
                string trashDir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trash");
                Directory.CreateDirectory(trashDir);

                string newPath = Path.Combine(trashDir, mail.mailName);

                // ★ ファイル移動（send も inbox も統一）
                if (File.Exists(oldPath))
                {
                    if (File.Exists(newPath))
                    {
                        string unique = Guid.NewGuid().ToString() + "_" + mail.mailName;
                        newPath = Path.Combine(trashDir, unique);
                        File.Move(oldPath, newPath);
                        mail.mailName = unique;
                    }
                    else
                    {
                        File.Move(oldPath, newPath);
                    }
                }

                // ★ collectionMail の整合性
                if (folder == "inbox")
                    collectionMail[RECEIVE].Remove(mail);
                else if (folder == "send")
                    collectionMail[SEND].Remove(mail);
                else if (folder.StartsWith("inbox/"))
                {
                    // サブフォルダは collectionMail に入っていないので何もしない
                }

                // ★ Undo 用
                mail.lastFolder = folder;

                // ★ ごみ箱へ
                mail.folder = "trash";
                collectionMail[DELETE].Add(mail);
            }

            // ★ ごみ箱内の完全削除（確認ダイアログあり）
            if (trashMails.Count > 0)
            {
                var msg = $"選択したごみ箱内メール {trashMails.Count} 件を完全に削除します。";
                if (MessageBox.Show(msg, "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    foreach (var mail in trashMails)
                    {
                        string path = ResolveMailPath(mail);
                        if (File.Exists(path))
                            File.Delete(path);

                        collectionMail[DELETE].Remove(mail);
                    }
                }
            }
            UpdateView();
        }

        private void menuNotReadYet_Click(object sender, EventArgs e)
        {
            if (currentRow < 0 || currentRow >= listMain.Items.Count)
                return;

            var selItem = listMain.Items[currentRow];
            if (!(selItem.Tag is Mail mail))
                return;

            // ★ 未読にする
            if (!mail.notReadYet)
            {
                string oldPath = folderManager.ResolveMailPath(mail);
                string newPath = oldPath.Replace(".eml", "_unread.eml");

                File.Move(oldPath, newPath);

                mail.notReadYet = true;
                mail.mailName = Path.GetFileName(newPath);
            }

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
                    mail.folder = "send";
                    collectionMail[SEND].Add(mail);
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
            Mail.popPortNo = 110;
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
            // ファイルを開くかの確認をする
            DialogResult result = MessageBox.Show(e.ClickedItem.Text + "を開きますか？\nファイルによってはウイルスの可能性もあるため\n注意してファイルを開いてください。", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (result == DialogResult.Yes)
            {
                // 添付ファイルを開く
                System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\" + e.ClickedItem.Text);
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
            toolReplyButton.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            toolDeleteButton.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuNotReadYet.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuDelete.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuClearTrash.Enabled = collectionMail[DELETE].Count > 0;
            menuMailDelete.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuMailReply.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            menuFileClearTrash.Enabled = collectionMail[DELETE].Count > 0;
            menuUndoMail.Enabled = listMain.SelectedItems.Count > 0 && collectionMail[DELETE].Count > 0;
            menuSaveAs.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            menuSpeechMail.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
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
            form.textMailTo.Text = currentMail.address;
            if (currentMail.body.Trim() != string.Empty)
            {
                form.textMailBody.Text = "\r\n\r\n------------------------------\r\n" + currentMail.body.TrimEnd('\r', '\n');
            }

            // 添付ファイルを設定する
            if (currentMail.atach != "")
            {
                string[] atachFiles = currentMail.atach.Split(';');
                foreach (string atachFile in atachFiles)
                {
                    Icon atachIcon = System.Drawing.Icon.ExtractAssociatedIcon(atachFile);
                    form.buttonAttachList.DropDownItems.Add(atachFile, atachIcon.ToBitmap());
                }
                form.buttonAttachList.Visible = true;
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
                    mail.folder = "send";
                    collectionMail[SEND].Add(mail);
                    SaveMail(mail);
                }

                // ツリービューとリストビューの表示を更新する
                UpdateView();
            }
        }

        private void menuClearTrash_Click(object sender, EventArgs e)
        {
            string trashDir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trash");

            if (MessageBox.Show("ごみ箱内のメールをすべて完全に削除します。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
                return;

            // ★ 1. trash フォルダ内のファイルを削除
            if (Directory.Exists(trashDir))
            {
                foreach (var file in Directory.GetFiles(trashDir, "*.eml"))
                {
                    try { File.Delete(file); }
                    catch { /* ログに残しても良い */ }
                }
            }

            // ★ 2. collectionMail[DELETE] の中で folder == "trash" のメールを削除
            // （folder ベース化のポイント）
            collectionMail[DELETE].RemoveAll(m =>
                (m.folder ?? "").Replace("\\", "/") == "trash"
            );

            // ★ 3. 表示更新
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
                sfd.FileName = currentMail.subject + ".eml";
                sfd.Filter = "EMLファイル (*.eml)|*.eml|すべてのファイル (*.*)|*.*";
                sfd.FilterIndex = 0;
                sfd.RestoreDirectory = true;
                sfd.OverwritePrompt = true;

                if (sfd.ShowDialog() != DialogResult.OK)
                    return;

                string folder = (currentMail.folder ?? "inbox").Replace("\\", "/");

                // ★ 送信メール（send / send/...）は .mail なので再構築
                if (folder == "send" || folder.StartsWith("send/"))
                {
                    var message = new MimeMessage();

                    // From / To / Cc / Bcc
                    if (!string.IsNullOrWhiteSpace(Mail.userAddress))
                        message.From.Add(MailboxAddress.Parse(Mail.userAddress));

                    if (!string.IsNullOrWhiteSpace(currentMail.address))
                        message.To.Add(MailboxAddress.Parse(currentMail.address));

                    if (!string.IsNullOrWhiteSpace(currentMail.ccaddress))
                        message.Cc.AddRange(
                            currentMail.ccaddress.Split(';')
                                .Select(x => x.Trim())
                                .Where(x => x != "")
                                .Select(MailboxAddress.Parse)
                        );

                    if (!string.IsNullOrWhiteSpace(currentMail.bccaddress))
                        message.Bcc.AddRange(
                            currentMail.bccaddress.Split(';')
                                .Select(x => x.Trim())
                                .Where(x => x != "")
                                .Select(MailboxAddress.Parse)
                        );

                    message.Subject = currentMail.subject ?? "";

                    var textPart = new TextPart(TextFormat.Text)
                    {
                        Text = currentMail.body ?? ""
                    };

                    var files = (currentMail.atach ?? "")
                        .Split(';')
                        .Select(f => f.Trim())
                        .Where(f => File.Exists(f));

                    if (!string.IsNullOrWhiteSpace(currentMail.atach) && files.Any())
                    {
                        var multipart = new Multipart("mixed");
                        multipart.Add(textPart);

                        foreach (var file in files)
                        {
                            var attachment = new MimePart()
                            {
                                Content = new MimeContent(File.OpenRead(file)),
                                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                                ContentTransferEncoding = ContentEncoding.Base64,
                                FileName = Path.GetFileName(file)
                            };
                            multipart.Add(attachment);
                        }
                        message.Body = multipart;
                    }
                    else
                    {
                        message.Body = textPart;
                    }

                    using (var stream = File.Create(sfd.FileName))
                    {
                        message.WriteTo(stream);
                    }
                }
                else
                {
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
            var mailCount = 0;
            try
            {
                // ステータスバーに状況を表示する
                labelMessage.Text = "メール受信中...";
                statusStrip1.Refresh();

                using (var client = new Pop3Client())
                {
                    // 非同期で受信サーバに接続・認証する
                    await client.ConnectAsync(Mail.popServerName, Mail.popPortNo, Mail.useSsl);
                    await client.AuthenticateAsync(Mail.userName, Mail.password);

                    // 取得したメールの件数を表示する
                    labelMessage.Text = client.Count + "件のメッセージがあります";
                    statusStrip1.Refresh();

                    toolMailProgress.Visible = true;
                    int total = client.Count;

                    for (int i = 0; i < client.Count; i++)
                    {
                        // 進捗バー更新
                        int percent = (int)((i + 1) * 100.0 / total);
                        toolMailProgress.Value = percent;
                        statusStrip1.Refresh();

                        // 修正: 同一 Pop3Client を別スレッドで操作しない（Task.Run を除去）
                        string uidl = client.GetMessageUid(i);

                        // 受信済みのメールか確認して未受信なら受信する
                        if (!localUidls.Contains(uidl))
                        {
                            labelMessage.Text = (i + 1) + "件目のメール受信中";
                            statusStrip1.Refresh();

                            // 非同期でメッセージを取得
                            var message = await client.GetMessageAsync(i);
                            string mailName = message.MessageId + ".eml";

                            // ファイルへの保存（重たい処理は別スレッドへ）
                            var inboxPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", mailName);
                            await Task.Run(() => message.WriteTo(inboxPath));

                            Mail mail = new Mail(message.From.ToString(), message.Cc.ToString(), message.Bcc.ToString(), message.Subject, message.TextBody, null, message.Date.ToString(), mailName, uidl, true);
                            mail.folder = "inbox";
                            localUidls.Add(uidl);
                            collectionMail[RECEIVE].Add(mail);
                            mailCount++;

                            // メール削除設定が有効な場合
                            if (Mail.deleteMail)
                            {
                                // 受信したメールをサーバから削除する
                                await client.DeleteMessageAsync(i);
                            }
                        }
                    }

                    // 切断（非同期）
                    await client.DisconnectAsync(true);
                }

                // ステータスバーに状況を表示する
                labelMessage.Text = "メール受信完了";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                // ステータスバーに状況を表示する
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
                    // 非同期で受信サーバに接続・認証する
                    await client.ConnectAsync(Mail.popServerName, Mail.popPortNo, Mail.useSsl);
                    await client.AuthenticateAsync(Mail.userName, Mail.password);

                    var inbox = client.Inbox;
                    await inbox.OpenAsync(FolderAccess.ReadOnly);

                    labelMessage.Text = $"{inbox.Count}件のメッセージがあります";
                    statusStrip1.Refresh();

                    // 0～最後までのメッセージについて UID を取得
                    var summaries = await inbox.FetchAsync(0, -1, MessageSummaryItems.UniqueId);

                    toolMailProgress.Visible = true;
                    int total = summaries.Count;

                    for (int i = 0; i < summaries.Count; i++)
                    {
                        // 進捗バー更新
                        int percent = (int)((i + 1) * 100.0 / total);
                        toolMailProgress.Value = percent;
                        statusStrip1.Refresh();

                        var summary = summaries[i];
                        var uniqueId = summary.UniqueId;
                        var uid = uniqueId.Id.ToString();   // long → string

                        // 未受信メールのみ処理
                        if (!localUidls.Contains(uid))
                        {
                            labelMessage.Text = $"{i + 1}件目のメール受信中";
                            statusStrip1.Refresh();

                            // UID 指定でメッセージ取得
                            var message = await inbox.GetMessageAsync(uniqueId);
                            string mailName = message.MessageId + ".eml";
                            var inboxPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", mailName);

                            await Task.Run(() =>
                            {
                                using (var stream = File.Create(inboxPath))
                                {
                                    message.WriteTo(stream);
                                }
                            });

                            Mail mail = new Mail(message.From.ToString(), message.Cc.ToString(), message.Bcc.ToString(), message.Subject, message.TextBody, null, message.Date.ToString(), mailName, uid, true);
                            mail.folder = "inbox";

                            localUidls.Add(uid);
                            collectionMail[RECEIVE].Add(mail);
                            mailCount++;

                            // メール削除設定が有効な場合
                            if (Mail.deleteMail)
                            {
                                // ★ 受信したメールに削除フラグを付ける
                                await inbox.AddFlagsAsync(uniqueId, MessageFlags.Deleted, true);
                            }
                        }
                    }

                    // ★ 削除フラグが付いたメールをサーバから完全削除
                    if (Mail.deleteMail)
                    {
                        await inbox.ExpungeAsync();
                    }

                    await client.DisconnectAsync(true);
                }

                labelMessage.Text = "メール受信完了";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                labelMessage.Text = "メール受信エラー : " + ex.Message;
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

            UpdateView();
        }

        private void menuUndoMail_Click(object sender, EventArgs e)
        {
            UndoMail();
        }

        private void UndoMail()
        {
            if (collectionMail[DELETE].Count == 0)
                return;

            // ★ 最後に削除したメールを取り出す
            Mail mail = collectionMail[DELETE].Last();
            collectionMail[DELETE].Remove(mail);

            string lastFolder = mail.lastFolder;   // inbox / inbox/xxx / send
            string currentPath = ResolveMailPath(mail);

            // ★ 元のフォルダの実パス
            string restoreDir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", lastFolder.Replace("/", "\\"));

            Directory.CreateDirectory(restoreDir);

            string restorePath = Path.Combine(restoreDir, mail.mailName);

            // ★ ファイル移動（send も inbox も統一）
            if (File.Exists(currentPath))
            {
                if (File.Exists(restorePath))
                {
                    // 名前衝突回避
                    string unique = Guid.NewGuid().ToString() + "_" + mail.mailName;
                    restorePath = Path.Combine(restoreDir, unique);
                    mail.mailName = unique;
                }

                File.Move(currentPath, restorePath);
            }

            // ★ collectionMail の整合性
            if (lastFolder == "inbox")
                collectionMail[RECEIVE].Add(mail);
            else if (lastFolder == "send")
                collectionMail[SEND].Add(mail);
            else if (lastFolder.StartsWith("inbox/"))
            {
                // サブフォルダは collectionMail に入れない
            }

            // ★ folder を元に戻す
            mail.folder = lastFolder;

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
                // 空なら通常表示に戻す
                UpdateListView();
                return;
            }

            // 現在選択中のフォルダ（TreeView の Tag）を取得
            string folder = treeMain.SelectedNode?.Tag as string ?? "inbox";

            List<Mail> sourceList = new List<Mail>();

            // ★ inbox / send / trash は collectionMail から
            if (folder == "inbox")
            {
                sourceList = collectionMail[RECEIVE].ToList();
            }
            else if (folder == "send")
            {
                sourceList = collectionMail[SEND].ToList();
            }
            else if (folder == "trash")
            {
                sourceList = collectionMail[DELETE].ToList();
            }
            else
            {
                // ★ サブフォルダは listMain に表示されている Mail をそのまま使う
                sourceList = listMain.Items
                    .Cast<ListViewItem>()
                    .Select(it => it.Tag as Mail)
                    .Where(m => m != null)
                    .ToList();
            }

            // ★ キーワードでフィルタ（address / subject / body）
            var filtered = sourceList
                .Where(m =>
                    (!string.IsNullOrEmpty(m.address) && m.address.Contains(keyword)) ||
                    (!string.IsNullOrEmpty(m.subject) && m.subject.Contains(keyword)) ||
                    (!string.IsNullOrEmpty(m.body) && m.body.Contains(keyword))
                )
                .ToList();

            // ★ 検索結果の表示（元の構造に合わせて）
            ShowSearchResult(filtered);
        }

        private void ShowSearchResult(List<Mail> list)
        {
            listMain.BeginUpdate();
            listMain.Items.Clear();

            var baseFont = listMain.Font;

            foreach (var mail in list)
            {
                // 0: 差出人 or 宛先
                ListViewItem item = new ListViewItem(mail.address);

                // 1: 件名
                item.SubItems.Add(mail.subject);

                // 2: 日付
                item.SubItems.Add(FormatReceivedDate(mail.date));

                // 3: サイズ（フォルダ対応済み）
                long sizeBytes = GetMailFileSize(mail);
                item.SubItems.Add(FormatSize(sizeBytes));

                // 4: ファイル名
                item.SubItems.Add(mail.mailName);

                // ★ 添付アイコン（HasAttach はフォルダ対応済み）
                if (HasAttachment(mail))
                {
                    item.ImageKey = "attach";   // ImageList に登録したキー
                }

                // Mail を保持
                item.Tag = mail;

                // 未読は太字
                item.Font = mail.notReadYet
                    ? new Font(baseFont, FontStyle.Bold)
                    : new Font(baseFont, FontStyle.Regular);

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

        private void ShowMailPreview(Mail mail)
        {
            if (mail == null)
                return;

            // 添付ファイル展開フォルダの削除
            string tmpPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "tmp");
            if (Directory.Exists(tmpPath))
                Directory.Delete(tmpPath, true);

            // 添付メニュー初期化
            buttonAtachMenu.DropDownItems.Clear();
            buttonAtachMenu.Visible = false;

            buttonAtachMenu.DropDownItemClicked -= buttonAtachMenu_DropDownItemClicked;
            buttonAtachMenu.DropDownItemClicked += buttonAtachMenu_DropDownItemClicked;

            // ★ まず mailPath を一括で決める
            string mailPath = ResolveMailPath(mail);

            // 未読解除（送信メール以外）
            if (listMain.Columns[0].Text != "宛先")
                mail.notReadYet = false;

            // ★ .eml（受信・送信どちらも）
            if (!string.IsNullOrEmpty(mailPath) && File.Exists(mailPath) && mailPath.EndsWith(".eml"))
            {
                MimeMessage message = MimeMessage.Load(mailPath);

                // ★ 本文抽出（multipart 対応）
                string body = "";

                if (message.Body is TextPart tp)
                {
                    body = tp.Text;
                }
                else if (message.Body is Multipart mp)
                {
                    var textPart = mp.OfType<TextPart>().FirstOrDefault();
                    if (textPart != null)
                        body = textPart.Text;
                }

                // HTMLメール
                if (!string.IsNullOrEmpty(message.HtmlBody))
                {
                    browserMail.Visible = true;
                    richTextBody.Visible = false;
                    browserMail.DocumentText = message.HtmlBody;
                }
                else
                {
                    browserMail.Visible = false;
                    richTextBody.Visible = true;
                    richTextBody.Text = body;
                    ColorizeQuoteLines();
                }

                // ★ 添付ファイル処理（multipart 対応）
                var attachments = message.Attachments.ToList();
                if (attachments.Any())
                {
                    buttonAtachMenu.Visible = true;
                    Directory.CreateDirectory(tmpPath);

                    foreach (var attachment in attachments)
                    {
                        if (attachment is MimePart part)
                        {
                            string fileName = part.FileName;
                            string savePath = Path.Combine(tmpPath, fileName);

                            using (var stream = File.Create(savePath))
                                part.Content.DecodeTo(stream);

                            var menuItem = new ToolStripMenuItem(fileName);
                            //menuItem.Click += (s, e) => System.Diagnostics.Process.Start(savePath);
                            buttonAtachMenu.DropDownItems.Add(menuItem);
                        }
                    }
                }

                return;
            }

            // ★ .mail（旧仕様）→ もう不要だが互換のため残す
            if (!string.IsNullOrEmpty(mailPath) && File.Exists(mailPath) && mailPath.EndsWith(".mail"))
            {
                Mail sendMail = LoadSendMail(mailPath);
                if (sendMail == null)
                {
                    MessageBox.Show("送信メールの読み込みに失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                browserMail.Visible = false;
                richTextBody.Visible = true;
                richTextBody.Text = sendMail.body;
                ColorizeQuoteLines();

                if (!string.IsNullOrEmpty(sendMail.atach))
                {
                    string[] atachFiles = sendMail.atach.Split(';');
                    foreach (string atach in atachFiles)
                    {
                        if (File.Exists(atach))
                        {
                            Icon appIcon = Icon.ExtractAssociatedIcon(atach);
                            buttonAtachMenu.DropDownItems.Add(atach, appIcon.ToBitmap());
                        }
                    }
                    buttonAtachMenu.Visible = true;
                }

                return;
            }

            // ★ ファイルが見つからない場合
            browserMail.Visible = false;
            richTextBody.Visible = true;
            richTextBody.Text = "(メールファイルが見つかりません)";
        }

        public string ResolveMailPath(Mail mail)
        {
            if (mail == null || mail.Folder == null)
                return null;

            return Path.Combine(mail.Folder.FullPath, mail.mailName);
        }

        private void listMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listMain.Columns[0].Text == "メールボックス名")
            {
                labelMessage.Text = "現在の状況";
                return;
            }

            int count = listMain.SelectedItems.Count;
            labelMessage.Text = $"{count} 件選択中";

            if (count == 0)
                return;

            if (listMain.SelectedItems.Count == 0)
            {
                menuUndoMail.Enabled = false;
                return;
            }

            Mail mail = listMain.SelectedItems[0].Tag as Mail;
            if (mail == null)
            {
                menuUndoMail.Enabled = false;
                return;
            }

            // ★ lastFolder があるときだけ Undo 可能
            menuUndoMail.Enabled = !string.IsNullOrEmpty(mail.lastFolder);

            ShowMailPreview(mail);
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

            // ★ 保存先フォルダが未設定なら inbox に入れる
            if (mail.Folder == null)
                mail.Folder = folderManager.Inbox;

            // ★ 旧方式互換（folder = "inbox" など）
            mail.folder = mail.Folder.Name;

            // ★ 保存前にファイル名を決める
            if (mail.notReadYet)
                mail.mailName = mail.mailName.Replace(".eml", "_unread.eml");
            else
                mail.mailName = mail.mailName.Replace("_unread.eml", ".eml");

            // ★ 保存先パスを決定
            string savePath = Path.Combine(mail.Folder.FullPath, mail.mailName);

            try
            {
                // MimeMessage を構築
                var message = new MimeKit.MimeMessage();

                // 宛先
                if (!string.IsNullOrEmpty(mail.address))
                {
                    foreach (var addr in mail.address.Split(';'))
                        if (!string.IsNullOrWhiteSpace(addr))
                            message.To.Add(MimeKit.MailboxAddress.Parse(addr));
                }

                // CC
                if (!string.IsNullOrEmpty(mail.ccaddress))
                {
                    foreach (var addr in mail.ccaddress.Split(';'))
                        if (!string.IsNullOrWhiteSpace(addr))
                            message.Cc.Add(MimeKit.MailboxAddress.Parse(addr));
                }

                // BCC
                if (!string.IsNullOrEmpty(mail.bccaddress))
                {
                    foreach (var addr in mail.bccaddress.Split(';'))
                        if (!string.IsNullOrWhiteSpace(addr))
                            message.Bcc.Add(MimeKit.MailboxAddress.Parse(addr));
                }

                // 件名
                message.Subject = mail.subject ?? "";

                // 本文
                var builder = new MimeKit.BodyBuilder();
                builder.TextBody = mail.body ?? "";

                // 添付ファイル
                if (!string.IsNullOrEmpty(mail.atach))
                {
                    foreach (var file in mail.atach.Split(';'))
                    {
                        if (File.Exists(file))
                            builder.Attachments.Add(file);
                    }
                }

                message.Body = builder.ToMessageBody();

                // 日付
                if (DateTime.TryParse(mail.date, out DateTime dt))
                {
                    message.Date = new DateTimeOffset(dt);
                }
                else
                {
                    message.Date = DateTimeOffset.Now;
                }

                // ★ 保存
                message.WriteTo(savePath);

                logger.Info($"SaveMail: {savePath}");
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
            if (parent.Type == FolderType.Send || parent.Type == FolderType.Trash)
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

            // ★ Tag は MailFolder
            MailFolder folder = node.Tag as MailFolder;
            if (folder == null)
            {
                e.CancelEdit = true;
                return;
            }

            // inbox / send / trash はリネーム禁止
            if (folder.Type == FolderType.Inbox ||
                folder.Type == FolderType.Send ||
                folder.Type == FolderType.Trash)
            {
                MessageBox.Show("このフォルダ名は変更できません。");
                e.CancelEdit = true;
                return;
            }

            string parentDir = Directory.GetParent(folder.FullPath).FullName;
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

            // ★ MailFolder を更新
            var newFolder = new MailFolder(newName, newDir, FolderType.InboxSub);
            node.Tag = newFolder;

            // ★ 子フォルダの MailFolder も更新
            UpdateChildFolderTagsRecursive(node, folder.FullPath, newDir);

            // ★ メールの保存先も更新
            UpdateMailFolderPaths(folder.FullPath, newDir);

            UpdateListView();
        }

        private void UpdateChildFolderTagsRecursive(TreeNode node, string oldBase, string newBase)
        {
            foreach (TreeNode child in node.Nodes)
            {
                MailFolder oldFolder = child.Tag as MailFolder;
                if (oldFolder == null)
                    continue;

                string relative = oldFolder.FullPath.Substring(oldBase.Length).TrimStart('\\');
                string newPath = Path.Combine(newBase, relative);

                var newFolder = new MailFolder(oldFolder.Name, newPath, FolderType.InboxSub);
                child.Tag = newFolder;

                UpdateChildFolderTagsRecursive(child, oldBase, newBase);
            }
        }

        private void UpdateChildFolderTags(TreeNode parent, string oldBase, string newBase)
        {
            foreach (TreeNode child in parent.Nodes)
            {
                string oldTag = child.Tag.ToString();
                string newTag = oldTag.Replace(oldBase, newBase);
                child.Tag = newTag;

                UpdateChildFolderTags(child, oldBase, newBase);
            }
        }

        private void UpdateMailFolderPaths(string oldDir, string newDir)
        {
            // oldDir, newDir は絶対パスで渡される前提
            // 例: C:\...\mbox\inbox\仕事 → C:\...\mbox\inbox\開発
            foreach (var list in collectionMail)
            {
                foreach (var mail in list)
                {
                    // ★ MailFolder が設定されていないメールは無視
                    if (mail.Folder == null)
                        continue;

                    // ★ このメールが移動対象フォルダに属しているか？
                    if (!mail.Folder.FullPath.StartsWith(oldDir, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // ★ 新しいフォルダパスを計算
                    string relative = mail.Folder.FullPath.Substring(oldDir.Length).TrimStart('\\');
                    string newFolderPath = Path.Combine(newDir, relative);

                    // ★ MailFolder を更新
                    mail.Folder = new MailFolder(
                        name: Path.GetFileName(newFolderPath),
                        fullPath: newFolderPath,
                        type: FolderType.InboxSub
                    );

                    // ★ ファイル名も更新（未読フラグ含む）
                    string oldMailPath = Path.Combine(oldDir, relative, mail.mailName);
                    string newMailPath = Path.Combine(newFolderPath, mail.mailName);

                    try
                    {
                        Directory.CreateDirectory(newFolderPath);

                        if (File.Exists(oldMailPath))
                        {
                            File.Move(oldMailPath, newMailPath);
                        }

                        mail.mailName = Path.GetFileName(newMailPath);
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"UpdateMailFolderPaths: {ex.Message}");
                    }
                }
            }
        }

        private void MoveEmlToFolder(Mail mail, string newFolder)
        {
            // ★ oldFolder を正規化
            string oldFolder = (mail.folder ?? "inbox").Replace("\\", "/");

            // ★ Undo 用
            mail.lastFolder = oldFolder;

            // ★ oldPath は ResolveMailPath を使う
            string oldPath = ResolveMailPath(mail);

            // ★ newFolder は "inbox/仕事" のような実パスが渡される前提
            string newDir = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", newFolder.Replace("/", "\\"));
            Directory.CreateDirectory(newDir);

            string newPath = Path.Combine(newDir, mail.mailName);

            // ★ ファイル移動
            if (File.Exists(oldPath))
                File.Move(oldPath, newPath);

            // ★ collectionMail の整合性（元の場所から削除）
            if (oldFolder == "inbox")
                collectionMail[RECEIVE].Remove(mail);
            else if (oldFolder == "send")
                collectionMail[SEND].Remove(mail);
            else if (oldFolder == "trash")
                collectionMail[DELETE].Remove(mail);

            // ★ folder 更新（最重要）
            mail.folder = newFolder;

            // ★ サブフォルダは collectionMail に入れない
            // → LoadEmlFolder がファイルから読み込むため

            UpdateListView();
            UpdateTreeView();
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

            Mail mail = (Mail)e.Data.GetData(typeof(Mail));

            // ★ ドロップ先ノード
            TreeNode targetNode = treeMain.GetNodeAt(treeMain.PointToClient(new Point(e.X, e.Y)));
            if (targetNode == null)
                return;

            // ★ ドロップ先フォルダ（MailFolder）
            MailFolder targetFolder = targetNode.Tag as MailFolder;
            if (targetFolder == null)
                return;

            // inbox / send / trash 以外は自由に移動可能
            // ただし send → inbox-sub などは許可する
            // 特別な制限は不要

            // ★ 移動元
            string oldPath = folderManager.ResolveMailPath(mail);

            // ★ 移動先
            string newPath = Path.Combine(targetFolder.FullPath, mail.mailName);

            logger.Info($"oldPath = {oldPath}");
            logger.Info($"newPath = {newPath}");

            try
            {
                // ★ ファイル移動
                File.Move(oldPath, newPath);

                // ★ Mail オブジェクト更新
                mail.Folder = targetFolder;
                mail.folder = targetFolder.Name; // 旧互換
            }
            catch (Exception ex)
            {
                MessageBox.Show("メールを移動できませんでした。\n" + ex.Message);
                return;
            }

            UpdateListView();
        }

        public IEnumerable<Mail> LoadEmlFolder(MailFolder folder)
        {
            var list = new List<Mail>();

            foreach (var file in Directory.GetFiles(folder.FullPath, "*.eml"))
            {
                try
                {
                    var mail = ParseMailWithMimeKit(file);
                    mail.Folder = folder;
                    list.Add(mail);
                }
                catch { }
            }

            return list;
        }

        private Mail LoadMailHeaderOnly(string path)
        {
            var mail = new Mail("", "", "", "", "", "", "", "", "", false);
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
            // ★ Mail を受け取れるようにする
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

            // ★ ドロップ先フォルダ
            MailFolder folder = node.Tag as MailFolder;
            if (folder == null)
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            // ★ inbox / send / trash 以外はすべて許可
            if (folder.Type == FolderType.Inbox ||
                folder.Type == FolderType.Send ||
                folder.Type == FolderType.Trash)
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            // ★ サブフォルダ無限階層に対応
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

            // 添付ファイル
            if (!string.IsNullOrEmpty(mail.atach))
            {
                string[] atachFiles = mail.atach.Split(';');
                foreach (string atachFile in atachFiles)
                {
                    if (File.Exists(atachFile))
                    {
                        Icon icon = Icon.ExtractAssociatedIcon(atachFile);
                        form.buttonAttachList.DropDownItems.Add(atachFile, icon.ToBitmap());
                    }
                }
                form.buttonAttachList.Visible = true;
            }

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
                    mail.Folder = folderManager.Send;
                    mail.folder = "send"; // ← 旧方式も残す

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

            // ★ Tag は MailFolder
            MailFolder folder = node.Tag as MailFolder;
            if (folder == null)
                return;

            // inbox / send / trash は削除禁止
            if (folder.Type == FolderType.Inbox ||
                folder.Type == FolderType.Send ||
                folder.Type == FolderType.Trash)
            {
                MessageBox.Show("このフォルダは削除できません。");
                return;
            }

            var result = MessageBox.Show(
                $"フォルダ「{node.Text}」を削除しますか？\n中のメールはごみ箱へ移動されます。",
                "フォルダ削除",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result != DialogResult.Yes)
                return;

            // ★ 中のメールをゴミ箱へ移動（再帰対応）
            MoveFolderMailsToTrashRecursive(folder);

            // ★ 実フォルダ削除
            try
            {
                Directory.Delete(folder.FullPath, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("フォルダを削除できませんでした。\n" + ex.Message);
                return;
            }

            // ★ FolderManager の第一階層リストから削除（InboxSubFolders）
            folderManager.InboxSubFolders.Remove(folder);

            // ★ TreeView から削除
            node.Remove();

            UpdateListView();
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

        private void LoadInboxFolders()
        {
            // inbox ノードを取得
            TreeNode root = treeMain.Nodes[0];
            TreeNode inboxNode = root.Nodes[0];

            inboxNode.Nodes.Clear();
            folderManager.InboxSubFolders.Clear();

            string inboxPath = folderManager.Inbox.FullPath;

            if (!Directory.Exists(inboxPath))
                return;

            // ★ 第一階層のサブフォルダを読み込む
            foreach (var dir in Directory.GetDirectories(inboxPath))
            {
                string name = Path.GetFileName(dir);
                var folder = new MailFolder(name, dir, FolderType.InboxSub);

                folderManager.InboxSubFolders.Add(folder);

                TreeNode node = new TreeNode(name);
                node.Tag = folder;
                inboxNode.Nodes.Add(node);

                // ★ 第二階層以降を再帰的に読み込む
                LoadSubFoldersRecursive(node, folder);
            }
        }

        private void LoadSubFoldersRecursive(TreeNode parentNode, MailFolder parentFolder)
        {
            foreach (var dir in Directory.GetDirectories(parentFolder.FullPath))
            {
                string name = Path.GetFileName(dir);
                var subFolder = new MailFolder(name, dir, FolderType.InboxSub);

                TreeNode node = new TreeNode(name);
                node.Tag = subFolder;
                parentNode.Nodes.Add(node);

                // ★ 再帰
                LoadSubFoldersRecursive(node, subFolder);
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

            var mail = new Mail("", "", "", "", "", "", "", "", "", false);
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

            var mail = new Mail("", "", "", "", "", "", "", "", "", false);
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

    }
}
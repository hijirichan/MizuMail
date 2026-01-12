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
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MizuMail
{
    public partial class FormMain : Form
    {
        // メールを格納するコレクションの配列
        public ArrayList[] collectionMail = new ArrayList[3];

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
            collectionMail[RECEIVE] = new ArrayList();
            collectionMail[SEND] = new ArrayList();
            collectionMail[DELETE] = new ArrayList();

            // 初期化
            currentRow = -1;
            currentMail = null;

            System.Windows.Forms.Application.Idle += Application_Idle;
            listMain.ColumnClick += listMain_ColumnClick;
            toolFilterCombo.SelectedIndex = 0;

            listMain.SmallImageList = new ImageList { ImageSize = new Size(1, 20) };
            listViewItemSorter = ListViewItemComparer.Default;
            listMain.ListViewItemSorter = listViewItemSorter;
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
            // 可能性のある保存場所をチェックしてファイルサイズを取得
            if (!string.IsNullOrEmpty(mail.mailName))
            {
                string[] paths = new[]
                {
                    Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", mail.mailName),
                    Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName),
                    Path.Combine(System.Windows.Forms.Application.StartupPath, mail.mailName)
                };
                foreach (var p in paths)
                {
                    if (File.Exists(p))
                    {
                        try
                        {
                            return new FileInfo(p).Length;
                        }
                        catch
                        {
                            // 無視して次へ
                        }
                    }
                }
            }

            // ファイルが見つからない場合は本文から概算サイズを返す
            if (!string.IsNullOrEmpty(mail.body))
            {
                try
                {
                    return Encoding.Default.GetByteCount(mail.body);
                }
                catch
                {
                    return mail.body.Length;
                }
            }

            return 0;
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
            ArrayList list = null;

            // 現在のソータを退避して、一時的に無効化する（追加中の自動ソートによるずれを防ぐ）
            var prevSorter = listMain.ListViewItemSorter;
            listMain.ListViewItemSorter = null;

            // 描画ブロック
            listMain.BeginUpdate();
            try
            {
                // リストビューの内容をクリアする
                listMain.Items.Clear();

                if (listMain.Columns[0].Text == "メールボックス名")
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
                    return;
                }
                if (listMain.Columns[0].Text == "差出人")
                {
                    // 受信メールの場合
                    list = collectionMail[RECEIVE];
                }
                else if (listMain.Columns[0].Text == "宛先")
                {
                    // 送信メールの場合
                    list = collectionMail[SEND];
                }
                else if (listMain.Columns[0].Text == "差出人または宛先")
                {
                    // ごみ箱の場合
                    list = collectionMail[DELETE];
                }

                // ★ 検索語があるならフィルタリング
                IEnumerable<Mail> displayList = list.Cast<Mail>();

                if (!string.IsNullOrEmpty(currentKeyword))
                {
                    displayList = displayList.Where(m =>
                        (m.subject?.Contains(currentKeyword) ?? false) ||
                        (m.body?.Contains(currentKeyword) ?? false) ||
                        (m.address?.Contains(currentKeyword) ?? false)
                    );
                }

                // ★ フィルタリング処理を追加
                string filter = toolFilterCombo.SelectedItem?.ToString();

                if (filter == "未読")
                {
                    displayList = displayList.Where(m => m.notReadYet);
                }
                else if (filter == "添付あり")
                {
                    if (list == collectionMail[SEND])
                    {
                        // ★ 送信メール（送信フォルダ）
                        displayList = displayList.Where(m =>
                            !string.IsNullOrEmpty(m.atach) &&
                            m.atach.Split(';').Any(f => !string.IsNullOrWhiteSpace(f))
                        );
                    }
                    else if (list == collectionMail[RECEIVE])
                    {
                        // ★ 受信メール（受信フォルダ）
                        displayList = displayList.Where(m => HasAttachment(m));
                    }
                    else if (list == collectionMail[DELETE])
                    {
                        // ★ ごみ箱（送信メール型と受信メール型が混在）
                        displayList = displayList.Where(m =>
                        {
                            bool isSendMail = !string.IsNullOrEmpty(m.atach); // 送信メール由来
                            if (isSendMail)
                            {
                                return m.atach.Split(';').Any(f => !string.IsNullOrWhiteSpace(f));
                            }
                            else
                            {
                                return HasAttachment(m); // 受信メール由来
                            }
                        });
                    }
                }
                else if (filter == "今日")
                {
                    displayList = displayList.Where(m =>
                    {
                        if (DateTime.TryParse(m.date, out DateTime dt))
                            return dt.Date == DateTime.Now.Date;
                        return false;
                    });
                }

                mailBoxViewFlag = false;

                // ベースフォントを ListView のフォントにする（フォームのフォントではなく）
                var baseFont = listMain.Font;

                // コレクションの内容をリストビューに表示する
                foreach (Mail mail in displayList)
                {
                    ListViewItem item = new ListViewItem(mail.address);
                    item.SubItems.Add(mail.subject);

                    // 日付表示をオフセット付きで追加
                    string displayDate = FormatReceivedDate(mail.date);
                    item.SubItems.Add(displayDate);

                    // メールサイズを計算してサブアイテムに追加
                    long sizeBytes = GetMailFileSize(mail);
                    string sizeText = FormatSize(sizeBytes);
                    item.SubItems.Add(sizeText);
                    item.SubItems.Add(mail.mailName);

                    // Mail オブジェクトを Tag に保持（表示順とデータを対応させる）
                    item.Tag = mail;

                    // アイテムを追加してから直接 item.Font を設定する（インデックスではなく参照で安全に設定）
                    listMain.Items.Add(item);

                    if (mail.notReadYet)
                    {
                        item.Font = new Font(baseFont, FontStyle.Bold);
                        item.BackColor = Color.FromArgb(0xE8, 0xF4, 0xFF);
                    }
                    else
                    {
                        item.Font = new Font(baseFont, FontStyle.Regular);
                    }

                    // デバッグ出力（必要なら Visual Studio の出力ウィンドウで確認）
                    logger.Debug($"UpdateListView: addr={mail.address}, mailName={mail.mailName}, size={sizeBytes}, notReadYet={mail.notReadYet}");
                }
            }
            finally
            {
                listMain.EndUpdate();

                // 追加処理が終わったら元のソータを復元する
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
            Mail mail = null;
            string mailPath = "";

            // 添付ファイル展開フォルダの削除
            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\") == true)
            {
                Directory.Delete(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\", true);
            }

            // 添付ファイルメニューに登録されている要素を破棄する
            buttonAtachMenu.DropDownItems.Clear();
            buttonAtachMenu.Visible = false;

            // 選択中アイテムから Mail を取得（表示順に依存しない）
            if (listMain.SelectedItems.Count > 0 && listMain.SelectedItems[0].Tag is Mail)
            {
                mail = (Mail)listMain.SelectedItems[0].Tag;
            }
            else
            {
                return;
            }

            // mailPath は現在表示しているフォルダに応じて決定
            if (listMain.Columns[0].Text == "差出人")
            {
                // 受信メールの場合
                mailPath = System.Windows.Forms.Application.StartupPath + "\\mbox\\inbox\\";
            }
            else if (listMain.Columns[0].Text == "宛先")
            {
                // 送信メールの場合
                mailPath = "";
            }
            else if (listMain.Columns[0].Text == "差出人または宛先")
            {
                // ごみ箱の場合
                mailPath = System.Windows.Forms.Application.StartupPath + "\\mbox\\trush\\";
            }

            // 受信メールとごみ箱の場合は未読を解除
            if (listMain.Columns[0].Text == "差出人" || listMain.Columns[0].Text == "差出人または宛先")
            {
                mail.notReadYet = false;
            }

            if (!string.IsNullOrEmpty(mailPath) && File.Exists(mailPath + mail.mailName) && mail.mailName.Contains(".eml"))
            {
                MimeMessage message = MimeMessage.Load(mailPath + mail.mailName);
                // HTMLメールの場合はHTML表示
                if (message.HtmlBody != null)
                {
                    browserMail.Visible = true;
                    browserMail.DocumentText = message.HtmlBody;
                }
                else
                {
                    browserMail.Visible = false;
                    richTextBody.Text = message.TextBody;
                    ColorizeQuoteLines();
                }

                // 添付ファイルが存在する場合
                if (message.Attachments.Count() > 0)
                {
                    // 添付ファイル展開フォルダの作成
                    if (Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\") == false)
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\");
                    }

                    buttonAtachMenu.Visible = true;

                    // 添付ファイルの展開
                    foreach (var atachment in message.Attachments)
                    {
                        MimePart part = (MimePart)atachment;
                        string atachFileName = part.FileName;
                        FileStream stream = new FileStream(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\" + atachFileName, FileMode.Create);
                        part.Content.DecodeTo(stream);
                        stream.Close();
                        Icon appIcon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\" + atachFileName);
                        buttonAtachMenu.DropDownItems.Add(atachFileName, appIcon.ToBitmap());
                    }
                }
            }
            else
            {
                if (mail.atach != string.Empty)
                {
                    // 送信メールで添付ファイルが存在する場合
                    string[] atachFiles = mail.atach.Split(';');
                    foreach (string atach in atachFiles)
                    {
                        Icon appIcon = System.Drawing.Icon.ExtractAssociatedIcon(atach);
                        buttonAtachMenu.DropDownItems.Add(atach, appIcon.ToBitmap());
                    }
                    buttonAtachMenu.Visible = true;
                }
                richTextBody.Text = mail.body.TrimEnd('\r', '\n');
                ColorizeQuoteLines();

                // 編集画面を表示する
                // メール作成ウィンドウを作成する
                FormMailCreate form = new FormMailCreate();

                // ウィンドウのタイトルを設定する
                form.Text = "編集";

                // 編集するメールを設定する
                form.textMailSubject.Text = currentMail.subject;
                form.textMailTo.Text = currentMail.address;
                form.textMailCc.Text = currentMail.ccaddress;
                form.textMailBcc.Text = currentMail.bccaddress;
                form.textMailBody.Text = currentMail.body.TrimEnd('\r', '\n');

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
                        Mail editmail = new Mail(to, cc, bcc, subject, body, atach, "未送信", "", "", true);
                        collectionMail[SEND].Add(editmail);
                        SaveMail(mail);
                    }
                }
            }

            // ツリービューとリストビューの表示を更新する
            if (string.IsNullOrEmpty(currentKeyword))
            {
                UpdateView();
            }
            else
            {
                // 検索中なら検索状態を維持したまま更新
                UpdateListView();
            }
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
            // ファイルストリームを作成する
            FileStream stream = new FileStream(System.Windows.Forms.Application.StartupPath + @"\MizuMail.dat", FileMode.Create);

            // ファイルストリームをストリームライタに関連付ける
            StreamWriter writer = new StreamWriter(stream, Encoding.Default);

            // メールの件数とデータを書き込む
            for (int i = 0; i < collectionMail.Length; i++)
            {
                writer.WriteLine(collectionMail[i].Count.ToString());
                foreach (Mail mail in collectionMail[i])
                {
                    writer.WriteLine(mail.address);
                    writer.WriteLine(mail.ccaddress);
                    writer.WriteLine(mail.bccaddress);
                    writer.WriteLine(mail.subject);
                    writer.WriteLine(mail.body.TrimEnd('\r', '\n'));
                    writer.WriteLine("\x03");
                    writer.WriteLine(mail.atach);
                    writer.WriteLine(mail.mailName);
                    writer.WriteLine(mail.uidl);
                    writer.WriteLine(mail.date);
                    // 保存時は True/False ではなく "1"/"0" にして互換性を高める
                    writer.WriteLine(mail.notReadYet ? "1" : "0");
                }
            }

            // ストリームライタとファイルストリームを閉じる
            writer.Close();
            stream.Close();

            // 設定を保存する
            SaveSettings();

        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\mbox\") == false)
            {
                // メールファイルフォルダを作成
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\mbox\");
            }

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\mbox\inbox\") == false)
            {
                // メールファイルフォルダを作成
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\mbox\inbox\");
            }

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\mbox\trush\") == false)
            {
                // メールファイルフォルダを作成
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\mbox\trush\");
            }

            // 設定を読み込む
            LoadSettings();

            if (File.Exists(System.Windows.Forms.Application.StartupPath + @"\MizuMail.dat") == true)
            {
                // ファイルストリームを作成する
                FileStream stream = new FileStream(System.Windows.Forms.Application.StartupPath + @"\\MizuMail.dat", FileMode.Open);

                // ファイルストリームをストリームリーダに関連付ける
                StreamReader reader = new StreamReader(stream, Encoding.Default);

                bool abortReading = false;

                // データを読み出す
                for (int i = 0; i < collectionMail.Length; i++)
                {
                    if (abortReading) break;

                    // メールの件数を読み出す
                    string countLine = reader.ReadLine();
                    if (countLine == null)
                    {
                        logger.Error("Load error: unexpected EOF when reading counts");
                        break;
                    }
                    int n = Int32.Parse(countLine);
                    // メールのデータを取得する
                    for (int j = 0; j < n; j++)
                    {
                        string address = reader.ReadLine();
                        string ccaddress = reader.ReadLine();
                        string bccaddress = reader.ReadLine();
                        string subject = reader.ReadLine();
                        if (address == null || subject == null)
                        {
                            logger.Error("Load error: unexpected EOF when reading address/subject");
                            abortReading = true;
                            break;
                        }

                        string body = "";
                        string b = reader.ReadLine();
                        // EOF とセパレータを両方考慮してループ
                        while (b != null && b != "\x03")
                        {
                            body = body + b + "\r\n";
                            b = reader.ReadLine();
                        }

                        if (b == null)
                        {
                            // ファイルが途中で終わっている -> 読み込み中止してログ出力
                            logger.Error("Load error: unexpected EOF while reading mail body");
                            abortReading = true;
                            break;
                        }

                        string atach = reader.ReadLine();
                        string mailName = reader.ReadLine();
                        string uidl = reader.ReadLine();
                        string date = reader.ReadLine();
                        string notReadLine = reader.ReadLine();
                        if (mailName == null || uidl == null || date == null)
                        {
                            logger.Error("Load error: unexpected EOF when reading mail metadata");
                            abortReading = true;
                            break;
                        }

                        // notReadLine を堅牢に処理する
                        bool notReadYet = false;
                        string rawNotRead = notReadLine;
                        if (notReadLine != null)
                        {
                            // 不可視文字や BOM、ゼロ文字を除去してトリム
                            notReadLine = notReadLine.Trim().Trim('\uFEFF', '\u200B', '\u00A0', '\0');
                        }

                        // デバッグで生データと文字コードを出す（問題解析用）
                        if (rawNotRead != null)
                        {
                            var codes = string.Join(",", rawNotRead.Select(c => ((int)c).ToString()));
                            logger.Debug($"Raw notReadLine: '{rawNotRead}' codes: {codes}");
                        }
                        else
                        {
                            logger.Debug("Raw notReadLine: null");
                        }

                        if (!string.IsNullOrEmpty(notReadLine))
                        {
                            // "1"/"0" を優先的に扱い、続いて bool.TryParse を試す
                            var t = notReadLine;
                            if (t == "1")
                            {
                                notReadYet = true;
                            }
                            else if (t == "0")
                            {
                                notReadYet = false;
                            }
                            else
                            {
                                // 空白やケース違いに強いパース
                                if (!bool.TryParse(t, out notReadYet))
                                {
                                    logger.Warn($"Load warning: could not parse notReadYet value '{rawNotRead}'");
                                    notReadYet = false;
                                }
                            }
                        }

                        logger.Debug($"Load mail: uidl={uidl} notReadYet={notReadYet} (raw='{rawNotRead}')");

                        // UIDL は受信コレクションだけに登録する
                        if (i == RECEIVE && !string.IsNullOrEmpty(uidl))
                        {
                            localUidls.Add(uidl);
                        }

                        Mail mail = new Mail(address, ccaddress, bccaddress, subject, body, atach, date, mailName, uidl, notReadYet);
                        collectionMail[i].Add(mail);
                    }
                }

                // ストリームリーダとファイルストリームを閉じる
                reader.Close();
                stream.Close();

                listViewItemSorter = ListViewItemComparer.Default;
                listMain.ListViewItemSorter = listViewItemSorter;
                richTextBody.DetectUrls = true;

                // ツリービューとリストビューの表示を更新する
                UpdateView();
            }

            // 定期受信タイマーを設定する
            SetTimer(Mail.checkMail, Mail.checkInterval);

            // ツリービューを展開する
            treeMain.ExpandAll();
        }

        private void menuDelete_Click(object sender, EventArgs e)
        {
            // 選択アイテムをスナップショット（変更中のコレクション破壊を防ぐ）
            var selectedItems = listMain.SelectedItems.Cast<ListViewItem>().ToList();
            if (selectedItems.Count == 0)
                return;

            // ごみ箱内の削除対象を集める（確認ダイアログ用）
            var trashMails = new List<Mail>();

            // まずごみ箱以外の移動処理（受信→ごみ箱、送信→ごみ箱）
            foreach (var item in selectedItems)
            {
                if (!(item.Tag is Mail))
                    continue;

                var mail = (Mail)item.Tag;

                // ごみ箱内の項目は後でまとめて処理
                if (selectedItems.Count == 1 && listMain.Columns[0].Text != "差出人または宛先")
                {
                    // 単一選択の場合は個別確認ダイアログを表示
                    if (MessageBox.Show("選択したメール「" + mail.subject + "」を削除しますか？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                        continue;
                }

                if (listMain.Columns[0].Text == "差出人")
                {
                    // 受信メールの場合：ファイルを移動してごみ箱コレクションへ
                    string src = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", mail.mailName);
                    string dst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName);
                    try
                    {
                        if (File.Exists(src))
                        {
                            // 既に同名ファイルがごみ箱にあれば上書きを避けるためユニーク化
                            if (File.Exists(dst))
                            {
                                string uniqueDst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", Guid.NewGuid().ToString() + "_" + mail.mailName);
                                File.Move(src, uniqueDst);
                                mail.mailName = Path.GetFileName(uniqueDst);
                            }
                            else
                            {
                                File.Move(src, dst);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"Move error: {ex.Message}");
                    }

                    collectionMail[DELETE].Add(mail);
                    collectionMail[RECEIVE].Remove(mail);
                }
                else if (listMain.Columns[0].Text == "宛先")
                {
                    // 送信メールの場合：ごみ箱コレクションへ
                    string src = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "send", mail.mailName);
                    string dst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName);
                    try
                    {
                        if (File.Exists(src))
                        {
                            // 既に同名ファイルがごみ箱にあれば上書きを避けるためユニーク化
                            if (File.Exists(dst))
                            {
                                string uniqueDst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", Guid.NewGuid().ToString() + "_" + mail.mailName);
                                File.Move(src, uniqueDst);
                                mail.mailName = Path.GetFileName(uniqueDst);
                            }
                            else
                            {
                                File.Move(src, dst);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"Move error: {ex.Message}");
                    }

                    collectionMail[DELETE].Add(mail);
                    collectionMail[SEND].Remove(mail);
                }
                else if (listMain.Columns[0].Text == "差出人または宛先")
                {
                    // ごみ箱内の項目は後で確認して完全削除
                    trashMails.Add(mail);
                }
            }

            // ごみ箱内削除はまとめて確認
            if (trashMails.Count > 0)
            {
                var msg = $"選択したごみ箱内メール {trashMails.Count} 件は完全に削除されます。\r\nよろしいですか？";
                if (MessageBox.Show(msg, "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    foreach (var mail in trashMails)
                    {
                        try
                        {
                            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName);
                            if (File.Exists(path))
                            {
                                File.Delete(path);
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Error($"Delete error: {ex.Message}");
                        }
                        collectionMail[DELETE].Remove(mail);
                    }
                }
            }

            // 表示を更新
            UpdateView();
        }

        private void menuNotReadYet_Click(object sender, EventArgs e)
        {
            if (currentRow < 0 || currentRow >= listMain.Items.Count)
                return;

            var selItem = listMain.Items[currentRow];
            if (!(selItem.Tag is Mail))
                return;

            Mail mail = (Mail)selItem.Tag;
            mail.notReadYet = true;

            // ツリービューとリストビューの表示を更新する
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
            // URLを開くかの確認をする
            if (MessageBox.Show("クリックしたURL\n" + e.LinkText + "\nを開きますか？\nページによっては詐欺のページやウイルスの可能性もあるため\n注意してURLを開いてください。", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(e.LinkText);
            }
        }

        private void buttonAtachMenu_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            // ファイルを開くかの確認をする
            if (MessageBox.Show(e.ClickedItem.Text + "を開きますか？\nファイルによってはウイルスの可能性もあるため\n注意してファイルを開いてください。", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                // 送信メールで添付ファイルが存在する場合
                if (currentMail.atach != string.Empty)
                {
                    System.Diagnostics.Process.Start(e.ClickedItem.Text);
                }
                else
                {
                    System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + "\\mbox\\tmp\\" + e.ClickedItem.Text);
                }
            }
        }

        // 受信日時をローカル時刻に変換して表示する（オフセット表記は表示しない）
        private string FormatReceivedDate(string dateText)
        {
            if (string.IsNullOrEmpty(dateText))
                return "";

            // 未送信はそのまま
            if (dateText == "未送信")
                return dateText;

            // まず DateTimeOffset としてパースを試みる（オフセット情報が含まれる場合はこちら）
            DateTimeOffset dto;
            if (DateTimeOffset.TryParse(dateText, out dto))
            {
                // ローカル時刻に変換して表示（オフセットは表示しない）
                var local = dto.ToLocalTime();
                return local.ToString("yyyy/MM/dd HH:mm:ss");
            }

            // DateTime としてパースできる場合
            DateTime dt;
            if (DateTime.TryParse(dateText, out dt))
            {
                // Kind に応じてローカルに変換またはそのまま扱う
                if (dt.Kind == DateTimeKind.Utc)
                {
                    return dt.ToLocalTime().ToString("yyyy/MM/dd HH:mm:ss");
                }
                if (dt.Kind == DateTimeKind.Unspecified)
                {
                    // 未指定ならローカルとして扱う
                    var localSpecified = DateTime.SpecifyKind(dt, DateTimeKind.Local);
                    return localSpecified.ToString("yyyy/MM/dd HH:mm:ss");
                }
                // Local の場合
                return dt.ToString("yyyy/MM/dd HH:mm:ss");
            }

            // 解析不能なら元文字列を返す
            return dateText;
        }

        private void Application_Idle(object sender, EventArgs e)
        {
            toolReplyButton.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            toolDeleteButton.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuNotReadYet.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuDelete.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuClearTrush.Enabled = collectionMail[DELETE].Count > 0;
            menuMailDelete.Enabled = listMain.SelectedItems.Count > 0 && mailBoxViewFlag == false;
            menuMailReply.Enabled = listMain.SelectedItems.Count == 1 && mailBoxViewFlag == false;
            menuFileClearTrush.Enabled = collectionMail[DELETE].Count > 0;
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
                    collectionMail[SEND].Add(mail);
                    SaveMail(mail);
                }

                // ツリービューとリストビューの表示を更新する
                UpdateView();
            }
        }

        private void menuClearTrush_Click(object sender, EventArgs e)
        {
            // ごみ箱内の全削除確認
            if (MessageBox.Show("ごみ箱内のメールをすべて完全に削除します。\r\nよろしいですか？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                // ArrayListはToList()を持たないため、Cast<Mail>()でIEnumerable<Mail>にし、ToArray()でスナップショットを作成
                foreach (Mail mail in collectionMail[DELETE].Cast<Mail>().ToArray())
                {
                    try
                    {
                        string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName);
                        if (File.Exists(path))
                        {
                            File.Delete(path);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"Delete error: {ex.Message}");
                    }
                    collectionMail[DELETE].Remove(mail);
                }
                // 表示を更新
                UpdateView();
            }
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
            // 選択中のメールをファイルに保存する
            if (currentMail == null)
                return;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = currentMail.subject + ".eml";
            sfd.Filter = "EMLファイル (*.eml)|*.eml|すべてのファイル (*.*)|*.*";
            sfd.FilterIndex = 0;
            sfd.RestoreDirectory = true;
            sfd.OverwritePrompt = true;

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string mailPath = "";
                // mailPath は現在表示しているフォルダに応じて決定
                if (listMain.Columns[0].Text == "差出人")
                {
                    // 受信メールの場合
                    mailPath = System.Windows.Forms.Application.StartupPath + "\\mbox\\inbox\\";
                }
                else if (listMain.Columns[0].Text == "宛先")
                {
                    // 送信メールの場合
                    mailPath = "";
                }
                else if (listMain.Columns[0].Text == "差出人または宛先")
                {
                    // ごみ箱の場合
                    mailPath = System.Windows.Forms.Application.StartupPath + "\\mbox\\trush\\";
                }
                if (!string.IsNullOrEmpty(mailPath) && File.Exists(mailPath + currentMail.mailName))
                {
                    File.Copy(mailPath + currentMail.mailName, sfd.FileName, true);
                }
                else
                {
                    // 送信メールの場合は新規に EML ファイルを作成する
                    var message = new MimeMessage();
                    message.From.Add(MailboxAddress.Parse(Mail.userAddress));
                    message.To.Add(MailboxAddress.Parse(currentMail.address));
                    message.Subject = currentMail.subject;

                    var textPart = new TextPart(TextFormat.Text)
                    {
                        Text = currentMail.body,
                    };

                    // 添付ファイルの確認
                    var files = currentMail.atach.Split(';').Select(f => f.Trim()).Where(f => File.Exists(f));

                    // 添付ファイルが複数ある場合
                    if (!string.IsNullOrWhiteSpace(currentMail.atach) && files.Count() > 0)
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
                    using (var stream = File.Create(sfd.FileName))
                    {
                        message.WriteTo(stream);
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
            // 選択中のメールを元に戻す
            foreach (ListViewItem selItem in listMain.SelectedItems)
            {
                if (!(selItem.Tag is Mail))
                    continue;
                Mail mail = (Mail)selItem.Tag;
                if (mail.mailName != string.Empty && mail.mailName.Contains(".eml"))
                {
                    // 受信メールの場合：受信コレクションへ（ファイルも移動）
                    string src = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName);
                    string dst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", mail.mailName);
                    try
                    {
                        if (File.Exists(src))
                        {
                            // 既に同名ファイルが受信箱にあれば上書きを避けるためユニーク化
                            if (File.Exists(dst))
                            {
                                string uniqueDst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", Guid.NewGuid().ToString() + "_" + mail.mailName);
                                File.Move(src, uniqueDst);
                                mail.mailName = Path.GetFileName(uniqueDst);
                            }
                            else
                            {
                                File.Move(src, dst);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"Move error: {ex.Message}");
                    }
                    collectionMail[RECEIVE].Add(mail);
                    collectionMail[DELETE].Remove(mail);
                }
                else
                {
                    // 送信メールの場合：送信コレクションへ
                    string src = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName);
                    string dst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "send", mail.mailName);
                    try
                    {
                        if (File.Exists(src))
                        {
                            // 既に同名ファイルが受信箱にあれば上書きを避けるためユニーク化
                            if (File.Exists(dst))
                            {
                                string uniqueDst = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "send", Guid.NewGuid().ToString() + "_" + mail.mailName);
                                File.Move(src, uniqueDst);
                                mail.mailName = Path.GetFileName(uniqueDst);
                            }
                            else
                            {
                                File.Move(src, dst);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"Move error: {ex.Message}");
                    }

                    collectionMail[SEND].Add(mail);
                    collectionMail[DELETE].Remove(mail);
                }
            }
            // 表示を更新
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

            ArrayList sourceList = null;

            // 現在のフォルダに応じて検索対象を切り替える
            if (listMain.Columns[0].Text == "差出人")
                sourceList = collectionMail[RECEIVE];
            else if (listMain.Columns[0].Text == "宛先")
                sourceList = collectionMail[SEND];
            else if (listMain.Columns[0].Text == "差出人または宛先")
                sourceList = collectionMail[DELETE];
            else
                return;

            // LINQ でフィルタ
            var filtered = sourceList
                .Cast<Mail>()
                .Where(m =>
                    (m.address != null && m.address.Contains(keyword)) ||
                    (m.subject != null && m.subject.Contains(keyword)) ||
                    (m.body != null && m.body.Contains(keyword))
                )
                .ToList();

            // 結果を表示
            ShowSearchResult(filtered);
        }

        private void ShowSearchResult(List<Mail> list)
        {
            listMain.BeginUpdate();
            listMain.Items.Clear();

            var baseFont = listMain.Font;

            foreach (var mail in list)
            {
                ListViewItem item = new ListViewItem(mail.address);
                item.SubItems.Add(mail.subject);
                item.SubItems.Add(FormatReceivedDate(mail.date));

                long sizeBytes = GetMailFileSize(mail);
                item.SubItems.Add(FormatSize(sizeBytes));
                item.SubItems.Add(mail.mailName);

                item.Tag = mail;

                listMain.Items.Add(item);

                item.Font = mail.notReadYet
                    ? new Font(baseFont, FontStyle.Bold)
                    : new Font(baseFont, FontStyle.Regular);
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
            try
            {
                // inbox / trush のどちらかにある
                string[] paths = new[]
                {
                    Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", mail.mailName),
                    Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName)
                };

                foreach (var path in paths)
                {
                    if (File.Exists(path))
                    {
                        using (var stream = File.OpenRead(path))
                        {
                            // ★ メール全体を読み込まず、ヘッダーだけ読む
                            var parser = new MimeParser(stream, MimeFormat.Entity);
                            var message = parser.ParseMessage();

                            return message.Attachments?.Any() == true;
                        }
                    }
                }
            }
            catch
            {
                // 読めなかった場合は添付なし扱い
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
            {
                Directory.Delete(tmpPath, true);
            }

            // 添付メニュー初期化
            buttonAtachMenu.DropDownItems.Clear();
            buttonAtachMenu.Visible = false;

            // メールファイルのパスを決定
            string mailPath = "";
            if (listMain.Columns[0].Text == "差出人") // 受信
                mailPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "inbox", mail.mailName);
            else if (listMain.Columns[0].Text == "差出人または宛先") // ごみ箱
                mailPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "mbox", "trush", mail.mailName);
            else if (listMain.Columns[0].Text == "宛先") // 送信
                mailPath = ""; // 送信メールはファイルがない場合もある

            // 未読解除（受信 or ごみ箱）
            if (listMain.Columns[0].Text != "宛先")
                mail.notReadYet = false;

            // メールファイルが存在する場合は読み込む
            if (!string.IsNullOrEmpty(mailPath) && File.Exists(mailPath) && mailPath.Contains(".eml"))
            {
                MimeMessage message = MimeMessage.Load(mailPath);

                // HTMLメール
                if (message.HtmlBody != null)
                {
                    browserMail.Visible = true;
                    richTextBody.Visible = false;
                    browserMail.DocumentText = message.HtmlBody;
                }
                else
                {
                    // テキストメール
                    browserMail.Visible = false;
                    richTextBody.Visible = true;
                    richTextBody.Text = message.TextBody;
                    ColorizeQuoteLines();
                }

                // 添付ファイル処理
                if (message.Attachments.Any())
                {
                    buttonAtachMenu.Visible = true;

                    // 展開フォルダ作成
                    Directory.CreateDirectory(tmpPath);

                    foreach (var attachment in message.Attachments)
                    {
                        var part = (MimePart)attachment;
                        string fileName = part.FileName;
                        string savePath = Path.Combine(tmpPath, fileName);

                        using (var stream = File.Create(savePath))
                        {
                            part.Content.DecodeTo(stream);
                        }

                        var menuItem = new ToolStripMenuItem(fileName);
                        menuItem.Click += (s, e) => System.Diagnostics.Process.Start(savePath);
                        buttonAtachMenu.DropDownItems.Add(menuItem);
                    }
                }
            }
            else
            {
                // 送信メール（ファイルがない場合）
                browserMail.Visible = false;
                richTextBody.Visible = true;
                richTextBody.Text = mail.body;
                ColorizeQuoteLines();

                // 送信メールの添付ファイルがある場合
                if (mail.atach != string.Empty)
                {
                    // 送信メールで添付ファイルが存在する場合
                    string[] atachFiles = mail.atach.Split(';');
                    foreach (string atach in atachFiles)
                    {
                        Icon appIcon = System.Drawing.Icon.ExtractAssociatedIcon(atach);
                        buttonAtachMenu.DropDownItems.Add(atach, appIcon.ToBitmap());
                    }
                    buttonAtachMenu.Visible = true;
                }
            }
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

            var mail = listMain.SelectedItems[0].Tag as Mail;
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
            string sendFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "mbox", "send");
            Directory.CreateDirectory(sendFolder);

            // mailName がまだ無い場合は新規生成
            if (string.IsNullOrEmpty(mail.mailName))
            {
                // ファイル名用の日時は必ず「現在時刻」を使う
                string safeDate = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                mail.mailName = $"{safeDate}_{mail.id}.mail";
            }

            string filePath = Path.Combine(sendFolder, mail.mailName);

            // JSON で保存（Newtonsoft.Json）
            string json = JsonConvert.SerializeObject(mail, Formatting.Indented);
            File.WriteAllText(filePath, json, Encoding.UTF8);

            // ArrayListに追加
            if (!collectionMail[SEND].Contains(mail))
            {
                collectionMail[SEND].Add(mail);
            }
        }

    }
}
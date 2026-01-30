using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace MizuMail
{
    public partial class FormMailCreate : Form
    {
        bool senderEmail = false;
        private FindDialog findDialog;
        private ReplaceDialog replaceDialog;

        public FormMailCreate()
        {
            InitializeComponent();

            // Application.Idle に登録してアイドル時にツールボタン／メニューの状態を更新する
            System.Windows.Forms.Application.Idle += OnApplicationIdle;
        }

        private void buttonSend_Click(object sender, EventArgs e)
        {
            // 送り先の必須チェック
            if (string.IsNullOrWhiteSpace(textMailTo.Text))
            {
                MessageBox.Show("宛先が指定されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textMailTo.Focus();
                return;
            }

            // 件名が空欄の場合は確認ダイアログを表示
            if (string.IsNullOrWhiteSpace(textMailSubject.Text))
            {
                var result = MessageBox.Show("件名が指定されていません。本当に送信しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result != DialogResult.Yes)
                {
                    textMailSubject.Focus();
                    return;
                }
            }

            // 本文の必須チェック
            if (string.IsNullOrWhiteSpace(textMailBody.Text))
            {
                MessageBox.Show("本文が入力されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textMailBody.Focus();
                return;
            }

            if (textMailBody.Text.Contains("添付") && buttonAttachList.DropDownItems.Count == 0)
            {
                if (MessageBox.Show("添付ファイルがありません。送信しますか？",
                    "確認", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
            }

            // ダイアログの戻り値を OK にしてフォームを閉じる
            senderEmail = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void menuCopy_Click(object sender, EventArgs e)
        {
            textMailBody.Copy();
        }

        private void menuPaste_Click(object sender, EventArgs e)
        {
            textMailBody.Paste();
        }

        private void menuDelete_Click(object sender, EventArgs e)
        {
            textMailBody.SelectedText = "";
        }

        private void menuCut_Click(object sender, EventArgs e)
        {
            textMailBody.Cut();
        }

        private void menuAllSelect_Click(object sender, EventArgs e)
        {
            textMailBody.SelectAll();
        }

        private void menuUndo_Click(object sender, EventArgs e)
        {
            textMailBody.Undo();
        }

        private void menuEdit_Click(object sender, EventArgs e)
        {
            // 既存処理は共通化した UpdateClipboardButtons に移譲
            UpdateClipboardButtons();
        }

        private void menuClose_Click(object sender, EventArgs e)
        {
            buttonCancel_Click(sender, e);
        }

        private void menuSend_Click(object sender, EventArgs e)
        {
            buttonSend_Click(sender, e);
        }

        /// <summary>
        /// UI の状態を更新する共通関数（メニュー／ツールボタン）
        /// 軽量にし、クリップボードアクセスは例外安全に行う
        /// </summary>
        private void UpdateClipboardButtons()
        {
            if (textMailBody == null) return;

            // Undo
            bool canUndo = false;
            try
            {
                canUndo = textMailBody.CanUndo;
            }
            catch
            {
                canUndo = false;
            }

            // Selection-based commands
            int selectionLength = 0;
            try
            {
                selectionLength = textMailBody.SelectionLength;
            }
            catch
            {
                selectionLength = 0;
            }

            // Clipboard check を例外安全に行う（別プロセスがクリップボードをロックしている可能性がある）
            bool clipboardHasText = false;
            try
            {
                clipboardHasText = Clipboard.ContainsText();
            }
            catch
            {
                clipboardHasText = false;
            }

            // メニュー項目の更新
            menuUndo.Enabled = canUndo;
            menuCut.Enabled = selectionLength > 0;
            menuCopy.Enabled = selectionLength > 0;
            menuDelete.Enabled = selectionLength > 0;
            menuPaste.Enabled = clipboardHasText;

            // ツールストリップボタンの更新（Designer 側で定義されていることが前提）
            try
            {
                if (toolStripButtonCut != null) toolStripButtonCut.Enabled = selectionLength > 0;
                if (toolStripButtonCopy != null) toolStripButtonCopy.Enabled = selectionLength > 0;
                if (toolStripButtonPaste != null) toolStripButtonPaste.Enabled = clipboardHasText;
            }
            catch
            {
                // ツールストリップ更新で例外が出ても無視（安全策）
            }
        }

        /// <summary>
        /// Application.Idle イベントハンドラ。頻繁に呼ばれるため軽量に保つ。
        /// </summary>
        private void OnApplicationIdle(object sender, EventArgs e)
        {
            UpdateClipboardButtons();
        }

        /// <summary>
        /// レイアウト調整：TextBox の幅を更新し、
        /// Panel1 の高さは直接設定せず SplitContainer.SplitterDistance を設定する。
        /// Panel1 の高さは textMailSubject のボトムから 5px 下とする（範囲内にクランプ）。
        /// </summary>
        private void AdjustLayout()
        {
            // --- 追加: ステータスバー高さ分の余白を Panel2 に確保（スクロールバーが隠れないようにする）
            try
            {
                if (splitContainer1 != null && splitContainer1.Panel2 != null && statusStrip1 != null)
                {
                    int extra = statusStrip1.Height + 2; // 余裕 +2px
                    splitContainer1.Panel2.Padding = new Padding(0, 0, 0, extra);
                }
            }
            catch
            {
                // 無視
            }

            // テキスト幅調整（デザイナで Anchor を使う方が望ましい）
            textMailTo.Width = this.ClientSize.Width - textMailTo.Left - 5;
            textMailCc.Width = this.ClientSize.Width - textMailCc.Left - 5;
            textMailBcc.Width = this.ClientSize.Width - textMailBcc.Left - 5;
            textMailSubject.Width = this.ClientSize.Width - textMailSubject.Left - 5;

            // splitContainer1 が未生成または高さが 0 の場合は処理しない
            if (splitContainer1 == null || splitContainer1.Height <= 0) return;

            // Panel1 の目標高さ：textMailSubject の下端 + 5px
            int desiredPanel1Height = textMailSubject.Bottom + 5;
            int minDistance = Math.Max(0, splitContainer1.Panel1MinSize);
            int maxDistance = splitContainer1.Height - splitContainer1.Panel2MinSize - splitContainer1.SplitterWidth;

            if (maxDistance < minDistance)
            {
                // レイアウトが確定していないか不整合。設定を保留して再実行する。
                // 例: UIスレッドの次回ループで再実行
                this.BeginInvoke((Action)(() => AdjustLayout()));
                return;
            }

            int distance = Math.Min(Math.Max(desiredPanel1Height, minDistance), maxDistance);
            splitContainer1.SplitterDistance = distance;
        }

        private void FormMailCreate_SizeChanged(object sender, EventArgs e)
        {
            AdjustLayout();
        }

        // (既存ファイルに追加するコード部分)
        private void FormMailCreate_Load(object sender, EventArgs e)
        {
            AdjustLayout();

            // Controls の Z-順／Dock 計算が崩れていることがあるため
            // ステータスバーを先にレイアウトに反映させ、Fill が残り領域を占めるようにする
            try
            {
                // インデックスの調整は例外安全に行う
                // 0: 最前面、Controls.Count-1: 最背面。ここでは statusStrip を最前面にして Dock の計算を安定させる
                this.Controls.SetChildIndex(this.statusStrip1, 0);
                this.Controls.SetChildIndex(this.toolStripContainer1, 1);
            }
            catch
            {
                // 無視
            }

            var addressBook = AddressBook.LoadAddressBook();
            var auto = new AutoCompleteStringCollection();
            auto.AddRange(addressBook.Entries.Select(a => a.Email).ToArray());

            SetupAutoComplete(textMailTo, auto);
            SetupAutoComplete(textMailCc, auto);
            SetupAutoComplete(textMailBcc, auto);

            // 初回ロード時にも状態更新
            UpdateClipboardButtons();

            // レイアウトを再計算して強制描画
            statusStrip1.PerformLayout();
            statusStrip1.Refresh();
            EnsureAttachButtonVisible();
        }

        private void toolAddAttachment_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "添付ファイルを選択してください";
                openFileDialog.Filter = "すべてのファイル (*.*)|*.*";
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (!string.IsNullOrEmpty(openFileDialog.FileName))
                    {
                        // ボタンが誤った Owner に入っている場合は一旦除去して statusStrip1 に追加し直す
                        try
                        {
                            if (buttonAttachList.Owner != statusStrip1)
                            {
                                if (buttonAttachList.Owner != null)
                                {
                                    try { buttonAttachList.Owner.Items.Remove(buttonAttachList); } catch { }
                                }
                                statusStrip1.Items.Add(buttonAttachList);
                            }
                        }
                        catch
                        {
                            // 無視して続行
                        }

                        buttonAttachList.Visible = true;
                        labelMessage.Text = openFileDialog.FileNames.Length + "個のファイルをメールに添付しました。";

                        // 選択されたファイルを添付リストに追加する
                        foreach (var file in openFileDialog.FileNames)
                        {
                            // 既に同じファイルが添付されている場合はスキップ
                            if (buttonAttachList.DropDownItems.Cast<ToolStripItem>().Any(item => item.Text.Equals(file, StringComparison.OrdinalIgnoreCase)))
                            {
                                continue;
                            }
                            var appIcon = System.Drawing.Icon.ExtractAssociatedIcon(file);
                            buttonAttachList.DropDownItems.Add(file, appIcon.ToBitmap());
                        }
                        EnsureAttachButtonVisible();

                        // レイアウトと描画を強制
                        statusStrip1.PerformLayout();
                        statusStrip1.Refresh();
                    }
                }
            }
        }

        // 追加: ステータスバーに添付ボタンを確実に配置して再レイアウトするヘルパー
        private void EnsureAttachButtonVisible()
        {
            try
            {
                // レイアウト方式をオーバーフロー許容にしてアイテムが切られないようにする
                statusStrip1.LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow;

                // Button を statusStrip に所属させる（重複追加は防止）
                if (buttonAttachList.Owner != statusStrip1)
                {
                    try { buttonAttachList.Owner?.Items.Remove(buttonAttachList); } catch { }
                    if (!statusStrip1.Items.Contains(buttonAttachList))
                    {
                        statusStrip1.Items.Add(buttonAttachList);
                    }
                }

                // 右端に寄せる・表示スタイルを調整（幅を小さくすると表示されやすい）
                buttonAttachList.Alignment = ToolStripItemAlignment.Right;
                buttonAttachList.DisplayStyle = ToolStripItemDisplayStyle.Image; // 必要なら ImageAndText に戻す

                // 強制レイアウトと描画
                statusStrip1.PerformLayout();
                statusStrip1.Refresh();

                // デバッグ出力（実行時に Immediate/Output で確認）
                Debug.WriteLine($"EnsureAttachButtonVisible: Owner={buttonAttachList.Owner?.Name}, Index={statusStrip1.Items.IndexOf(buttonAttachList)}, BtnBounds={buttonAttachList.Bounds}, StatusBounds={statusStrip1.Bounds}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("EnsureAttachButtonVisible error: " + ex.Message);
            }
        }

        private void buttonAttachList_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (MessageBox.Show(e.ClickedItem.Text + "を削除しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            // 選択した添付ファイルメニューを削除する
            buttonAttachList.DropDownItems.Remove(e.ClickedItem);
            // 添付ファイルの数が0になったらリストを閉じる
            buttonAttachList.Visible = buttonAttachList.DropDownItems.Count > 0;

            labelMessage.Text = "";
        }

        private void menuFind_Click(object sender, EventArgs e)
        {
            if (findDialog == null || findDialog.IsDisposed)
            {
                findDialog = new FindDialog();
                findDialog.FindNextRequested += FindNext;
            }
            findDialog.Show();
            findDialog.Activate();
        }

        private void menuReplace_Click(object sender, EventArgs e)
        {
            if (replaceDialog == null || replaceDialog.IsDisposed)
            {
                replaceDialog = new ReplaceDialog();
                replaceDialog.FindNextRequested += FindNext;
                replaceDialog.ReplaceRequested += ReplaceCurrent;
                replaceDialog.ReplaceAllRequested += ReplaceAll;
            }
            replaceDialog.Show();
            replaceDialog.Activate();
        }

        private void FindNext(string text, bool matchCase, bool searchDown)
        {
            StringComparison cmp = matchCase ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase;

            int start = textMailBody.SelectionStart;
            int index;

            if (searchDown)
            {
                index = textMailBody.Text.IndexOf(text, start + textMailBody.SelectionLength, cmp);
            }
            else
            {
                string before = textMailBody.Text.Substring(0, start);
                index = before.LastIndexOf(text, cmp);
            }

            if (index >= 0)
            {
                textMailBody.SelectionStart = index;
                textMailBody.SelectionLength = text.Length;
                textMailBody.ScrollToCaret();
            }
            else
            {
                MessageBox.Show("検索文字列は見つかりませんでした。");
            }
        }

        private void ReplaceCurrent(string findText, string replaceText, bool matchCase)
        {
            StringComparison cmp = matchCase ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase;

            if (textMailBody.SelectionLength > 0 &&
                string.Equals(textMailBody.SelectedText, findText, cmp))
            {
                textMailBody.SelectedText = replaceText;
            }

            FindNext(findText, matchCase, true);
        }

        private void ReplaceAll(string findText, string replaceText, bool matchCase)
        {
            StringComparison cmp = matchCase ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase;

            int count = 0;
            int index = 0;

            while (true)
            {
                index = textMailBody.Text.IndexOf(findText, index, cmp);
                if (index < 0) break;

                textMailBody.Select(index, findText.Length);
                textMailBody.SelectedText = replaceText;

                index += replaceText.Length;
                count++;
            }

            MessageBox.Show($"{count} 件置換しました。");
        }

        private void textMailBody_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var file in files)
            {
                // 既に同じファイルが添付されている場合はスキップ
                if (buttonAttachList.DropDownItems.Cast<ToolStripItem>().Any(item => item.Text.Equals(file, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }
                var appIcon = System.Drawing.Icon.ExtractAssociatedIcon(file);
                buttonAttachList.DropDownItems.Add(file, appIcon.ToBitmap());
            }

            if (files.Length > 0)
            {
                buttonAttachList.Visible = true;
                labelMessage.Text = $"{files.Length} 個のファイルをメールに添付しました。";
                EnsureAttachButtonVisible();
                // レイアウトと描画を強制
                statusStrip1.PerformLayout();
                statusStrip1.Refresh();
            }
        }

        private void textMailBody_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void menuSelectFromAddressBook_Click(object sender, EventArgs e)
        {
            // どの TextBox から呼ばれたか
            var target = contextAddress.SourceControl as System.Windows.Forms.TextBox;
            if (target == null)
                return;

            // アドレス帳を読み込む
            var book = AddressBook.LoadAddressBook();

            // 編集ダイアログを開く（選択モード）
            using (var dlg = new AddressBookEditorForm(book))
            {
                dlg.SelectMode = true;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // 選択されたアドレスを取得
                    var entry = dlg.SelectedEntry;
                    if (entry != null)
                    {
                        // 挿入形式：「表示名 <メールアドレス>」
                        string insert = $"{entry.DisplayName} <{entry.Email}>";

                        // TextBox に挿入
                        if (string.IsNullOrWhiteSpace(target.Text))
                            target.Text = insert;
                        else
                            target.Text += "; " + insert;
                    }

                    // 保存
                    AddressBook.SaveAddressBook(book);
                }
            }
        }

        private void SetupAutoComplete(System.Windows.Forms.TextBox box, AutoCompleteStringCollection auto)
        {
            box.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            box.AutoCompleteSource = AutoCompleteSource.CustomSource;
            box.AutoCompleteCustomSource = auto;
        }

        private void textMailTo_Leave(object sender, EventArgs e)
        {
            var book = AddressBook.LoadAddressBook();
            foreach (var entry in book.Entries)
            {
                if (textMailTo.Text.Trim().Equals(entry.Email, StringComparison.OrdinalIgnoreCase))
                {
                    textMailTo.Text = $"{entry.DisplayName} <{entry.Email}>";
                    break;
                }
            }
        }

        private void textMailCc_Leave(object sender, EventArgs e)
        {
            var book = AddressBook.LoadAddressBook();
            foreach (var entry in book.Entries)
            {
                if (textMailCc.Text.Trim().Equals(entry.Email, StringComparison.OrdinalIgnoreCase))
                {
                    textMailCc.Text = $"{entry.DisplayName} <{entry.Email}>";
                    break;
                }
            }
        }

        private void textMailBcc_Leave(object sender, EventArgs e)
        {
            var book = AddressBook.LoadAddressBook();
            foreach (var entry in book.Entries)
            {
                if (textMailBcc.Text.Trim().Equals(entry.Email, StringComparison.OrdinalIgnoreCase))
                {
                    textMailBcc.Text = $"{entry.DisplayName} <{entry.Email}>";
                    break;
                }
            }
        }

        private void menuInsertSignature_Click(object sender, EventArgs e)
        {
            var sig = FormMain.LoadSignature();

            if (sig.Enabled && !string.IsNullOrWhiteSpace(sig.Signature))
            {
                textMailBody.SelectedText = "\r\n" + sig.Signature;
            }
        }

        private void menuInsertAttachment_Click(object sender, EventArgs e)
        {
            toolAddAttachment_Click(sender, e);
        }

        /// <summary>
        /// フォーム終了時に Application.Idle の登録を解除
        /// </summary>
        private void FormMailCreate_FormClosing(object sender, FormClosingEventArgs e)
        {
            var isEdit = (!string.IsNullOrWhiteSpace(textMailTo.Text) || !string.IsNullOrWhiteSpace(textMailSubject.Text) || !string.IsNullOrWhiteSpace(textMailBody.Text) || buttonAttachList?.DropDownItems?.Count > 0);
            // 送信フラグがfalseかつ宛先、件名、本文、添付ファイルが1件以上のいずれかが入力されている状態で閉じる場合、確認ダイアログを表示する
            if (senderEmail == false && isEdit)
            {
                var result = MessageBox.Show("編集中の内容が失われます。本当に閉じますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result != DialogResult.Yes)
                {
                    // 閉じるのをキャンセル
                    this.DialogResult = DialogResult.None;
                    e.Cancel = true;
                    return;
                }
            }

            try
            {
                System.Windows.Forms.Application.Idle -= OnApplicationIdle;
            }
            catch
            {
                // 無視
            }
        }
    }
}

using MailKit;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace MizuMail
{
    public class Mail
    {
        // 静的フィールド(メールアカウント情報)
        public static string fromName;          // 差出人名
        public static string userAddress;       // ユーザのアドレス
        public static string userName;          // ユーザ名
        public static string smtpServerName;    // 送信サーバの名前
        public static string popServerName;     // 受信サーバの名前
        public static string imapServerName;    // IMAPサーバの名前
        public static int imapPortNo;           // IMAPサーバのポート番号
        public static int smtpPortNo;           // 送信サーバのポート番号
        public static int popPortNo;            // 受信サーバのポート番号
        public static string password;          // 受信サーバのパスワード
        public static bool deleteMail;          // 受信後メール削除フラグ
        public static bool alertSound;          // 受信通知音を鳴らすかどうかのフラグ
        public static string alertSoundFile;    // 受信通知音ファイル名
        public static int checkInterval;        // メールチェック間隔(分)
        public static bool checkMail;           // メールチェックを行うかどうかのフラグ
        public static bool useSsl;              // SSLを使用するかどうかのフラグ
        public static bool receiveMethod_Pop3;  // 受信方法がPOP3かどうかのフラグ(true:POP3、false:IMAP4)

        // インスタンスフィールド(メールの情報)
        public string from;                     // 差出人
        public MailAddress From { get; set; }   // MailboxAddressオブジェクト
        public string address;                  // 宛先のアドレス
        public string ccaddress;                // CCアドレス
        public string bccaddress;               // BCCアドレス
        public string subject;                  // 件名
        public string body;                     // 本文
        public MimeMessage message;             // MimeMessageオブジェクト
        public bool hasAtach;                   // 添付ファイルがあるかどうか
        public string atach;                    // 添付ファイル名
        public List<string> attachList { get; set; } = new List<string>();  // 添付ファイル名リスト
        public string date;                     // 受信日時またはメール送信日時
        public string mailName;                 // メールファイル名
        public string uidl;                     // UIDL
        public bool notReadYet;                 // 未読、未送信ならtrue
        public bool isDraft;                    // 下書き or 未送信
        public string id { get; set; } = Guid.NewGuid().ToString(); // メールの一意識別子
        public MailFolder Folder { get; set; }  // メールフォルダ情報
        public MailFolder LastFolder;           // 直前のフォルダ
        public string LastMailName;             // 直前のファイル名
        public string mailPath;                 // メールのフルパス
        public bool isHtml;                     // 本文がHTML形式かどうか
        public long sizeBytes;                  // ファイルサイズ（バイト）
        public string preview;                  // プレビュー用の1行テキスト

        // コンストラクタ
        public Mail()
        {
            notReadYet = false;
            isDraft = false;
        }

        public Mail(string address, string cc, string bcc, string subject, string body, string atach, string date, string mailName, string uidl, bool notReadYet)
        {
            this.address = address;
            this.ccaddress = cc;
            this.bccaddress = bcc;
            this.subject = subject;
            this.body = body;
            this.atach = atach;
            this.date = date;
            this.mailName = mailName;
            this.uidl = uidl;
            this.notReadYet = notReadYet;
        }

        public static Mail FromMimeMessage(MimeMessage msg)
        {
            if (msg == null)
                return null;

            Mail mail = new Mail();

            // ★ MimeMessage を保持（後で本文や添付を読むために必須）
            mail.message = msg;

            // ★ 差出人・宛先
            mail.From = msg.From?.Mailboxes?.FirstOrDefault() != null ? new MailAddress(msg.From.Mailboxes.First().Address, msg.From.Mailboxes.First().Name) : null;
            mail.from = msg.From?.ToString() ?? "";
            mail.address = msg.To?.ToString() ?? "";
            mail.ccaddress = msg.Cc?.ToString() ?? "";
            mail.bccaddress = msg.Bcc?.ToString() ?? "";

            // ★ 件名
            mail.subject = msg.Subject ?? "";

            // ★ 本文（安全版）
            mail.body =
                msg.GetTextBody(MimeKit.Text.TextFormat.Plain) ??
                msg.GetTextBody(MimeKit.Text.TextFormat.Html) ??
                "";

            // ★ 日付
            mail.date = msg.Date.ToString("yyyy/MM/dd HH:mm:ss");

            // ★ 添付ファイル
            mail.hasAtach = msg.Attachments?.Any() ?? false;
            mail.attachList = msg.Attachments?
                .Select(a => a.ContentDisposition?.FileName ?? a.ContentType.Name)
                .Where(n => !string.IsNullOrEmpty(n))
                .ToList()
                ?? new List<string>();

            // ★ mailName は後で上書きされるので仮でOK
            mail.mailName = Guid.NewGuid().ToString() + ".eml";

            // ★ Folder は受信後に Pop3Receive 側で必ず設定する
            // mail.Folder = folderManager.Inbox; ← ここでは設定しない

            return mail;
        }

    }
}

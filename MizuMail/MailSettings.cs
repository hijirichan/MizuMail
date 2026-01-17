using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MizuMail
{
    public class MailSettings
    {
        public string m_fromName = string.Empty;                        // ユーザ(差出人)の名前 
        public string m_mailAddress = string.Empty;                     // ユーザのメールアドレス
        public string m_userName = string.Empty;                        // ユーザ名
        public string m_passWord = string.Empty;                        // POP3のパスワード
        public string m_popServer = string.Empty;                       // POP3サーバ名
        public string m_imapServer = string.Empty;                      // IMAP4サーバ名
        public int m_imapPortNo = 143;                                  // IMAP4のポート番号
        public string m_smtpServer = string.Empty;                      // SMTPサーバ名
        public int m_popPortNo = 110;                                   // POP3のポート番号
        public int m_smtpPortNo = 25;                                   // SMTPのポート番号
        public bool m_deleteMail = false;                               // POP受信時メール削除フラグ
        public bool m_useSsl = false;                                   // SSLを使用するかどうかのフラグ
        public bool m_ReceiveMethod_Pop3 = true;                        // 受信方法がPOP3かどうかのフラグ(true:POP3、false:IMAP4)
        public int m_windowLeft = 0;                                    // ウィンドウの左上のLeft座標
        public int m_windowTop = 0;                                     // ウィンドウの左上のTop座標
        public int m_windowWidth = 0;                                   // ウィンドウの幅
        public int m_windowHeight = 0;                                  // ウィンドウの高さ
        public FormWindowState m_windowStat = FormWindowState.Normal;   // ウィンドウの状態
        public bool m_alertSound = false;                               // 受信通知音を鳴らすかどうかのフラグ
        public string m_alertSoundFile = string.Empty;                  // 受信通知音ファイル名
        public int m_checkInterval = 5;                                 // メールチェック間隔(分)
        public bool m_checkMail = false;                                // メールチェックを行うかどうかのフラグ
    }
}

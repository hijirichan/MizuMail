using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MizuMail
{
    public enum FolderType
    {
        Inbox,
        Send,
        Trash,
        InboxSub
    }

    public class MailFolder
    {
        public string Name { get; }
        public string FullPath { get; }
        public FolderType Type { get; }

        public MailFolder(string name, string fullPath, FolderType type)
        {
            Name = name;
            FullPath = fullPath;
            Type = type;
        }

        public override string ToString() => Name;
    }

    public class FolderManager
    {
        public MailFolder Inbox { get; }
        public MailFolder Send { get; }
        public MailFolder Trash { get; }

        public List<MailFolder> InboxSubFolders { get; } = new List<MailFolder>();

        private readonly string baseDir;

        public FolderManager()
        {
            baseDir = Path.Combine(Application.StartupPath, "mbox");

            Inbox = new MailFolder("inbox", Path.Combine(baseDir, "inbox"), FolderType.Inbox);
            Send = new MailFolder("send", Path.Combine(baseDir, "send"), FolderType.Send);
            Trash = new MailFolder("trash", Path.Combine(baseDir, "trash"), FolderType.Trash);

            Directory.CreateDirectory(Inbox.FullPath);
            Directory.CreateDirectory(Send.FullPath);
            Directory.CreateDirectory(Trash.FullPath);

            LoadInboxSubFolders();
        }

        private void LoadInboxSubFolders()
        {
            InboxSubFolders.Clear();

            if (!Directory.Exists(Inbox.FullPath))
                return;

            foreach (var dir in Directory.GetDirectories(Inbox.FullPath))
            {
                string name = Path.GetFileName(dir);
                InboxSubFolders.Add(new MailFolder(name, dir, FolderType.InboxSub));
            }
        }

        public MailFolder GetFolder(FolderType type)
        {
            switch (type)
            {
                case FolderType.Inbox:
                    return Inbox;

                case FolderType.Send:
                    return Send;

                case FolderType.Trash:
                    return Trash;

                default:
                    return null;
            }
        }

        public MailFolder FindFolderByName(string name)
        {
            if (name == "inbox") return Inbox;
            if (name == "send") return Send;
            if (name == "trash") return Trash;

            return InboxSubFolders.FirstOrDefault(f => f.Name == name);
        }

        public string ResolveMailPath(Mail mail)
        {
            if (mail == null)
                return null;

            // ★ 新方式（MailFolder）優先
            if (mail.Folder != null)
                return Path.Combine(mail.Folder.FullPath, mail.mailName);

            // ★ 旧方式（string folder）も互換性のため残す
            if (!string.IsNullOrEmpty(mail.folder))
            {
                var folder = FindFolderByName(mail.folder);
                if (folder != null)
                    return Path.Combine(folder.FullPath, mail.mailName);
            }

            // ★ 最後の fallback（inbox 扱い）
            return Path.Combine(Inbox.FullPath, mail.mailName);
        }
    }
}
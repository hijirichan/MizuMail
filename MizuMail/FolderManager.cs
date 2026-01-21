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
        Draft,
        Spam,
        InboxSub
    }

    public class MailFolder
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string FullPath { get; private set; }
        public FolderType Type { get; }
        public List<MailFolder> SubFolders { get; set; } = new List<MailFolder>();

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
        public MailFolder Draft { get; }
        public MailFolder Spam { get; }

        public List<MailFolder> InboxSubFolders { get; } = new List<MailFolder>();

        private readonly string baseDir;

        public FolderManager()
        {
            baseDir = Path.Combine(Application.StartupPath, "mbox");

            Inbox = new MailFolder("Inbox", Path.Combine(baseDir, "inbox"), FolderType.Inbox)
            {
                DisplayName = "受信メール"
            };
            Spam = new MailFolder("Spam", Path.Combine(baseDir, "spam"), FolderType.Spam)
            {
                DisplayName = "迷惑メール"
            };
            Send = new MailFolder("Send", Path.Combine(baseDir, "send"), FolderType.Send)
            {
                DisplayName = "送信メール"
            };
            Draft = new MailFolder("Draft", Path.Combine(baseDir, "draft"), FolderType.Draft)
            {
                DisplayName = "下書き"
            };
            Trash = new MailFolder("Trash", Path.Combine(baseDir, "trash"), FolderType.Trash)
            {
                DisplayName = "ごみ箱"
            };

            Directory.CreateDirectory(Inbox.FullPath);
            Directory.CreateDirectory(Spam.FullPath);
            Directory.CreateDirectory(Send.FullPath);
            Directory.CreateDirectory(Trash.FullPath);
            Directory.CreateDirectory(Draft.FullPath);

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

                case FolderType.Spam:
                    return Spam;

                case FolderType.Send:
                    return Send;

                case FolderType.Draft:
                    return Draft;

                case FolderType.Trash:
                    return Trash;

                default:
                    return null;
            }
        }

        public MailFolder FindFolderByName(string name)
        {
            if (name == "inbox") return Inbox;
            if (name == "spam") return Spam;
            if (name == "send") return Send;
            if (name == "draft") return Draft;
            if (name == "trash") return Trash;

            return InboxSubFolders.FirstOrDefault(f => f.Name == name);
        }

        public MailFolder FindByPath(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath))
                return null;

            // 1. ルートフォルダ
            if (string.Equals(Inbox.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return Inbox;

            if (string.Equals(Spam.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return Spam;

            if (string.Equals(Send.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return Send;

            if (string.Equals(Draft.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return Draft;

            if (string.Equals(Trash.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return Trash;

            // 2. Inbox のサブフォルダ
            foreach (var sub in InboxSubFolders)
            {
                if (string.Equals(sub.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                    return sub;
            }

            return null;
        }

        public MailFolder GetFolderByType(string type)
        {
            switch (type)
            {
                case "Inbox":
                    return Inbox;
                case "Spam":
                    return Spam;
                case "Send":
                    return Send;
                case "Draft":
                    return Draft;
                case "Trash":
                    return Trash;
                default:
                    return Inbox;
            }
        }

        public MailFolder FindFolder(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            // ルート5フォルダから検索開始
            return FindFolderRecursive(Inbox, name)
                ?? FindFolderRecursive(Spam, name)
                ?? FindFolderRecursive(Send, name)
                ?? FindFolderRecursive(Draft, name)
                ?? FindFolderRecursive(Trash, name);
        }

        private MailFolder FindFolderRecursive(MailFolder folder, string name)
        {
            if (folder == null)
                return null;

            // 名前一致（大文字小文字無視）
            if (string.Equals(folder.Name, name, StringComparison.OrdinalIgnoreCase))
                return folder;

            // サブフォルダを再帰的に検索
            foreach (var sub in folder.SubFolders)
            {
                var found = FindFolderRecursive(sub, name);
                if (found != null)
                    return found;
            }

            return null;
        }

        public MailFolder CreateSubFolder(MailFolder parent, string name)
        {
            string newPath = Path.Combine(parent.FullPath, name);

            if (!Directory.Exists(newPath))
                Directory.CreateDirectory(newPath);

            var folder = new MailFolder(name, newPath, FolderType.InboxSub);

            parent.SubFolders.Add(folder);

            return folder;
        }

        public MailFolder GetOrCreateFolderByPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;

            var parts = path.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);

            MailFolder current = null;
            string root = parts[0].ToLower();

            switch (root)
            {
                case "inbox":
                    current = Inbox;
                    break;
                case "spam":
                    current = Spam;
                    break;
                case "send":
                    current = Send;
                    break;
                case "draft":
                    current = Draft;
                    break;
                case "trash":
                    current = Trash;
                    break;
                default:
                    return null;
            }

            for (int i = 1; i < parts.Length; i++)
            {
                string name = parts[i];

                var next = current.SubFolders
                    .FirstOrDefault(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

                if (next == null)
                {
                    string newPath = Path.Combine(current.FullPath, name);
                    Directory.CreateDirectory(newPath);

                    next = new MailFolder(name, newPath, FolderType.InboxSub);
                    current.SubFolders.Add(next);

                    // ★ UI 更新は FormMain 側で行う
                }

                current = next;
            }

            return current;
        }

    }
}
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
        public MailFolder Draft { get; }

        public List<MailFolder> InboxSubFolders { get; } = new List<MailFolder>();

        private readonly string baseDir;

        public FolderManager()
        {
            baseDir = Path.Combine(Application.StartupPath, "mbox");

            Inbox = new MailFolder("受信トレイ", Path.Combine(baseDir, "inbox"), FolderType.Inbox);
            Send = new MailFolder("送信トレイ", Path.Combine(baseDir, "send"), FolderType.Send);
            Trash = new MailFolder("ごみ箱", Path.Combine(baseDir, "trash"), FolderType.Trash);
            Draft = new MailFolder("下書き", Path.Combine(baseDir, "draft"), FolderType.Draft);

            Directory.CreateDirectory(Inbox.FullPath);
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

        public MailFolder FindByPath(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath))
                return null;

            // 1. ルートフォルダ
            if (string.Equals(Inbox.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return Inbox;

            if (string.Equals(Send.FullPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return Send;

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
    }
}
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace MizuMail
{
    public static class TagStorage
    {
        public static string GetJsonPath(Mail mail)
        {
            if (mail == null || string.IsNullOrEmpty(mail.mailPath))
                return null;

            // ★ Message-ID を安全なファイル名に変換
            string id = NormalizeId(mail.MessageId);
            if (string.IsNullOrEmpty(id))
                return null;

            // ★ タグ JSON はメールと同じフォルダに置く（現状維持）
            string folder = Path.GetDirectoryName(mail.mailPath);
            return Path.Combine(folder, id + ".json");
        }

        public static List<string> LoadTags(Mail mail)
        {
            string jsonPath = GetJsonPath(mail);
            if (jsonPath == null || !File.Exists(jsonPath))
                return new List<string>();

            try
            {
                string json = File.ReadAllText(jsonPath);
                var data = JsonConvert.DeserializeObject<TagData>(json);
                return data?.Tags ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }

        public static void SaveTags(Mail mail)
        {
            string jsonPath = GetJsonPath(mail);
            if (jsonPath == null)
                return;

            if (mail.Labels == null)
                return;

            var data = new TagData { Tags = mail.Labels };
            string json = JsonConvert.SerializeObject(data, Formatting.Indented);

            File.WriteAllText(jsonPath, json);
        }

        public static void DeleteTags(Mail mail)
        {
            string jsonPath = GetJsonPath(mail);
            if (jsonPath != null && File.Exists(jsonPath))
                File.Delete(jsonPath);
        }

        public static void MoveTags(string messageId, string oldMailPath, string newMailPath)
        {
            if (string.IsNullOrEmpty(messageId))
                return;

            string id = NormalizeId(messageId);

            string oldJson = Path.Combine(Path.GetDirectoryName(oldMailPath), id + ".json");
            string newJson = Path.Combine(Path.GetDirectoryName(newMailPath), id + ".json");

            if (File.Exists(oldJson))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(newJson));
                File.Move(oldJson, newJson);
            }
        }

        private static string NormalizeId(string id)
        {
            if (string.IsNullOrEmpty(id))
                return null;

            // ★ ファイル名に使えない文字を除去
            foreach (char c in Path.GetInvalidFileNameChars())
                id = id.Replace(c, '_');

            return id.Replace("<", "").Replace(">", "").Replace(":", "_");
        }

        public static bool Exists(Mail mail)
        {
            string jsonPath = GetJsonPath(mail);
            return File.Exists(jsonPath);
        }

        private class TagData
        {
            public List<string> Tags { get; set; }
        }
    }
}
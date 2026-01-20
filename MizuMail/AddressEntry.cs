using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace MizuMail
{
    public class AddressEntry
    {
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public string Note { get; set; }
    }

    public class AddressBook
    {
        public List<AddressEntry> Entries { get; set; } = new List<AddressEntry>();

        private static string GetPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "addressbook.json");
        }

        public static AddressBook LoadAddressBook()
        {
            string path = GetPath();

            if (!File.Exists(path))
                return new AddressBook();

            try
            {
                string json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<AddressBook>(json) ?? new AddressBook();
            }
            catch
            {
                return new AddressBook(); // 壊れていたら空で返す
            }
        }

        public static void SaveAddressBook(AddressBook book)
        {
            string path = GetPath();

            string json = JsonConvert.SerializeObject(book, Formatting.Indented);
            File.WriteAllText(path, json);
        }
    }
}
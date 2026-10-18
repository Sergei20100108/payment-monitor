using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace PaymentMonitor.Core
{
    /// <summary>Настройки приложения (settings.xml рядом с программой).</summary>
    [Serializable]
    public class AppSettings
    {
        public string Bank = "Не указан";
        public string DbFileName = "payments.db";
        public string WatchFolder = "";
        public int WatchIntervalSec = 60;
    }

    public static class SettingsStore
    {
        public static string SettingsPath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.xml"); }
        }

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var xs = new XmlSerializer(typeof(AppSettings));
                    using (var fs = File.OpenRead(SettingsPath))
                        return (AppSettings)xs.Deserialize(fs) ?? new AppSettings();
                }
            }
            catch
            {
                // повреждённый файл настроек — не повод падать, используем значения по умолчанию
            }
            return new AppSettings();
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                var xs = new XmlSerializer(typeof(AppSettings));
                using (var w = new StreamWriter(SettingsPath, false, Encoding.UTF8))
                    xs.Serialize(w, settings);
            }
            catch
            {
                // настройки не критичны для работы
            }
        }

        /// <summary>Полный путь к файлу базы данных (рядом с exe или в AppData, если туда нельзя писать).</summary>
        public static string ResolveDbPath(AppSettings settings)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            string path = Path.Combine(dir, settings.DbFileName ?? "payments.db");

            try
            {
                string probe = Path.Combine(dir, ".write-probe");
                File.WriteAllText(probe, "ok");
                File.Delete(probe);
                return path;
            }
            catch
            {
                string fallbackDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PaymentMonitor");
                Directory.CreateDirectory(fallbackDir);
                return Path.Combine(fallbackDir, settings.DbFileName ?? "payments.db");
            }
        }
    }
}

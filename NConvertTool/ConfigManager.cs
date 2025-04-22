using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NCreoConvertTool
{    public class ConfigManager
    {
        private static Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        private static AppSettingsSection appSettings = config.AppSettings;

        public static string GetAppSetting(string key)
        {
            return appSettings.Settings[key]?.Value;
        }

        public static void SetAppSetting(string key, string value)
        {
            appSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
    }
}

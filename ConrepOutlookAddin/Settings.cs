using System;
using System.IO;
using Newtonsoft.Json;

namespace ConrepOutlookAddin
{
    public class Settings
    {
        private Settings()
        {
            
        }

        public string LoginUrl { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public bool LogIncomingEmails { get; set; }
        public bool LogOutgoingEmails { get; set; }
        public string CompanyId { get; set; }
        public string[] OutlookAccounts { get; set; }
        public string PaneHeaderMessage { get; set; }
        public string PaneFooterMessage { get; set; }
        public string RightPaneName { get; set; }
        public int CalendarSyncInterval { get; set; }
        public bool ShowConrepPaneAutomatically { get; set; }

        #region Methods
        public void LoadSettings()
        {
            try
            {
                string settingsFilePath = GetSettingsFilePath();
                if (!File.Exists(settingsFilePath))
                    return;

                using (var textReader = new StreamReader(settingsFilePath))
                {
                    var settings = JsonConvert.DeserializeObject<Settings>(textReader.ReadToEnd());
                    _currentSettings = settings;
                }
            }
            catch(System.Exception ex)
            {
                Logger.Error(ex);
            }
        }

        public void SaveSettings()
        {
            try
            {
                string appDataPath = GetAppDataFolder();
                if (!Directory.Exists(appDataPath))
                    Directory.CreateDirectory(appDataPath);

                string settingsFilePath = GetSettingsFilePath();
                string settings = JsonConvert.SerializeObject(this);
                File.WriteAllText(settingsFilePath, settings);
            }
            catch(Exception ex)
            {
                Logger.Error(ex);
            }
        }

        private string GetSettingsFilePath()
        {
            string appDataPath = GetAppDataFolder();
            return Path.Combine(appDataPath, "settings.json");
        }

        public static string GetAppDataFolder()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appDataPath, "Conrep Outlook Addin");
        }
        #endregion

        private static Settings _currentSettings;
        public static Settings CurrentSettings
        {
            get
            {
                if(_currentSettings == null)
                    _currentSettings = new Settings();

                return _currentSettings;
            }
            set => _currentSettings = value;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ConrepOutlookAddin
{
    public partial class SettingsForm : Form
    {
        private ServerSettings _serverSettings;
        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            UpdateSettings();
            Settings.CurrentSettings.SaveSettings();
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            OnCancel();
        }

        private void UpdateSettings()
        {
            Settings.CurrentSettings.LoginUrl = txtLoginUrl.Text;
            Settings.CurrentSettings.UserName = txtUserName.Text;
            Settings.CurrentSettings.Password = txtPassword.Text;
            Settings.CurrentSettings.CompanyId = txtCompanyId.Text;
            Settings.CurrentSettings.LogIncomingEmails = chkLogIncomingEmails.Checked;
            Settings.CurrentSettings.LogOutgoingEmails = chkLogOutgoingEmails.Checked;

            if (_serverSettings != null)
            {
                Settings.CurrentSettings.CalendarSyncInterval = _serverSettings.CalendarSyncInterval;
                Settings.CurrentSettings.RightPaneName = _serverSettings.RightPaneName;
                Settings.CurrentSettings.PaneHeaderMessage = _serverSettings.PaneHeaderMessage;
                Settings.CurrentSettings.PaneFooterMessage = _serverSettings.PaneFooterMessage;
            }

            List<string> outlookAccounts = new List<string>();
            foreach (var item in cboAccounts.CheckBoxItems)
            {
                if (item.Checked)
                {
                    outlookAccounts.Add(item.Text);
                }
            }

            Settings.CurrentSettings.ShowConrepPaneAutomatically = chkShowPane.Checked;
            Settings.CurrentSettings.OutlookAccounts = outlookAccounts.ToArray();
            Globals.ThisAddIn.UpdateCalendarSyncInterval(Settings.CurrentSettings.CalendarSyncInterval);
        }

        private void OnCancel()
        {
            this.Close();
        }

        private void LoadSettings()
        {
            //login section
            txtLoginUrl.Text = Settings.CurrentSettings.LoginUrl;
            txtCompanyId.Text = Settings.CurrentSettings.CompanyId;
            txtUserName.Text = Settings.CurrentSettings.UserName;
            txtPassword.Text = Settings.CurrentSettings.Password;

            //log section
            chkLogIncomingEmails.Checked = Settings.CurrentSettings.LogIncomingEmails;
            chkLogOutgoingEmails.Checked = Settings.CurrentSettings.LogOutgoingEmails;

            //misc section
            chkShowPane.Checked = Settings.CurrentSettings.ShowConrepPaneAutomatically;

            //calendar sync section
            txtCalendarSyncInterval.Text = Settings.CurrentSettings.CalendarSyncInterval.ToString();
            if (Settings.CurrentSettings.OutlookAccounts != null)
            {
                LoadOutlookAccounts(Settings.CurrentSettings.OutlookAccounts);
            }
        }

        private void LoadOutlookAccounts(string[] accounts)
        {
            cboAccounts.Clear();
            foreach (var account in accounts)
            {
                if (!string.IsNullOrEmpty(account))
                {
                    cboAccounts.Items.Add(account);
                }
            }

            foreach (var checkboxItem in cboAccounts.CheckBoxItems)
            {
                checkboxItem.Checked = true;
                checkboxItem.Enabled = false;
            }
        }

        private void btnUpdateSettings_Click(object sender, EventArgs e)
        {
            var mailImporter = new ApiHandler();
            _serverSettings = mailImporter.GetServerSettings(txtUserName.Text, txtPassword.Text, txtCompanyId.Text);

            if (_serverSettings == null)
            {
                MessageBox.Show(@"Cannot get settings from the server. Please make sure your login url, username, password or company id is correct.");
                return;
            }

            chkLogOutgoingEmails.Checked = _serverSettings.LogOutgoingEmails;
            chkLogIncomingEmails.Checked = _serverSettings.LogIncomingEmails;
            txtCalendarSyncInterval.Text = _serverSettings.CalendarSyncInterval.ToString();
            LoadOutlookAccounts(_serverSettings.EmailAccounts.ToArray());

            UpdateSettings();
            MessageBox.Show(_serverSettings.SuccessMessage ?? "Updated successfully!");
        }
    }
}

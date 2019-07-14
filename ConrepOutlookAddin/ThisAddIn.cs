using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ConrepOutlookAddin.Enums;
using ConrepOutlookAddin.ExtensionMethods;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using log4net;
using log4net.Appender;
using log4net.Core;
using log4net.Layout;
using log4net.Repository.Hierarchy;

namespace ConrepOutlookAddin
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _conrepTaskPane;
        private Microsoft.Office.Core.IRibbonExtensibility _conrepRibbon;
        private List<Items> _outlookItems = new List<Items>();
        private Explorer _activeExplorer;
        private Inspectors _inspectors;
        private Timer _calendarSyncTimer;


        #region Properties
        #endregion

        #region Event Handlers
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            EnableTLS();
            ConfigLogger();
            RegisterEvents();
            LoadSettings();
            CreateConrepTaskPane();
            CreateCalendarSyncScheduler();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void OnItemSend(object item)
        {
            MailItem mailItem = item as MailItem;
            if (mailItem == null)
                return;

            if (AllowImportBackgroundProcess(mailItem, RequestMethod.SendEmail))
            {
                var taskPane = ConrepTaskPanes.GetTaskPane(_activeExplorer);
                ImportEmail(mailItem, RequestMethod.SendEmail, ApiInvokeMode.OutgoingEmail, taskPane);
            }
        }

        private void OnItemReceive(object item)
        {
            MailItem mailItem = item as MailItem;
            if (mailItem == null)
                return;

            if (AllowImportBackgroundProcess(mailItem, RequestMethod.ReceiveEmail))
            {
                var taskPane = ConrepTaskPanes.GetTaskPane(_activeExplorer);
                ImportEmail(mailItem, RequestMethod.ReceiveEmail, ApiInvokeMode.IncomingEmail, taskPane, true);
            }
        }

        private void OnItemSelectionChange()
        {
            var selectedEmail = _activeExplorer.GetSelectedEmail();
            var taskPane = ConrepTaskPanes.GetTaskPane(_activeExplorer);
            if (taskPane != null && taskPane.Visible)
            {
                CheckAndImportEmail(selectedEmail, taskPane);
            }
        }

        private void OnExplorerActive()
        {
            var taskPane = ConrepTaskPanes.GetTaskPane(_activeExplorer);
            if (taskPane != null)
            {
                ShowTaskPane(taskPane.Visible);
            }
        }

        #endregion

        #region Private Methods

        private void LoadSettings()
        {
            Settings.CurrentSettings.LoadSettings();
        }

        private void EnableTLS()
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        //This function get called when user selects an email in any folder or he wants to show the Conrep taskpane
        private void CheckAndImportEmail(MailItem mailItem, ConrepTaskPane taskPane)
        {
            if (mailItem == null)
                return;
            
            string outlookAccount = mailItem.GetAssociatedAccount();
            if (Settings.CurrentSettings.OutlookAccounts != null &&
                Settings.CurrentSettings.OutlookAccounts.Contains(outlookAccount))
            {
                var header = mailItem.GetRawHeader();
                RequestMethod requestMethod = string.IsNullOrEmpty(header)
                    ? RequestMethod.SendEmail
                    : RequestMethod.ReceiveEmail;

                ImportEmail(mailItem, requestMethod, ApiInvokeMode.RightPane, taskPane);
            }
            else
            {
                taskPane.DisplayManualImportText();
            }
        }

        public void ImportEmail(MailItem mailItem, RequestMethod requestMethod, ApiInvokeMode mode, ConrepTaskPane taskPane, bool background = false)
        {
            taskPane.ChangeStatus("Importing...");
            Task.Run(() =>
            {
                var mailImporter = new ApiHandler();
                var validateResult = mailImporter.ValidateEmail(mailItem, mode, requestMethod);

                if (validateResult == null)
                {
                    taskPane.ChangeStatus("Import failed: Cannot validate user.");
                    return;
                }

                if (validateResult.SendData)
                {
                    var importSettings = new MailImportSettings()
                    {
                        MailItem = mailItem,
                        SendHeaderOnly = false,
                        SendAttachments = validateResult.SendAttachments,
                        MaxAttachmentSize = validateResult.MaxAttachmentSize
                    };

                    mailImporter.ImportEmails(new List<MailImportSettings>() { importSettings },
                        requestMethod, validateResult.TransToken, mode);
                }

                if (taskPane.Visible && mode == ApiInvokeMode.RightPane)
                {
                    ShowEmailDetails(mailItem.GetMessageId(), validateResult.TransToken, requestMethod, taskPane, validateResult.PostUrl);
                }

                if (background && !string.IsNullOrEmpty(validateResult.MoveTo))
                {
                    try
                    {
                        string[] folderPaths = validateResult.MoveTo.Split(new [] {@"\"}, StringSplitOptions.RemoveEmptyEntries);
                        MAPIFolder folder = mailItem.SendUsingAccount.DeliveryStore.GetRootFolder();

                        for (int i = 0; i < folderPaths.Length; i++)
                        {
                            folder = folder.Folders[folderPaths[i]];
                        }
                        
                        mailItem.Move(folder);
                    }
                    catch(System.Exception ex)
                    {
                        Logger.Error(ex);
                    }
                }

                taskPane.ChangeStatus("Import completed.");
            });
        }

        private void ShowEmailDetails(string mailId, string transToken, RequestMethod requestMethod, ConrepTaskPane taskPane, string postUrl = "")
        {
            string baseUrl = string.IsNullOrEmpty(postUrl) ? Settings.CurrentSettings.LoginUrl : postUrl;
            string requestUrl = $"{baseUrl.EnsureStartsWithHttps()}/conrep/outlook/web/email_details.php?" +
                                $"MessageId={WebUtility.UrlEncode(mailId)}" +
                                $"&TransToken={WebUtility.UrlEncode(transToken)}" +
                                $"&EmailType={requestMethod}";

            taskPane.OpenUrl(requestUrl);
        }

        private void CreateConrepTaskPane()
        {
            ConrepTaskPanes.Add(_activeExplorer, true);
        }

        private void RegisterEvents()
        {
            ConrepInspectors.Handle();

            _activeExplorer = this.Application.ActiveExplorer();
            ((ExplorerEvents_Event)_activeExplorer).Activate += OnExplorerActive;
            _activeExplorer.SelectionChange += OnItemSelectionChange;

            foreach (Account account in Application.Session.Accounts)
            {
                try
                {
                    var store = account.DeliveryStore;
                    var inboxFolder = store.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                    var inboxItems = inboxFolder.Items;
                    inboxItems.ItemAdd += OnItemReceive;
                    _outlookItems.Add(inboxItems);

                    var sentFolder = store.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
                    var sentItems = sentFolder.Items;
                    sentItems.ItemAdd += OnItemSend;
                    _outlookItems.Add(sentItems);
                }
                catch(System.Exception ex)
                {
                    Logger.Error(ex);
                }
            }
        }

        private bool AllowImportBackgroundProcess(MailItem mailItem, RequestMethod requestMethod)
        {
            string outLookAccount = mailItem.SendUsingAccount.SmtpAddress;
            return (Settings.CurrentSettings.OutlookAccounts != null &&
                    Settings.CurrentSettings.OutlookAccounts.Contains(outLookAccount))
                   && (requestMethod == RequestMethod.SendEmail
                       ? Settings.CurrentSettings.LogOutgoingEmails
                       : Settings.CurrentSettings.LogIncomingEmails);
        }

        private bool IsValidEmailAddress(string address)
        {
            const string pattern = @"^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?$";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(address);

            return match.Success;
        }

        private void SyncCalendars()
        {
            var httpHandler = new ApiHandler();
            var calendarItems = httpHandler.GetCalendarItems();
            MAPIFolder calendarFolder = this.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            foreach (var calendarItem in calendarItems)
            {
                AppointmentItem appointment = FindAppointment(calendarFolder, calendarItem.CalendarId);
                if (appointment == null)
                {
                    appointment = this.Application.CreateItem(OlItemType.olAppointmentItem);
                    var calenderIdProperty =
                        appointment.UserProperties.Add("ConrepCalendarId", OlUserPropertyType.olInteger, true, 1);
                    calenderIdProperty.Value = calendarItem.CalendarId;
                }

                appointment.Start = calendarItem.StartTime ?? DateTime.Now;
                appointment.End = calendarItem.EndTime ?? DateTime.Now.AddHours(1);
                appointment.Location = calendarItem.Location;
                appointment.Subject = calendarItem.Subject;
                appointment.Body = calendarItem.Description;
                appointment.AllDayEvent = calendarItem.AllDayEvent;

                string[] attendees = string.IsNullOrEmpty(calendarItem.Attendees)
                    ? new string[0]
                    : calendarItem.Attendees.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

                bool needSend = false;
                foreach (string attendee in attendees)
                {
                    if (IsValidEmailAddress(attendee))
                    {
                        var recipient = appointment.Recipients.Add(attendee);
                        recipient.Type = (int)OlMeetingRecipientType.olRequired;
                        needSend = true;
                    }
                }

                appointment.Recipients.ResolveAll();
                appointment.Save();

                if (needSend)
                {
                    appointment.Send();
                }
            }
        }

        private AppointmentItem FindAppointment(MAPIFolder folder, int calendarId)
        {
            var appointmentItems = folder.Items.Restrict(
                "@SQL=(http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/ConrepCalendarId IS NOT NULL)");

            foreach (AppointmentItem appointment in appointmentItems)
            {
                var calendarIdProperty = appointment.UserProperties.Find("ConrepCalendarId");
                if (calendarIdProperty != null)
                {
                    int value = (int)calendarIdProperty.Value;
                    if (value == calendarId)
                        return appointment;
                }
            }

            return null;
        }

        private void CreateCalendarSyncScheduler()
        {
            int interval = Settings.CurrentSettings.CalendarSyncInterval;
            if (interval <= 0)
                interval = 30; //default is 30 mins

            _calendarSyncTimer = new Timer();
            _calendarSyncTimer.Interval = interval * 60 * 1000;
            _calendarSyncTimer.Tick += OnCalendarSyncHappens;
            _calendarSyncTimer.Start();
        }

        private void OnCalendarSyncHappens(object sender, EventArgs e)
        {
            var taskPane = ConrepTaskPanes.GetTaskPane(_activeExplorer);
            SyncCalendarsAsync(taskPane);
        }

        private void ConfigLogger()
        {
            Hierarchy hierarchy = (Hierarchy)LogManager.GetRepository();

            PatternLayout patternLayout = new PatternLayout();
            patternLayout.ConversionPattern = "%newline%date [%thread] %-5level %logger - %message%newline";
            patternLayout.ActivateOptions();

            RollingFileAppender roller = new RollingFileAppender();
            roller.LockingModel = new FileAppender.MinimalLock();
            roller.AppendToFile = true;
            roller.File = Path.Combine(Settings.GetAppDataFolder(), "log.txt");
            roller.Layout = patternLayout;
            roller.MaxSizeRollBackups = 5;
            roller.MaximumFileSize = "10MB";
            roller.RollingStyle = RollingFileAppender.RollingMode.Size;
            roller.StaticLogFileName = true;
            roller.ActivateOptions();
            hierarchy.Root.AddAppender(roller);

            hierarchy.Root.Level = Level.Info;
            hierarchy.Configured = true;
        }
        #endregion

        #region Public Methods

        public void BulkImport(ApiInvokeMode mode, List<MailItem> mailItems, ConrepTaskPane taskPane, bool sendHeaderOnly = false)
        {
            if (string.IsNullOrEmpty(Settings.CurrentSettings.UserName) ||
                string.IsNullOrEmpty(Settings.CurrentSettings.Password))
            {
                taskPane.ChangeStatus("Can not bulk import. Invalid username or password.");
                return;
            }

            if (mailItems.Count == 0)
            {
                taskPane.ChangeStatus("No items to import.");
                return;
            }

            taskPane.ChangeStatus("Importing...");
            Task.Run(() =>
            {
                var mailImporter = new ApiHandler();
                var importSettings = new List<MailImportSettings>();

                string transToken = string.Empty;
                foreach (var mailItem in mailItems)
                {
                    var validateResult = mailImporter.ValidateEmail(mailItem, mode, RequestMethod.ReceiveEmail);
                    if (validateResult != null && validateResult.SendData)
                    {
                        importSettings.Add(new MailImportSettings()
                        {
                            MailItem = mailItem,
                            SendHeaderOnly = sendHeaderOnly,
                            SendAttachments = validateResult.SendAttachments,
                            MaxAttachmentSize = validateResult.MaxAttachmentSize
                        });

                        transToken = validateResult.TransToken;
                    }
                }

                if (importSettings.Count > 0)
                {
                    mailImporter.ImportEmails(importSettings, RequestMethod.ReceiveEmail, transToken, mode);
                    taskPane.ChangeStatus("Import completed.");
                }
                else
                {
                    taskPane.ChangeStatus("No items to import.");
                }
            });
        }

        public void ImportSelectedEmailWhenShowingTaskPane(MailItem mailItem, ConrepTaskPane taskPane)
        {
            CheckAndImportEmail(mailItem, taskPane);
        }

        public void SyncCalendarsAsync(ConrepTaskPane taskPane)
        {
            if (string.IsNullOrEmpty(Settings.CurrentSettings.UserName) ||
                string.IsNullOrEmpty(Settings.CurrentSettings.Password))
            {
                taskPane.ChangeStatus("Can not sync calendar. Invalid username or password.");
                return;
            }

            taskPane.ChangeStatus("Syncing calendar...");
            Task.Run(() =>
            {
                SyncCalendars();
            }).ContinueWith(o => taskPane.ChangeStatus("Sync calendar completed."));
        }

        public void ShowTaskPane(bool show)
        {
            ((AddinRibbon)_conrepRibbon).ShowTaskPane(show);
        }

        public void UpdateCalendarSyncInterval(int minutes)
        {
            if (minutes <= 0)
                return;

            _calendarSyncTimer.Stop();
            _calendarSyncTimer.Interval = minutes * 60 * 1000;
            _calendarSyncTimer.Start();
        }

        #endregion

        #region Ribbon
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _conrepRibbon = new AddinRibbon();
            return _conrepRibbon;
        }
        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

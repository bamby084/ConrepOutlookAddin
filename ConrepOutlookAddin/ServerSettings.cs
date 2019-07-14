using System.Collections.Generic;

namespace ConrepOutlookAddin
{
    /// <summary>
    /// Settings returned from Conrep server
    /// </summary>
    public class ServerSettings
    {
        public List<string> EmailAccounts { get; set; }
        public bool LogIncomingEmails { get; set; }
        public bool LogOutgoingEmails { get; set; }
        public string PaneHeaderMessage { get; set; }
        public string PaneFooterMessage { get; set; }
        public string RightPaneName { get; set; }
        public int CalendarSyncInterval { get; set; }
        public string SuccessMessage { get; set; }
    }
}

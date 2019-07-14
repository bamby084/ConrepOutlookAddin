
namespace ConrepOutlookAddin
{
    public class ValidateEmailResult
    {
        public string UserDisplayName { get; set; }
        public string RecordFound { get; set; }
        public int RecordId { get; set; }
        public int RecordApplication { get; set; }
        public string TransToken { get; set; }
        public bool SendData { get; set; }
        public bool SendAttachments { get; set; }
        public int MaxAttachmentSize { get; set; }
        public string PostUrl { get; set; }
        public string EmailsAccounts { get; set; }
        public bool LogIncomingEmails { get; set; }
        public bool LogOutgoingEmails { get; set; }
        public string PaneHeaderMessage { get; set; }
        public string PaneFooterMessage { get; set; }
        public string RightPaneName { get; set; }
        public int CalendarSyncInterval { get; set; }
        public string SuccessMessage { get; set; }
        public string MoveTo { get; set; }
    }
}

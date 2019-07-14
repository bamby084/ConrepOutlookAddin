using Microsoft.Office.Interop.Outlook;

namespace ConrepOutlookAddin
{
    public class MailImportSettings
    {
        public MailItem MailItem { get; set; }
        public bool SendAttachments { get; set; }
        public int MaxAttachmentSize { get; set; }
        public bool SendHeaderOnly { get; set; }
    }
}

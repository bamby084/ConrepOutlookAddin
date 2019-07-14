using System;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using stdole;

namespace ConrepOutlookAddin.ExtensionMethods
{
    public static class AttachmentExtensions
    {
        private const string PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102";

        public static bool IsMailAttachment(this Attachment attachment)
        {
            var flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");
            if (flags != 4 && attachment.Type != OlAttachmentType.olOLE)
            {
                return true;
            }

            return false;
        }

        public static byte[] GetContent(this Attachment attachment)
        {
            try
            {
                string appDataPath = Settings.GetAppDataFolder();
                string tempFile = Path.Combine(appDataPath, $"{Guid.NewGuid()}.tmp");
                attachment.SaveAsFile(tempFile);

                byte[] data = File.ReadAllBytes(tempFile);
                File.Delete(tempFile);

                return data;
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex);
                return null;
            }
        }
    }
}

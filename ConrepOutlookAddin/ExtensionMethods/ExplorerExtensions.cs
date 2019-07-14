using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace ConrepOutlookAddin.ExtensionMethods
{
    public static class ExplorerExtensions
    {
        public static IEnumerable<MailItem> GetSelectedEmails(this Explorer explorer)
        {
            foreach (MailItem email in explorer.Selection)
            {
                yield return email;
            }
        }

        public static MailItem GetSelectedEmail(this Explorer explorer)
        {
            var selectedEmails = explorer.GetSelectedEmails().ToList();
            if (selectedEmails.Count == 0)
                return null;

            return selectedEmails[0];
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;

namespace ConrepOutlookAddin.ExtensionMethods
{
    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
            "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string GetSender(this MailItem mailItem)
        {
            if (mailItem.SenderEmailType.ToLower() == "ex")
            {
                var recipient = Globals.ThisAddIn.Application.GetNamespace("MAPI").CreateRecipient(mailItem.SenderEmailAddress);
                var sender = recipient.AddressEntry.GetExchangeUser();

                return sender.PrimarySmtpAddress;
            }

            return mailItem.SenderEmailAddress;
        }

        public static string GetRawHeader(this MailItem mailItem)
        {
            return mailItem.PropertyAccessor.GetProperty(TransportMessageHeadersSchema);
        }

        public static List<MailItemHeader> GetHeaders(this MailItem mailItem)
        {
            var headerLookup = mailItem.GetHeaderLookup();
            var headers = new List<MailItemHeader>();

            foreach (var lookup in headerLookup)
            {
                foreach (var value in lookup)
                {
                    headers.Add(new MailItemHeader()
                    {
                        Key = lookup.Key,
                        Value = string.IsNullOrEmpty(value) ? value : value.Replace("\r\n", "")
                    });
                }
            }

            return headers;
        }

        public static string[] GetHeader(this MailItem mailItem, string name)
        {
            var headers = mailItem.GetHeaderLookup();
            if (headers.Contains(name))
                return headers[name].ToArray();
            
            return new string[0];
        }

        public static string GetMessageId(this MailItem mailItem)
        {
            string[] messageIdArray = mailItem.GetHeader("Message-ID");
            if (messageIdArray.Length == 1)
                return messageIdArray[0].Replace("\r\n", "").Replace("<", "").Replace(">", "");

            return string.Empty;
        }

        public static string GetAssociatedAccount(this MailItem mailItem)
        {
            return mailItem.SendUsingAccount.SmtpAddress;
        }

        public static string[] GetMailItemRecipientsAddress(this MailItem mailItem)
        {
            string[] recipients = new string[4];
            foreach (Recipient recipient in mailItem.Recipients)
            {
                recipients[recipient.Type] += GetRecipientEmailAddress(recipient) + ";";
            }

            return recipients;
        }

        private static ILookup<string, string> GetHeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.GetRawHeader();

            //since sent items are created at local, so we need to create header by its common properties
            if (string.IsNullOrEmpty(headerString))
            {
                
                Dictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("Message-ID", mailItem.EntryID);
                headers.Add("From", mailItem.SendUsingAccount.SmtpAddress);
                headers.Add("Subject", mailItem.Subject);
                headers.Add("Date", mailItem.SentOn.ToString());

                string[] recipients = GetMailItemRecipientsAddress(mailItem);
                headers.Add("To", recipients[(int)OlMailRecipientType.olTo]);
                headers.Add("CC", recipients[(int)OlMailRecipientType.olCC]);
                headers.Add("BCC", recipients[(int)OlMailRecipientType.olBCC]);

                return headers.ToLookup(item => item.Key, item => item.Value);
            }

            var headerMatches = Regex.Matches
                (headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();

            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value,
                StringComparer.InvariantCultureIgnoreCase);
        }

        private static string GetRecipientEmailAddress(Recipient recipient)
        {
            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            string smtpAddress = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS).ToString();

            return smtpAddress;
        }
    }
}

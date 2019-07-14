using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Xml.Serialization;
using ConrepOutlookAddin.Enums;
using ConrepOutlookAddin.ExtensionMethods;
using ConrepOutlookAddin.Win32;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Exception = System.Exception;

namespace ConrepOutlookAddin
{
    public class ApiHandler
    {
        private static HttpClient _httpClient = new HttpClient();

        public ValidateEmailResult ValidateEmail(MailItem mailItem, ApiInvokeMode mode, RequestMethod requestMethod)
        {
            try
            {
                string requestUrl = Settings.CurrentSettings.LoginUrl.EnsureStartsWithHttps() +
                                    "/conrep/outlook/web/email_requests.php?RequestMethod=ValidateUserEmail" +
                                    $"&UserName={WebUtility.UrlEncode(Settings.CurrentSettings.UserName)}" +
                                    $"&Password={WebUtility.UrlEncode(Settings.CurrentSettings.Password)}" +
                                    $"&From={WebUtility.UrlEncode(mailItem.GetSender())}" +
                                    $"&MailId={WebUtility.UrlEncode(mailItem.GetMessageId())}" +
                                    $"&OutlookAccount={WebUtility.UrlEncode(mailItem.GetAssociatedAccount())}" +
                                    $"&CompanyId={WebUtility.UrlEncode(Settings.CurrentSettings.CompanyId)}" +
                                    $"&Mode={mode}" +
                                    $"&EmailType={requestMethod}" +
                                    $"&MAC={Network.GetMACAdrress()}";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUrl);
                request.Method = "POST";
                using (var stream = request.GetRequestStream())
                {
                    var mailItemInfoCollection = new MailItemInfoCollection();
                    string[] recipients = mailItem.GetMailItemRecipientsAddress();

                    mailItemInfoCollection.MailItems.Add(new MailItemInfo()
                    {
                        Headers = mailItem.GetHeaders(),
                        From = mailItem.GetSender(),
                        To = recipients[(int)OlMailRecipientType.olTo],
                        CC = recipients[(int)OlMailRecipientType.olCC],
                        BCC = recipients[(int)OlMailRecipientType.olBCC],
                        Subject = mailItem.Subject
                    });

                    string postDataJson = JsonConvert.SerializeObject(mailItemInfoCollection);
                    byte[] postData = Encoding.ASCII.GetBytes(postDataJson);

                    stream.Write(postData, 0, postData.Length);
                }

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                return GetValidateEmailResultFromStream(response.GetResponseStream());
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                return null;
            }
        }

        public ServerSettings GetServerSettings(string userName, string password, string companyId)
        {
            try
            {
                string requestUrl = Settings.CurrentSettings.LoginUrl.EnsureStartsWithHttps() +
                                    "/conrep/outlook/web/email_requests.php?RequestMethod=ValidateUserEmail" +
                                    $"&UserName={WebUtility.UrlEncode(userName)}" +
                                    $"&Password={WebUtility.UrlEncode(password)}" +
                                    $"&CompanyId={WebUtility.UrlEncode(companyId)}" +
                                    $"&Mode={ApiInvokeMode.Settings}";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUrl);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                var validateResult = GetValidateEmailResultFromStream(response.GetResponseStream());

                return new ServerSettings()
                {
                    EmailAccounts = validateResult.EmailsAccounts
                        .Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries).ToList(),
                    LogIncomingEmails = validateResult.LogIncomingEmails,
                    LogOutgoingEmails = validateResult.LogOutgoingEmails,
                    CalendarSyncInterval = validateResult.CalendarSyncInterval,
                    RightPaneName = validateResult.RightPaneName,
                    PaneHeaderMessage = validateResult.PaneHeaderMessage,
                    PaneFooterMessage = validateResult.PaneFooterMessage,
                    SuccessMessage = validateResult.SuccessMessage
                };
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                return null;
            }
        }

        public bool ImportEmails(IList<MailImportSettings> settings, RequestMethod requestMethod, string token, ApiInvokeMode mode)
        {
            try
            {
                string requestUrl = Settings.CurrentSettings.LoginUrl.EnsureStartsWithHttps() +
                                    "/conrep/outlook/web/email_requests.php?" +
                                    $"RequestMethod={requestMethod}" +
                                    $"&TransToken={WebUtility.UrlEncode(token)}" +
                                    $"&Mode={mode}";

                var formData = new MultipartFormDataContent();
                var mailItemInfoCollection = new MailItemInfoCollection();

                foreach (var importSettings in settings)
                {
                    var mailItemInfo = new MailItemInfo();
                    var mailItem = importSettings.MailItem;

                    mailItemInfo.Headers = mailItem.GetHeaders();
                    mailItemInfo.MailItemId = mailItem.GetMessageId();

                    if (!importSettings.SendHeaderOnly)
                    {
                        mailItemInfo.HtmlBody = mailItem.HTMLBody;
                        if (importSettings.SendAttachments)
                        {
                            foreach (Attachment attachment in mailItem.Attachments)
                            {
                                if (attachment.IsMailAttachment() &&
                                    attachment.Size <= importSettings.MaxAttachmentSize)
                                {
                                    byte[] attachmentData = attachment.GetContent();
                                    if (attachmentData != null)
                                    {
                                        formData.Add(new ByteArrayContent(attachmentData), attachment.FileName,
                                            $"{attachment.FileName}::{mailItemInfo.MailItemId}");
                                    }
                                }
                            }
                        }
                    }

                    mailItemInfoCollection.MailItems.Add(mailItemInfo);
                }

                var data = JsonConvert.SerializeObject(mailItemInfoCollection);
                formData.Add(new StringContent(data), "emails");

                var response = _httpClient.PostAsync(requestUrl, formData).Result;
                string responseMessage = response.Content.ReadAsStringAsync().Result;

                return response.StatusCode == HttpStatusCode.OK;
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                return false;
            }
        }

        public List<CalendarItem> GetCalendarItems()
        {
            try
            {
                string token = GetTransToken();
                var requestUrl = Settings.CurrentSettings.LoginUrl.EnsureStartsWithHttps() +
                                 "/conrep/outlook/web/email_requests.php?" +
                                 "RequestMethod=CalendarSync" +
                                 $"&TransToken={WebUtility.UrlEncode(token)}";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUrl);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                using (StreamReader streamReader = new StreamReader(response.GetResponseStream()))
                {
                    string result = streamReader.ReadToEnd();
                    List<CalendarItem> calendars = JsonConvert.DeserializeObject<List<CalendarItem>>(result);

                    return calendars;
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                return new List<CalendarItem>();
            }
        }

        private string GetTransToken()
        {
            try
            {
                string requestUrl = Settings.CurrentSettings.LoginUrl.EnsureStartsWithHttps() +
                                    "/conrep/outlook/web/email_requests.php?RequestMethod=ValidateUserEmail" +
                                    $"&UserName={WebUtility.UrlEncode(Settings.CurrentSettings.UserName)}" +
                                    $"&Password={WebUtility.UrlEncode(Settings.CurrentSettings.Password)}" +
                                    $"&CompanyId={WebUtility.UrlEncode(Settings.CurrentSettings.CompanyId)}";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUrl);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                var validateResult = GetValidateEmailResultFromStream(response.GetResponseStream());

                if (validateResult != null)
                {
                    return validateResult.TransToken;
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                return string.Empty;
            }
        }

        private ValidateEmailResult GetValidateEmailResultFromStream(Stream stream)
        {
            try
            {
                using (var streamReader = new StreamReader(stream))
                {
                    var xml = streamReader.ReadToEnd();
                    XmlRootAttribute rootAttribute = new XmlRootAttribute("source");
                    rootAttribute.IsNullable = true;
                    var xmlSerializer = new XmlSerializer(typeof(ValidateEmailResult), rootAttribute);

                    using (var stringReader = new StringReader(xml))
                    {
                        var result = (ValidateEmailResult)xmlSerializer.Deserialize(stringReader);
                        result.MaxAttachmentSize = result.MaxAttachmentSize * 1024 * 1024;

                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                return null;
            }
        }
    }
}

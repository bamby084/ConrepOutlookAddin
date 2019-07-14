using System.Collections.Generic;
using Newtonsoft.Json;

namespace ConrepOutlookAddin
{
    public class MailItemInfo
    {
        public string MailItemId { get; set; }

        [JsonConverter(typeof(MailItemHeaderConverter))]
        public List<MailItemHeader> Headers { get; set; }

        public string HtmlBody { get; set; }

        public string From { get; set; }

        public string To { get; set; }

        public string CC { get; set; }

        public string BCC { get; set; }

        public string Subject { get; set; }
    }

    public class MailItemInfoCollection
    {
        public MailItemInfoCollection()
        {
            MailItems = new List<MailItemInfo>();
        }

        public IList<MailItemInfo> MailItems { get; set; }
    }
}

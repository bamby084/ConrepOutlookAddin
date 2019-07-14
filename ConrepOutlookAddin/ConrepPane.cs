using System;
using System.Windows.Forms;

namespace ConrepOutlookAddin
{
    public partial class ConrepPane : UserControl
    {
        private delegate void ChangeStatusTextDelegate(string status);
        public event EventHandler ProcessButtonClick;

        public ConrepPane()
        {
            InitializeComponent();
            lblStatus.Text = "";
        }

        private void OnDocumentClick(object sender, HtmlElementEventArgs e)
        {
            if (WebBrowser.Document.ActiveElement.TagName == "BUTTON"
                && WebBrowser.Document.ActiveElement.Id == "processButton")
            {
                ProcessButtonClick?.Invoke(sender, new EventArgs());
            }
        }

        public void OpenUrl(string url)
        {
            WebBrowser.Navigate(url);
        }

        public void ChangeStatusText(string status)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new ChangeStatusTextDelegate(ChangeStatusText), status);
            }
            else
            {
                lblStatus.Text = status;
            }

        }

        public void DisplayManualImportText()
        {
            WebBrowser.Navigate("about:blank");
            if (WebBrowser.Document != null)
            {
                WebBrowser.Document.Write(string.Empty);
                WebBrowser.Document.Click -= OnDocumentClick;
                WebBrowser.Document.Click += OnDocumentClick;
            }

            string header = Settings.CurrentSettings.PaneHeaderMessage ??
                            "Email account associated with this email is not listed in your settings, to import email contents. Click button below to process the email contents.";
            string footer = Settings.CurrentSettings.PaneFooterMessage ??
                            "You can close this pane if you do not want to see this message.";

            string htmlContent = "<html><body>" +
                                 $"<div>{header}</div>" +
                                 "<div style='margin: 10 0px;'><button id='processButton' type='button'>Process Email Contents</button></div>" +
                                 $"<div>{footer}</div>" +
                                 "</body></html>";

            WebBrowser.DocumentText = htmlContent;
            ChangeStatusText(string.Empty);
        }
    }
}

using System;
using Microsoft.Office.Tools;

namespace ConrepOutlookAddin
{
    public class ConrepTaskPane
    {
        public event EventHandler VisibleChanged;
        public event EventHandler ProcessEmailManually;

        public CustomTaskPane TaskPane { get; set; }

        public ConrepPane ConrepPane { get; }

        public string Title => Settings.CurrentSettings.RightPaneName ?? "Conrep Pane";

        public bool Visible
        {
            get => TaskPane.Visible;
            set => TaskPane.Visible = value;
        }

        public object Parent { get; }

        public ConrepTaskPane(object parent)
        {
            ConrepPane = new ConrepPane();
            ConrepPane.ProcessButtonClick += (sender, e) =>
            {
                ProcessEmailManually?.Invoke(this, new EventArgs());
            };
            Parent = parent;
        }

        public void ChangeStatus(string status)
        {
            ConrepPane.ChangeStatusText(status);
        }

        public void DisplayManualImportText()
        {
            ConrepPane.DisplayManualImportText();
        }

        public void OnVisibleChanged()
        {
            VisibleChanged?.Invoke(this, new EventArgs());
        }

        public void OpenUrl(string url)
        {
            ConrepPane.OpenUrl(url);
        }
    }

}

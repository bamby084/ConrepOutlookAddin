using Microsoft.Office.Interop.Outlook;

namespace ConrepOutlookAddin
{
    public class ReadingInspector
    {
        private readonly Inspector _inspector;
        
        public ReadingInspector(Inspector inspector)
        {
            _inspector = inspector;
        }

        public void Handle()
        {
            var mailItem = _inspector.CurrentItem as MailItem;
            if (mailItem == null)
                return;

            //not in read mode?
            if (mailItem.Sent == false)
                return;

            ((InspectorEvents_Event)_inspector).Activate += OnActivate;
            ((InspectorEvents_Event)_inspector).Close += OnClose;

            var taskPane = ConrepTaskPanes.Add(_inspector, Settings.CurrentSettings.ShowConrepPaneAutomatically);
            taskPane.VisibleChanged += (sender, e) =>
            {
                Globals.ThisAddIn.ShowTaskPane(taskPane.Visible);
            };

            Globals.ThisAddIn.ImportSelectedEmailWhenShowingTaskPane(mailItem, taskPane);
        }

        private void OnClose()
        {
            ((InspectorEvents_Event)_inspector).Close -= OnClose;
            ((InspectorEvents_Event)_inspector).Activate -= OnActivate;

            var taskPane = ConrepTaskPanes.GetTaskPane(_inspector);
            if (taskPane != null)
            {
                ConrepTaskPanes.Remove(taskPane);
            }
        }

        private void OnActivate()
        {
            var taskPane = ConrepTaskPanes.GetTaskPane(_inspector);
            if (taskPane != null)
            {
                Globals.ThisAddIn.ShowTaskPane(taskPane.Visible);
            }
        }
    }
}

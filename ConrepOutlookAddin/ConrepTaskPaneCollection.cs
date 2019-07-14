using System;
using System.Collections.Generic;
using System.Linq;
using ConrepOutlookAddin.Enums;
using ConrepOutlookAddin.ExtensionMethods;
using Microsoft.Office.Interop.Outlook;
using Exception = Microsoft.Office.Interop.Outlook.Exception;

namespace ConrepOutlookAddin
{
    public static class ConrepTaskPanes
    {
        private static readonly List<ConrepTaskPane> _taskPanes = new List<ConrepTaskPane>();

        public static ConrepTaskPane Add(object window, bool visible)
        {
            var taskPane = new ConrepTaskPane(window);

            var customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(taskPane.ConrepPane, taskPane.Title, taskPane.Parent);
            customTaskPane.Width = 500;
            customTaskPane.Visible = visible;
            customTaskPane.VisibleChanged += (sender, e) =>
            {
                taskPane.OnVisibleChanged();
            };
            taskPane.TaskPane = customTaskPane;
            taskPane.ProcessEmailManually += OnProcessEmailManually;
            _taskPanes.Add(taskPane);

            return taskPane;
        }

        //Function for importing email when user hasnot setup email account in settings yet but he wants to process anyway
        private static void OnProcessEmailManually(object sender, System.EventArgs e)
        {
            var taskPane = sender as ConrepTaskPane;
            var inspector = taskPane.Parent as Inspector;
            var explorer = taskPane.Parent as Explorer;

            try
            {
                if (inspector != null)
                {
                    var mailItem = inspector.CurrentItem as MailItem;
                    Globals.ThisAddIn.ImportEmail(mailItem, RequestMethod.ReceiveEmail, ApiInvokeMode.RightPane, taskPane);
                }
                else if (explorer != null)
                {
                    var mailItem = explorer.GetSelectedEmail();
                    Globals.ThisAddIn.ImportEmail(mailItem, RequestMethod.ReceiveEmail, ApiInvokeMode.RightPane, taskPane);
                }
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex);
            }
        }

        public static void Remove(ConrepTaskPane taskPane)
        {
            Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane.TaskPane);
            _taskPanes.Remove(taskPane);
        }

        public static ConrepTaskPane GetTaskPane(object window)
        {
            return _taskPanes.FirstOrDefault(t => t.Parent == window);
        }
    }
}

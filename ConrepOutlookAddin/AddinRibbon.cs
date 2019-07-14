using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ConrepOutlookAddin.Enums;
using ConrepOutlookAddin.ExtensionMethods;
using ConrepOutlookAddin.Properties;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new AddinRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ConrepOutlookAddin
{
    [ComVisible(true)]
    public class AddinRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private bool _showTaskPane = true;

        #region ctor
        public AddinRibbon()
        {
        }
        #endregion

        #region Custom Code
        public void ShowSettingsPage(Office.IRibbonControl control)
        {
            try
            {
                var settingsForm = new SettingsForm();
                settingsForm.ShowDialog();
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex);
            }
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            Size imageSize = new Size(32, 32);

            switch (control.Id)
            {
                case "BulkImportButton":
                {
                    return new Bitmap(Resources.conrep_bulk_import, imageSize);
                }
                case "CalendarSyncButton":
                {
                    return new Bitmap(Resources.conrep_calendar_sync, imageSize);
                }
                case "SettingsButton":
                {
                    return new Bitmap(Resources.conrep_settings, imageSize);
                }
                case "ShowHideConrepPaneButton":
                {
                    return new Bitmap(Resources.conrep_logo, imageSize);
                }
            }

            return null;
        }

        public void BulkImport(Office.IRibbonControl control)
        {
            BulkImport(control, ApiInvokeMode.BulkAll);
        }

        public void BulkImportHeaderOnly(Office.IRibbonControl control)
        {
            BulkImport(control, ApiInvokeMode.BulkHeaders, true);
        }

        public bool ConrepPaneGetPressed(Office.IRibbonControl control)
        {
            return _showTaskPane;
        }

        public void OnConrepPaneToggleButtonClick(Office.IRibbonControl control, bool isPressed)
        {
            var inspector = control.Context as Inspector;
            var explorer = control.Context as Explorer;

            //we are in inspector?
            if (inspector != null)
            {
                var taskPane = ConrepTaskPanes.GetTaskPane(inspector);
                if (taskPane != null)
                    taskPane.Visible = isPressed;

                if (isPressed)
                {
                    var mailItem = inspector.CurrentItem as MailItem;
                    Globals.ThisAddIn.ImportSelectedEmailWhenShowingTaskPane(mailItem, taskPane);
                }
            }
            //or in main explorer
            else if (explorer != null)
            {
                var taskPane = ConrepTaskPanes.GetTaskPane(explorer);
                if (taskPane != null)
                    taskPane.Visible = isPressed;

                if (isPressed)
                {
                    var mailItem = explorer.GetSelectedEmail();
                    Globals.ThisAddIn.ImportSelectedEmailWhenShowingTaskPane(mailItem, taskPane);
                }
            }
        }

        public void SyncCalendars(Office.IRibbonControl control)
        {
            var inspector = control.Context as Inspector;
            var explorer = control.Context as Explorer;

            if (inspector != null)
            {
                var taskPane = ConrepTaskPanes.GetTaskPane(inspector);
                Globals.ThisAddIn.SyncCalendarsAsync(taskPane);
            }
            else if (explorer != null)
            {
                var taskPane = ConrepTaskPanes.GetTaskPane(explorer);
                Globals.ThisAddIn.SyncCalendarsAsync(taskPane);
            }
        }

        public void ShowTaskPane(bool show)
        {
            _showTaskPane = show;
            ribbon.InvalidateControl("ShowHideConrepPaneButton");
        }

        private void BulkImport(Office.IRibbonControl control, ApiInvokeMode mode, bool sendHeadersOnly = false)
        {
            var inspector = control.Context as Inspector;
            var explorer = control.Context as Explorer;

            if (inspector != null)
            {
                var mailItems = new List<MailItem>() { inspector.CurrentItem as MailItem };
                var taskPane = ConrepTaskPanes.GetTaskPane(inspector);
                Globals.ThisAddIn.BulkImport(mode, mailItems, taskPane, sendHeadersOnly);
            }
            else if (explorer != null)
            {
                var mailItems = explorer.GetSelectedEmails().ToList();
                var taskPane = ConrepTaskPanes.GetTaskPane(explorer);

                Globals.ThisAddIn.BulkImport(mode, mailItems, taskPane, sendHeadersOnly);
            }
        }
        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string ribbonXml = string.Empty;
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                case "Microsoft.Outlook.Mail.Read":
                {
                    ribbonXml = GetResourceText("ConrepOutlookAddin.AddinRibbon.xml");
                    break;
                }
            }

            return ribbonXml;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}

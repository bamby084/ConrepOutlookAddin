using Microsoft.Office.Interop.Outlook;

namespace ConrepOutlookAddin
{
    public static class ConrepInspectors
    {
        private static Inspectors _inspectors;

        public static void Handle()
        {
            _inspectors = Globals.ThisAddIn.Application.Inspectors;
            _inspectors.NewInspector += OnNewInspector;
        }

        //only handle reading mail inspector
        private static void OnNewInspector(Inspector inspector)
        {
            new ReadingInspector(inspector).Handle();
        }
    }
}

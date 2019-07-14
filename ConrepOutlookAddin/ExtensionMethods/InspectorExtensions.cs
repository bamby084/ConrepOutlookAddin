using System;
using ConrepOutlookAddin.Win32;
using Microsoft.Office.Interop.Outlook;

namespace ConrepOutlookAddin.ExtensionMethods
{
    public static class InspectorExtensions
    {
        public static IntPtr GetWindowHandle(this Inspector inspector)
        {
            var window = inspector as IOleWindow;
            if (window == null)
                return IntPtr.Zero;

            IntPtr hWnd;
            window.GetWindow(out hWnd);

            return hWnd;
        }
    }
}

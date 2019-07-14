﻿using System;
using System.Runtime.InteropServices;

namespace ConrepOutlookAddin.Win32
{
    [ComImport]
    [Guid("00000114-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleWindow
    {
        /// <summary>
        /// Returns the window handle to one of the windows participating in in-place activation 
        /// (frame, document, parent, or in-place object window).
        /// </summary>
        /// <param name="phwnd">Pointer to where to return the window handle.</param>
        void GetWindow(out IntPtr phwnd);

        /// <summary>
        /// Determines whether context-sensitive help mode should be entered during an 
        /// in-place activation session.
        /// </summary>
        /// <param name="fEnterMode"><c>true</c> if help mode should be entered; 
        /// <c>false</c> if it should be exited.</param>
        void ContextSensitiveHelp([In, MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
    }
}

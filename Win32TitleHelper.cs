using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitBranchTitlebar
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.Shell;
    using EnvDTE80;

    public static class Win32TitleHelper
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool SetWindowText(IntPtr hwnd, string lpString);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public static bool SetVSMainWindowTitle(DTE2 dte, string newTitle)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string currentTitle = dte.MainWindow.Caption;
            IntPtr hwnd = FindWindow(null, currentTitle);

            if (hwnd != IntPtr.Zero)
            {
                return SetWindowText(hwnd, newTitle);
            }

            return false;
        }
    }

}

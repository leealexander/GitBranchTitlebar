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

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        public static bool SetVSMainWindowTitle(string currentTitle, string newTitle)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IntPtr hwnd = FindWindow(null, currentTitle);

            if (hwnd != IntPtr.Zero)
            {
                return SetWindowText(hwnd, newTitle);
            }

            return false;
        }

        public static string GetVSMainWindowTitle(string currentTitle)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IntPtr hwnd = FindWindow(null, currentTitle);

            if (hwnd != IntPtr.Zero)
            {
                var sb = new StringBuilder(256);
                if (GetWindowText(hwnd, sb, sb.Capacity) > 0)
                {
                    return sb.ToString();
                }
            }
            return null;
        }
    }

}

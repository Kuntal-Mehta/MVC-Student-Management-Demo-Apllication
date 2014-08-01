using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows;

namespace ViewportAdornment1
{
    public static class clsSHELLEXECUTEINFO
    {
        const int SEE_MASK_INVOKEIDLIST = 0x00000C;
        const int SEE_MASK_NOCLOSEPROCESS = 0x000040;
        const int SEE_MASK_FLAG_NO_UI = 0x000400;

        [StructLayout(LayoutKind.Sequential)]
        internal struct SHELLEXECUTEINFO
        {
            internal int cbSize;
            internal uint fMask;
            internal IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string lpVerb;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string lpFile;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string lpParameters;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string lpDirectory;
            internal int nShow;
            internal IntPtr hInstApp;
            internal IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string lpClass;
            internal IntPtr hkeyClass;
            internal uint dwHotKey;
            internal IntPtr hIcon;
            internal IntPtr hProcess;
        }

        public static void ShowFilePropertyDlg(string path)
        {
            SHELLEXECUTEINFO info =
                new SHELLEXECUTEINFO();

            info.cbSize = Marshal.SizeOf(info);
            info.hwnd = IntPtr.Zero;

            info.fMask = (SEE_MASK_INVOKEIDLIST |
                SEE_MASK_NOCLOSEPROCESS |
                SEE_MASK_FLAG_NO_UI);

            //Properties dialog
            //for file or folder
            //to be displayed
            info.lpVerb = "properties";

            //using file path which is passed as a parameter         
            info.lpFile = path;

            if (!ShellExecuteEx(ref info))
            {
                int lastError = Marshal.GetLastWin32Error();

                MessageBox.Show(
                    string.Format("Could not open properties window.\nShellExecuteEx failed with error code: {0}",
                    lastError));
            }
        }

        [DllImport("shell32.dll", SetLastError = true,
            EntryPoint = "ShellExecuteExW",
            CharSet = CharSet.Unicode)]
        static extern bool ShellExecuteEx(
            ref SHELLEXECUTEINFO lpExecInfo);
    }
}
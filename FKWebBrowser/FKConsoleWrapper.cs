using FKWebBrowser.Source;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace FKWebBrowser
{
    public class FKConsoleWrapper
    {
        #region WIN32接口
        [DllImport("User32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);
        [DllImport("User32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        #endregion

        public enum ENUM_UserMode
        {
            eUserMode_Developer,        // 开发者/Admin
            eUserMode_Guest,            // 小游客
        }

        private ENUM_UserMode m_eUserMode = ENUM_UserMode.eUserMode_Developer;
        private Form_Main m_MainForm = null;
        private Log m_Log = null;

        public FKConsoleWrapper()
        {
            m_Log = new Log(null);
            m_Log.AddInfo(Log.ENUM_Level.UserInfo, "================= APP START =================");
            HideConsoleWindow();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (m_MainForm == null)
            {
                m_MainForm = new Form_Main(this);
            }
        }
        public void ShowForm()
        {
            if (m_MainForm != null)
            {
                try
                {
                    m_MainForm.ShowDialog();
                }
                catch { }
            }
        }
        public Log GetLog()
        {
            return m_Log;
        }
        public ENUM_UserMode GetUserMode()
        {
            return m_eUserMode;
        }
        private void HideConsoleWindow()
        {
            string strConsoleTitle = Console.Title;
            IntPtr hWnd = FindWindow("ConsoleWindowClass", strConsoleTitle);
            if (hWnd != IntPtr.Zero)
                ShowWindow(hWnd, 0);
        }

        // 退出程序
        public void ExitApp()
        {
            m_Log.SetForm(null);
            m_Log.AddInfo(Log.ENUM_Level.UserInfo, "================= APP EXIT =================");
        }
    }
}

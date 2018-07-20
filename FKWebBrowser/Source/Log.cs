using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

namespace FKWebBrowser.Source
{
    public class Log
    {
        public enum ENUM_Level
        {
            Error,      // 写入文件，控制台输出，界面输出，必要时弹出窗口和关闭进程
            UserInfo,   // 写入文件，控制台输出，界面输出
            Info,       // 写入文件，控制台输出，
            Debug,      // 仅控制台输出
        }

        private static TraceSource m_Trace = new TraceSource("LOG");
        private object m_Lock = new object();
        private int m_nLogID = 0;
        private bool m_bIsLogEnable = true;
        private string m_strFileName = "DefaultLog.log";
        private ENUM_Level m_eLevel = ENUM_Level.Debug;
        private Form m_Form = null;

        // 构造函数
        public Log(Form form = null)
        {
            if (form != null)
                m_Form = form;

            m_Trace.Switch = new SourceSwitch("SourceSwitch", "Verbose");
        }

        // 添加一行Log
        public void AddInfo(ENUM_Level eLevel, string strMessage)
        {
            if (strMessage == null)
                strMessage = "";

            string strCurLogName = DateTime.Now.ToString("yyyy-MM-dd");
            m_strFileName = strCurLogName + ".log";

            m_nLogID++;
            if (m_nLogID >= 99999)
                m_nLogID = 0;
            string strTotalMsg = GetStrLine(eLevel, m_nLogID, strMessage);
            lock (m_Lock)
            {
                // 写入日志
                if (NeedLogFile(eLevel))
                {
                    if (m_bIsLogEnable)
                    {
                        m_Trace.TraceEvent(GetFormTypeAnalyzer(eLevel), m_nLogID, strMessage);
                        using (StreamWriter w = File.AppendText(m_strFileName))
                        {
                            w.WriteLine(strTotalMsg);
                        }
                    }
                }

                // 写入控制台
                Console.WriteLine(strTotalMsg);

                // 写入用户界面
                if (m_Form != null)
                {
                    (m_Form as Form_Main).AddInfoToForm(strTotalMsg);
                }
            }
        }

        // 简易接口
        public void AddInfo(string strMessage)
        {
            AddInfo(ENUM_Level.Info, strMessage);
        }
        public void SetForm(Form hForm)
        {
            m_Form = hForm;
        }
        public void SetLevel(ENUM_Level eLevel)
        {
            m_eLevel = eLevel;
        }
        public void DisableLog()
        {
            m_bIsLogEnable = false;
        }
        public void EnableLog()
        {
            m_bIsLogEnable = true;
        }
        private string GetStrLine(ENUM_Level eLevel, int nLogID, string strMessage)
        {
            string t = "";
            if (eLevel == ENUM_Level.Debug)
                t = "Debug";
            if (eLevel == ENUM_Level.Error)
                t = "Error";
            if (eLevel == ENUM_Level.UserInfo)
                t = "Info";
            if (eLevel == ENUM_Level.Info)
                t = "Info";

            if (nLogID > 999990)
                nLogID = 0;

            string strTimeStamp = DateTime.Now.ToString("yyyy-MM-dd|HH:mm:ss:ffff");
            return nLogID.ToString("000000") + "|" + strTimeStamp + "|" + t + "|" + strMessage;
        }
        private TraceEventType GetFormTypeAnalyzer(ENUM_Level eLevel)
        {
            TraceEventType e = TraceEventType.Information;
            if (eLevel == ENUM_Level.Debug)
                e = TraceEventType.Verbose;
            if (eLevel == ENUM_Level.Error)
                e = TraceEventType.Error;
            if (eLevel == ENUM_Level.UserInfo)
                e = TraceEventType.Information;
            if (eLevel == ENUM_Level.Info)
                e = TraceEventType.Information;

            return e;
        }

        private bool NeedLogFile(ENUM_Level eLevel)
        {
            if (((int)eLevel) <= ((int)ENUM_Level.Info))
                return true;
            return false;
        }
    }
}

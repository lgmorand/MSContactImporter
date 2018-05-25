using System;
using System.Windows.Forms;
using log4net;

namespace Microsoft.Internal.MSContactImporter
{
    public static class Logger
    {
        private static TextBox controlWhereToLog;
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));

        public static void SetControl(TextBox ctrl)
        {
            controlWhereToLog = ctrl;
        }

        public static void LogMessageToConsole(string message, Exception ex = null)
        {
            controlWhereToLog.Text += string.Format("\r\n{0:dd/MM/yyyy HH:mm:ss}: {1}", DateTime.Now, message);
            controlWhereToLog.Select(controlWhereToLog.TextLength, 0);
            controlWhereToLog.ScrollToCaret();
            Application.DoEvents();
            LogMessageToLogFile(message, ex);
        }

        public static void LogMessageToLogFile(string message, Exception ex = null)
        {
            if (ex == null)
            {
                log.Info(message);
            }
            else
            {
                log.Error(message, ex);
            }
        }
    }
}
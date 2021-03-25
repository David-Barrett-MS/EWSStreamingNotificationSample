/*
 * By David Barrett, Microsoft Ltd. 2013. Use at your own risk.  No warranties are given.
 * 
 * DISCLAIMER:
 * THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
 * MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
 * A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
 * MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
 * BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
 * SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
 * OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 * */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace EWSStreamingNotificationSample
{
    public class LoggerEventArgs : EventArgs
    {
        private DateTime _logTime;
        private string _logDetails;

        public LoggerEventArgs(DateTime LogTime, string LogDetails)
        {
            _logTime = LogTime;
            _logDetails = LogDetails;
        }

        public DateTime LogTime
        {
            get { return _logTime; }
        }

        public string LogDetails
        {
            get { return _logDetails; }
        }
    }

    public class ClassLogger
    {
        private StreamWriter _logStream = null;
        private string _logPath = "";
        private bool _logDateAndTime = true;

        public delegate void LoggerEventHandler(object sender, LoggerEventArgs a);
        public event LoggerEventHandler LogAdded;

        public ClassLogger(string LogFile)
        {
            try
            {
                _logStream = File.AppendText(LogFile);
                _logPath = LogFile;
            }
            catch { }
        }

        ~ClassLogger()
        {
            try
            {
                _logStream.Flush();
                _logStream.Close();
            }
            catch { }
        }

        protected virtual void OnLogAdded(LoggerEventArgs e)
        {
            LoggerEventHandler handler = LogAdded;
            if (handler != null)
                handler(this, e);
        }

        public bool LogDateAndTime
        {
            get { return _logDateAndTime; }
            set { _logDateAndTime = value; }
        }

        public void ClearFile()
        {
        }

        public void Log(string Details, string Description = "", bool SuppressEvent = false)
        {
            try
            {
                DateTime oLogTime = DateTime.Now;

                if (String.IsNullOrEmpty(Description))
                {
                    if (_logDateAndTime)
                        _logStream.WriteLine(String.Format("{0:dd/MM/yy HH:mm:ss}", oLogTime) + " ==> " + Details);
                }
                else
                {
                    _logStream.WriteLine("");
                    if (_logDateAndTime)
                        _logStream.WriteLine(String.Format("{0:dd/MM/yy HH:mm:ss}", oLogTime) + " ==> " + Description);
                    _logStream.WriteLine(Details);
                }
                _logStream.Flush();

                if (!SuppressEvent)
                    OnLogAdded(new LoggerEventArgs(oLogTime, Description + " " + Details));
            }
            catch {}
        }
    }
}

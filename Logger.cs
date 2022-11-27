/*
 * By David Barrett, Microsoft Ltd. 2013-2022. Use at your own risk.  No warranties are given.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * */

using System;
using System.Globalization;
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

    public class Logger
    {
        private StreamWriter _logStream = null;
        private string _logPath = "";
        private bool _logDateAndTime = true;
        private int _logFlushCounter = 0;

        public delegate void LoggerEventHandler(object sender, LoggerEventArgs a);
        public event LoggerEventHandler LogAdded;
        static public Logger DefaultLogger = null;

        public Logger(string LogFile)
        {
            try
            {
                _logStream = File.AppendText(LogFile);
                _logPath = LogFile;
            }
            catch { }
            if (DefaultLogger == null)
                DefaultLogger = this;
        }

        ~Logger()
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

                lock (_logStream)
                {
                    if (String.IsNullOrEmpty(Description))
                    {
                        
                        if (_logDateAndTime)
                            _logStream.WriteLine(oLogTime.ToString("O", DateTimeFormatInfo.InvariantInfo) + " ==> " + Details);
                    }
                    else
                    {
                        _logStream.WriteLine("");
                        if (_logDateAndTime)
                            _logStream.WriteLine(oLogTime.ToString("O", DateTimeFormatInfo.InvariantInfo) + " ==> " + Description);
                        _logStream.WriteLine(Details);
                    }
                }

                _logFlushCounter++;
                if (_logFlushCounter > 9)
                {
                    _logStream.Flush();
                    _logFlushCounter = 0;
                }

                if (!SuppressEvent)
                    OnLogAdded(new LoggerEventArgs(oLogTime, Description + " " + Details));
            }
            catch { }
        }
    }
}

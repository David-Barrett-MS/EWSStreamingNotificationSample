/*
 * By David Barrett, Microsoft Ltd. 2022. Use at your own risk.  No warranties are given.
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
using Microsoft.Exchange.WebServices.Data;
using System.IO;

namespace EWSStreamingNotificationSample
{
    public class TraceListener : ITraceListener
    {
        string _traceFile = "";
        private StreamWriter _traceStream = null;
        object _writeLock = new object();

        public TraceListener(string traceFile)
        {
            try
            {
                _traceStream = File.AppendText(traceFile);
                _traceFile = traceFile;
            }
            catch { }
        }

        ~TraceListener()
        {
            CloseTraceFile();
        }

        public void CloseTraceFile()
        {
            if (_traceStream == null)
                return;

            try
            {
                _traceStream.Flush();
                _traceStream.Close();
            }
            catch { }
            _traceStream = null;
        }

        public void Trace(string traceType, string traceMessage)
        {
            if (String.IsNullOrEmpty(traceMessage))
                return;

            if (_traceStream == null)
            {
                if (String.IsNullOrEmpty(_traceFile))
                    return;
                try
                {
                    _traceStream = File.AppendText(_traceFile);
                }
                catch
                {
                    return;
                }
            }

            lock (_writeLock)
            {
                try
                {
                    _traceStream.WriteLine(traceMessage);
                    _traceStream.Flush();
                }
                catch { }
            }
        }
    }
}

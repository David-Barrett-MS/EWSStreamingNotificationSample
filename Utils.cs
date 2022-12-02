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
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;

namespace EWSStreamingNotificationSample
{
    internal class Utils
    {
        private static TraceListener _traceListener;
        private static Auth.CredentialHandler _credentialHandler;

        /// <summary>
        /// Dictionary of ConnectionTrackers, indexed by group name
        /// </summary>
        public static Dictionary<string, ConnectionTracker> Connections = new Dictionary<string, ConnectionTracker>();

        /// <summary>
        /// Set a new random Guid as the client request Id
        /// </summary>
        /// <param name="exchangeService">The ExchangeService object to which to apply the ClientRequestId</param>
        public static void SetClientRequestId(ExchangeService exchangeService)
        {
            exchangeService.ClientRequestId = Guid.NewGuid().ToString();
        }

        /// <summary>
        /// Set a new random Guid as the client request Id
        /// </summary>
        /// <param name="autodiscoverService">The AutodiscoverService object to which to apply the ClientRequestId</param>
        public static void SetClientRequestId(AutodiscoverService autodiscoverService)
        {
            autodiscoverService.ClientRequestId = Guid.NewGuid().ToString();
        }

        /// <summary>
        /// Returns a trace listener object (if one not already created, trace.log is created in the current folder)
        /// </summary>
        public static ITraceListener TraceListener
        {
            get {
                if (_traceListener != null)
                    return _traceListener;
                CreateTraceListener("trace.log");
                return _traceListener;
            }
        }

        /// <summary>
        /// Create a trace listener that write to a specific file
        /// </summary>
        /// <param name="TraceFile">The file to which the trace is written</param>
        public static void CreateTraceListener(string TraceFile)
        {
            if (_traceListener != null)
                _traceListener.CloseTraceFile();
            _traceListener = new TraceListener(TraceFile);
        }

        /// <summary>
        /// The CredentialHandler object to be used to apply credentials to other objects (e.g. ExchangeService)
        /// </summary>
        public static Auth.CredentialHandler CredentialHandler
        {
            get { return _credentialHandler; }
            set { _credentialHandler = value; }
        }

        public static void SetXAnchorMailbox(ExchangeService exchange, string xAnchorMailbox)
        {
            if (exchange.HttpHeaders.ContainsKey("X-AnchorMailbox"))
                exchange.HttpHeaders.Remove("X-AnchorMailbox");
            exchange.HttpHeaders.Add("X-AnchorMailbox", xAnchorMailbox);
        }

        /// <summary>
        /// Create a new exchange service
        /// </summary>
        /// <param name="Mailbox">The mailbox being impersonated (set as ImpersonatedUserId)</param>
        /// <param name="EwsUrl">The EWS URL</param>
        /// <param name="AnchorMailbox">Anchor mailbox (will be set as X-Anchormailbox if provided)</param>
        /// <returns>ExchangeService object</returns>
        public static ExchangeService NewExchangeService(string Mailbox, string EwsUrl, string AnchorMailbox = "")
        {
            if (String.IsNullOrEmpty(EwsUrl))
                return null;

            // Create exchange service for this group
            ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2016);
            exchangeService.UserAgent = "StreamingSampleApp";
            exchangeService.HttpHeaders.Add("X-PreferServerAffinity", "true");

            exchangeService.Url = new Uri(EwsUrl);
            if (TraceListener != null)
            {
                exchangeService.ReturnClientRequestId = true;
                exchangeService.TraceListener = Utils.TraceListener;
                exchangeService.TraceFlags = TraceFlags.All;
                exchangeService.TraceEnabled = true;
            }

            exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Mailbox);
            
            if (!String.IsNullOrEmpty(AnchorMailbox))
                SetXAnchorMailbox(exchangeService, AnchorMailbox);

            return exchangeService;
        }
    }
}

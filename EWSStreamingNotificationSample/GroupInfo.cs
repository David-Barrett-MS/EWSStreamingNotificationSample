/*
 * By David Barrett, Microsoft Ltd. 2014. Use at your own risk.  No warranties are given.
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
using Microsoft.Exchange.WebServices.Data;

namespace EWSStreamingNotificationSample
{
    public class GroupInfo
    {
        private string _name = "";
        private string _primaryMailbox = "";
        private List<String> _mailboxes;
        //private List<StreamingSubscriptionConnection> _streamingConnection;
        private ExchangeService _exchangeService = null;
        private ITraceListener _traceListener = null;
        private string _ewsUrl = "";
        private ExchangeVersion _exchangeVersion = ExchangeVersion.Exchange2013;

        public GroupInfo(string Name, string PrimaryMailbox, string EWSUrl, ITraceListener TraceListener = null)
        {
            // initialise the group information
            _name = Name;
            _primaryMailbox = PrimaryMailbox;
            _ewsUrl = EWSUrl;
            _traceListener = TraceListener;
            _mailboxes = new List<String>();
            _mailboxes.Add(PrimaryMailbox);
        }

        public string Name
        {
            get { return _name; }
        }

        public ExchangeVersion ExchangeVersion
        {
            get { return _exchangeVersion; }
            set { _exchangeVersion = value; }
        }

        public string PrimaryMailbox
        {
            get { return _primaryMailbox; }
            set
            {
                // If the primary mailbox changes, we need to ensure that it is in the mailbox list also
                _primaryMailbox = value;
                if (!_mailboxes.Contains(_primaryMailbox))
                    _mailboxes.Add(_primaryMailbox);
            }
        }

        public ExchangeService ExchangeService
        {
            get
            {
                if (_exchangeService != null)
                    return _exchangeService;

                // Create exchange service for this group
                ExchangeService exchange = new ExchangeService(_exchangeVersion);
                exchange.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, _primaryMailbox);
                exchange.HttpHeaders.Add("X-AnchorMailbox", _primaryMailbox);
                exchange.HttpHeaders.Add("X-PreferServerAffinity", "true");
                exchange.Url = new Uri(_ewsUrl);
                if (_traceListener != null)
                {
                    exchange.TraceListener = _traceListener;
                    exchange.TraceFlags = TraceFlags.All;
                    exchange.TraceEnabled = true;
                }
                return exchange;
            }
        }

        public void SetXAnchorToPrimary()
        {
            if (_exchangeService == null)
                return;

            if (_exchangeService.HttpHeaders.ContainsKey("X-AnchorMailbox"))
                _exchangeService.HttpHeaders.Remove("X-AnchorMailbox");
            _exchangeService.HttpHeaders.Add("X-AnchorMailbox", _primaryMailbox);
        }

        public List<String> Mailboxes
        {
            get { return _mailboxes; }
        }

        public int NumberOfGroups
        {
            // The maximum number of mailboxes in a group shouldn't exceed 200, which means that this group may consist
            // of several groups. 
            get { return ((_mailboxes.Count / 200))+1; }
        }

        public List<List<String>> MailboxesGrouped
        {
            get
            {
                // Return a list of lists (the group split into lists of 200)
                List<List<String>> groupedMailboxes = new List<List<String>>();
                for (int i=0; i<NumberOfGroups; i++)
                {
                    List<String> mailboxes = _mailboxes.GetRange(i * 200, 200);
                    groupedMailboxes.Add(mailboxes);
                }
                return groupedMailboxes;
            }
        }
    }
}

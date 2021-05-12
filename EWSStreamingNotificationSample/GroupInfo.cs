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
        private string _xBackendOverrideCookie = "";
        //private List<StreamingSubscriptionConnection> _streamingConnection;
        private ExchangeService _exchangeService = null;
        private ITraceListener _traceListener = null;
        private string _ewsUrl = "";
        private ExchangeVersion _exchangeVersion = ExchangeVersion.Exchange2016;
        private Auth.CredentialHandler _credentialHandler;

        public GroupInfo(string Name, string PrimaryMailbox, string EWSUrl, Auth.CredentialHandler credentialHandler, ITraceListener TraceListener = null)
        {
            // initialise the group information
            _name = Name;
            _primaryMailbox = PrimaryMailbox;
            _ewsUrl = EWSUrl;
            _traceListener = TraceListener;
            _mailboxes = new List<String>();
            _mailboxes.Add(PrimaryMailbox);
            _credentialHandler = credentialHandler;
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

        public void ApplyHeadersToConnection(StreamingSubscriptionConnection connection)
        {
        }

        public ExchangeService ExchangeService
        {
            get
            {
                if (_exchangeService != null)
                    return _exchangeService;

                // Create exchange service for this group
                _exchangeService = new ExchangeService(_exchangeVersion);
                _credentialHandler.ApplyCredentialsToExchangeService(_exchangeService);
                _exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, _primaryMailbox);

                _exchangeService.HttpHeaders.Add("X-AnchorMailbox", _primaryMailbox);
                _exchangeService.HttpHeaders.Add("X-PreferServerAffinity", "true");
                _exchangeService.HttpHeaders.Add("return-client-request-id", "true");

                _exchangeService.Url = new Uri(_ewsUrl);
                if (_traceListener != null)
                {
                    _exchangeService.TraceListener = _traceListener;
                    _exchangeService.TraceFlags = TraceFlags.All;
                    _exchangeService.TraceEnabled = true;
                }
                return _exchangeService;
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
                // This isn't implemented, and would need completion for applications that subscribe to
                // large numbers of mailboxes
                List<List<String>> groupedMailboxes = new List<List<String>>();
                for (int i=0; i<NumberOfGroups; i++)
                {
                    List<String> mailboxes = _mailboxes.GetRange(i * 200, 200);
                    groupedMailboxes.Add(mailboxes);
                }
                return groupedMailboxes;
            }
        }

        private void SetClientRequestId(ExchangeService exchangeService)
        {
            if (exchangeService.HttpHeaders.ContainsKey("client-request-id"))
                exchangeService.HttpHeaders.Remove("client-request-id");
            exchangeService.HttpHeaders.Add("client-request-id", Guid.NewGuid().ToString());
        }

        private StreamingSubscription AddSubscription(string Mailbox,
            ref Dictionary<string, StreamingSubscription> SubscriptionList, EventType[] SubscribeEvents, FolderId[] SubscribeFolders = null, ExchangeService exchange = null)
        {
            // Return the subscription, or create a new one if we don't already have one

            if (SubscriptionList.ContainsKey(Mailbox))
                SubscriptionList.Remove(Mailbox);

            if (exchange==null)
                exchange = ExchangeService;
            exchange.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Mailbox);

            //FolderId[] selectedFolders = SelectedFolders();
            StreamingSubscription subscription;
            SetClientRequestId(exchange);
            if (SubscribeFolders==null)
                subscription = exchange.SubscribeToStreamingNotificationsOnAllFolders(SubscribeEvents);
            else
                subscription = exchange.SubscribeToStreamingNotifications(SubscribeFolders, SubscribeEvents);
            SubscriptionList.Add(Mailbox, subscription);


            if (_primaryMailbox.Equals(Mailbox))
            {
                // Check for the X-BackendOverride cookie
                // Set-Cookie: X-BackEndOverrideCookie=DB7PR04MB4764.EURPRD04.PROD.OUTLOOK.COM~1943310282; path=/; secure; HttpOnly
                // We don't actually need this as we rely on the ExchangeService object to keep track of cookies

                System.Net.CookieCollection cookies = exchange.CookieContainer.GetCookies(new Uri("https://outlook.office365.com"));
                _xBackendOverrideCookie = $"X-BackEndOverrideCookie={cookies["X-BackEndOverrideCookie"].Value}";
            }
            return subscription;
        }

        public StreamingSubscriptionConnection AddGroupSubscriptions(
            ref Dictionary<string, StreamingSubscriptionConnection> Connections,
            ref Dictionary<string, StreamingSubscription> SubscriptionList,
            EventType[] SubscribeEvents,
            FolderId[] SubscribeFolders,
            ClassLogger Logger,
            int TimeOut = 30)
        {
            if (Connections.ContainsKey(_name))
            {
                foreach (StreamingSubscription subscription in Connections[_name].CurrentSubscriptions)
                {
                    try
                    {
                        subscription.Unsubscribe();
                    }
                    catch { }
                }
                try
                {
                    if (Connections[_name].IsOpen)
                        Connections[_name].Close();
                }
                catch { }
                Connections.Remove(_name);
            }

            StreamingSubscriptionConnection groupConnection = null;
            try
            {
                // Create the subscription to the primary mailbox, then create the subscription connection
                if (SubscriptionList.ContainsKey(_primaryMailbox))
                    SubscriptionList.Remove(_primaryMailbox);
                StreamingSubscription subscription = AddSubscription(_primaryMailbox, ref SubscriptionList, SubscribeEvents, SubscribeFolders);
                groupConnection = new StreamingSubscriptionConnection(subscription.Service, TimeOut);
                Connections.Add(_name, groupConnection);

                //SubscribeConnectionEvents(groupConnection);
                groupConnection.AddSubscription(subscription);
                Logger.Log($"{_primaryMailbox} (primary mailbox) subscription created in group {_name}");

                // Now add any further subscriptions in this group
                foreach (string sMailbox in _mailboxes)
                {
                    if (!sMailbox.Equals(_primaryMailbox))
                    {
                        try
                        {
                            if (SubscriptionList.ContainsKey(sMailbox))
                                SubscriptionList.Remove(sMailbox);
                            subscription = AddSubscription(sMailbox, ref SubscriptionList, SubscribeEvents, SubscribeFolders);
                            groupConnection.AddSubscription(subscription);
                            Logger.Log($"{sMailbox} subscription created in group {_name}");
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(String.Format("ERROR when subscribing {0} in group {1}: {2}", sMailbox, _name, ex.Message));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"ERROR when creating subscription connection group {_name}: {ex.Message}");
            }
            return groupConnection;
        }
    }
}

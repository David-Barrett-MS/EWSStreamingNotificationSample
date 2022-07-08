/*
 * By David Barrett, Microsoft Ltd. 2014-2022. Use at your own risk.  No warranties are given.
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
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace EWSStreamingNotificationSample
{
    public class GroupInfo
    {
        private string _name = "";
        private List<String> _mailboxes;
        private static int _maxGroupSize = 200;
        //private string _xBackendOverrideCookie = "";
        private ExchangeService _exchangeService = null;
        private ITraceListener _traceListener = null;
        private string _ewsUrl = "";
        private ExchangeVersion _exchangeVersion = ExchangeVersion.Exchange2016;
        private Auth.CredentialHandler _credentialHandler;

        public GroupInfo(string Name, string FirstMailbox, string EWSUrl, Auth.CredentialHandler credentialHandler, ITraceListener TraceListener = null)
        {
            // initialise the group information
            _name = Name;
            _ewsUrl = EWSUrl;
            _traceListener = TraceListener;
            _mailboxes = new List<String>();
            _mailboxes.Add(FirstMailbox);
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

        public static int MaxGroupSize
        {
            get { return _maxGroupSize; }
            set
            {
                if (value > 0 && value < 201)
                    _maxGroupSize = value;
                else
                    throw new ArgumentException("Group size must be between 1 and 200");
            }
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

        public List<String> Mailboxes
        {
            get { return _mailboxes; }
        }

        private void SetClientRequestId(ExchangeService exchangeService)
        {
            if (exchangeService.HttpHeaders.ContainsKey("client-request-id"))
                exchangeService.HttpHeaders.Remove("client-request-id");
            exchangeService.HttpHeaders.Add("client-request-id", Guid.NewGuid().ToString());
        }

        private StreamingSubscription AddSubscription(string Mailbox,
            ref Dictionary<string,StreamingSubscription> SubscriptionList,
            EventType[] SubscribeEvents,
            FolderId[] SubscribeFolders = null,
            ExchangeService exchange = null,
            string AnchorMailbox = null)
        {
            // Return the subscription, or create a new one if we don't already have one

            if (SubscriptionList.ContainsKey(Mailbox))
                SubscriptionList.Remove(Mailbox);

            if (exchange==null)
                exchange = ExchangeService;
            if (String.IsNullOrEmpty(AnchorMailbox))
                AnchorMailbox = Mailbox;

            exchange.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Mailbox);
            if (exchange.HttpHeaders.ContainsKey("X-AnchorMailbox"))
                exchange.HttpHeaders.Remove("X-AnchorMailbox");
            exchange.HttpHeaders.Add("X-AnchorMailbox", AnchorMailbox);

            //FolderId[] selectedFolders = SelectedFolders();
            StreamingSubscription subscription;
            SetClientRequestId(exchange);
            if (SubscribeFolders==null)
                subscription = exchange.SubscribeToStreamingNotificationsOnAllFolders(SubscribeEvents);
            else
                subscription = exchange.SubscribeToStreamingNotifications(SubscribeFolders, SubscribeEvents);
            SubscriptionList.Add(Mailbox, subscription);


            //if (_primaryMailbox.Equals(Mailbox))
            //{
            //    // Check for the X-BackendOverride cookie
            //    // Set-Cookie: X-BackEndOverrideCookie=DB7PR04MB4764.EURPRD04.PROD.OUTLOOK.COM~1943310282; path=/; secure; HttpOnly
            //    // We don't actually need this as we rely on the ExchangeService object to keep track of cookies

            //    System.Net.CookieCollection cookies = exchange.CookieContainer.GetCookies(new Uri("https://outlook.office365.com"));
            //    _xBackendOverrideCookie = $"X-BackEndOverrideCookie={cookies["X-BackEndOverrideCookie"].Value}";
            //}
            return subscription;
        }

        // Creates the streaming connection for this group with all the subscriptions
        // If there are too many subscriptions for a single connection, SubscriptionIndex will return a positive number
        // and the caller should repeatedly call this method until 0 is returned
        public StreamingSubscriptionConnection AddGroupSubscriptions(
            ref Dictionary<string, StreamingSubscriptionConnection> Connections,
            ref Dictionary<string, StreamingSubscription> SubscriptionList,
            EventType[] SubscribeEvents,
            FolderId[] SubscribeFolders,
            ClassLogger Logger,
            ref int SubscriptionIndex,
            int TimeOut = 30)
        {

            if (SubscriptionIndex == 0)
            {
                // Sort the mailboxes alphabetically
                _mailboxes.Sort();

                // Clear out any old connections
                List<string> activeConnections = Connections.Keys.ToList<string>();
                foreach (string connectionName in activeConnections)
                {
                    if (connectionName.StartsWith(_name))
                    {
                        foreach (StreamingSubscription subscription in Connections[connectionName].CurrentSubscriptions)
                        {
                            try
                            {
                                subscription.Unsubscribe();
                            }
                            catch { }
                        }
                        try
                        {
                            if (Connections[connectionName].IsOpen)
                                Connections[connectionName].Close();
                        }
                        catch { }
                        Connections.Remove(connectionName);
                    }
                }
            }

            StreamingSubscriptionConnection groupConnection = null;

            try
            {
                // The first mailbox of the group is what we set X-AnchorMailbox to for this connection
                string anchorMailbox = _mailboxes[SubscriptionIndex];
                if (SubscriptionList.ContainsKey(anchorMailbox))
                    SubscriptionList.Remove(anchorMailbox);
                StreamingSubscription subscription = AddSubscription(anchorMailbox, ref SubscriptionList, SubscribeEvents, SubscribeFolders);

                string groupName = $"{_name}{anchorMailbox}";
                groupConnection = new StreamingSubscriptionConnection(subscription.Service, TimeOut);
                Connections.Add(groupName, groupConnection);

                groupConnection.AddSubscription(subscription);
                Logger.Log($"{anchorMailbox} (anchor mailbox) subscription created in group {groupName}");

                int i = 1;
                if (_mailboxes.Count <= SubscriptionIndex + i)
                    i = 0;
                while (i < _maxGroupSize && i>0)
                {
                    string sMailbox = _mailboxes[SubscriptionIndex + i++];
                    try
                    {
                        if (SubscriptionList.ContainsKey(sMailbox))
                            SubscriptionList.Remove(sMailbox);
                        subscription = AddSubscription(sMailbox, ref SubscriptionList, SubscribeEvents, SubscribeFolders, null, anchorMailbox);
                        groupConnection.AddSubscription(subscription);
                        Logger.Log($"{sMailbox} subscription created in group {groupName}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"ERROR when subscribing {sMailbox} in group {groupName}: {ex.Message}");
                    }
                    if (SubscriptionIndex + i >= _mailboxes.Count)
                        i = 0;
                }
                if (i > 0)
                    SubscriptionIndex = SubscriptionIndex + i;
                else
                    SubscriptionIndex = 0;

            }
            catch (Exception ex)
            {
                Logger.Log($"ERROR when creating subscription connection group {_name}: {ex.Message}");
            }
            return groupConnection;
        }
    }
}

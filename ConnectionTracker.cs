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
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace EWSStreamingNotificationSample
{
    internal class ConnectionTracker
    {
        public string GroupName { get; internal set; }
        public string AnchorMailbox { get; internal set; }
        public string BackEndOverrideCookie { get; internal set; }
        public StreamingSubscriptionConnection StreamingSubscriptionConnection { get; internal set; }
        private ExchangeService _exchangeService;
        /// <summary>
        /// Lifetime for the streaming connection
        /// </summary>
        public static int ConnectionLifetime = 5;
        /// <summary>
        /// The number of open connections
        /// </summary>
        public static int TotalOpenConnections = 0;

        public ConnectionTracker(string GroupName)
        {
            this.GroupName = GroupName;
        }

        public bool IsConnected
        {
            get
            {
                if (StreamingSubscriptionConnection != null)
                    return StreamingSubscriptionConnection.IsOpen;
                return false;
            }
        }

        public ExchangeService ExchangeService
        {
            get {  return _exchangeService; }
        }

        public void SetBackendOverrideCookie(string cookieValue)
        {
            BackEndOverrideCookie = cookieValue;
        }

        /// <summary>
        /// Adds the subscription to this connection
        /// </summary>
        /// <param name="subscriptionTracker">The SubscriptionTracker object containing the subscription information</param>
        /// <returns>True if successful, otherwise false</returns>
        public bool AddSubscription(SubscriptionTracker subscriptionTracker)
        {
            if (_exchangeService == null)
                return false;

            if (subscriptionTracker.Subscription == null)
            {
                Logger.DefaultLogger?.Log($"Add subscription failed (subscription not found): {subscriptionTracker.SMTPAddress}");
                return false;
            }

            if (this.StreamingSubscriptionConnection == null)
            {
                // Ensure the connection lifetime is not longer than the OAuth access token validity period
                //int lifeTime = (int)Utils.CredentialHandler.TimeUntilAccessTokenExpires().TotalMinutes;
                //if (ConnectionLifetime<lifeTime)
                //    lifeTime = ConnectionLifetime;
                //
                // Turns out this is pointless as you can't set Lifetime on GetEvents with the EWS API, you can only set it when you first create the StreamingSubscriptionConnection.
                // In practise it doesn't matter too much as an error will occur if the token expires, and renewal/reconnection will be triggered at that point.
                // It would be neater if we could control lifetime on GetEvents and sync with OAuth token renewal, but that would require modification of the EWS API.
                //

                this.StreamingSubscriptionConnection = new StreamingSubscriptionConnection(_exchangeService, ConnectionLifetime);
                StreamingSubscriptionConnection.OnDisconnect += StreamingSubscriptionConnection_OnDisconnect;
            }

            StreamingSubscriptionConnection.AddSubscription(subscriptionTracker.Subscription);
            return true;
        }

        /// <summary>
        /// Configure the supplied SubscriptionTracker (applies group AnchorMailbox)
        /// </summary>
        /// <param name="subscriptionTracker">The SubscriptionTracker to be configured</param>
        public void ConfigureSubscription(SubscriptionTracker subscriptionTracker)
        {
            subscriptionTracker.Connection = this;
            if (_exchangeService == null)
            {
                // This is the first subscription to be configured, and so becomes the anchor mailbox
                if (String.IsNullOrEmpty(AnchorMailbox))
                    AnchorMailbox = subscriptionTracker.SMTPAddress;
                if (String.IsNullOrEmpty(subscriptionTracker.AnchorMailbox))
                    subscriptionTracker.AnchorMailbox = AnchorMailbox;
                _exchangeService = Utils.NewExchangeService(subscriptionTracker.SMTPAddress, subscriptionTracker.EwsUrl, subscriptionTracker.AnchorMailbox);
                Logger.DefaultLogger?.Log($"ExchangeService created for {GroupName}, anchor mailbox {AnchorMailbox}");
            }

            if (String.IsNullOrEmpty(subscriptionTracker.AnchorMailbox))
                subscriptionTracker.AnchorMailbox = AnchorMailbox;

            if (!String.IsNullOrEmpty(BackEndOverrideCookie))
            {
                _exchangeService.CookieContainer.Add(_exchangeService.Url, new System.Net.Cookie("X-BackEndOverrideCookie", BackEndOverrideCookie));
            }
        }

        private void StreamingSubscriptionConnection_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            // Handle OnDisconnect event
            TotalOpenConnections--;
            StringBuilder log = new StringBuilder("OnDisconnect received");
            try
            {
                // If Subscription is null, then the error applies to the whole connection (this is the usual case)
                if (args.Subscription != null)
                    log.Append($" for {args.Subscription.Service.ImpersonatedUserId.Id}");
                if (args.Exception != null)
                    log.Append($" with error {args.Exception.Message}");
            }
            catch { }
            //Logger.DefaultLogger?.Log(log.ToString());
        }

        /// <summary>
        /// Connect to the subscriptions and start receiving events
        /// </summary>
        public bool Connect()
        {
            if (_exchangeService == null || StreamingSubscriptionConnection == null || StreamingSubscriptionConnection.CurrentSubscriptions.Count()<1)
                return false;

            try
            {
                Utils.CredentialHandler?.ApplyCredentialsToExchangeService(_exchangeService);
                Utils.SetClientRequestId(_exchangeService);
                _exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, AnchorMailbox);
                StreamingSubscriptionConnection.Open();
                Logger.DefaultLogger?.Log($"Opened connection: {GroupName}");
                TotalOpenConnections++;
                return true;
            }
            catch (Exception ex)
            {
                Logger.DefaultLogger?.Log($"Failed to connect {GroupName}: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Disconnect from subscriptions (stop receiving events)
        /// </summary>
        public void Disconnect()
        {
            if (!IsConnected)
                return;

            try
            {
                StreamingSubscriptionConnection.Close();
            }
            catch (Exception ex)
            {
                Logger.DefaultLogger?.Log($"Failed to close connection {GroupName}: {ex.Message}");
            }
        }
    }
}

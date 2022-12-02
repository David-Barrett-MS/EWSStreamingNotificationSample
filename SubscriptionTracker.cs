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
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;

namespace EWSStreamingNotificationSample
{
    /// <summary>
    /// Configures and tracks EWS streaming subscriptions
    /// </summary>
    public class SubscriptionTracker
    {
        /// <summary>
        /// Set when the class is instantiated, so that we know when we obtained the information
        /// </summary>
        private DateTime _timeInfoSet;
        /// <summary>
        /// SMTP address of the mailbox this instance is tracking
        /// </summary>
        public string SMTPAddress { get; private set; }
        /// <summary>
        /// EWS Url for this mailbox
        /// </summary>
        public string EwsUrl { get; private set; }
        /// <summary>
        /// GroupingInformation for this mailbox
        /// </summary>
        public string GroupingInformation { get; private set; }
        /// <summary>
        /// The group name for the subscription group
        /// </summary>
        public string GroupName { get; set; }
        /// <summary>
        /// X-AnchorMailbox for requests (is set on CreateSubscription when grouping information is known)
        /// </summary>
        public string AnchorMailbox { get; set; }
        /// <summary>
        /// SubscriptionId for this mailbox
        /// </summary>
        public StreamingSubscription Subscription { get; private set; }
        /// <summary>
        /// Total number of active subscriptions (have been subscribed)
        /// </summary>
        public static int TotalActiveSubscriptions { get; private set; } = 0;
        /// <summary>
        /// Autodiscover service required for autodiscover requests
        /// </summary>
        private AutodiscoverService _autodiscover;
        /// <summary>
        /// The ConnectionTracker this subscription is tied to
        /// </summary>
        internal ConnectionTracker Connection;
        /// <summary>
        /// The ExchangeService object to use for EWS requests
        /// </summary>
        //private ExchangeService _exchangeService = null;
        /// <summary>
        /// Holds information to subscriptions to mailbox.  Key is subscription ID, value is mailbox SMTP address
        /// </summary>
        private static Dictionary<string,string> _subscriptionIdToMailbox = new Dictionary<string,string>();

        private SubscriptionTracker()
        {
            _timeInfoSet = DateTime.Now;
        }

        /// <summary>
        /// Returns the SMTP address of the mailbox to which this subscription is connected
        /// </summary>
        /// <param name="SubscriptionId">The subscription Id to be searched</param>
        /// <returns></returns>
        public static string MailboxOwningSubscription(string SubscriptionId)
        {
            if (_subscriptionIdToMailbox.ContainsKey(SubscriptionId))
                return _subscriptionIdToMailbox[SubscriptionId];
            return "";
        }

        /// <summary>
        /// Clear the subscription from all tracking (it failed and needs to be rebuilt)
        /// </summary>
        /// <param name="SubscriptionId">The SubscriptionId to clear/remove</param>
        public static void ClearSubscription(string SubscriptionId)
        {
            if (_subscriptionIdToMailbox.ContainsKey(SubscriptionId))
                _subscriptionIdToMailbox.Remove(SubscriptionId);
        }



        /// <summary>
        /// Attempt to create a new subscription tracker for the given mailbox
        /// </summary>
        /// <returns>The SubscriptionTracker if successful, null otherwise</returns>
        public static SubscriptionTracker CreateTrackerForMailbox(string Mailbox, string AutodiscoverURL = "")
        {
            // Retrieve the autodiscover information
            Logger.DefaultLogger?.Log($"Retrieving user settings for {Mailbox}");
            AutodiscoverService autoDiscoverService = new AutodiscoverService();
            if (Utils.TraceListener != null)
            {
                autoDiscoverService.ReturnClientRequestId = true;
                autoDiscoverService.TraceListener = Utils.TraceListener;
                autoDiscoverService.TraceFlags = TraceFlags.All;
                autoDiscoverService.TraceEnabled = true;
            }
            Utils.CredentialHandler?.ApplyCredentialsToAutodiscoverService(autoDiscoverService);
            GetUserSettingsResponse userSettings = null;

            if (!String.IsNullOrEmpty(AutodiscoverURL))
            {
                // Use the supplied Autodiscover URL
                try
                {
                    autoDiscoverService.Url = new Uri(AutodiscoverURL);
                    Utils.SetClientRequestId(autoDiscoverService);
                    userSettings = autoDiscoverService.GetUserSettings(Mailbox, UserSettingName.InternalEwsUrl, UserSettingName.ExternalEwsUrl, UserSettingName.GroupingInformation);
                }
                catch
                {
                    autoDiscoverService.Url = null;
                }
            }

            if (userSettings == null)
            {
                // If we don't have an AutoDiscover URL, or the supplied one fails, we start without one (which triggers full AutoD)
                try
                {
                    Utils.SetClientRequestId(autoDiscoverService);
                    userSettings = autoDiscoverService.GetUserSettings(Mailbox, UserSettingName.InternalEwsUrl, UserSettingName.ExternalEwsUrl, UserSettingName.GroupingInformation);
                }
                catch (Exception ex)
                {
                    Logger.DefaultLogger?.Log($"Failed to autodiscover for {Mailbox}: {ex.Message}");
                    return null;
                }
            }

            string ewsUrl = "";
            try
            {
                ewsUrl = (string)userSettings.Settings[UserSettingName.ExternalEwsUrl];
            }
            catch
            {
                try
                {
                    ewsUrl = (string)userSettings.Settings[UserSettingName.InternalEwsUrl];
                }
                catch { }
            }
            string groupingInformation = "none";
            try
            {
                groupingInformation = (string)userSettings.Settings[UserSettingName.GroupingInformation];
            }
            catch { }

            if (!String.IsNullOrEmpty(ewsUrl) && !String.IsNullOrEmpty(groupingInformation))
            {
                SubscriptionTracker tracker = new SubscriptionTracker();
                tracker._autodiscover = autoDiscoverService;
                tracker.SMTPAddress = Mailbox;
                tracker.EwsUrl = ewsUrl;
                tracker.GroupingInformation = groupingInformation;
                tracker.GroupName = $"{groupingInformation}{ewsUrl}";
                return tracker;
            }
            return null;
        }


        /// <summary>
        /// Returns true if AutoDiscover information for this mailbox is more than 24 hours old
        /// </summary>
        private bool IsAutodiscoverStale
        {
            // We assume that the information is stale if it is over 24 hours old
            get { return DateTime.Now.Subtract(_timeInfoSet) > new TimeSpan(0, 24, 0, 0); }
        }



        /// <summary>
        /// Create new StreamingSubscription for this mailbox
        /// </summary>
        /// <param name="SubscribeEvents">Array of events to subscribe to</param>
        /// <param name="SubscribeFolders">Array of folders to subscribe to (if missing, all folders are subscribed)</param>
        /// <param name="GroupAnchorMailbox">Value for X-AnchorMailbox (based on grouping)</param>
        /// <returns></returns>
        public bool Subscribe(EventType[] SubscribeEvents,
            FolderId[] SubscribeFolders = null,
            string GroupAnchorMailbox = null)
        {
            if (!String.IsNullOrEmpty(GroupAnchorMailbox))
                AnchorMailbox = GroupAnchorMailbox;
            if (String.IsNullOrEmpty(AnchorMailbox))
                AnchorMailbox = SMTPAddress;
            //if (_exchangeService == null)
            //    _exchangeService = NewExchangeService();

            ExchangeService exchangeService = Connection.ExchangeService;
            Utils.SetClientRequestId(exchangeService);
            Utils.CredentialHandler?.ApplyCredentialsToExchangeService(exchangeService);

            try
            {
                if (SMTPAddress != AnchorMailbox)
                    exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, SMTPAddress);

                const string cookieName = "X-BackEndOverrideCookie";
                if (!String.IsNullOrEmpty(Connection.BackEndOverrideCookie))
                {
                    // Add the X-BackEndOverrideCookie to the request
                    string domain = EwsUrl.Substring(8, EwsUrl.IndexOf("/", 8)-8);
                    exchangeService.CookieContainer.Add(new System.Net.Cookie(cookieName, Connection.BackEndOverrideCookie,"/",domain));
                }

                if (SubscribeFolders == null)
                    Subscription = exchangeService.SubscribeToStreamingNotificationsOnAllFolders(SubscribeEvents);
                else
                    Subscription = exchangeService.SubscribeToStreamingNotifications(SubscribeFolders, SubscribeEvents);

                // Extract the X-BackEndOverrideCookie
                if (exchangeService.HttpResponseHeaders.ContainsKey("Set-Cookie"))
                {
                    string cookies = exchangeService.HttpResponseHeaders["Set-Cookie"];
                    if (cookies.Contains(cookieName))
                    {
                        int cookieStart = cookies.IndexOf(cookieName) + cookieName.Length+1;
                        int cookieEnd = cookies.IndexOf(";", cookieStart);
                        string cookie = cookies.Substring(cookieStart, cookieEnd - cookieStart);
                        Logger.DefaultLogger?.Log($"{cookieName} received for {SMTPAddress}: {cookie}");
                        Connection.SetBackendOverrideCookie(cookie);
                    }
                }

                _subscriptionIdToMailbox.Add(Subscription.Id, SMTPAddress);
                Logger.DefaultLogger?.Log($"Subscription created for {SMTPAddress}: {Subscription.Id}");
                TotalActiveSubscriptions++;
                return true;
            }
            catch (Exception ex)
            {
                Logger.DefaultLogger?.Log($"Failed to subscribe for {SMTPAddress}: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Unsubscribe from (release) this subscription
        /// </summary>
        public void Unsubscribe()
        {
            if (Subscription == null)
                return;

            Utils.SetClientRequestId(Connection.ExchangeService);
            Utils.CredentialHandler?.ApplyCredentialsToExchangeService(Connection.ExchangeService);

            try
            {
                if (_subscriptionIdToMailbox.ContainsKey(Subscription.Id))
                    _subscriptionIdToMailbox.Remove(Subscription.Id);
                Subscription.Unsubscribe();
                TotalActiveSubscriptions--;
                Logger.DefaultLogger?.Log($"Unsubscribed from {SMTPAddress}: {Subscription.Id}");
            }
            catch (Exception ex)
            {
                Logger.DefaultLogger?.Log($"Failed to unsubscribe for {SMTPAddress}: {ex.Message}");
            }
        }

        /// <summary>
        /// Reset the subscription (nullify without unsubscribing)
        /// </summary>
        public void Reset()
        {
            if (Subscription != null && !String.IsNullOrEmpty(Subscription.Id))
                TotalActiveSubscriptions--;
            Subscription = null;
        }
    }
}

﻿/*
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
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;

namespace EWSStreamingNotificationSample
{
    public class Mailboxes
    {
        // This class handles the autodiscover and exposes information needed for grouping

        private Dictionary<string, MailboxInfo> _mailboxes;  // Stores information for each mailbox, as returned by autodiscover
        private AutodiscoverService _autodiscover;
        private Auth.CredentialHandler _credentialHandler;
        private ITraceListener _traceListener;
        private ClassLogger _logger;
        private bool _useGrouping = true;
        //private string _lastKnownAutodiscoverUrl = "";


        public Mailboxes(ClassLogger Logger, ITraceListener TraceListener = null, Auth.CredentialHandler CredentialHandler = null)
        {
            _logger = Logger;
            _credentialHandler = CredentialHandler;
            _mailboxes = new Dictionary<string, MailboxInfo>();
            _traceListener = TraceListener;
            CreateAutodiscoverService();
        }

        private void CreateAutodiscoverService()
        {
            _autodiscover = new AutodiscoverService(ExchangeVersion.Exchange2013);  // Minimum version we need is 2013

            _autodiscover.RedirectionUrlValidationCallback = RedirectionCallback;
            if (_traceListener != null)
            {
                _autodiscover.TraceListener = _traceListener;
                _autodiscover.TraceFlags = TraceFlags.All;
                _autodiscover.TraceEnabled = true;
            }
            if (CredentialHandler != null)
            {
                _credentialHandler = CredentialHandler;
                _credentialHandler.ApplyCredentialsToAutodiscoverService(_autodiscover);
            }
        }

        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        public Auth.CredentialHandler CredentialHandler
        {
            get { return _credentialHandler; }
            set {
                _credentialHandler = value;
                _credentialHandler.ApplyCredentialsToAutodiscoverService(_autodiscover);
            }
        }


        public List<string> AllMailboxes
        {
            get
            {
                return _mailboxes.Keys.ToList<string>();
            }
        }

        public bool GroupMailboxes
        {
            get { return _useGrouping; }
            set { _useGrouping = value; }
        }

        public bool AddMailbox(string SMTPAddress, string AutodiscoverURL = "")
        {
            // Perform autodiscover for the mailbox and store the information

            if (_mailboxes.ContainsKey(SMTPAddress))
            {
                // We already have autodiscover information for this mailbox, if it is recent enough we don't bother retrieving it again
                if (_mailboxes[SMTPAddress].IsStale)
                    _mailboxes.Remove(SMTPAddress);
                else
                    return true;
            }

            MailboxInfo info = null;

            // Retrieve the autodiscover information
            _logger.Log($"Retrieving user settings for {SMTPAddress}");
            GetUserSettingsResponse userSettings = null;
            if (!String.IsNullOrEmpty(AutodiscoverURL))
            {
                // Use the supplied Autodiscover URL
                try
                {
                    _autodiscover.Url = new Uri(AutodiscoverURL);
                    userSettings = _autodiscover.GetUserSettings(SMTPAddress, UserSettingName.InternalEwsUrl, UserSettingName.ExternalEwsUrl, UserSettingName.GroupingInformation);
                }
                catch
                {
                }
            }
                
            if (userSettings == null)
            {
                try
                {
                    // Try full autodiscover
                    userSettings = GetUserSettings(SMTPAddress);
                }
                catch (Exception ex)
                {
                    _logger.Log(String.Format("Failed to autodiscover for {0}: {1}", SMTPAddress, ex.Message));
                    return false;
                }
            }

            // Store the autodiscover result, and check that we have what we need for subscriptions
            info = new MailboxInfo(SMTPAddress, userSettings);
            

            if (!info.HaveSubscriptionInformation)
            {
                _logger.Log(String.Format("Autodiscover succeeded, but EWS Url was not returned for {0}", SMTPAddress));
                return false;
            }

            // Add the mailbox to our list, and if it will be part of a new group add that to the group list (with this mailbox as the primary mailbox)
            _mailboxes.Add(info.SMTPAddress, info);
            return true;
        }

        private GetUserSettingsResponse GetUserSettings(string Mailbox, string startUrl = "")
        {
            // Attempt autodiscover, with maximum of 10 hops
            // As per docs: https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.autodiscover.autodiscoverservice.getusersettings?view=exchange-ews-api

            Uri url = null;
            GetUserSettingsResponse response = null;
            if (!String.IsNullOrEmpty(startUrl))
                url = new Uri(startUrl);

            if (!_credentialHandler.ApplyCredentialsToAutodiscoverService(_autodiscover))
                throw new Exception("Failed to apply credentials to Autodiscover service");
                

            for (int attempt = 0; attempt < 10; attempt++)
            {
                if (url != null)
                    _autodiscover.Url = url;
                _autodiscover.EnableScpLookup = false;// (attempt < 2);

                response = _autodiscover.GetUserSettings(Mailbox, UserSettingName.InternalEwsUrl, UserSettingName.ExternalEwsUrl, UserSettingName.GroupingInformation);

                if (response.ErrorCode == AutodiscoverErrorCode.RedirectAddress)
                {
                    // Redirecting to different mail address (can occur in hybrid)
                    return GetUserSettings(response.RedirectTarget);
                }
                else if (response.ErrorCode == AutodiscoverErrorCode.RedirectUrl)
                {
                    // Redirecting to another AutoDiscover Url
                    url = new Uri(response.RedirectTarget);
                }
                else
                {
                    _logger.Log($"Autodiscover Url: {_autodiscover.Url}");
                    return response;
                }
            }

            throw new Exception("No suitable Autodiscover endpoint was found.");
        }

        public MailboxInfo Mailbox(string SMTPAddress)
        {
            if (_mailboxes.ContainsKey(SMTPAddress))
                return _mailboxes[SMTPAddress];
            return null;
        }
    }
}

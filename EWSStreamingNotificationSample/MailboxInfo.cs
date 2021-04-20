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
using Microsoft.Exchange.WebServices.Autodiscover;

namespace EWSStreamingNotificationSample
{
    public class MailboxInfo
    {
        // This class stores the autodiscover information for a mailbox and exposes the subscription/grouping data

        private DateTime _timeInfoSet; // Set when the class is instantiated, so that we know when we obtained the information
        public string SMTPAddress { get; private set; }
        public string EwsUrl { get; private set; }
        public string GroupingInformation { get; private set; }
        public string Watermark { get; set; }
        
        public MailboxInfo()
        {
            _timeInfoSet = DateTime.Now;
        }

        public MailboxInfo(string Mailbox, GetUserSettingsResponse UserSettings):this()
        {
            SMTPAddress = Mailbox;
            try
            {
                EwsUrl = (string)UserSettings.Settings[UserSettingName.ExternalEwsUrl];
            }
            catch
            {
                try
                {
                    EwsUrl = (string)UserSettings.Settings[UserSettingName.InternalEwsUrl];
                }
                catch { }
            }
            try
            {
                GroupingInformation = (string)UserSettings.Settings[UserSettingName.GroupingInformation];
            }
            catch { }
            if ( String.IsNullOrEmpty(GroupingInformation) )
                GroupingInformation="all";
        }

        public MailboxInfo(string Mailbox, string EWSUri, string Group = "all"):this()
        {
            SMTPAddress = Mailbox;
            EwsUrl = EWSUri;
            GroupingInformation = Group;
        }

        public bool HaveSubscriptionInformation
        {
            get { return (!String.IsNullOrEmpty(EwsUrl) && !String.IsNullOrEmpty(GroupingInformation));  }
        }

        public string GroupName
        {
            get { return String.Format("{1}{0}", EwsUrl, GroupingInformation); }
        }

        public bool IsStale
        {
            // We assume that the information is stale if it is over 24 hours old
            get { return DateTime.Now.Subtract(_timeInfoSet) > new TimeSpan(0, 24, 0, 0);  }
        }
    }
}

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

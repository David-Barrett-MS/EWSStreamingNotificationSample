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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Net;
using EWS = Microsoft.Exchange.WebServices.Data;

namespace EWSStreamingNotificationSample
{
    struct NotificationInfo
    {
        public string Mailbox;
        public object Event;
        public EWS.ExchangeService Service;
    }

    public partial class FormMain : Form
    {
        /// <summary>
        /// Dictionary of all SubscriptionTrackers, indexed by mailbox address
        /// </summary>
        private Dictionary<string, SubscriptionTracker> _subscriptions = new Dictionary<string,SubscriptionTracker>();

        /// <summary>
        /// List of any subscriptions that need recreating due to fatal error
        /// </summary>
        private List<string> _subscriptionsInError = new List<string>();
        private Logger _logger;
        /// <summary>
        /// Count of item notifications received
        /// </summary>
        private long _itemNotificationsReceived = 0;
        /// <summary>
        /// Count of folder notifications received
        /// </summary>
        private long _folderNotificationsReceived = 0;
        private object _subscriptionsLock = new object();
        private bool _initialising = false;

        public FormMain()
        {
            InitializeComponent();

            // Create our logger
            _logger = new Logger("Notifications.log");
            _logger.LogAdded += _logger_LogAdded;
            Utils.CreateTraceListener("Trace.log");

            // Increase default connection limit - this is CRUCIAL when supporting multiple subscriptions, as otherwise only 2 connections will work
            ServicePointManager.DefaultConnectionLimit = 255;

            // Set default UI options
            comboBoxSubscribeTo.SelectedIndex = 0; // Select All Folders
            buttonUnsubscribe.Enabled = false;
            checkBoxSelectAll.CheckState = CheckState.Checked;
            checkBoxSelectAll_CheckedChanged(this, null);

            ReadMailboxes();

            UpdateUriUI();
            UpdateAuthUI();
        }

        /// <summary>
        /// Update UI dependent upon AutoDiscover options
        /// </summary>
        private void UpdateUriUI()
        {
            textBoxEWSUri.Enabled = !radioButtonAutodiscover.Checked;
            if (radioButtonOffice365.Checked)
                textBoxEWSUri.Text = "https://outlook.office365.com/EWS/Exchange.asmx";
            textBoxEWSUri.ReadOnly = !radioButtonSpecificUri.Checked;
        }

        /// <summary>
        /// Update Auth UI elements
        /// </summary>
        private void UpdateAuthUI()
        {
            textBoxAuthCertificate.Enabled = radioButtonAuthWithCertificate.Checked;
            buttonSelectCertificate.Enabled = radioButtonAuthWithCertificate.Checked;
            textBoxClientSecret.Enabled = radioButtonAuthWithClientSecret.Checked;

            textBoxTenantId.Enabled = radioButtonAuthOAuth.Checked;
            textBoxApplicationId.Enabled = radioButtonAuthOAuth.Checked;
            groupBoxOAuthAuthMethod.Enabled = radioButtonAuthOAuth.Checked;

            textBoxUsername.Enabled = radioButtonAuthBasic.Checked;
            textBoxPassword.Enabled = radioButtonAuthBasic.Checked;
            textBoxDomain.Enabled = radioButtonAuthBasic.Checked;
        }

        /// <summary>
        /// Event handler to update ListBox with new log entries
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="a"></param>
        private void _logger_LogAdded(object sender, LoggerEventArgs a)
        {
            ShowEvent(a.LogDetails);
        }

        /// <summary>
        /// Read configuration (mailboxes to subscribe to and potentially auth) into the UI
        /// </summary>
        /// <param name="MailboxFile">File containing list of mailboxes (and optional auth)</param>
        private void ReadMailboxes(string MailboxFile = "")
        {
            // We read config from a file.  If no filename provided, we check for one for this machine name
            string sMailboxFile = MailboxFile;
            if (String.IsNullOrEmpty(MailboxFile))
                sMailboxFile = "Mailboxes " + Environment.MachineName.ToUpper() + ".txt";
            if (!System.IO.File.Exists(sMailboxFile))
                return;

            checkedListBoxMailboxes.Items.Clear();
            using (System.IO.StreamReader reader = new System.IO.StreamReader(sMailboxFile))
            {
                while (!reader.EndOfStream)
                {
                    string sMailbox = reader.ReadLine();
                    if (sMailbox.ToLower().StartsWith("admin="))
                    {
                        textBoxUsername.Text = sMailbox.Substring(6);
                    }
                    else if (sMailbox.ToLower().StartsWith("password="))
                    {
                        textBoxPassword.Text = sMailbox.Substring(9);
                    }
                    else if (sMailbox.ToLower().StartsWith("tenantid="))
                    {
                        textBoxTenantId.Text = sMailbox.Substring(9);
                    }
                    else if (sMailbox.ToLower().StartsWith("appid="))
                    {
                        textBoxApplicationId.Text = sMailbox.Substring(6);
                    }
                    else if (sMailbox.ToLower().StartsWith("secret="))
                    {
                        textBoxClientSecret.Text = sMailbox.Substring(7);
                        radioButtonAuthOAuth.Checked = true;
                    }
                    else
                        checkedListBoxMailboxes.Items.Add(sMailbox);
                }
            }
            buttonSelectAllMailboxes_Click(this, null);
        }

        /// <summary>
        /// Create CredentialHandler object based on current configuration
        /// </summary>
        /// <returns></returns>
        private Auth.CredentialHandler CreateCredentialHandler()
        {
            // Instantiate the relevant CredentialHandler if not already done

            if (Utils.CredentialHandler != null)
                return Utils.CredentialHandler;

            if (radioButtonAuthBasic.Checked)
            {
                // We are using Basic auth (also covers NTLM, etc.)
                Utils.CredentialHandler = new Auth.CredentialHandler(Auth.AuthType.Basic, _logger);
                Utils.CredentialHandler.Username = textBoxUsername.Text;
                Utils.CredentialHandler.Password = textBoxPassword.Text;
                Utils.CredentialHandler.Domain = textBoxDomain.Text;
            }
            else
            {
                // OAuth
                Utils.CredentialHandler = new Auth.CredentialHandler(Auth.AuthType.OAuth, _logger);
                Utils.CredentialHandler.ApplicationId = textBoxApplicationId.Text;
                Utils.CredentialHandler.TenantId = textBoxTenantId.Text;
                if (radioButtonAuthWithClientSecret.Checked)
                    Utils.CredentialHandler.ClientSecret = textBoxClientSecret.Text;
                else if (textBoxAuthCertificate.Tag != null)
                    Utils.CredentialHandler.Certificate = (System.Security.Cryptography.X509Certificates.X509Certificate2)textBoxAuthCertificate.Tag;
            }
            return Utils.CredentialHandler;
        }

        /// <summary>
        /// Split the subscriptions into groups (each group will be one streaming connection)
        /// </summary>
        private void ConfigureSubscriptionGroups()
        {
            // First we work out the group for each subscription
            Dictionary<string,int> groupNameCount = new Dictionary<string,int>();
            foreach (string sMailbox in _subscriptions.Keys)
            {
                SubscriptionTracker sTracker = _subscriptions[sMailbox];
                sTracker.GroupName = $"{sTracker.GroupingInformation}{sTracker.EwsUrl}";
                if (groupNameCount.ContainsKey(sTracker.GroupName))
                {
                    if (groupNameCount[sTracker.GroupName] >= (int)numericUpDownMaxMailboxesInGroup.Value)
                    {
                        // We already have enough subscriptions in this group, so we create a new one
                        int subGroups = 2;
                        while (groupNameCount[sTracker.GroupName] >= (int)numericUpDownMaxMailboxesInGroup.Value)
                        {
                            sTracker.GroupName = $"{sTracker.GroupingInformation}{sTracker.EwsUrl}-{subGroups}";
                            if (!groupNameCount.ContainsKey(sTracker.GroupName))
                                groupNameCount.Add(sTracker.GroupName, 1);
                            subGroups++;
                        }
                    }
                    else
                        groupNameCount[sTracker.GroupName]++;
                }
                else
                    groupNameCount.Add(sTracker.GroupName, 1);
                _logger.Log($"{sTracker.GroupName}: {sTracker.SMTPAddress} added");
            }

            // Now create the ConnectionTracker, one per group, and configure the subscription
            // Subscriptions should be added in alphabetical order
            List<string> smtpAddresses = _subscriptions.Keys.ToList();
            smtpAddresses.Sort();
            for (int i=0; i<smtpAddresses.Count; i++)
            {
                SubscriptionTracker sTracker = _subscriptions[smtpAddresses[i]];
                if (!Utils.Connections.ContainsKey(sTracker.GroupName))
                {
                    ConnectionTracker cTracker = new ConnectionTracker(sTracker.GroupName);
                    cTracker.ConfigureSubscription(sTracker);
                    Utils.Connections.Add(sTracker.GroupName, cTracker);
                }
                else
                    Utils.Connections[sTracker.GroupName].ConfigureSubscription(sTracker);
            }
        }

        private void StreamingSubscriptionConnection_OnSubscriptionError(object sender, EWS.SubscriptionErrorEventArgs args)
        {
            // These errors all indicate that the subscription should be recreated (which can be done immediately, but not in this event handler)
            string[] resubscribeErrors = { "The specified subscription was not found.",
                "The subscription must be recreated.",
                "The Service instance doesn't have sufficient permissions to perform the request." };

            // Need to add delayed resubscribe mechanism for mailbox moves (not implemented) and temporary errors.
            // (AutoDiscover renewal will also be required):
            //  <m:MessageText>Mailbox move in progress. Try again later., Cannot open mailbox. Server = VI1PR04MB5440.eurprd04.prod.outlook.com, user = /o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=3693ebc023a445428b521a5a448d9dd0-113, maiboxGuid = f4d3b438-94bd-4449-b16e-cf9a7077c99e</m:MessageText>
            //  <m:ResponseCode>ErrorMailboxMoveInProgress</m:ResponseCode>
            //
            //  <m:MessageText>The mailbox database is temporarily unavailable., Cannot open mailbox. Server = HE1PR04MB3099.eurprd04.prod.outlook.com, user = /o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=eb9df0127efb4851a7563deae7dd6ad4-119, maiboxGuid = f1912d48-a20f-417d-a876-8dc322c1b406</m:MessageText>
            //  <m:ResponseCode>ErrorMailboxStoreUnavailable</m:ResponseCode>
            //
            // Handled error responses (immediately resubscribe):
            //
            //      <m:MessageText>The caller does not have sufficient permissions to perform the request., The Service instance doesn't have sufficient permissions to perform the request.</m:MessageText>
            //      <m:ResponseCode>ErrorProxyRequestNotAllowed</m:ResponseCode>
            //
            //      <m:MessageText>Unable to retrieve events for this subscription.  The subscription must be recreated., The events couldn't be read.</m:MessageText>
            //      <m:ResponseCode>ErrorReadEventsFailed</m:ResponseCode>
            //
            //      <m:MessageText>The specified subscription was not found.</m:MessageText>
            //      <m:ResponseCode>ErrorSubscriptionNotFound</m:ResponseCode>
            //
            // https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/handling-notification-related-errors-in-ews-in-exchange#recovering-from-lost-subscriptions
            //

            try
            {
                _logger.Log($"OnSubscriptionError received for {SubscriptionTracker.MailboxOwningSubscription(args.Subscription.Id)}: {args.Exception.Message}");                
                foreach (string error in resubscribeErrors)
                    if (args.Exception.Message.Contains(error))
                    {
                        _subscriptionsInError.Add(args.Subscription.Id);
                        break;
                    }
            }
            catch
            {
                _logger.Log("OnSubscriptionError received");
            }
        }


        /// <summary>
        /// Handle incoming EWS notifications
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void StreamingSubscriptionConnection_OnNotificationEvent(object sender, EWS.NotificationEventArgs args)
        {
            foreach (EWS.NotificationEvent e in args.Events)
                ProcessNotification(e, args.Subscription);
        }

        /// <summary>
        /// Process the received EWS notification
        /// </summary>
        /// <param name="e"></param>
        /// <param name="Subscription"></param>
        void ProcessNotification(object e, EWS.StreamingSubscription Subscription)
        {
            // We have the subscription Id, so we need to get the mailbox
            string sMailbox = SubscriptionTracker.MailboxOwningSubscription(Subscription.Id);

            string sEvent = sMailbox + ": ";

            if (e is EWS.ItemEvent)
            {
                _itemNotificationsReceived++;
                if (!checkBoxShowItemEvents.Checked) return; // We're ignoring item events
                sEvent += "Item " + (e as EWS.ItemEvent).EventType.ToString() + ": ";
            }
            else if (e is EWS.FolderEvent)
            {
                _folderNotificationsReceived++;
                if (!checkBoxShowFolderEvents.Checked) return; // We're ignoring folder events
                sEvent += "Folder " + (e as EWS.FolderEvent).EventType.ToString() + ": ";
            }

            try
            {
                if (checkBoxQueryMore.Checked)
                {
                    // We want more information, we'll get this on a new thread
                    NotificationInfo info;
                    info.Mailbox = sMailbox;
                    info.Event = e;
                    info.Service = null;

                    ThreadPool.QueueUserWorkItem(new WaitCallback(ShowMoreInfo), info);
                }
                else
                {
                    // Just log event and ID
                    if (e is EWS.ItemEvent)
                    {
                        sEvent += "ItemId = " + (e as EWS.ItemEvent).ItemId.UniqueId;
                    }
                    else if (e is EWS.FolderEvent)
                    {
                        sEvent += "FolderId = " + (e as EWS.FolderEvent).FolderId.UniqueId;
                    }
                }
            }
            catch { }
            UpdateStats();

            if (checkBoxQueryMore.Checked)
                return; // The event is being handled in a worker thread

            ShowEvent(sEvent);
        }

        /// <summary>
        /// Returns the EWS URL for the given mailbox
        /// </summary>
        /// <param name="Mailbox">Mailbox for which EWS URL is required</param>
        /// <returns>EWS URL</returns>
        public string MailboxEWSUrl(string Mailbox)
        {
            if (!_subscriptions.ContainsKey(Mailbox))
                return null;
            return _subscriptions[Mailbox].EwsUrl;
        }

        /// <summary>
        /// Get more info for the given item or folder.
        /// </summary>
        /// <param name="e"></param>
        void ShowMoreInfo(object e)
        {
            NotificationInfo n = (NotificationInfo)e;

            EWS.ExchangeService ewsMoreInfoService = Utils.NewExchangeService(n.Mailbox, MailboxEWSUrl(n.Mailbox), n.Mailbox);

            string sEvent = "";
            if (n.Event is EWS.ItemEvent)
            {
                sEvent = $"{n.Mailbox}: Item {(n.Event as EWS.ItemEvent).EventType}: {MoreItemInfo(n.Event as EWS.ItemEvent, ewsMoreInfoService)}";
            }
            else
                sEvent = $"{n.Mailbox}: Folder {(n.Event as EWS.FolderEvent).EventType}: {MoreFolderInfo(n.Event as EWS.FolderEvent, ewsMoreInfoService)}";

            ShowEvent(sEvent);
        }

        /// <summary>
        /// Retrieve further information for the given item Id
        /// </summary>
        /// <param name="e">ItemEvent containing Item Id</param>
        /// <param name="service">ExchangeService object to be used for EWS requests</param>
        /// <returns></returns>
        private string MoreItemInfo(EWS.ItemEvent e, EWS.ExchangeService service)
        {
            string sMoreInfo = "";

            if (service != null)
            {
                if (e.EventType == EWS.EventType.Deleted)
                {
                    // We cannot get more info for a deleted item by binding to it, so skip item details
                }
                else
                    sMoreInfo += GetItemInfo(e.ItemId, service);
                if (e.ParentFolderId != null)
                {
                    if (!String.IsNullOrEmpty(sMoreInfo)) sMoreInfo += ", ";
                    sMoreInfo += "Parent Folder Name=" + GetFolderName(e.ParentFolderId, service);
                }
            }
            return sMoreInfo;
        }

        /// <summary>
        /// Retrieve further information for the given folder Id
        /// </summary>
        /// <param name="e">ItemEvent containing Folder Id</param>
        /// <param name="service">ExchangeService object to be used for EWS requests</param>
        /// <returns></returns>
        private string MoreFolderInfo(EWS.FolderEvent e, EWS.ExchangeService service)
        {
            // Retrieve some more information about the specified folder

            string sMoreInfo = "";
            if (e.EventType == EWS.EventType.Deleted)
            {
                // We cannot get more info for a deleted item by binding to it, so skip item details
            }
            else
                sMoreInfo += "Folder name=" + GetFolderName(e.FolderId, service);
            if (e.ParentFolderId != null)
            {
                if (!String.IsNullOrEmpty(sMoreInfo)) sMoreInfo += ", ";
                sMoreInfo += "Parent Folder Name=" + GetFolderName(e.ParentFolderId, service);
            }
            return sMoreInfo;
        }

        /// <summary>
        /// Perform the EWS requests to obtain further information for the given item
        /// </summary>
        /// <param name="itemId"></param>
        /// <param name="service"></param>
        /// <returns>string containing further item details (e.g. subject, etc.)</returns>
        private string GetItemInfo(EWS.ItemId itemId, EWS.ExchangeService service)
        {
            string sItemInfo = "";
            EWS.Item oItem;
            EWS.PropertySet oPropertySet;

            if (checkBoxIncludeMime.Checked)
            {
                oPropertySet = new EWS.PropertySet(EWS.BasePropertySet.FirstClassProperties, EWS.ItemSchema.MimeContent);
            }
            else
                oPropertySet = new EWS.PropertySet(EWS.ItemSchema.Subject);

            try
            {
                EWS.ItemId cleanItemId = new EWS.ItemId(itemId.UniqueId);
                Utils.SetClientRequestId(service);
                Utils.CredentialHandler.ApplyCredentialsToExchangeService(service);
                oItem = EWS.Item.Bind(service, cleanItemId, oPropertySet);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            if (oItem is EWS.Appointment)
            {
                sItemInfo += "Appointment subject=" + oItem.Subject;
                // Show attendee information
                EWS.Appointment oAppt = oItem as EWS.Appointment;
                sItemInfo += ",RequiredAttendees=" + GetAttendees(oAppt.RequiredAttendees);
                sItemInfo += ",OptionalAttendees=" + GetAttendees(oAppt.OptionalAttendees);
            }
            else
                sItemInfo += "Item subject=" + oItem.Subject;
            if (checkBoxIncludeMime.Checked)
                sItemInfo += ", MIME length=" + oItem.MimeContent.Content.Length.ToString() + " bytes";
            return sItemInfo;
        }

        /// <summary>
        /// Read and return a string listing the attendees from the given collection
        /// </summary>
        /// <param name="attendees">collection of attendees</param>
        /// <returns>Comma separated list of attendees</returns>
        private string GetAttendees(EWS.AttendeeCollection attendees)
        {
            if (attendees.Count == 0) return "none";

            string sAttendees = "";
            foreach (EWS.Attendee attendee in attendees)
            {
                if (!String.IsNullOrEmpty(sAttendees))
                    sAttendees += ", ";
                sAttendees += attendee.Name;
            }

            return sAttendees;
        }

        /// <summary>
        /// Retrieve the folder name for the given folder Id
        /// </summary>
        /// <param name="folderId">Id of the folder to retrieve</param>
        /// <param name="service">ExchangeService object through which EWS request will be sent</param>
        /// <returns>DisplayName of the folder</returns>
        private string GetFolderName(EWS.FolderId folderId, EWS.ExchangeService service)
        {
            // Retrieve display name of the given folder
            try
            {
                Utils.SetClientRequestId(service);
                Utils.CredentialHandler.ApplyCredentialsToExchangeService(service);
                EWS.Folder oFolder = EWS.Folder.Bind(service, folderId, new EWS.PropertySet(EWS.FolderSchema.DisplayName));
                return oFolder.DisplayName;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        /// <summary>
        /// Add the given event to the events listbox
        /// </summary>
        /// <param name="eventDetails"></param>
        private void ShowEvent(string eventDetails)
        {
            Action action = new Action(() =>
            {
                listBoxEvents.Items.Add(eventDetails);
                while (listBoxEvents.Items.Count > 1000)
                    listBoxEvents.Items.RemoveAt(0);
                if (listBoxEvents.SelectedIndex<0)
                    listBoxEvents.TopIndex = listBoxEvents.Items.Count - 1;
            });
            if (listBoxEvents.InvokeRequired)
                listBoxEvents.Invoke(action);
            else
                action();
        }

        /// <summary>
        /// Update UI with current connection and subscription stats
        /// </summary>
        private void UpdateStats()
        {
            Action action = new Action(() =>
            {
                textBoxNumConnections.Text = $"{ConnectionTracker.TotalOpenConnections}/{Utils.Connections.Count}";
            });
            if (textBoxNumConnections.InvokeRequired)
                textBoxNumConnections.Invoke(action);
            else
                action();

            action = new Action(() =>
            {
                textBoxNumSubscriptions.Text = $"{SubscriptionTracker.TotalActiveSubscriptions}/{_subscriptions.Count}";
            });
            if (textBoxNumSubscriptions.InvokeRequired)
                textBoxNumSubscriptions.Invoke(action);
            else
                action();

            action = new Action(() =>
            {
                textBoxNotificationCount.Text = $"{_folderNotificationsReceived}/{_itemNotificationsReceived}";
            });
            if (textBoxNotificationCount.InvokeRequired)
                textBoxNotificationCount.Invoke(action);
            else
                action();
        }

        /// <summary>
        /// Create the subscriptions for each mailbox
        /// </summary>
        private void InitialiseSubscriptions()
        {
            string autodiscoverUrl = "";
            if (radioButtonOffice365.Checked)
                autodiscoverUrl = "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc";
            List<Task> initTasks = new List<Task>();

            DateTime timeTrack = DateTime.Now;
            foreach (string sMailbox in checkedListBoxMailboxes.CheckedItems)
            {
                if (!_subscriptions.ContainsKey(sMailbox))
                {
                    Action action = new Action(() =>
                    {
                        SubscriptionTracker subscriptionTracker = SubscriptionTracker.CreateTrackerForMailbox(sMailbox, autodiscoverUrl);
                        if (subscriptionTracker != null)
                            lock (_subscriptionsLock)
                                _subscriptions.Add(sMailbox, subscriptionTracker);
                    });
                    initTasks.Add(Task.Run(() => action()));
                }
            }

            int i = 0;
            while (i < initTasks.Count)
            {
                while (initTasks[i].Status == TaskStatus.Running || initTasks[i].Status == TaskStatus.WaitingToRun)
                    Thread.Yield();
                i++;
            }
            TimeSpan subscriptionInitTime = DateTime.Now - timeTrack;
            _logger.Log($"Autodiscover completed in {subscriptionInitTime.TotalSeconds} seconds");
            
            timeTrack = DateTime.Now;
            ConfigureSubscriptionGroups();
            _logger.Log($"Group configuration completed in {(DateTime.Now - timeTrack).TotalSeconds} seconds");
        }

        /// <summary>
        /// Create the subscriptions
        /// </summary>
        private void CreateSubscriptions()
        {
            DateTime timeTrack = DateTime.Now;
            List<Task> subscribeTasks = new List<Task>();
            EWS.EventType[] selectedEvents = SelectedEvents();
            EWS.FolderId[] selectedFolders = SelectedFolders();
            List<string> subscriptionMailboxes = _subscriptions.Keys.ToList<string>();
            subscriptionMailboxes.Sort();

            foreach (ConnectionTracker cTracker in Utils.Connections.Values)
            {
                Action action = new Action(() =>
                {
                    for (int s=0; s< subscriptionMailboxes.Count; s++)
                    {
                        SubscriptionTracker sTracker = _subscriptions[subscriptionMailboxes[s]];
                        if (sTracker.GroupName==cTracker.GroupName)
                        {
                            try
                            {
                                sTracker.Subscribe(selectedEvents, selectedFolders, sTracker.AnchorMailbox);
                                if (sTracker.Connection.AddSubscription(sTracker))
                                    if (cTracker.AnchorMailbox == sTracker.SMTPAddress)
                                    {
                                        //  Anchor mailbox, we hook into subscription events at this point (i.e. only once)
                                        sTracker.Connection.StreamingSubscriptionConnection.OnNotificationEvent += StreamingSubscriptionConnection_OnNotificationEvent;
                                        sTracker.Connection.StreamingSubscriptionConnection.OnSubscriptionError += StreamingSubscriptionConnection_OnSubscriptionError;
                                    }
                            }
                            catch (Exception ex)
                            {
                                _logger.Log($"Error subscribing for {sTracker.SMTPAddress}: {ex.Message}");
                            }
                        }
                    }
                    UpdateStats();
                });
                subscribeTasks.Add(Task.Run(() => action()));
            }

            int i = 0;
            while (i < subscribeTasks.Count)
            {
                while (subscribeTasks[i].Status == TaskStatus.Running || subscribeTasks[i].Status == TaskStatus.WaitingToRun)
                    Thread.Yield();
                i++;
            }

            _logger.Log($"Subscription creation completed in {(DateTime.Now - timeTrack).TotalSeconds} seconds");
        }

        /// <summary>
        /// Connect to the subscriptions and start receiving events
        /// </summary>
        private void ConnectSubscriptions()
        {
            DateTime timeTrack = DateTime.Now;
            _logger.Log($"Opening connections");
            List<Task> connectTasks = new List<Task>();

            foreach (ConnectionTracker cTracker in Utils.Connections.Values)
            {
                Action action = new Action(() =>
                {
                    if (!cTracker.IsConnected)
                        cTracker.Connect();
                    UpdateStats();
                });
                connectTasks.Add(Task.Run(() => action()));
            }

            int i = 0;
            while (i < connectTasks.Count)
            {
                while (connectTasks[i].Status == TaskStatus.Running || connectTasks[i].Status == TaskStatus.WaitingToRun)
                    Thread.Yield();
                i++;
            }

            _logger.Log($"All connections opened in {(DateTime.Now - timeTrack).TotalSeconds} seconds");
        }

        /// <summary>
        /// Disconnect from all subscriptions (does not unsubscribe)
        /// </summary>
        private void DisconnectSubscriptions()
        {
            DateTime timeTrack = DateTime.Now;
            List<Task> disconnectTasks = new List<Task>();

            foreach (ConnectionTracker cTracker in Utils.Connections.Values)
            {
                Action action = new Action(() =>
                {
                    cTracker.Disconnect();
                    UpdateStats();
                });
                disconnectTasks.Add(Task.Run(() => action()));
            }

            int i = 0;
            while (i < disconnectTasks.Count)
            {
                while (disconnectTasks[i].Status == TaskStatus.Running || disconnectTasks[i].Status == TaskStatus.WaitingToRun)
                    Thread.Yield();
                i++;
            }

            _logger.Log($"All connections closed in {(DateTime.Now - timeTrack).TotalSeconds} seconds");
        }

        /// <summary>
        /// Unsubscribe and nullify all subscriptions
        /// </summary>
        private void UnsubscribeSubscriptions()
        {
            DateTime timeTrack = DateTime.Now;
            List<Task> unsubscribeTasks = new List<Task>();

            foreach (ConnectionTracker cTracker in Utils.Connections.Values)
                if (cTracker.StreamingSubscriptionConnection != null)
                    cTracker.StreamingSubscriptionConnection.OnSubscriptionError -= StreamingSubscriptionConnection_OnSubscriptionError;

            foreach (SubscriptionTracker sTracker in _subscriptions.Values)
            {
                Action action = new Action(() =>
                {
                    sTracker.Unsubscribe();
                    UpdateStats();
                });
                unsubscribeTasks.Add(Task.Run(() => action()));
            }

            int i = 0;
            while (i < unsubscribeTasks.Count)
            {
                while (unsubscribeTasks[i].Status == TaskStatus.Running || unsubscribeTasks[i].Status == TaskStatus.WaitingToRun)
                    Thread.Yield();
                i++;
            }

            _logger.Log($"All subscriptions unsubscribed in {(DateTime.Now - timeTrack).TotalSeconds} seconds");
        }

        /// <summary>
        /// Recreate any broken subscriptions and/or connections
        /// </summary>
        private void RecoverBrokenSubscriptions()
        {
            EWS.EventType[] selectedEvents = SelectedEvents();
            EWS.FolderId[] selectedFolders = SelectedFolders();

            List<string> groupsToReconnect = new List<string>();
            while (_subscriptionsInError.Count>0)
            {
                SubscriptionTracker sTracker = _subscriptions[SubscriptionTracker.MailboxOwningSubscription(_subscriptionsInError[0])];
                if (!groupsToReconnect.Contains(sTracker.GroupName))
                {
                    groupsToReconnect.Add(sTracker.GroupName);
                    sTracker.Connection.Disconnect();
                }

                try
                {
                    sTracker.Connection.StreamingSubscriptionConnection.RemoveSubscription(sTracker.Subscription);
                }
                catch { }
                sTracker.Reset();
                sTracker.Subscribe(selectedEvents, selectedFolders, sTracker.AnchorMailbox);
                sTracker.Connection.AddSubscription(sTracker);
                _subscriptionsInError.RemoveAt(0);
            }
            UpdateStats();

            foreach (string groupName in groupsToReconnect)
            {
                Utils.Connections[groupName].Connect();
            }
            UpdateStats();
        }

        /// <summary>
        /// Check each of the connections, and if disconnected, connect it
        /// </summary>
        private void EnsureConnectionsAreOpen()
        {
            foreach (ConnectionTracker cTracker in Utils.Connections.Values)
                if (!cTracker.IsConnected)
                    cTracker.Connect();
            UpdateStats();
        }

        /// <summary>
        /// Read the currently selected events (to be subscribed to) from the UI
        /// </summary>
        /// <returns>Array of selected events</returns>
        private EWS.EventType[] SelectedEvents()
        {
            // Read the selected events and return as an array

            EWS.EventType[] events = null;
            Action action = new Action(() =>
            {
                if (checkedListBoxEvents.CheckedItems.Count > 0)
                {
                    events = new EWS.EventType[checkedListBoxEvents.CheckedItems.Count];

                    for (int i = 0; i < checkedListBoxEvents.CheckedItems.Count; i++)
                    {
                        switch (checkedListBoxEvents.CheckedItems[i].ToString())
                        {
                            case "NewMail": { events[i] = EWS.EventType.NewMail; break; }
                            case "Deleted": { events[i] = EWS.EventType.Deleted; break; }
                            case "Modified": { events[i] = EWS.EventType.Modified; break; }
                            case "Moved": { events[i] = EWS.EventType.Moved; break; }
                            case "Copied": { events[i] = EWS.EventType.Copied; break; }
                            case "Created": { events[i] = EWS.EventType.Created; break; }
                            case "FreeBusyChanged": { events[i] = EWS.EventType.FreeBusyChanged; break; }
                        }
                    }
                }
            });

            if (comboBoxSubscribeTo.InvokeRequired)
                checkedListBoxEvents.Invoke(action);
            else
                action();

            return events;
        }

        /// <summary>
        /// Read the folders to be subscribed to from the UI
        /// </summary>
        /// <returns>Array of FolderIds</returns>
        private EWS.FolderId[] SelectedFolders()
        {
            string sSubscribeFolder = "";
            bool allFolders = false;
            Action action = new Action(() =>
            {
                if (comboBoxSubscribeTo.SelectedIndex < 1)
                    allFolders = true; // Subscribe to all folders

                sSubscribeFolder = comboBoxSubscribeTo.SelectedItem.ToString();

            });
            if (comboBoxSubscribeTo.InvokeRequired)
                comboBoxSubscribeTo.Invoke(action);
            else
                action();

            if (allFolders)
                return null;

            EWS.FolderId[] folders = new EWS.FolderId[1];


            switch (sSubscribeFolder)
            {
                case "Calendar":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Calendar); break;

                case "Contacts":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Contacts); break;

                case "DeletedItems":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.DeletedItems); break;

                case "Drafts":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Drafts); break;

                case "Inbox":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Inbox); break;

                case "Journal":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Journal); break;

                case "Notes":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Notes); break;

                case "Outbox":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Outbox); break;

                case "SentItems":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.SentItems); break;

                case "Tasks":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.Tasks); break;

                case "MsgFolderRoot":
                    folders[0] = new EWS.FolderId(EWS.WellKnownFolderName.MsgFolderRoot); break;

                case "All Folders":
                    folders[0] = new EWS.FolderId("AllFolders"); break;
            }
            return folders;
        }

        #region Control events

        private void checkBoxSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool bChecked = true;
            if (checkBoxSelectAll.CheckState == CheckState.Unchecked)
                bChecked = false;

            for (int i = 0; i < checkedListBoxEvents.Items.Count; i++)
                checkedListBoxEvents.SetItemChecked(i, bChecked);
            if (bChecked)
            {
                checkBoxSelectAll.CheckState = CheckState.Checked;
            }
            else
                checkBoxSelectAll.CheckState = CheckState.Unchecked;
        }

        private void buttonSelectAllMailboxes_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxMailboxes.Items.Count; i++)
                checkedListBoxMailboxes.SetItemChecked(i, true);
        }

        private void buttonDeselectAllMailboxes_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxMailboxes.Items.Count; i++)
                checkedListBoxMailboxes.SetItemChecked(i, false);
        }

        private void buttonLoadMailboxes_Click(object sender, EventArgs e)
        {
            OpenFileDialog oDialog = new OpenFileDialog();
            oDialog.Filter = "Text files (*.txt)|*.txt|All Files|*.*";
            oDialog.DefaultExt = "txt";
            oDialog.Title = "Select mailbox file";
            oDialog.CheckFileExists = true;
            if (oDialog.ShowDialog() != DialogResult.OK)
                return;

            ReadMailboxes(oDialog.FileName);
        }

        private void buttonEditMailboxes_Click(object sender, EventArgs e)
        {
            // Allow the list of mailboxes to be edited

            StringBuilder allMailboxes = new StringBuilder();
            foreach (object mbx in checkedListBoxMailboxes.Items)
            {
                allMailboxes.AppendLine(mbx.ToString());
            }
            FormEditMailboxes form = new FormEditMailboxes();
            string mailboxes = form.EditMailboxes(allMailboxes.ToString());

            // Were there any changes?
            if (mailboxes.Equals(allMailboxes.ToString()))
                return;

            checkedListBoxMailboxes.Items.Clear();
            string[] mbxList = mailboxes.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string mbx in mbxList)
                checkedListBoxMailboxes.Items.Add(mbx);
        }

        private void radioButtonAutodiscover_CheckedChanged(object sender, EventArgs e)
        {
            UpdateUriUI();
        }

        private void radioButtonOffice365_CheckedChanged(object sender, EventArgs e)
        {
            UpdateUriUI();
        }

        private void radioButtonSpecificUri_CheckedChanged(object sender, EventArgs e)
        {
            UpdateUriUI();
        }

        private void radioButtonAuthOAuth_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAuthUI();
        }

        private void radioButtonAuthWithClientSecret_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAuthUI();
        }

        private void radioButtonAuthWithCertificate_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAuthUI();
        }

        private void radioButtonAuthBasic_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAuthUI();
        }

        private void buttonSubscribe_Click(object sender, EventArgs e)
        {
            // Create the subscriptions and connect

            _initialising = true;
            _itemNotificationsReceived = 0;
            _folderNotificationsReceived = 0;
            Utils.CredentialHandler = null; // We do this in case the credentials have been changed since last run
            CreateCredentialHandler();
            if (radioButtonAuthOAuth.Checked)
                Utils.CredentialHandler.AcquireToken();

            buttonSubscribe.Enabled = false;

            ConnectionTracker.ConnectionLifetime = (int)numericUpDownTimeout.Value;

            Action subscribeAction = new Action(() =>
            {
                InitialiseSubscriptions();
                CreateSubscriptions();
                ConnectSubscriptions();
                _initialising = false;
            });
            Task.Run(() => subscribeAction());
            
            buttonUnsubscribe.Enabled = true;
            timerMonitorConnections.Start();
        }

        private void buttonUnsubscribe_Click(object sender, EventArgs e)
        {
            buttonUnsubscribe.Enabled = false;

            Action unsubscribeAction = new Action(() =>
            {
                DisconnectSubscriptions();
                UnsubscribeSubscriptions();

            });
            Task.Run(() => unsubscribeAction());
            buttonSubscribe.Enabled = true;
        }

        private void timerMonitorConnections_Tick(object sender, EventArgs e)
        {
            if (buttonSubscribe.Enabled || _initialising)
                return;

            timerMonitorConnections.Stop();
            if (_subscriptionsInError.Count > 0)
                RecoverBrokenSubscriptions();

            EnsureConnectionsAreOpen();
            
            timerMonitorConnections.Start();
        }

        private void listBoxEvents_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxEvents.SelectedIndex < 0 || listBoxEvents.SelectedItems.Count>1)
                return;

            if (MessageBox.Show(this, listBoxEvents.Items[listBoxEvents.SelectedIndex].ToString(), "Event information", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                listBoxEvents.SelectedIndex = -1;
        }
    }
    #endregion
}

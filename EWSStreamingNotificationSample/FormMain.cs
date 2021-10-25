/*
 * By David Barrett, Microsoft Ltd. 2013-2021. Use at your own risk.  No warranties are given.
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
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;
using System.Threading;
using System.Net;

namespace EWSStreamingNotificationSample
{
    struct NotificationInfo
    {
        public string Mailbox;
        public object Event;
        public ExchangeService Service;
    }

    public partial class FormMain : Form
    {
        private ClassLogger _logger =null;
        private ClassTraceListener _traceListener = null;
        private Dictionary<string,StreamingSubscriptionConnection> _connections = null;
        private Dictionary<string, StreamingSubscription> _subscriptions = null;
        private Dictionary<String, String> _subscriptionIdToMailboxMapping = null;
        private Dictionary<string, GroupInfo> _groups = null;
        private Mailboxes _mailboxes = null;
        private bool _reconnect = false;
        private Object _reconnectLock = new Object();
        private Auth.CredentialHandler _credentialHandler = null;

        public FormMain()
        {
            InitializeComponent();

            // Create our logger
            _logger = new ClassLogger("Notifications.log");
            _logger.LogAdded += new ClassLogger.LoggerEventHandler(_logger_LogAdded);
            _traceListener = new ClassTraceListener("Trace.log");

            // Increase default connection limit - this is CRUCIAL when supporting multiple subscriptions, as otherwise only 2 connections will work
            ServicePointManager.DefaultConnectionLimit = 255;
            _logger.Log("Default connection limit increased to 255");

            comboBoxSubscribeTo.SelectedIndex = 0; // Select All Folders
            buttonUnsubscribe.Enabled = false;
            checkBoxSelectAll.CheckState = CheckState.Checked;
            checkBoxSelectAll_CheckedChanged(this, null);

            _connections = new Dictionary<string, StreamingSubscriptionConnection>();
            ReadMailboxes();

            _mailboxes = new Mailboxes(_logger, _traceListener);

            UpdateUriUI();
            UpdateAuthUI();
        }

        private void ReadMailboxes(string MailboxFile="")
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

        void _logger_LogAdded(object sender, LoggerEventArgs a)
        {
            try
            {
                if (listBoxEvents.InvokeRequired)
                {
                    // Need to invoke
                    listBoxEvents.Invoke(new MethodInvoker(delegate()
                    {
                        listBoxEvents.Items.Add(a.LogDetails);
                        listBoxEvents.SelectedIndex = listBoxEvents.Items.Count - 1;
                    }));
                }
                else
                {
                    listBoxEvents.Items.Add(a.LogDetails);
                    listBoxEvents.SelectedIndex = listBoxEvents.Items.Count - 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private ExchangeService ExchangeServiceForMailboxAccess(MailboxInfo Mailbox)
        {
            // Create an ExchangeService object for mailbox access (used to retrieve further information about notified items)

            ExchangeService mailboxAccessService = new ExchangeService(ExchangeVersion.Exchange2016);
            CredentialHandler().ApplyCredentialsToExchangeService(mailboxAccessService);
            mailboxAccessService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Mailbox.SMTPAddress);
            mailboxAccessService.HttpHeaders.Add("X-AnchorMailbox", Mailbox.SMTPAddress);
            mailboxAccessService.HttpHeaders.Add("client-request-id", Guid.NewGuid().ToString());
            mailboxAccessService.HttpHeaders.Add("return-client-request-id", "true");
            mailboxAccessService.Url = new Uri(Mailbox.EwsUrl);
            mailboxAccessService.TraceListener = _traceListener;
            mailboxAccessService.TraceFlags = TraceFlags.All;
            mailboxAccessService.TraceEnabled = true;
            return mailboxAccessService;
        }

        void ProcessNotification(object e, StreamingSubscription Subscription)
        {
            // We have received a notification

            // We have the subscription Id, so we need to get the mailbox
            string sMailbox = "Unknown mailbox"; 
            if (_subscriptionIdToMailboxMapping.ContainsKey(Subscription.Id))
                sMailbox = _subscriptionIdToMailboxMapping[Subscription.Id];

            string sEvent = sMailbox + ": ";

            if (e is ItemEvent)
            {
                if (!checkBoxShowItemEvents.Checked) return; // We're ignoring item events
                sEvent += "Item " + (e as ItemEvent).EventType.ToString() + ": ";
            }
            else if (e is FolderEvent)
            {
                if (!checkBoxShowFolderEvents.Checked) return; // We're ignoring folder events
                sEvent += "Folder " + (e as FolderEvent).EventType.ToString() + ": ";
            }

            try
            {
                if (checkBoxQueryMore.Checked)
                {
                    // We want more information, we'll get this on a new thread
                    NotificationInfo info;
                    info.Mailbox = sMailbox;
                    info.Event = e;
                    info.Service = Subscription.Service;
                    
                    ThreadPool.QueueUserWorkItem(new WaitCallback(ShowMoreInfo), info);
                }
                else
                {
                    // Just log event and ID
                    if (e is ItemEvent)
                    {
                        sEvent += "ItemId = " + (e as ItemEvent).ItemId.UniqueId;
                    }
                    else if (e is FolderEvent)
                    {
                        sEvent += "FolderId = " + (e as FolderEvent).FolderId.UniqueId;
                    }
                }
            }
            catch { }

            if (checkBoxQueryMore.Checked)
                return;

            ShowEvent(sEvent);
        }

        void ShowMoreInfo(object e)
        {
            // Get more info for the given item.  This will run on it's own thread
            // so that the main program can continue as usual (we won't hold anything up)

            NotificationInfo n = (NotificationInfo)e;

            ExchangeService ewsMoreInfoService = ExchangeServiceForMailboxAccess(_mailboxes.Mailbox(n.Mailbox));

            string sEvent = "";
            if (n.Event is ItemEvent)
            {
                sEvent = n.Mailbox + ": Item " + (n.Event as ItemEvent).EventType.ToString() + ": " + MoreItemInfo(n.Event as ItemEvent, ewsMoreInfoService);
            }
            else
                sEvent = n.Mailbox + ": Folder " + (n.Event as FolderEvent).EventType.ToString() + ": " + MoreFolderInfo(n.Event as FolderEvent, ewsMoreInfoService);

            ShowEvent(sEvent);
        }

        private string MoreItemInfo(ItemEvent e, ExchangeService service)
        {
            string sMoreInfo = "";
            if (e.EventType == EventType.Deleted)
            {
                // We cannot get more info for a deleted item by binding to it, so skip item details
            }
            else
                sMoreInfo += "Item subject=" + GetItemInfo(e.ItemId, service);
            if (e.ParentFolderId != null)
            {
                if (!String.IsNullOrEmpty(sMoreInfo)) sMoreInfo += ", ";
                sMoreInfo += "Parent Folder Name=" + GetFolderName(e.ParentFolderId, service);
            }
            return sMoreInfo;
        }

        private string MoreFolderInfo(FolderEvent e, ExchangeService service)
        {
            string sMoreInfo = "";
            if (e.EventType == EventType.Deleted)
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

        private string GetItemInfo(ItemId itemId, ExchangeService service)
        {
            // Retrieve the subject for a given item
            string sItemInfo = "";
            Item oItem;
            PropertySet oPropertySet;

            if (checkBoxIncludeMime.Checked)
            {
                oPropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.MimeContent);
            }
            else
                oPropertySet = new PropertySet(ItemSchema.Subject);

            try
            {
                ItemId cleanItemId = new ItemId(itemId.UniqueId);
                oItem = Item.Bind(service, cleanItemId, oPropertySet);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            if (oItem is Appointment)
            {
                sItemInfo += "Appointment subject=" + oItem.Subject;
                // Show attendee information
                Appointment oAppt=oItem as Appointment;
                sItemInfo += ",RequiredAttendees=" + GetAttendees(oAppt.RequiredAttendees);
                sItemInfo += ",OptionalAttendees=" + GetAttendees(oAppt.OptionalAttendees);
            }
            else
                sItemInfo += "Item subject=" + oItem.Subject;
            if (checkBoxIncludeMime.Checked)
                sItemInfo += ", MIME length=" + oItem.MimeContent.Content.Length.ToString() + " bytes";
            return sItemInfo;
        }

        private string GetAttendees(AttendeeCollection attendees)
        {
            if (attendees.Count == 0) return "none";

            string sAttendees = "";
            foreach (Attendee attendee in attendees)
            {
                if (!String.IsNullOrEmpty(sAttendees))
                    sAttendees += ", ";
                sAttendees += attendee.Name;
            }

            return sAttendees;
        }

        private string GetFolderName(FolderId folderId, ExchangeService service)
        {
            // Retrieve display name of the given folder
            try
            {
                Folder oFolder = Folder.Bind(service, folderId, new PropertySet(FolderSchema.DisplayName));
                return oFolder.DisplayName;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private void ShowEvent(string eventDetails)
        {
            try
            {
                if (listBoxEvents.InvokeRequired)
                {
                    // Need to invoke
                    listBoxEvents.Invoke(new MethodInvoker(delegate()
                    {
                        listBoxEvents.Items.Add(eventDetails);
                        listBoxEvents.SelectedIndex = listBoxEvents.Items.Count - 1;
                    }));
                }
                else
                {
                    listBoxEvents.Items.Add(eventDetails);
                    listBoxEvents.SelectedIndex = listBoxEvents.Items.Count - 1;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error");
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private EventType[] SelectedEvents()
        {
            // Read the selected events
            EventType[] events = null;
            if (comboBoxSubscribeTo.InvokeRequired)
            {
                checkedListBoxEvents.Invoke(new MethodInvoker(delegate()
                {
                    if (checkedListBoxEvents.CheckedItems.Count > 0)
                    {
                        events = new EventType[checkedListBoxEvents.CheckedItems.Count];

                        for (int i = 0; i < checkedListBoxEvents.CheckedItems.Count; i++)
                        {
                            switch (checkedListBoxEvents.CheckedItems[i].ToString())
                            {
                                case "NewMail": { events[i] = EventType.NewMail; break; }
                                case "Deleted": { events[i] = EventType.Deleted; break; }
                                case "Modified": { events[i] = EventType.Modified; break; }
                                case "Moved": { events[i] = EventType.Moved; break; }
                                case "Copied": { events[i] = EventType.Copied; break; }
                                case "Created": { events[i] = EventType.Created; break; }
                                case "FreeBusyChanged": { events[i] = EventType.FreeBusyChanged; break; }
                            }
                        }
                    }
                }));
            }
            else
            {
                if (checkedListBoxEvents.CheckedItems.Count < 1)
                    return null;
                events = new EventType[checkedListBoxEvents.CheckedItems.Count];

                for (int i = 0; i < checkedListBoxEvents.CheckedItems.Count; i++)
                {
                    switch (checkedListBoxEvents.CheckedItems[i].ToString())
                    {
                        case "NewMail": { events[i] = EventType.NewMail; break; }
                        case "Deleted": { events[i] = EventType.Deleted; break; }
                        case "Modified": { events[i] = EventType.Modified; break; }
                        case "Moved": { events[i] = EventType.Moved; break; }
                        case "Copied": { events[i] = EventType.Copied; break; }
                        case "Created": { events[i] = EventType.Created; break; }
                        case "FreeBusyChanged": { events[i] = EventType.FreeBusyChanged; break; }
                    }
                }
            }

            return events;
        }

        private void buttonSubscribe_Click(object sender, EventArgs e)
        {
            if (radioButtonAuthOAuth.Checked)
                CredentialHandler().AcquireToken();
            CreateGroups();

            if (ConnectToSubscriptions())
            {
                buttonUnsubscribe.Enabled = true;
                buttonSubscribe.Enabled = false;
                _logger.Log("Connected to subscription(s)");
            }
            else
                _logger.Log("Failed to create any subscriptions");
        }


        private FolderId[] SelectedFolders()
        {
            if (comboBoxSubscribeTo.SelectedIndex < 1)
                return null; // Subscribe to all folders

            FolderId[] folders = new FolderId[1];
            string sSubscribeFolder = "";
            if (comboBoxSubscribeTo.InvokeRequired)
            {
                comboBoxSubscribeTo.Invoke(new MethodInvoker(delegate()
                {
                    sSubscribeFolder = comboBoxSubscribeTo.SelectedItem.ToString();
                }));
            }
            else
                sSubscribeFolder = comboBoxSubscribeTo.SelectedItem.ToString();

            switch (sSubscribeFolder)
            {
                case "Calendar":
                    folders[0] = new FolderId(WellKnownFolderName.Calendar); break;

                case "Contacts":
                    folders[0] = new FolderId(WellKnownFolderName.Contacts); break;

                case "DeletedItems":
                    folders[0] = new FolderId(WellKnownFolderName.DeletedItems); break;

                case "Drafts":
                    folders[0] = new FolderId(WellKnownFolderName.Drafts); break;

                case "Inbox":
                    folders[0] = new FolderId(WellKnownFolderName.Inbox); break;

                case "Journal":
                    folders[0] = new FolderId(WellKnownFolderName.Journal); break;

                case "Notes":
                    folders[0] = new FolderId(WellKnownFolderName.Notes); break;

                case "Outbox":
                    folders[0] = new FolderId(WellKnownFolderName.Outbox); break;

                case "SentItems":
                    folders[0] = new FolderId(WellKnownFolderName.SentItems); break;

                case "Tasks":
                    folders[0] = new FolderId(WellKnownFolderName.Tasks); break;

                case "MsgFolderRoot":
                    folders[0] = new FolderId(WellKnownFolderName.MsgFolderRoot); break;

                case "All Folders":
                    folders[0] = new FolderId("AllFolders"); break;
            }
            return folders;
        }

        private void SubscribeConnectionEvents(StreamingSubscriptionConnection connection)
        {
            // Subscribe to events for this connection

            connection.OnNotificationEvent += connection_OnNotificationEvent;
            connection.OnDisconnect += connection_OnDisconnect;
            connection.OnSubscriptionError += connection_OnSubscriptionError;
        }

        void connection_OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            try
            {
                _logger.Log(String.Format("OnSubscriptionError received for {0}: {1}", args.Subscription.Service.ImpersonatedUserId.Id, args.Exception.Message));
            }
            catch
            {
                _logger.Log("OnSubscriptionError received");
            }
        }

        void connection_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            try
            {
                if (args.Subscription != null)
                    _logger.Log(String.Format("OnDisconnection received for {0}", args.Subscription.Service.ImpersonatedUserId.Id));
            }
            catch
            {
                _logger.Log("OnDisconnection received");
            }
            _reconnect = true;  // We can't reconnect in the disconnect event, so we set a flag for the timer to pick this up and check all the connections
        }

        void connection_OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            foreach (NotificationEvent e in args.Events)
            {
                ProcessNotification(e, args.Subscription);
            }
        }

        private void AddGroupSubscriptions(string sGroup)
        {
            if (_groups.ContainsKey(sGroup))
                _groups[sGroup].AddGroupSubscriptions(ref _connections, ref _subscriptions, SelectedEvents(), SelectedFolders(), _logger);
        }

        private void AddAllSubscriptions()
        {
            if (_subscriptions == null)
                _subscriptions = new Dictionary<string, StreamingSubscription>();
            if (_connections == null)
                _connections = new Dictionary<string, StreamingSubscriptionConnection>();

            foreach (string sGroup in _groups.Keys)
                AddGroupSubscriptions(sGroup);

            // Now we build a reverse mapping so that we know which subscription Id is for which mailbox
            _subscriptionIdToMailboxMapping = new Dictionary<string, string>();
            foreach (String mailbox in _subscriptions.Keys)
                _subscriptionIdToMailboxMapping.Add(_subscriptions[mailbox].Id, mailbox);
        }

        private Auth.CredentialHandler CredentialHandler()
        {
            if (_credentialHandler != null)
                return _credentialHandler;

            if (radioButtonAuthBasic.Checked)
            {
                // We are using Basic auth (also covers NTLM, etc.)
                _credentialHandler = new Auth.CredentialHandler(Auth.AuthType.Basic, _logger);
                _credentialHandler.Username = textBoxUsername.Text;
                _credentialHandler.Password = textBoxPassword.Text;
                _credentialHandler.Domain = textBoxDomain.Text;
            }
            else
            {
                // OAuth
                _credentialHandler = new Auth.CredentialHandler(Auth.AuthType.OAuth, _logger);
                _credentialHandler.ApplicationId = textBoxApplicationId.Text;
                _credentialHandler.TenantId = textBoxTenantId.Text;
                if (radioButtonAuthWithClientSecret.Checked)
                    _credentialHandler.ClientSecret = textBoxClientSecret.Text;
                else if (textBoxAuthCertificate.Tag != null)
                    _credentialHandler.Certificate = (System.Security.Cryptography.X509Certificates.X509Certificate2)textBoxAuthCertificate.Tag;
            }
            return _credentialHandler;
        }

        private void CreateGroups()
        {
            // Go through all the mailboxes and organise into groups based on grouping information
            _groups = new Dictionary<string, GroupInfo>();  // Clear any existing groups
            string ewsUrl = textBoxEWSUri.Text;
            if (radioButtonAutodiscover.Checked)
                ewsUrl = "";
            _mailboxes.GroupMailboxes = !radioButtonSpecificUri.Checked;
            _mailboxes.CredentialHandler = CredentialHandler();
            //Auth.CredentialHandler credentialHandler = CredentialHandler();

            foreach (string sMailbox in checkedListBoxMailboxes.CheckedItems)
            {
                _mailboxes.AddMailbox(sMailbox, ewsUrl);
                MailboxInfo mailboxInfo = _mailboxes.Mailbox(sMailbox);
                if (mailboxInfo != null)
                {
                    GroupInfo groupInfo = null;
                    if (_groups.ContainsKey(mailboxInfo.GroupName))
                    {
                        groupInfo = _groups[mailboxInfo.GroupName];
                    }
                    else
                    {
                        groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, _credentialHandler, _traceListener);
                        _groups.Add(mailboxInfo.GroupName, groupInfo);
                    }
                    if (groupInfo.Mailboxes.Count > 199)
                    {
                        // We already have enough mailboxes in this group, so we rename it and create a new one
                        // Renaming it means that we can still put new mailboxes into the correct group based on GroupingInformation
                        int i = 1;
                        while (_groups.ContainsKey(String.Format("{0}{1}", groupInfo.Name, i)))
                            i++;
                        _groups.Remove(groupInfo.Name);
                        _groups.Add(String.Format("{0}{1}", groupInfo.Name, i), groupInfo);
                        groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, _credentialHandler, _traceListener);
                        _groups.Add(mailboxInfo.GroupName, groupInfo);
                    }

                    groupInfo.Mailboxes.Add(sMailbox);
                }
            }
        }

        private bool ConnectToSubscriptions()
        {
            AddAllSubscriptions();
            foreach (StreamingSubscriptionConnection connection in _connections.Values)
            {
                SubscribeConnectionEvents(connection);
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    _logger.Log($"Error on connect: {ex.Message}");
                }
            }
            timerMonitorConnections.Start();
            
            return true;
        }

        private void ReconnectToSubscriptions()
        {
            // Go through our connections and reconnect any that have closed
            _reconnect = false;
            List<string> groupsToRecreate = new List<string>();
            lock (_reconnectLock)  // Prevent this code being run concurrently (i.e. if an event fires in the middle of the processing)
            {
                foreach (string sConnectionGroup in _connections.Keys)
                {
                    StreamingSubscriptionConnection connection = _connections[sConnectionGroup];
                    if (!connection.IsOpen)
                    {
                        try
                        {
                            if (radioButtonAuthOAuth.Checked)
                            {
                                // If we are using OAuth, we need to ensure that the token is still valid
                                _logger.Log($"Updating OAuth token on all subscriptions in group {sConnectionGroup}");
                                foreach (StreamingSubscription subscription in connection.CurrentSubscriptions)
                                    CredentialHandler().ApplyCredentialsToExchangeService(subscription.Service);
                            }
                            connection.Open();
                            _logger.Log(String.Format("Re-opened connection for group {0}", sConnectionGroup));
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.StartsWith("You must add at least one subscription to this connection before it can be opened"))
                            {
                                // Try recreating this group                                    
                                groupsToRecreate.Add(sConnectionGroup);
                                //AddGroupSubscriptions(sConnectionGroup);  This won't currently work as it would involve modifying our _connections collection while reading it
                            }
                            else
                                _logger.Log($"Failed to reopen connection: {ex.Message}");
                        }
                    }
                }
                foreach (string sConnectionGroup in groupsToRecreate)
                {
                    _logger.Log($"Rebuilding subscriptions for group: {sConnectionGroup}");
                    AddGroupSubscriptions(sConnectionGroup);
                }
            }
        }

        private void CloseConnections()
        {
            foreach (StreamingSubscriptionConnection connection in _connections.Values)
            {
                if (connection.IsOpen) connection.Close();
            }
        }

        void _connection_OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            if (args.Exception == null)
                return;

            _logger.Log("Subscription error: " + args.Exception.Message);
        }

        private void comboBoxSubscribeTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxSubscribeTo.SelectedIndex != 0) return;
        }

        private void checkBoxSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool bChecked = true;
            if (checkBoxSelectAll.CheckState == CheckState.Unchecked)
                bChecked=false;

            for (int i = 0; i < checkedListBoxEvents.Items.Count; i++)
                checkedListBoxEvents.SetItemChecked(i, bChecked);
            if (bChecked)
            {
                checkBoxSelectAll.CheckState = CheckState.Checked;
            }
            else
                checkBoxSelectAll.CheckState = CheckState.Unchecked;
        }

        private void checkedListBoxEvents_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (checkedListBoxEvents.Items.Count != checkedListBoxEvents.SelectedItems.Count)
                checkBoxSelectAll.CheckState = CheckState.Indeterminate;
        }

        private void checkBoxQueryMore_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxIncludeMime.Enabled = checkBoxQueryMore.Checked;
        }

        private void listBoxEvents_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                string sInfo = listBoxEvents.SelectedItem.ToString();
                System.Windows.Forms.MessageBox.Show(sInfo, "Event", MessageBoxButtons.OK);
            }
            catch { }
        }

        private void UnsubscribeAll()
        {
            // Unsubscribe all
            if (_subscriptions == null)
                return;

            for (int i = _subscriptions.Keys.Count - 1; i>=0; i-- )
            {
                string sMailbox = _subscriptions.Keys.ElementAt<string>(i);
                StreamingSubscription subscription = _subscriptions[sMailbox];
                try
                {
                    subscription.Unsubscribe();
                    _logger.Log(String.Format("Unsubscribed from {0}", sMailbox));
                }
                catch (Exception ex)
                {
                    _logger.Log(String.Format("Error when unsubscribing {0}: {1}", sMailbox, ex.Message));
                }
                _subscriptions.Remove(sMailbox);
            }
        }

        private void buttonUnsubscribe_Click(object sender, EventArgs e)
        {
            timerMonitorConnections.Stop();
            CloseConnections();
            UnsubscribeAll();
            _reconnect = false;
            buttonUnsubscribe.Enabled = false;
            buttonSubscribe.Enabled = true;
        }


        private void buttonDeselectAllMailboxes_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxMailboxes.Items.Count; i++)
                checkedListBoxMailboxes.SetItemChecked(i, false);
        }

        private void buttonSelectAllMailboxes_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxMailboxes.Items.Count; i++)
                checkedListBoxMailboxes.SetItemChecked(i, true);
        }

        private void timerMonitorConnections_Tick(object sender, EventArgs e)
        {
            if (!_reconnect)
                return;

            timerMonitorConnections.Stop();
            ReconnectToSubscriptions();
            timerMonitorConnections.Start();
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            timerMonitorConnections.Stop();
            CloseConnections();
            UnsubscribeAll();
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
            string[] mbxList = mailboxes.Split(new string[]{Environment.NewLine},StringSplitOptions.RemoveEmptyEntries);
            foreach (string mbx in mbxList)
                checkedListBoxMailboxes.Items.Add(mbx);
        }

        private void UpdateUriUI()
        {
            textBoxEWSUri.Enabled = !radioButtonAutodiscover.Checked;
            if (radioButtonOffice365.Checked)
                textBoxEWSUri.Text = "https://outlook.office365.com/EWS/Exchange.asmx";
            textBoxEWSUri.ReadOnly = !radioButtonSpecificUri.Checked;
        }

        private void radioButtonOffice365_CheckedChanged(object sender, EventArgs e)
        {
            UpdateUriUI();
        }

        private void radioButtonAutodiscover_CheckedChanged(object sender, EventArgs e)
        {
            UpdateUriUI();
        }

        private void radioButtonSpecificUri_CheckedChanged(object sender, EventArgs e)
        {
            UpdateUriUI();
        }

        private void textBoxClientSecret_TextChanged(object sender, EventArgs e)
        {
            _credentialHandler = null;
        }

        private void textBoxApplicationId_TextChanged(object sender, EventArgs e)
        {
            _credentialHandler = null;
        }

        private void textBoxTenantId_TextChanged(object sender, EventArgs e)
        {
            _credentialHandler = null;
        }

        private void buttonSelectCertificate_Click(object sender, EventArgs e)
        {
            using (Auth.FormChooseAuthCertificate formChooseCert = new Auth.FormChooseAuthCertificate())
            {
                formChooseCert.ShowDialog(this);
                if (formChooseCert.Certificate != null)
                {
                    textBoxAuthCertificate.Text = formChooseCert.Certificate.FriendlyName;
                    if (_credentialHandler != null)
                        _credentialHandler.Certificate = formChooseCert.Certificate;
                    textBoxAuthCertificate.Tag = formChooseCert.Certificate;
                }
            }
        }
        
        private void UpdateAuthUI()
        {
            textBoxAuthCertificate.Enabled = radioButtonAuthWithCertificate.Checked;
            buttonSelectCertificate.Enabled = radioButtonAuthWithCertificate.Checked;
            textBoxClientSecret.Enabled = radioButtonAuthWithClientSecret.Checked;
        }

        private void radioButtonAuthWithCertificate_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAuthUI();
        }

        private void radioButtonAuthWithClientSecret_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAuthUI();
        }
    }
}

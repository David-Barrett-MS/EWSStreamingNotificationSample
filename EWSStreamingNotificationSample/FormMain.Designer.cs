namespace EWSStreamingNotificationSample
{
    partial class FormMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxEWSUri = new System.Windows.Forms.TextBox();
            this.radioButtonSpecificUri = new System.Windows.Forms.RadioButton();
            this.radioButtonOffice365 = new System.Windows.Forms.RadioButton();
            this.radioButtonAutodiscover = new System.Windows.Forms.RadioButton();
            this.buttonEditMailboxes = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.comboBoxExchangeVersion = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.textBoxDomain = new System.Windows.Forms.TextBox();
            this.buttonLoadMailboxes = new System.Windows.Forms.Button();
            this.buttonDeselectAllMailboxes = new System.Windows.Forms.Button();
            this.buttonSelectAllMailboxes = new System.Windows.Forms.Button();
            this.checkedListBoxMailboxes = new System.Windows.Forms.CheckedListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.textBoxUsername = new System.Windows.Forms.TextBox();
            this.listBoxEvents = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonSubscribe = new System.Windows.Forms.Button();
            this.checkBoxQueryMore = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.buttonUnsubscribe = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.numericUpDownTimeout = new System.Windows.Forms.NumericUpDown();
            this.checkBoxIncludeMime = new System.Windows.Forms.CheckBox();
            this.checkBoxSelectAll = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.checkBoxShowFolderEvents = new System.Windows.Forms.CheckBox();
            this.checkBoxShowItemEvents = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBoxSubscribeTo = new System.Windows.Forms.ComboBox();
            this.checkedListBoxEvents = new System.Windows.Forms.CheckedListBox();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.timerMonitorConnections = new System.Windows.Forms.Timer(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownTimeout)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxEWSUri);
            this.groupBox1.Controls.Add(this.radioButtonSpecificUri);
            this.groupBox1.Controls.Add(this.radioButtonOffice365);
            this.groupBox1.Controls.Add(this.radioButtonAutodiscover);
            this.groupBox1.Controls.Add(this.buttonEditMailboxes);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.comboBoxExchangeVersion);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.textBoxDomain);
            this.groupBox1.Controls.Add(this.buttonLoadMailboxes);
            this.groupBox1.Controls.Add(this.buttonDeselectAllMailboxes);
            this.groupBox1.Controls.Add(this.buttonSelectAllMailboxes);
            this.groupBox1.Controls.Add(this.checkedListBoxMailboxes);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBoxPassword);
            this.groupBox1.Controls.Add(this.textBoxUsername);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(611, 163);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "EWS Configuration";
            // 
            // textBoxEWSUri
            // 
            this.textBoxEWSUri.Enabled = false;
            this.textBoxEWSUri.Location = new System.Drawing.Point(226, 18);
            this.textBoxEWSUri.Name = "textBoxEWSUri";
            this.textBoxEWSUri.Size = new System.Drawing.Size(375, 20);
            this.textBoxEWSUri.TabIndex = 22;
            this.textBoxEWSUri.Text = "https://<exchange>/EWS/Exchange.asmx";
            // 
            // radioButtonSpecificUri
            // 
            this.radioButtonSpecificUri.AutoSize = true;
            this.radioButtonSpecificUri.Checked = true;
            this.radioButtonSpecificUri.Location = new System.Drawing.Point(179, 19);
            this.radioButtonSpecificUri.Name = "radioButtonSpecificUri";
            this.radioButtonSpecificUri.Size = new System.Drawing.Size(41, 17);
            this.radioButtonSpecificUri.TabIndex = 21;
            this.radioButtonSpecificUri.TabStop = true;
            this.radioButtonSpecificUri.Text = "Uri:";
            this.radioButtonSpecificUri.UseVisualStyleBackColor = true;
            this.radioButtonSpecificUri.CheckedChanged += new System.EventHandler(this.radioButtonSpecificUri_CheckedChanged);
            // 
            // radioButtonOffice365
            // 
            this.radioButtonOffice365.AutoSize = true;
            this.radioButtonOffice365.Location = new System.Drawing.Point(99, 19);
            this.radioButtonOffice365.Name = "radioButtonOffice365";
            this.radioButtonOffice365.Size = new System.Drawing.Size(74, 17);
            this.radioButtonOffice365.TabIndex = 20;
            this.radioButtonOffice365.Text = "Office 365";
            this.radioButtonOffice365.UseVisualStyleBackColor = true;
            this.radioButtonOffice365.CheckedChanged += new System.EventHandler(this.radioButtonOffice365_CheckedChanged);
            // 
            // radioButtonAutodiscover
            // 
            this.radioButtonAutodiscover.AutoSize = true;
            this.radioButtonAutodiscover.Location = new System.Drawing.Point(6, 19);
            this.radioButtonAutodiscover.Name = "radioButtonAutodiscover";
            this.radioButtonAutodiscover.Size = new System.Drawing.Size(87, 17);
            this.radioButtonAutodiscover.TabIndex = 19;
            this.radioButtonAutodiscover.Text = "Autodiscover";
            this.radioButtonAutodiscover.UseVisualStyleBackColor = true;
            this.radioButtonAutodiscover.CheckedChanged += new System.EventHandler(this.radioButtonAutodiscover_CheckedChanged);
            // 
            // buttonEditMailboxes
            // 
            this.buttonEditMailboxes.Location = new System.Drawing.Point(340, 134);
            this.buttonEditMailboxes.Name = "buttonEditMailboxes";
            this.buttonEditMailboxes.Size = new System.Drawing.Size(50, 23);
            this.buttonEditMailboxes.TabIndex = 18;
            this.buttonEditMailboxes.Text = "Edit...";
            this.buttonEditMailboxes.UseVisualStyleBackColor = true;
            this.buttonEditMailboxes.Click += new System.EventHandler(this.buttonEditMailboxes_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(8, 125);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(95, 13);
            this.label9.TabIndex = 17;
            this.label9.Text = "Exchange version:";
            // 
            // comboBoxExchangeVersion
            // 
            this.comboBoxExchangeVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxExchangeVersion.FormattingEnabled = true;
            this.comboBoxExchangeVersion.Items.AddRange(new object[] {
            "Exchange2013",
            "Exchange2013_SP1",
            "Exchange2016"});
            this.comboBoxExchangeVersion.Location = new System.Drawing.Point(109, 122);
            this.comboBoxExchangeVersion.Name = "comboBoxExchangeVersion";
            this.comboBoxExchangeVersion.Size = new System.Drawing.Size(157, 21);
            this.comboBoxExchangeVersion.TabIndex = 16;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 99);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(46, 13);
            this.label7.TabIndex = 15;
            this.label7.Text = "Domain:";
            // 
            // textBoxDomain
            // 
            this.textBoxDomain.Location = new System.Drawing.Point(70, 96);
            this.textBoxDomain.Name = "textBoxDomain";
            this.textBoxDomain.Size = new System.Drawing.Size(196, 20);
            this.textBoxDomain.TabIndex = 14;
            // 
            // buttonLoadMailboxes
            // 
            this.buttonLoadMailboxes.Location = new System.Drawing.Point(396, 134);
            this.buttonLoadMailboxes.Name = "buttonLoadMailboxes";
            this.buttonLoadMailboxes.Size = new System.Drawing.Size(50, 23);
            this.buttonLoadMailboxes.TabIndex = 13;
            this.buttonLoadMailboxes.Text = "Load...";
            this.toolTips.SetToolTip(this.buttonLoadMailboxes, "Load mailboxes from text file");
            this.buttonLoadMailboxes.UseVisualStyleBackColor = true;
            this.buttonLoadMailboxes.Click += new System.EventHandler(this.buttonLoadMailboxes_Click);
            // 
            // buttonDeselectAllMailboxes
            // 
            this.buttonDeselectAllMailboxes.Location = new System.Drawing.Point(535, 134);
            this.buttonDeselectAllMailboxes.Name = "buttonDeselectAllMailboxes";
            this.buttonDeselectAllMailboxes.Size = new System.Drawing.Size(70, 23);
            this.buttonDeselectAllMailboxes.TabIndex = 10;
            this.buttonDeselectAllMailboxes.Text = "Deselect all";
            this.buttonDeselectAllMailboxes.UseVisualStyleBackColor = true;
            this.buttonDeselectAllMailboxes.Click += new System.EventHandler(this.buttonDeselectAllMailboxes_Click);
            // 
            // buttonSelectAllMailboxes
            // 
            this.buttonSelectAllMailboxes.Location = new System.Drawing.Point(459, 134);
            this.buttonSelectAllMailboxes.Name = "buttonSelectAllMailboxes";
            this.buttonSelectAllMailboxes.Size = new System.Drawing.Size(70, 23);
            this.buttonSelectAllMailboxes.TabIndex = 9;
            this.buttonSelectAllMailboxes.Text = "Select all";
            this.buttonSelectAllMailboxes.UseVisualStyleBackColor = true;
            this.buttonSelectAllMailboxes.Click += new System.EventHandler(this.buttonSelectAllMailboxes_Click);
            // 
            // checkedListBoxMailboxes
            // 
            this.checkedListBoxMailboxes.CheckOnClick = true;
            this.checkedListBoxMailboxes.FormattingEnabled = true;
            this.checkedListBoxMailboxes.Location = new System.Drawing.Point(336, 44);
            this.checkedListBoxMailboxes.Name = "checkedListBoxMailboxes";
            this.checkedListBoxMailboxes.Size = new System.Drawing.Size(265, 79);
            this.checkedListBoxMailboxes.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(273, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Mailboxes:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Password:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Username:";
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(70, 70);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.Size = new System.Drawing.Size(196, 20);
            this.textBoxPassword.TabIndex = 1;
            this.textBoxPassword.UseSystemPasswordChar = true;
            // 
            // textBoxUsername
            // 
            this.textBoxUsername.Location = new System.Drawing.Point(70, 44);
            this.textBoxUsername.Name = "textBoxUsername";
            this.textBoxUsername.Size = new System.Drawing.Size(196, 20);
            this.textBoxUsername.TabIndex = 0;
            // 
            // listBoxEvents
            // 
            this.listBoxEvents.FormattingEnabled = true;
            this.listBoxEvents.Location = new System.Drawing.Point(12, 306);
            this.listBoxEvents.Name = "listBoxEvents";
            this.listBoxEvents.Size = new System.Drawing.Size(611, 225);
            this.listBoxEvents.TabIndex = 1;
            this.listBoxEvents.DoubleClick += new System.EventHandler(this.listBoxEvents_DoubleClick);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 290);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Received Events:";
            // 
            // buttonSubscribe
            // 
            this.buttonSubscribe.Location = new System.Drawing.Point(449, 77);
            this.buttonSubscribe.Name = "buttonSubscribe";
            this.buttonSubscribe.Size = new System.Drawing.Size(75, 23);
            this.buttonSubscribe.TabIndex = 3;
            this.buttonSubscribe.Text = "Subscribe";
            this.buttonSubscribe.UseVisualStyleBackColor = true;
            this.buttonSubscribe.Click += new System.EventHandler(this.buttonSubscribe_Click);
            // 
            // checkBoxQueryMore
            // 
            this.checkBoxQueryMore.AutoSize = true;
            this.checkBoxQueryMore.Checked = true;
            this.checkBoxQueryMore.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxQueryMore.Location = new System.Drawing.Point(234, 51);
            this.checkBoxQueryMore.Name = "checkBoxQueryMore";
            this.checkBoxQueryMore.Size = new System.Drawing.Size(100, 17);
            this.checkBoxQueryMore.TabIndex = 5;
            this.checkBoxQueryMore.Text = "Query extra info";
            this.checkBoxQueryMore.UseVisualStyleBackColor = true;
            this.checkBoxQueryMore.CheckedChanged += new System.EventHandler(this.checkBoxQueryMore_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.buttonUnsubscribe);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.numericUpDownTimeout);
            this.groupBox2.Controls.Add(this.checkBoxIncludeMime);
            this.groupBox2.Controls.Add(this.checkBoxSelectAll);
            this.groupBox2.Controls.Add(this.buttonSubscribe);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.checkBoxQueryMore);
            this.groupBox2.Controls.Add(this.checkBoxShowFolderEvents);
            this.groupBox2.Controls.Add(this.checkBoxShowItemEvents);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.comboBoxSubscribeTo);
            this.groupBox2.Controls.Add(this.checkedListBoxEvents);
            this.groupBox2.Location = new System.Drawing.Point(12, 181);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(611, 106);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Subscription Information";
            // 
            // buttonUnsubscribe
            // 
            this.buttonUnsubscribe.Location = new System.Drawing.Point(530, 77);
            this.buttonUnsubscribe.Name = "buttonUnsubscribe";
            this.buttonUnsubscribe.Size = new System.Drawing.Size(75, 23);
            this.buttonUnsubscribe.TabIndex = 21;
            this.buttonUnsubscribe.Text = "Unsubscribe";
            this.buttonUnsubscribe.UseVisualStyleBackColor = true;
            this.buttonUnsubscribe.Click += new System.EventHandler(this.buttonUnsubscribe_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(140, 78);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(43, 13);
            this.label11.TabIndex = 20;
            this.label11.Text = "minutes";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 78);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(48, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "Timeout:";
            // 
            // numericUpDownTimeout
            // 
            this.numericUpDownTimeout.Location = new System.Drawing.Point(60, 76);
            this.numericUpDownTimeout.Maximum = new decimal(new int[] {
            30,
            0,
            0,
            0});
            this.numericUpDownTimeout.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownTimeout.Name = "numericUpDownTimeout";
            this.numericUpDownTimeout.Size = new System.Drawing.Size(74, 20);
            this.numericUpDownTimeout.TabIndex = 18;
            this.numericUpDownTimeout.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // checkBoxIncludeMime
            // 
            this.checkBoxIncludeMime.AutoSize = true;
            this.checkBoxIncludeMime.Checked = true;
            this.checkBoxIncludeMime.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxIncludeMime.Location = new System.Drawing.Point(234, 74);
            this.checkBoxIncludeMime.Name = "checkBoxIncludeMime";
            this.checkBoxIncludeMime.Size = new System.Drawing.Size(92, 17);
            this.checkBoxIncludeMime.TabIndex = 17;
            this.checkBoxIncludeMime.Text = "Include MIME";
            this.checkBoxIncludeMime.UseVisualStyleBackColor = true;
            // 
            // checkBoxSelectAll
            // 
            this.checkBoxSelectAll.AutoSize = true;
            this.checkBoxSelectAll.Checked = true;
            this.checkBoxSelectAll.CheckState = System.Windows.Forms.CheckState.Indeterminate;
            this.checkBoxSelectAll.Location = new System.Drawing.Point(369, 74);
            this.checkBoxSelectAll.Name = "checkBoxSelectAll";
            this.checkBoxSelectAll.Size = new System.Drawing.Size(69, 17);
            this.checkBoxSelectAll.TabIndex = 16;
            this.checkBoxSelectAll.Text = "Select all";
            this.checkBoxSelectAll.ThreeState = true;
            this.checkBoxSelectAll.UseVisualStyleBackColor = true;
            this.checkBoxSelectAll.CheckedChanged += new System.EventHandler(this.checkBoxSelectAll_CheckedChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 52);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(28, 13);
            this.label8.TabIndex = 15;
            this.label8.Text = "Log:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(320, 19);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(43, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Events:";
            // 
            // checkBoxShowFolderEvents
            // 
            this.checkBoxShowFolderEvents.AutoSize = true;
            this.checkBoxShowFolderEvents.Location = new System.Drawing.Point(138, 51);
            this.checkBoxShowFolderEvents.Name = "checkBoxShowFolderEvents";
            this.checkBoxShowFolderEvents.Size = new System.Drawing.Size(90, 17);
            this.checkBoxShowFolderEvents.TabIndex = 14;
            this.checkBoxShowFolderEvents.Text = "Folder events";
            this.checkBoxShowFolderEvents.UseVisualStyleBackColor = true;
            // 
            // checkBoxShowItemEvents
            // 
            this.checkBoxShowItemEvents.AutoSize = true;
            this.checkBoxShowItemEvents.Checked = true;
            this.checkBoxShowItemEvents.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxShowItemEvents.Location = new System.Drawing.Point(51, 51);
            this.checkBoxShowItemEvents.Name = "checkBoxShowItemEvents";
            this.checkBoxShowItemEvents.Size = new System.Drawing.Size(81, 17);
            this.checkBoxShowItemEvents.TabIndex = 13;
            this.checkBoxShowItemEvents.Text = "Item events";
            this.checkBoxShowItemEvents.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 22);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(39, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Folder:";
            // 
            // comboBoxSubscribeTo
            // 
            this.comboBoxSubscribeTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxSubscribeTo.FormattingEnabled = true;
            this.comboBoxSubscribeTo.Items.AddRange(new object[] {
            "All Folders",
            "Calendar",
            "Contacts",
            "DeletedItems",
            "Drafts",
            "Inbox",
            "Journal",
            "Notes",
            "Outbox",
            "SentItems",
            "Tasks",
            "MsgFolderRoot"});
            this.comboBoxSubscribeTo.Location = new System.Drawing.Point(51, 19);
            this.comboBoxSubscribeTo.Name = "comboBoxSubscribeTo";
            this.comboBoxSubscribeTo.Size = new System.Drawing.Size(205, 21);
            this.comboBoxSubscribeTo.TabIndex = 10;
            this.comboBoxSubscribeTo.SelectedIndexChanged += new System.EventHandler(this.comboBoxSubscribeTo_SelectedIndexChanged);
            // 
            // checkedListBoxEvents
            // 
            this.checkedListBoxEvents.CheckOnClick = true;
            this.checkedListBoxEvents.FormattingEnabled = true;
            this.checkedListBoxEvents.Items.AddRange(new object[] {
            "NewMail",
            "Deleted",
            "Modified",
            "Moved",
            "Copied",
            "Created",
            "FreeBusyChanged"});
            this.checkedListBoxEvents.Location = new System.Drawing.Point(369, 19);
            this.checkedListBoxEvents.Name = "checkedListBoxEvents";
            this.checkedListBoxEvents.Size = new System.Drawing.Size(236, 49);
            this.checkedListBoxEvents.TabIndex = 9;
            this.toolTips.SetToolTip(this.checkedListBoxEvents, "Select which events to subscribe for");
            this.checkedListBoxEvents.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBoxEvents_ItemCheck);
            // 
            // timerMonitorConnections
            // 
            this.timerMonitorConnections.Interval = 5000;
            this.timerMonitorConnections.Tick += new System.EventHandler(this.timerMonitorConnections_Tick);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(635, 544);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.listBoxEvents);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormMain";
            this.Text = "EWS Streaming Notification Sample Application";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMain_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownTimeout)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.TextBox textBoxUsername;
        private System.Windows.Forms.ListBox listBoxEvents;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonSubscribe;
        private System.Windows.Forms.CheckBox checkBoxQueryMore;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBoxShowFolderEvents;
        private System.Windows.Forms.CheckBox checkBoxShowItemEvents;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBoxSubscribeTo;
        private System.Windows.Forms.CheckedListBox checkedListBoxEvents;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.CheckBox checkBoxSelectAll;
        private System.Windows.Forms.CheckBox checkBoxIncludeMime;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.NumericUpDown numericUpDownTimeout;
        private System.Windows.Forms.Button buttonUnsubscribe;
        private System.Windows.Forms.CheckedListBox checkedListBoxMailboxes;
        private System.Windows.Forms.Button buttonDeselectAllMailboxes;
        private System.Windows.Forms.Button buttonSelectAllMailboxes;
        private System.Windows.Forms.Button buttonLoadMailboxes;
        private System.Windows.Forms.ToolTip toolTips;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBoxDomain;
        private System.Windows.Forms.Timer timerMonitorConnections;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBoxExchangeVersion;
        private System.Windows.Forms.Button buttonEditMailboxes;
        private System.Windows.Forms.TextBox textBoxEWSUri;
        private System.Windows.Forms.RadioButton radioButtonSpecificUri;
        private System.Windows.Forms.RadioButton radioButtonOffice365;
        private System.Windows.Forms.RadioButton radioButtonAutodiscover;
    }
}


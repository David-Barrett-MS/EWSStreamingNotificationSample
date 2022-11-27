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
            this.timerMonitorConnections = new System.Windows.Forms.Timer(this.components);
            this.groupBoxOAuthAuthMethod = new System.Windows.Forms.GroupBox();
            this.buttonSelectCertificate = new System.Windows.Forms.Button();
            this.radioButtonAuthWithClientSecret = new System.Windows.Forms.RadioButton();
            this.textBoxAuthCertificate = new System.Windows.Forms.TextBox();
            this.textBoxClientSecret = new System.Windows.Forms.TextBox();
            this.radioButtonAuthWithCertificate = new System.Windows.Forms.RadioButton();
            this.buttonUnsubscribe = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.numericUpDownTimeout = new System.Windows.Forms.NumericUpDown();
            this.checkBoxIncludeMime = new System.Windows.Forms.CheckBox();
            this.checkBoxSelectAll = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.buttonSubscribe = new System.Windows.Forms.Button();
            this.checkBoxQueryMore = new System.Windows.Forms.CheckBox();
            this.textBoxNotificationCount = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.textBoxNumConnections = new System.Windows.Forms.TextBox();
            this.textBoxNumSubscriptions = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.label16 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.buttonLoadCertificate = new System.Windows.Forms.Button();
            this.textBoxTenantId = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.textBoxApplicationId = new System.Windows.Forms.TextBox();
            this.radioButtonAuthOAuth = new System.Windows.Forms.RadioButton();
            this.radioButtonAuthBasic = new System.Windows.Forms.RadioButton();
            this.textBoxUsername = new System.Windows.Forms.TextBox();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxDomain = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.numericUpDownMaxMailboxesInGroup = new System.Windows.Forms.NumericUpDown();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.checkedListBoxMailboxes = new System.Windows.Forms.CheckedListBox();
            this.buttonSelectAllMailboxes = new System.Windows.Forms.Button();
            this.buttonDeselectAllMailboxes = new System.Windows.Forms.Button();
            this.buttonLoadMailboxes = new System.Windows.Forms.Button();
            this.buttonEditMailboxes = new System.Windows.Forms.Button();
            this.checkBoxShowFolderEvents = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxEWSUri = new System.Windows.Forms.TextBox();
            this.radioButtonSpecificUri = new System.Windows.Forms.RadioButton();
            this.radioButtonOffice365 = new System.Windows.Forms.RadioButton();
            this.radioButtonAutodiscover = new System.Windows.Forms.RadioButton();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.checkedListBoxEvents = new System.Windows.Forms.CheckedListBox();
            this.listBoxEvents = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBoxShowItemEvents = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBoxSubscribeTo = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBoxOAuthAuthMethod.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownTimeout)).BeginInit();
            this.groupBox6.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownMaxMailboxesInGroup)).BeginInit();
            this.groupBox5.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // timerMonitorConnections
            // 
            this.timerMonitorConnections.Interval = 1000;
            this.timerMonitorConnections.Tick += new System.EventHandler(this.timerMonitorConnections_Tick);
            // 
            // groupBoxOAuthAuthMethod
            // 
            this.groupBoxOAuthAuthMethod.Controls.Add(this.buttonSelectCertificate);
            this.groupBoxOAuthAuthMethod.Controls.Add(this.radioButtonAuthWithClientSecret);
            this.groupBoxOAuthAuthMethod.Controls.Add(this.textBoxAuthCertificate);
            this.groupBoxOAuthAuthMethod.Controls.Add(this.textBoxClientSecret);
            this.groupBoxOAuthAuthMethod.Controls.Add(this.radioButtonAuthWithCertificate);
            this.groupBoxOAuthAuthMethod.Location = new System.Drawing.Point(34, 176);
            this.groupBoxOAuthAuthMethod.Name = "groupBoxOAuthAuthMethod";
            this.groupBoxOAuthAuthMethod.Size = new System.Drawing.Size(324, 67);
            this.groupBoxOAuthAuthMethod.TabIndex = 37;
            this.groupBoxOAuthAuthMethod.TabStop = false;
            this.groupBoxOAuthAuthMethod.Text = "Auth method";
            // 
            // buttonSelectCertificate
            // 
            this.buttonSelectCertificate.Location = new System.Drawing.Point(257, 38);
            this.buttonSelectCertificate.Name = "buttonSelectCertificate";
            this.buttonSelectCertificate.Size = new System.Drawing.Size(52, 20);
            this.buttonSelectCertificate.TabIndex = 37;
            this.buttonSelectCertificate.Text = "Select...";
            this.buttonSelectCertificate.UseVisualStyleBackColor = true;
            // 
            // radioButtonAuthWithClientSecret
            // 
            this.radioButtonAuthWithClientSecret.AutoSize = true;
            this.radioButtonAuthWithClientSecret.Checked = true;
            this.radioButtonAuthWithClientSecret.Location = new System.Drawing.Point(5, 17);
            this.radioButtonAuthWithClientSecret.Margin = new System.Windows.Forms.Padding(2, 1, 2, 1);
            this.radioButtonAuthWithClientSecret.Name = "radioButtonAuthWithClientSecret";
            this.radioButtonAuthWithClientSecret.Size = new System.Drawing.Size(86, 17);
            this.radioButtonAuthWithClientSecret.TabIndex = 32;
            this.radioButtonAuthWithClientSecret.TabStop = true;
            this.radioButtonAuthWithClientSecret.Tag = "NoTextSave";
            this.radioButtonAuthWithClientSecret.Text = "Client secret:";
            this.radioButtonAuthWithClientSecret.UseVisualStyleBackColor = true;
            this.radioButtonAuthWithClientSecret.CheckedChanged += new System.EventHandler(this.radioButtonAuthWithClientSecret_CheckedChanged);
            // 
            // textBoxAuthCertificate
            // 
            this.textBoxAuthCertificate.Location = new System.Drawing.Point(94, 38);
            this.textBoxAuthCertificate.Margin = new System.Windows.Forms.Padding(2, 1, 2, 1);
            this.textBoxAuthCertificate.Name = "textBoxAuthCertificate";
            this.textBoxAuthCertificate.Size = new System.Drawing.Size(162, 20);
            this.textBoxAuthCertificate.TabIndex = 36;
            // 
            // textBoxClientSecret
            // 
            this.textBoxClientSecret.Location = new System.Drawing.Point(94, 16);
            this.textBoxClientSecret.Name = "textBoxClientSecret";
            this.textBoxClientSecret.Size = new System.Drawing.Size(215, 20);
            this.textBoxClientSecret.TabIndex = 33;
            this.textBoxClientSecret.UseSystemPasswordChar = true;
            // 
            // radioButtonAuthWithCertificate
            // 
            this.radioButtonAuthWithCertificate.AutoSize = true;
            this.radioButtonAuthWithCertificate.Location = new System.Drawing.Point(5, 38);
            this.radioButtonAuthWithCertificate.Margin = new System.Windows.Forms.Padding(2, 1, 2, 1);
            this.radioButtonAuthWithCertificate.Name = "radioButtonAuthWithCertificate";
            this.radioButtonAuthWithCertificate.Size = new System.Drawing.Size(75, 17);
            this.radioButtonAuthWithCertificate.TabIndex = 34;
            this.radioButtonAuthWithCertificate.Tag = "NoTextSave";
            this.radioButtonAuthWithCertificate.Text = "Certificate:";
            this.radioButtonAuthWithCertificate.UseVisualStyleBackColor = true;
            this.radioButtonAuthWithCertificate.CheckedChanged += new System.EventHandler(this.radioButtonAuthWithCertificate_CheckedChanged);
            // 
            // buttonUnsubscribe
            // 
            this.buttonUnsubscribe.Enabled = false;
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
            this.label10.Size = new System.Drawing.Size(46, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "Lifetime:";
            this.toolTips.SetToolTip(this.label10, "Connection lifetime");
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
            30,
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
            // 
            // textBoxNotificationCount
            // 
            this.textBoxNotificationCount.Location = new System.Drawing.Point(155, 71);
            this.textBoxNotificationCount.Name = "textBoxNotificationCount";
            this.textBoxNotificationCount.ReadOnly = true;
            this.textBoxNotificationCount.Size = new System.Drawing.Size(100, 20);
            this.textBoxNotificationCount.TabIndex = 9;
            this.textBoxNotificationCount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 74);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(127, 13);
            this.label9.TabIndex = 8;
            this.label9.Text = "Notifications (folder/item):";
            // 
            // textBoxNumConnections
            // 
            this.textBoxNumConnections.Location = new System.Drawing.Point(155, 45);
            this.textBoxNumConnections.Name = "textBoxNumConnections";
            this.textBoxNumConnections.ReadOnly = true;
            this.textBoxNumConnections.Size = new System.Drawing.Size(100, 20);
            this.textBoxNumConnections.TabIndex = 7;
            this.textBoxNumConnections.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBoxNumSubscriptions
            // 
            this.textBoxNumSubscriptions.Location = new System.Drawing.Point(155, 19);
            this.textBoxNumSubscriptions.Name = "textBoxNumSubscriptions";
            this.textBoxNumSubscriptions.ReadOnly = true;
            this.textBoxNumSubscriptions.Size = new System.Drawing.Size(100, 20);
            this.textBoxNumSubscriptions.TabIndex = 6;
            this.textBoxNumSubscriptions.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(6, 22);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(136, 13);
            this.label15.TabIndex = 4;
            this.label15.Text = "Subscriptions (active/total):";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.textBoxNotificationCount);
            this.groupBox6.Controls.Add(this.label9);
            this.groupBox6.Controls.Add(this.textBoxNumConnections);
            this.groupBox6.Controls.Add(this.textBoxNumSubscriptions);
            this.groupBox6.Controls.Add(this.label16);
            this.groupBox6.Controls.Add(this.label15);
            this.groupBox6.Location = new System.Drawing.Point(793, 66);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(261, 98);
            this.groupBox6.TabIndex = 17;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Statistics";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(6, 48);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(127, 13);
            this.label16.TabIndex = 5;
            this.label16.Text = "Connections (open/total):";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.groupBoxOAuthAuthMethod);
            this.groupBox3.Controls.Add(this.buttonLoadCertificate);
            this.groupBox3.Controls.Add(this.textBoxTenantId);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Controls.Add(this.textBoxApplicationId);
            this.groupBox3.Controls.Add(this.radioButtonAuthOAuth);
            this.groupBox3.Controls.Add(this.radioButtonAuthBasic);
            this.groupBox3.Controls.Add(this.textBoxUsername);
            this.groupBox3.Controls.Add(this.textBoxPassword);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.textBoxDomain);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Location = new System.Drawing.Point(12, 66);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(425, 257);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Authentication";
            // 
            // buttonLoadCertificate
            // 
            this.buttonLoadCertificate.Location = new System.Drawing.Point(497, 196);
            this.buttonLoadCertificate.Margin = new System.Windows.Forms.Padding(2, 1, 2, 1);
            this.buttonLoadCertificate.Name = "buttonLoadCertificate";
            this.buttonLoadCertificate.Size = new System.Drawing.Size(50, 21);
            this.buttonLoadCertificate.TabIndex = 35;
            this.buttonLoadCertificate.Text = "Select...";
            this.buttonLoadCertificate.UseVisualStyleBackColor = true;
            // 
            // textBoxTenantId
            // 
            this.textBoxTenantId.Location = new System.Drawing.Point(119, 126);
            this.textBoxTenantId.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxTenantId.Name = "textBoxTenantId";
            this.textBoxTenantId.Size = new System.Drawing.Size(215, 20);
            this.textBoxTenantId.TabIndex = 24;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(31, 151);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(80, 13);
            this.label12.TabIndex = 22;
            this.label12.Text = "Application ID*:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(31, 129);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(62, 13);
            this.label13.TabIndex = 23;
            this.label13.Tag = "";
            this.label13.Text = "Tenant ID*:";
            // 
            // textBoxApplicationId
            // 
            this.textBoxApplicationId.Location = new System.Drawing.Point(119, 148);
            this.textBoxApplicationId.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxApplicationId.Name = "textBoxApplicationId";
            this.textBoxApplicationId.Size = new System.Drawing.Size(215, 20);
            this.textBoxApplicationId.TabIndex = 25;
            // 
            // radioButtonAuthOAuth
            // 
            this.radioButtonAuthOAuth.AutoSize = true;
            this.radioButtonAuthOAuth.Location = new System.Drawing.Point(14, 109);
            this.radioButtonAuthOAuth.Name = "radioButtonAuthOAuth";
            this.radioButtonAuthOAuth.Size = new System.Drawing.Size(55, 17);
            this.radioButtonAuthOAuth.TabIndex = 1;
            this.radioButtonAuthOAuth.Text = "OAuth";
            this.radioButtonAuthOAuth.UseVisualStyleBackColor = true;
            this.radioButtonAuthOAuth.CheckedChanged += new System.EventHandler(this.radioButtonAuthOAuth_CheckedChanged);
            // 
            // radioButtonAuthBasic
            // 
            this.radioButtonAuthBasic.AutoSize = true;
            this.radioButtonAuthBasic.Checked = true;
            this.radioButtonAuthBasic.Location = new System.Drawing.Point(14, 24);
            this.radioButtonAuthBasic.Name = "radioButtonAuthBasic";
            this.radioButtonAuthBasic.Size = new System.Drawing.Size(182, 17);
            this.radioButtonAuthBasic.TabIndex = 0;
            this.radioButtonAuthBasic.TabStop = true;
            this.radioButtonAuthBasic.Text = "Basic auth (inc. NTLM, Kerberos)";
            this.radioButtonAuthBasic.UseVisualStyleBackColor = true;
            this.radioButtonAuthBasic.CheckedChanged += new System.EventHandler(this.radioButtonAuthBasic_CheckedChanged);
            // 
            // textBoxUsername
            // 
            this.textBoxUsername.Location = new System.Drawing.Point(95, 47);
            this.textBoxUsername.Name = "textBoxUsername";
            this.textBoxUsername.Size = new System.Drawing.Size(128, 20);
            this.textBoxUsername.TabIndex = 0;
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(291, 47);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.Size = new System.Drawing.Size(117, 20);
            this.textBoxPassword.TabIndex = 1;
            this.textBoxPassword.UseSystemPasswordChar = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Username:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(229, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Password:";
            // 
            // textBoxDomain
            // 
            this.textBoxDomain.Location = new System.Drawing.Point(95, 73);
            this.textBoxDomain.Name = "textBoxDomain";
            this.textBoxDomain.Size = new System.Drawing.Size(128, 20);
            this.textBoxDomain.TabIndex = 14;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(31, 76);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(46, 13);
            this.label7.TabIndex = 15;
            this.label7.Text = "Domain:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(43, 135);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(170, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "Max number of mailboxes in group:";
            // 
            // numericUpDownMaxMailboxesInGroup
            // 
            this.numericUpDownMaxMailboxesInGroup.Location = new System.Drawing.Point(219, 133);
            this.numericUpDownMaxMailboxesInGroup.Maximum = new decimal(new int[] {
            200,
            0,
            0,
            0});
            this.numericUpDownMaxMailboxesInGroup.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownMaxMailboxesInGroup.Name = "numericUpDownMaxMailboxesInGroup";
            this.numericUpDownMaxMailboxesInGroup.Size = new System.Drawing.Size(53, 20);
            this.numericUpDownMaxMailboxesInGroup.TabIndex = 19;
            this.numericUpDownMaxMailboxesInGroup.Value = new decimal(new int[] {
            200,
            0,
            0,
            0});
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label4);
            this.groupBox5.Controls.Add(this.numericUpDownMaxMailboxesInGroup);
            this.groupBox5.Controls.Add(this.checkedListBoxMailboxes);
            this.groupBox5.Controls.Add(this.buttonSelectAllMailboxes);
            this.groupBox5.Controls.Add(this.buttonDeselectAllMailboxes);
            this.groupBox5.Controls.Add(this.buttonLoadMailboxes);
            this.groupBox5.Controls.Add(this.buttonEditMailboxes);
            this.groupBox5.Location = new System.Drawing.Point(443, 66);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(278, 161);
            this.groupBox5.TabIndex = 16;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Mailboxes";
            // 
            // checkedListBoxMailboxes
            // 
            this.checkedListBoxMailboxes.CheckOnClick = true;
            this.checkedListBoxMailboxes.FormattingEnabled = true;
            this.checkedListBoxMailboxes.Location = new System.Drawing.Point(6, 19);
            this.checkedListBoxMailboxes.Name = "checkedListBoxMailboxes";
            this.checkedListBoxMailboxes.Size = new System.Drawing.Size(265, 79);
            this.checkedListBoxMailboxes.TabIndex = 8;
            // 
            // buttonSelectAllMailboxes
            // 
            this.buttonSelectAllMailboxes.Location = new System.Drawing.Point(125, 104);
            this.buttonSelectAllMailboxes.Name = "buttonSelectAllMailboxes";
            this.buttonSelectAllMailboxes.Size = new System.Drawing.Size(70, 23);
            this.buttonSelectAllMailboxes.TabIndex = 9;
            this.buttonSelectAllMailboxes.Text = "Select all";
            this.buttonSelectAllMailboxes.UseVisualStyleBackColor = true;
            this.buttonSelectAllMailboxes.Click += new System.EventHandler(this.buttonSelectAllMailboxes_Click);
            // 
            // buttonDeselectAllMailboxes
            // 
            this.buttonDeselectAllMailboxes.Location = new System.Drawing.Point(201, 104);
            this.buttonDeselectAllMailboxes.Name = "buttonDeselectAllMailboxes";
            this.buttonDeselectAllMailboxes.Size = new System.Drawing.Size(70, 23);
            this.buttonDeselectAllMailboxes.TabIndex = 10;
            this.buttonDeselectAllMailboxes.Text = "Deselect all";
            this.buttonDeselectAllMailboxes.UseVisualStyleBackColor = true;
            this.buttonDeselectAllMailboxes.Click += new System.EventHandler(this.buttonDeselectAllMailboxes_Click);
            // 
            // buttonLoadMailboxes
            // 
            this.buttonLoadMailboxes.Location = new System.Drawing.Point(62, 104);
            this.buttonLoadMailboxes.Name = "buttonLoadMailboxes";
            this.buttonLoadMailboxes.Size = new System.Drawing.Size(50, 23);
            this.buttonLoadMailboxes.TabIndex = 13;
            this.buttonLoadMailboxes.Text = "Load...";
            this.toolTips.SetToolTip(this.buttonLoadMailboxes, "Load mailboxes from text file");
            this.buttonLoadMailboxes.UseVisualStyleBackColor = true;
            this.buttonLoadMailboxes.Click += new System.EventHandler(this.buttonLoadMailboxes_Click);
            // 
            // buttonEditMailboxes
            // 
            this.buttonEditMailboxes.Location = new System.Drawing.Point(6, 104);
            this.buttonEditMailboxes.Name = "buttonEditMailboxes";
            this.buttonEditMailboxes.Size = new System.Drawing.Size(50, 23);
            this.buttonEditMailboxes.TabIndex = 18;
            this.buttonEditMailboxes.Text = "Edit...";
            this.buttonEditMailboxes.UseVisualStyleBackColor = true;
            this.buttonEditMailboxes.Click += new System.EventHandler(this.buttonEditMailboxes_Click);
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
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxEWSUri);
            this.groupBox1.Controls.Add(this.radioButtonSpecificUri);
            this.groupBox1.Controls.Add(this.radioButtonOffice365);
            this.groupBox1.Controls.Add(this.radioButtonAutodiscover);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(614, 48);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "EWS/Autodiscover Configuration";
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
            this.radioButtonSpecificUri.Enabled = false;
            this.radioButtonSpecificUri.Location = new System.Drawing.Point(179, 19);
            this.radioButtonSpecificUri.Name = "radioButtonSpecificUri";
            this.radioButtonSpecificUri.Size = new System.Drawing.Size(41, 17);
            this.radioButtonSpecificUri.TabIndex = 21;
            this.radioButtonSpecificUri.Text = "Uri:";
            this.radioButtonSpecificUri.UseVisualStyleBackColor = true;
            this.radioButtonSpecificUri.CheckedChanged += new System.EventHandler(this.radioButtonSpecificUri_CheckedChanged);
            // 
            // radioButtonOffice365
            // 
            this.radioButtonOffice365.AutoSize = true;
            this.radioButtonOffice365.Checked = true;
            this.radioButtonOffice365.Location = new System.Drawing.Point(99, 19);
            this.radioButtonOffice365.Name = "radioButtonOffice365";
            this.radioButtonOffice365.Size = new System.Drawing.Size(74, 17);
            this.radioButtonOffice365.TabIndex = 20;
            this.radioButtonOffice365.TabStop = true;
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
            // 
            // listBoxEvents
            // 
            this.listBoxEvents.FormattingEnabled = true;
            this.listBoxEvents.Location = new System.Drawing.Point(12, 345);
            this.listBoxEvents.Name = "listBoxEvents";
            this.listBoxEvents.Size = new System.Drawing.Size(1042, 225);
            this.listBoxEvents.TabIndex = 12;
            this.listBoxEvents.DoubleClick += new System.EventHandler(this.listBoxEvents_DoubleClick);
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
            this.groupBox2.Location = new System.Drawing.Point(443, 233);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(611, 106);
            this.groupBox2.TabIndex = 14;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Subscription Information";
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
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 329);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Received Events:";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1066, 583);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.listBoxEvents);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.Text = "EWS Streaming Subscriptions Sample Application";
            this.groupBoxOAuthAuthMethod.ResumeLayout(false);
            this.groupBoxOAuthAuthMethod.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownTimeout)).EndInit();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownMaxMailboxesInGroup)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer timerMonitorConnections;
        private System.Windows.Forms.GroupBox groupBoxOAuthAuthMethod;
        private System.Windows.Forms.Button buttonSelectCertificate;
        private System.Windows.Forms.RadioButton radioButtonAuthWithClientSecret;
        private System.Windows.Forms.TextBox textBoxAuthCertificate;
        private System.Windows.Forms.TextBox textBoxClientSecret;
        private System.Windows.Forms.RadioButton radioButtonAuthWithCertificate;
        private System.Windows.Forms.Button buttonUnsubscribe;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.NumericUpDown numericUpDownTimeout;
        private System.Windows.Forms.CheckBox checkBoxIncludeMime;
        private System.Windows.Forms.CheckBox checkBoxSelectAll;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button buttonSubscribe;
        private System.Windows.Forms.CheckBox checkBoxQueryMore;
        private System.Windows.Forms.TextBox textBoxNotificationCount;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textBoxNumConnections;
        private System.Windows.Forms.TextBox textBoxNumSubscriptions;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button buttonLoadCertificate;
        private System.Windows.Forms.TextBox textBoxTenantId;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox textBoxApplicationId;
        private System.Windows.Forms.RadioButton radioButtonAuthOAuth;
        private System.Windows.Forms.RadioButton radioButtonAuthBasic;
        private System.Windows.Forms.TextBox textBoxUsername;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxDomain;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numericUpDownMaxMailboxesInGroup;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.CheckedListBox checkedListBoxMailboxes;
        private System.Windows.Forms.Button buttonSelectAllMailboxes;
        private System.Windows.Forms.Button buttonDeselectAllMailboxes;
        private System.Windows.Forms.Button buttonLoadMailboxes;
        private System.Windows.Forms.ToolTip toolTips;
        private System.Windows.Forms.Button buttonEditMailboxes;
        private System.Windows.Forms.CheckBox checkBoxShowFolderEvents;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBoxEWSUri;
        private System.Windows.Forms.RadioButton radioButtonSpecificUri;
        private System.Windows.Forms.RadioButton radioButtonOffice365;
        private System.Windows.Forms.RadioButton radioButtonAutodiscover;
        private System.Windows.Forms.CheckedListBox checkedListBoxEvents;
        private System.Windows.Forms.ListBox listBoxEvents;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBoxShowItemEvents;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBoxSubscribeTo;
        private System.Windows.Forms.Label label3;
    }
}


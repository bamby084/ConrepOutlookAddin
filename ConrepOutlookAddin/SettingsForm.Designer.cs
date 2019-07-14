namespace ConrepOutlookAddin
{
    partial class SettingsForm
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
            PresentationControls.CheckBoxProperties checkBoxProperties1 = new PresentationControls.CheckBoxProperties();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtCompanyId = new System.Windows.Forms.TextBox();
            this.txtLoginUrl = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.chkLogOutgoingEmails = new System.Windows.Forms.CheckBox();
            this.chkLogIncomingEmails = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtCalendarSyncInterval = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cboAccounts = new PresentationControls.CheckBoxComboBox();
            this.btnUpdateSettings = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.chkShowPane = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtCompanyId);
            this.groupBox1.Controls.Add(this.txtLoginUrl);
            this.groupBox1.Controls.Add(this.txtUserName);
            this.groupBox1.Controls.Add(this.txtPassword);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(696, 92);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Login";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 53);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(88, 17);
            this.label4.TabIndex = 5;
            this.label4.Text = "Company ID:";
            // 
            // txtCompanyId
            // 
            this.txtCompanyId.Location = new System.Drawing.Point(134, 50);
            this.txtCompanyId.MaxLength = 6;
            this.txtCompanyId.Name = "txtCompanyId";
            this.txtCompanyId.Size = new System.Drawing.Size(202, 22);
            this.txtCompanyId.TabIndex = 1;
            // 
            // txtLoginUrl
            // 
            this.txtLoginUrl.Location = new System.Drawing.Point(134, 21);
            this.txtLoginUrl.Name = "txtLoginUrl";
            this.txtLoginUrl.Size = new System.Drawing.Size(202, 22);
            this.txtLoginUrl.TabIndex = 0;
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(478, 23);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(202, 22);
            this.txtUserName.TabIndex = 2;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(478, 51);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(202, 22);
            this.txtPassword.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(369, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(101, 17);
            this.label3.TabIndex = 3;
            this.label3.Text = "Password Key:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Host/Login Url:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(369, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "User Name:";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(518, 445);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(93, 34);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(617, 445);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(93, 34);
            this.btnOk.TabIndex = 7;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // chkLogOutgoingEmails
            // 
            this.chkLogOutgoingEmails.AutoSize = true;
            this.chkLogOutgoingEmails.Enabled = false;
            this.chkLogOutgoingEmails.Location = new System.Drawing.Point(19, 117);
            this.chkLogOutgoingEmails.Name = "chkLogOutgoingEmails";
            this.chkLogOutgoingEmails.Size = new System.Drawing.Size(433, 21);
            this.chkLogOutgoingEmails.TabIndex = 6;
            this.chkLogOutgoingEmails.Text = "Automatically log outgoing emails for existing contacts in Conrep";
            this.chkLogOutgoingEmails.UseVisualStyleBackColor = true;
            // 
            // chkLogIncomingEmails
            // 
            this.chkLogIncomingEmails.AutoSize = true;
            this.chkLogIncomingEmails.Enabled = false;
            this.chkLogIncomingEmails.Location = new System.Drawing.Point(19, 89);
            this.chkLogIncomingEmails.Name = "chkLogIncomingEmails";
            this.chkLogIncomingEmails.Size = new System.Drawing.Size(434, 21);
            this.chkLogIncomingEmails.TabIndex = 5;
            this.chkLogIncomingEmails.Text = "Automatically log incoming emails for existing contacts in Conrep";
            this.chkLogIncomingEmails.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtCalendarSyncInterval);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Location = new System.Drawing.Point(12, 121);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(696, 61);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Calendar Sync Options";
            // 
            // txtCalendarSyncInterval
            // 
            this.txtCalendarSyncInterval.Enabled = false;
            this.txtCalendarSyncInterval.Location = new System.Drawing.Point(238, 24);
            this.txtCalendarSyncInterval.Name = "txtCalendarSyncInterval";
            this.txtCalendarSyncInterval.Size = new System.Drawing.Size(98, 22);
            this.txtCalendarSyncInterval.TabIndex = 1;
            this.txtCalendarSyncInterval.TabStop = false;
            this.txtCalendarSyncInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(19, 27);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(213, 17);
            this.label6.TabIndex = 0;
            this.label6.Text = "Auto sync Calendar in (minutes):";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.cboAccounts);
            this.groupBox3.Controls.Add(this.chkLogOutgoingEmails);
            this.groupBox3.Controls.Add(this.chkLogIncomingEmails);
            this.groupBox3.Location = new System.Drawing.Point(12, 199);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(696, 154);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Log Settings";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 32);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(448, 17);
            this.label5.TabIndex = 6;
            this.label5.Text = "Enable Auto-Logging incoming/outgoing emails for following accounts:";
            // 
            // cboAccounts
            // 
            checkBoxProperties1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cboAccounts.CheckBoxProperties = checkBoxProperties1;
            this.cboAccounts.DisplayMemberSingleItem = "";
            this.cboAccounts.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboAccounts.FormattingEnabled = true;
            this.cboAccounts.Location = new System.Drawing.Point(20, 52);
            this.cboAccounts.Name = "cboAccounts";
            this.cboAccounts.Size = new System.Drawing.Size(316, 24);
            this.cboAccounts.TabIndex = 4;
            this.cboAccounts.TabStop = false;
            // 
            // btnUpdateSettings
            // 
            this.btnUpdateSettings.Location = new System.Drawing.Point(291, 446);
            this.btnUpdateSettings.Name = "btnUpdateSettings";
            this.btnUpdateSettings.Size = new System.Drawing.Size(221, 33);
            this.btnUpdateSettings.TabIndex = 5;
            this.btnUpdateSettings.Text = "Validate && Update Settings";
            this.btnUpdateSettings.UseVisualStyleBackColor = true;
            this.btnUpdateSettings.Click += new System.EventHandler(this.btnUpdateSettings_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.chkShowPane);
            this.groupBox4.Location = new System.Drawing.Point(12, 368);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(696, 57);
            this.groupBox4.TabIndex = 16;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Misc";
            // 
            // chkShowPane
            // 
            this.chkShowPane.AutoSize = true;
            this.chkShowPane.Location = new System.Drawing.Point(19, 22);
            this.chkShowPane.Name = "chkShowPane";
            this.chkShowPane.Size = new System.Drawing.Size(358, 21);
            this.chkShowPane.TabIndex = 4;
            this.chkShowPane.Text = "Show Conrep Pane in Reading window automatically";
            this.chkShowPane.UseVisualStyleBackColor = true;
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(722, 491);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.btnUpdateSettings);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Conrep Settings";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtLoginUrl;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.CheckBox chkLogOutgoingEmails;
        private System.Windows.Forms.CheckBox chkLogIncomingEmails;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtCompanyId;
        private System.Windows.Forms.GroupBox groupBox3;
        private PresentationControls.CheckBoxComboBox cboAccounts;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnUpdateSettings;
        private System.Windows.Forms.TextBox txtCalendarSyncInterval;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckBox chkShowPane;
    }
}
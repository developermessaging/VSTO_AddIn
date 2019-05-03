namespace VSTO_AddIn
{
    partial class ControlPanel
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.verboseStoresLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseFolderLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseFoldersLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseNameSpaceEventsBox = new System.Windows.Forms.CheckBox();
            this.verboseItemsLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseInspectorLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseInspectorsLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseExplorerLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseExplorersLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseApplicationLoggingBox = new System.Windows.Forms.CheckBox();
            this.verboseItemLoggingBox = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.trackInboxBox = new System.Windows.Forms.CheckBox();
            this.trackOutboxBox = new System.Windows.Forms.CheckBox();
            this.trackSentItemsBox = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.gcButton = new System.Windows.Forms.Button();
            this.clearLogButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.verboseStoresLoggingBox);
            this.groupBox1.Controls.Add(this.verboseFolderLoggingBox);
            this.groupBox1.Controls.Add(this.verboseFoldersLoggingBox);
            this.groupBox1.Controls.Add(this.verboseNameSpaceEventsBox);
            this.groupBox1.Controls.Add(this.verboseItemsLoggingBox);
            this.groupBox1.Controls.Add(this.verboseInspectorLoggingBox);
            this.groupBox1.Controls.Add(this.verboseInspectorsLoggingBox);
            this.groupBox1.Controls.Add(this.verboseExplorerLoggingBox);
            this.groupBox1.Controls.Add(this.verboseExplorersLoggingBox);
            this.groupBox1.Controls.Add(this.verboseApplicationLoggingBox);
            this.groupBox1.Controls.Add(this.verboseItemLoggingBox);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(400, 130);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Enable logging for:";
            // 
            // verboseStoresLoggingBox
            // 
            this.verboseStoresLoggingBox.AutoSize = true;
            this.verboseStoresLoggingBox.Checked = true;
            this.verboseStoresLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseStoresLoggingBox.Location = new System.Drawing.Point(259, 65);
            this.verboseStoresLoggingBox.Name = "verboseStoresLoggingBox";
            this.verboseStoresLoggingBox.Size = new System.Drawing.Size(56, 17);
            this.verboseStoresLoggingBox.TabIndex = 10;
            this.verboseStoresLoggingBox.Text = "Stores";
            this.verboseStoresLoggingBox.UseVisualStyleBackColor = true;
            this.verboseStoresLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseStoresLoggingBox_CheckedChanged);
            // 
            // verboseFolderLoggingBox
            // 
            this.verboseFolderLoggingBox.AutoSize = true;
            this.verboseFolderLoggingBox.Checked = true;
            this.verboseFolderLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseFolderLoggingBox.Location = new System.Drawing.Point(259, 42);
            this.verboseFolderLoggingBox.Name = "verboseFolderLoggingBox";
            this.verboseFolderLoggingBox.Size = new System.Drawing.Size(55, 17);
            this.verboseFolderLoggingBox.TabIndex = 9;
            this.verboseFolderLoggingBox.Text = "Folder";
            this.verboseFolderLoggingBox.UseVisualStyleBackColor = true;
            this.verboseFolderLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseFolderLoggingBox_CheckedChanged);
            // 
            // verboseFoldersLoggingBox
            // 
            this.verboseFoldersLoggingBox.AutoSize = true;
            this.verboseFoldersLoggingBox.Checked = true;
            this.verboseFoldersLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseFoldersLoggingBox.Location = new System.Drawing.Point(259, 19);
            this.verboseFoldersLoggingBox.Name = "verboseFoldersLoggingBox";
            this.verboseFoldersLoggingBox.Size = new System.Drawing.Size(60, 17);
            this.verboseFoldersLoggingBox.TabIndex = 8;
            this.verboseFoldersLoggingBox.Text = "Folders";
            this.verboseFoldersLoggingBox.UseVisualStyleBackColor = true;
            this.verboseFoldersLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseFoldersLoggingBox_CheckedChanged);
            // 
            // verboseNameSpaceEventsBox
            // 
            this.verboseNameSpaceEventsBox.AutoSize = true;
            this.verboseNameSpaceEventsBox.Checked = true;
            this.verboseNameSpaceEventsBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseNameSpaceEventsBox.Location = new System.Drawing.Point(137, 88);
            this.verboseNameSpaceEventsBox.Name = "verboseNameSpaceEventsBox";
            this.verboseNameSpaceEventsBox.Size = new System.Drawing.Size(85, 17);
            this.verboseNameSpaceEventsBox.TabIndex = 7;
            this.verboseNameSpaceEventsBox.Text = "NameSpace";
            this.verboseNameSpaceEventsBox.UseVisualStyleBackColor = true;
            this.verboseNameSpaceEventsBox.CheckedChanged += new System.EventHandler(this.VerboseNameSpaceEventsBox_CheckedChanged);
            // 
            // verboseItemsLoggingBox
            // 
            this.verboseItemsLoggingBox.AutoSize = true;
            this.verboseItemsLoggingBox.Checked = true;
            this.verboseItemsLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseItemsLoggingBox.Location = new System.Drawing.Point(6, 19);
            this.verboseItemsLoggingBox.Name = "verboseItemsLoggingBox";
            this.verboseItemsLoggingBox.Size = new System.Drawing.Size(51, 17);
            this.verboseItemsLoggingBox.TabIndex = 6;
            this.verboseItemsLoggingBox.Text = "Items";
            this.verboseItemsLoggingBox.UseVisualStyleBackColor = true;
            this.verboseItemsLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseItemsLoggingBox_CheckedChanged);
            // 
            // verboseInspectorLoggingBox
            // 
            this.verboseInspectorLoggingBox.AutoSize = true;
            this.verboseInspectorLoggingBox.Checked = true;
            this.verboseInspectorLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseInspectorLoggingBox.Location = new System.Drawing.Point(137, 42);
            this.verboseInspectorLoggingBox.Name = "verboseInspectorLoggingBox";
            this.verboseInspectorLoggingBox.Size = new System.Drawing.Size(70, 17);
            this.verboseInspectorLoggingBox.TabIndex = 5;
            this.verboseInspectorLoggingBox.Text = "Inspector";
            this.verboseInspectorLoggingBox.UseVisualStyleBackColor = true;
            this.verboseInspectorLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseInspectorLoggingBox_CheckedChanged);
            // 
            // verboseInspectorsLoggingBox
            // 
            this.verboseInspectorsLoggingBox.AutoSize = true;
            this.verboseInspectorsLoggingBox.Checked = true;
            this.verboseInspectorsLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseInspectorsLoggingBox.Location = new System.Drawing.Point(137, 19);
            this.verboseInspectorsLoggingBox.Name = "verboseInspectorsLoggingBox";
            this.verboseInspectorsLoggingBox.Size = new System.Drawing.Size(75, 17);
            this.verboseInspectorsLoggingBox.TabIndex = 4;
            this.verboseInspectorsLoggingBox.Text = "Inspectors";
            this.verboseInspectorsLoggingBox.UseVisualStyleBackColor = true;
            this.verboseInspectorsLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseInspectorsLoggingBox_CheckedChanged);
            // 
            // verboseExplorerLoggingBox
            // 
            this.verboseExplorerLoggingBox.AutoSize = true;
            this.verboseExplorerLoggingBox.Checked = true;
            this.verboseExplorerLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseExplorerLoggingBox.Location = new System.Drawing.Point(6, 88);
            this.verboseExplorerLoggingBox.Name = "verboseExplorerLoggingBox";
            this.verboseExplorerLoggingBox.Size = new System.Drawing.Size(64, 17);
            this.verboseExplorerLoggingBox.TabIndex = 3;
            this.verboseExplorerLoggingBox.Text = "Explorer";
            this.verboseExplorerLoggingBox.UseVisualStyleBackColor = true;
            this.verboseExplorerLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseExplorerLoggingBox_CheckedChanged);
            // 
            // verboseExplorersLoggingBox
            // 
            this.verboseExplorersLoggingBox.AutoSize = true;
            this.verboseExplorersLoggingBox.Checked = true;
            this.verboseExplorersLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseExplorersLoggingBox.Location = new System.Drawing.Point(6, 65);
            this.verboseExplorersLoggingBox.Name = "verboseExplorersLoggingBox";
            this.verboseExplorersLoggingBox.Size = new System.Drawing.Size(69, 17);
            this.verboseExplorersLoggingBox.TabIndex = 2;
            this.verboseExplorersLoggingBox.Text = "Explorers";
            this.verboseExplorersLoggingBox.UseVisualStyleBackColor = true;
            this.verboseExplorersLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseExplorersLoggingBox_CheckedChanged);
            // 
            // verboseApplicationLoggingBox
            // 
            this.verboseApplicationLoggingBox.AutoSize = true;
            this.verboseApplicationLoggingBox.Checked = true;
            this.verboseApplicationLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseApplicationLoggingBox.Location = new System.Drawing.Point(137, 65);
            this.verboseApplicationLoggingBox.Name = "verboseApplicationLoggingBox";
            this.verboseApplicationLoggingBox.Size = new System.Drawing.Size(78, 17);
            this.verboseApplicationLoggingBox.TabIndex = 1;
            this.verboseApplicationLoggingBox.Text = "Application";
            this.verboseApplicationLoggingBox.UseVisualStyleBackColor = true;
            this.verboseApplicationLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseApplicationLoggingBox_CheckedChanged);
            // 
            // verboseItemLoggingBox
            // 
            this.verboseItemLoggingBox.AutoSize = true;
            this.verboseItemLoggingBox.Checked = true;
            this.verboseItemLoggingBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verboseItemLoggingBox.Location = new System.Drawing.Point(6, 42);
            this.verboseItemLoggingBox.Name = "verboseItemLoggingBox";
            this.verboseItemLoggingBox.Size = new System.Drawing.Size(46, 17);
            this.verboseItemLoggingBox.TabIndex = 0;
            this.verboseItemLoggingBox.Text = "Item";
            this.verboseItemLoggingBox.UseVisualStyleBackColor = true;
            this.verboseItemLoggingBox.CheckedChanged += new System.EventHandler(this.VerboseItemLoggingBox_CheckedChanged);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox3);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 130);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(400, 844);
            this.panel1.TabIndex = 1;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.trackSentItemsBox);
            this.groupBox2.Controls.Add(this.trackOutboxBox);
            this.groupBox2.Controls.Add(this.trackInboxBox);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(400, 97);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Tracked folders:";
            // 
            // trackInboxBox
            // 
            this.trackInboxBox.AutoSize = true;
            this.trackInboxBox.Checked = true;
            this.trackInboxBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.trackInboxBox.Location = new System.Drawing.Point(7, 20);
            this.trackInboxBox.Name = "trackInboxBox";
            this.trackInboxBox.Size = new System.Drawing.Size(52, 17);
            this.trackInboxBox.TabIndex = 0;
            this.trackInboxBox.Text = "Inbox";
            this.trackInboxBox.UseVisualStyleBackColor = true;
            this.trackInboxBox.CheckedChanged += new System.EventHandler(this.TrackInboxBox_CheckedChanged);
            // 
            // trackOutboxBox
            // 
            this.trackOutboxBox.AutoSize = true;
            this.trackOutboxBox.Checked = true;
            this.trackOutboxBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.trackOutboxBox.Location = new System.Drawing.Point(7, 43);
            this.trackOutboxBox.Name = "trackOutboxBox";
            this.trackOutboxBox.Size = new System.Drawing.Size(60, 17);
            this.trackOutboxBox.TabIndex = 1;
            this.trackOutboxBox.Text = "Outbox";
            this.trackOutboxBox.UseVisualStyleBackColor = true;
            this.trackOutboxBox.CheckedChanged += new System.EventHandler(this.TrackOutboxBox_CheckedChanged);
            // 
            // trackSentItemsBox
            // 
            this.trackSentItemsBox.AutoSize = true;
            this.trackSentItemsBox.Checked = true;
            this.trackSentItemsBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.trackSentItemsBox.Location = new System.Drawing.Point(6, 66);
            this.trackSentItemsBox.Name = "trackSentItemsBox";
            this.trackSentItemsBox.Size = new System.Drawing.Size(73, 17);
            this.trackSentItemsBox.TabIndex = 2;
            this.trackSentItemsBox.Text = "SentItems";
            this.trackSentItemsBox.UseVisualStyleBackColor = true;
            this.trackSentItemsBox.CheckedChanged += new System.EventHandler(this.TrackSentItemsBox_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.clearLogButton);
            this.groupBox3.Controls.Add(this.gcButton);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox3.Location = new System.Drawing.Point(0, 97);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(400, 100);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Utility";
            // 
            // gcButton
            // 
            this.gcButton.Location = new System.Drawing.Point(7, 20);
            this.gcButton.Name = "gcButton";
            this.gcButton.Size = new System.Drawing.Size(75, 23);
            this.gcButton.TabIndex = 0;
            this.gcButton.Text = "GC";
            this.gcButton.UseVisualStyleBackColor = true;
            this.gcButton.Click += new System.EventHandler(this.GcButton_Click);
            // 
            // clearLogButton
            // 
            this.clearLogButton.Location = new System.Drawing.Point(7, 49);
            this.clearLogButton.Name = "clearLogButton";
            this.clearLogButton.Size = new System.Drawing.Size(75, 23);
            this.clearLogButton.TabIndex = 1;
            this.clearLogButton.Text = "Clear Log";
            this.clearLogButton.UseVisualStyleBackColor = true;
            this.clearLogButton.Click += new System.EventHandler(this.ClearLogButton_Click);
            // 
            // ControlPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.groupBox1);
            this.Name = "ControlPanel";
            this.Size = new System.Drawing.Size(400, 974);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox verboseItemLoggingBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox verboseApplicationLoggingBox;
        private System.Windows.Forms.CheckBox verboseExplorerLoggingBox;
        private System.Windows.Forms.CheckBox verboseExplorersLoggingBox;
        private System.Windows.Forms.CheckBox verboseInspectorLoggingBox;
        private System.Windows.Forms.CheckBox verboseInspectorsLoggingBox;
        private System.Windows.Forms.CheckBox verboseItemsLoggingBox;
        private System.Windows.Forms.CheckBox verboseNameSpaceEventsBox;
        private System.Windows.Forms.CheckBox verboseFolderLoggingBox;
        private System.Windows.Forms.CheckBox verboseFoldersLoggingBox;
        private System.Windows.Forms.CheckBox verboseStoresLoggingBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox trackSentItemsBox;
        private System.Windows.Forms.CheckBox trackOutboxBox;
        private System.Windows.Forms.CheckBox trackInboxBox;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button clearLogButton;
        private System.Windows.Forms.Button gcButton;
    }
}

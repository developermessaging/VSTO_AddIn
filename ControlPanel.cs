using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSTO_AddIn
{
    public partial class ControlPanel : UserControl
    {
        private static System.Windows.Forms.TextBox logBox;

        internal static bool ItemLoggingEnabled
        {
            get
            {
                return itemLoggingEnabled;
            }
        }

        private static bool itemLoggingEnabled = false;

        internal static bool FoldersLoggingEnabled
        {
            get
            {
                return foldersLoggingEnabled;
            }
        }

        private static bool foldersLoggingEnabled = false;

        internal static bool FolderLoggingEnabled
        {
            get
            {
                return folderLoggingEnabled;
            }
        }

        private static bool folderLoggingEnabled = false;

        internal static bool ApplicationLoggingEnabled
        {
            get
            {
                return applicationLoggingEnabled;
            }
        }

        private static bool applicationLoggingEnabled = true;

        internal static bool ExplorersLoggingEnabled
        {
            get
            {
                return explorersLoggingEnabled;
            }
        }

        private static bool explorersLoggingEnabled = true;

        internal static bool ExplorerLoggingEnabled
        {
            get
            {
                return explorerLoggingEnabled;
            }
        }

        private static bool explorerLoggingEnabled = true;

        internal static bool InspectorsLoggingEnabled
        {
            get
            {
                return inspectorsLoggingEnabled;
            }
        }

        private static bool inspectorsLoggingEnabled = true;

        internal static bool InspectorLoggingEnabled
        {
            get
            {
                return inspectorLoggingEnabled;
            }
        }

        private static bool inspectorLoggingEnabled = true;

        internal static bool ItemsLoggingEnabled
        {
            get
            {
                return itemsLoggingEnabled;
            }
        }

        private static bool itemsLoggingEnabled = true;

        internal static bool NameSpaceLoggingEnabled
        {
            get
            {
                return nameSpaceLoggingEnabled;
            }
        }

        private static bool nameSpaceLoggingEnabled = true;

        internal static bool StoresLoggingEnabled
        {
            get
            {
                return storesLoggingEnabled;
            }
        }

        private static bool storesLoggingEnabled = true;

        internal static bool TrackInboxEnabled
        {
            get
            {
                return trackInboxEnabled;
            }
        }

        private static bool trackInboxEnabled = true;

        internal static bool TrackOutboxEnabled
        {
            get
            {
                return trackOutboxEnabled;
            }
        }

        private static bool trackOutboxEnabled = true;

        internal static bool TrackSentItemsEnabled
        {
            get
            {
                return trackSentItemsEnabled;
            }
        }

        private static bool trackSentItemsEnabled = true;

        public ControlPanel()
        {
            InitializeComponent();

            logBox = new System.Windows.Forms.TextBox();
            this.panel1.Controls.Add(logBox);
            // 
            // logBox
            // 
            logBox.Dock = System.Windows.Forms.DockStyle.Fill;
            logBox.Location = new System.Drawing.Point(0, 0);
            logBox.Multiline = true;
            logBox.Name = "logBox";
            logBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            logBox.Size = new System.Drawing.Size(408, 845);
            logBox.TabIndex = 0;
        }

        private void VerboseItemLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            itemLoggingEnabled = verboseItemLoggingBox.Checked;
        }

        private void VerboseApplicationLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            applicationLoggingEnabled = verboseApplicationLoggingBox.Checked;
        }

        static internal void AddText(string message)
        {
            logBox.AppendText(message);
        }

        private void VerboseExplorersLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            explorersLoggingEnabled = verboseExplorersLoggingBox.Checked;
        }

        private void VerboseExplorerLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            explorerLoggingEnabled = verboseExplorerLoggingBox.Checked;
        }

        private void VerboseItemsLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            itemsLoggingEnabled = verboseItemsLoggingBox.Checked;
        }

        private void VerboseInspectorsLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            inspectorsLoggingEnabled = verboseInspectorsLoggingBox.Checked;
        }

        private void VerboseInspectorLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            inspectorLoggingEnabled = verboseInspectorLoggingBox.Checked;
        }

        private void VerboseNameSpaceEventsBox_CheckedChanged(object sender, EventArgs e)
        {
            nameSpaceLoggingEnabled = verboseNameSpaceEventsBox.Checked;
        }

        private void VerboseFoldersLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            foldersLoggingEnabled = verboseFoldersLoggingBox.Checked;
        }

        private void VerboseFolderLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            folderLoggingEnabled = verboseFolderLoggingBox.Checked;
        }

        private void VerboseStoresLoggingBox_CheckedChanged(object sender, EventArgs e)
        {
            storesLoggingEnabled = verboseStoresLoggingBox.Checked;
        }

        private void TrackInboxBox_CheckedChanged(object sender, EventArgs e)
        {
            trackInboxEnabled = trackInboxBox.Checked;
        }

        private void TrackOutboxBox_CheckedChanged(object sender, EventArgs e)
        {
            trackOutboxEnabled = trackOutboxBox.Checked;
        }

        private void TrackSentItemsBox_CheckedChanged(object sender, EventArgs e)
        {
            trackSentItemsEnabled = trackSentItemsBox.Checked;
        }

        private void GcButton_Click(object sender, EventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private static void ClearLog()
        {
            logBox.Clear();
        }
        private void ClearLogButton_Click(object sender, EventArgs e)
        {
            ClearLog();
        }
    }
}

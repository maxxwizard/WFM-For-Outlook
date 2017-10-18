namespace WFM_For_Outlook
{
    partial class CalendarIntegrationRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CalendarIntegrationRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.components = new System.ComponentModel.Container();
            this.tabWfmForOutlook = this.Factory.CreateRibbonTab();
            this.grpMeetingOptions = this.Factory.CreateRibbonGroup();
            this.grpSyncOptions = this.Factory.CreateRibbonGroup();
            this.grpSyncMisc = this.Factory.CreateRibbonGroup();
            this.grpSyncStatus = this.Factory.CreateRibbonGroup();
            this.labelLastSyncTime = this.Factory.CreateRibbonLabel();
            this.labelLastSyncStatus = this.Factory.CreateRibbonLabel();
            this.labelNextSyncTime = this.Factory.CreateRibbonLabel();
            this.syncBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.syncTimer = new System.Windows.Forms.Timer(this.components);
            this.galleryReminder = this.Factory.CreateRibbonGallery();
            this.galleryAvailStatus = this.Factory.CreateRibbonGallery();
            this.galleryCategory = this.Factory.CreateRibbonGallery();
            this.gallerySyncMode = this.Factory.CreateRibbonGallery();
            this.btnSubject = this.Factory.CreateRibbonButton();
            this.galleryPollingInterval = this.Factory.CreateRibbonGallery();
            this.galleryDaysToPull = this.Factory.CreateRibbonGallery();
            this.btnSegmentName = this.Factory.CreateRibbonButton();
            this.btnSyncNow = this.Factory.CreateRibbonButton();
            this.btnSyncLog = this.Factory.CreateRibbonButton();
            this.btnResetSettings = this.Factory.CreateRibbonButton();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tabWfmForOutlook.SuspendLayout();
            this.grpMeetingOptions.SuspendLayout();
            this.grpSyncOptions.SuspendLayout();
            this.grpSyncMisc.SuspendLayout();
            this.grpSyncStatus.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabWfmForOutlook
            // 
            this.tabWfmForOutlook.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabWfmForOutlook.Groups.Add(this.grpMeetingOptions);
            this.tabWfmForOutlook.Groups.Add(this.grpSyncOptions);
            this.tabWfmForOutlook.Groups.Add(this.grpSyncMisc);
            this.tabWfmForOutlook.Groups.Add(this.grpSyncStatus);
            this.tabWfmForOutlook.Label = "WFM for Outlook";
            this.tabWfmForOutlook.Name = "tabWfmForOutlook";
            // 
            // grpMeetingOptions
            // 
            this.grpMeetingOptions.Items.Add(this.galleryReminder);
            this.grpMeetingOptions.Items.Add(this.galleryAvailStatus);
            this.grpMeetingOptions.Items.Add(this.galleryCategory);
            this.grpMeetingOptions.Items.Add(this.btnSubject);
            this.grpMeetingOptions.Label = "Meeting Options";
            this.grpMeetingOptions.Name = "grpMeetingOptions";
            // 
            // grpSyncOptions
            // 
            this.grpSyncOptions.Items.Add(this.gallerySyncMode);
            this.grpSyncOptions.Items.Add(this.galleryPollingInterval);
            this.grpSyncOptions.Items.Add(this.galleryDaysToPull);
            this.grpSyncOptions.Items.Add(this.btnSegmentName);
            this.grpSyncOptions.Label = "Sync Options";
            this.grpSyncOptions.Name = "grpSyncOptions";
            // 
            // grpSyncMisc
            // 
            this.grpSyncMisc.Items.Add(this.btnSyncNow);
            this.grpSyncMisc.Items.Add(this.btnSyncLog);
            this.grpSyncMisc.Items.Add(this.btnResetSettings);
            this.grpSyncMisc.Items.Add(this.btnHelp);
            this.grpSyncMisc.Name = "grpSyncMisc";
            // 
            // grpSyncStatus
            // 
            this.grpSyncStatus.Items.Add(this.labelLastSyncTime);
            this.grpSyncStatus.Items.Add(this.labelLastSyncStatus);
            this.grpSyncStatus.Items.Add(this.labelNextSyncTime);
            this.grpSyncStatus.Name = "grpSyncStatus";
            // 
            // labelLastSyncTime
            // 
            this.labelLastSyncTime.Label = "Last sync time:";
            this.labelLastSyncTime.Name = "labelLastSyncTime";
            // 
            // labelLastSyncStatus
            // 
            this.labelLastSyncStatus.Label = "Last sync status:";
            this.labelLastSyncStatus.Name = "labelLastSyncStatus";
            // 
            // labelNextSyncTime
            // 
            this.labelNextSyncTime.Label = "Next sync time:";
            this.labelNextSyncTime.Name = "labelNextSyncTime";
            // 
            // syncBackgroundWorker
            // 
            this.syncBackgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.syncBackgroundWorker_DoWork);
            // 
            // syncTimer
            // 
            this.syncTimer.Interval = 60000;
            this.syncTimer.Tick += new System.EventHandler(this.syncTimer_Tick);
            // 
            // galleryReminder
            // 
            this.galleryReminder.ColumnCount = 1;
            this.galleryReminder.Label = "Reminder";
            this.galleryReminder.Name = "galleryReminder";
            this.galleryReminder.OfficeImageId = "ReminderGallery";
            this.galleryReminder.ShowImage = true;
            this.galleryReminder.ShowItemImage = false;
            this.galleryReminder.ShowItemSelection = true;
            this.galleryReminder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryReminder_Click);
            // 
            // galleryAvailStatus
            // 
            this.galleryAvailStatus.ColumnCount = 1;
            this.galleryAvailStatus.Label = "Show As";
            this.galleryAvailStatus.Name = "galleryAvailStatus";
            this.galleryAvailStatus.OfficeImageId = "ShowTimeAsGallery";
            this.galleryAvailStatus.ShowImage = true;
            this.galleryAvailStatus.ShowItemImage = false;
            this.galleryAvailStatus.ShowItemSelection = true;
            this.galleryAvailStatus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryAvailStatus_Click);
            // 
            // galleryCategory
            // 
            this.galleryCategory.ColumnCount = 1;
            this.galleryCategory.Label = "Category";
            this.galleryCategory.Name = "galleryCategory";
            this.galleryCategory.OfficeImageId = "CategorizeGallery";
            this.galleryCategory.ShowImage = true;
            this.galleryCategory.ShowItemImage = false;
            this.galleryCategory.ShowItemSelection = true;
            this.galleryCategory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryCategory_Click);
            // 
            // gallerySyncMode
            // 
            this.gallerySyncMode.ColumnCount = 1;
            this.gallerySyncMode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallerySyncMode.Label = "Sync Mode";
            this.gallerySyncMode.Name = "gallerySyncMode";
            this.gallerySyncMode.OfficeImageId = "VideoContrastGallery";
            this.gallerySyncMode.ShowImage = true;
            this.gallerySyncMode.ShowItemSelection = true;
            this.gallerySyncMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallerySyncMode_Click);
            // 
            // btnSubject
            // 
            this.btnSubject.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSubject.Label = "Subject Prefix";
            this.btnSubject.Name = "btnSubject";
            this.btnSubject.OfficeImageId = "MemoSettingsMenu";
            this.btnSubject.ScreenTip = "Meetings created on calendar will have this prefix.";
            this.btnSubject.ShowImage = true;
            this.btnSubject.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSubject_Click);
            // 
            // galleryPollingInterval
            // 
            this.galleryPollingInterval.ColumnCount = 1;
            this.galleryPollingInterval.Label = "Sync Frequency";
            this.galleryPollingInterval.Name = "galleryPollingInterval";
            this.galleryPollingInterval.OfficeImageId = "SynchronizationStatus";
            this.galleryPollingInterval.ShowImage = true;
            this.galleryPollingInterval.ShowItemImage = false;
            this.galleryPollingInterval.ShowItemSelection = true;
            this.galleryPollingInterval.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryPollingInterval_Click);
            // 
            // galleryDaysToPull
            // 
            this.galleryDaysToPull.ColumnCount = 1;
            this.galleryDaysToPull.Label = "Days To Pull";
            this.galleryDaysToPull.Name = "galleryDaysToPull";
            this.galleryDaysToPull.OfficeImageId = "MeetingsToolToday";
            this.galleryDaysToPull.ShowImage = true;
            this.galleryDaysToPull.ShowItemImage = false;
            this.galleryDaysToPull.ShowItemSelection = true;
            this.galleryDaysToPull.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryDaysToPull_Click);
            // 
            // btnSegmentName
            // 
            this.btnSegmentName.Label = "Segment Filter";
            this.btnSegmentName.Name = "btnSegmentName";
            this.btnSegmentName.OfficeImageId = "Filters";
            this.btnSegmentName.ScreenTip = "List of segment name(s) you wish to include or exclude. Separate values using a s" +
    "emicolon delimiter.";
            this.btnSegmentName.ShowImage = true;
            this.btnSegmentName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSegmentName_Click);
            // 
            // btnSyncNow
            // 
            this.btnSyncNow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSyncNow.Label = "Sync Now";
            this.btnSyncNow.Name = "btnSyncNow";
            this.btnSyncNow.OfficeImageId = "Synchronize";
            this.btnSyncNow.ScreenTip = "Forces Outlook to do an immediate pull of WFM segments";
            this.btnSyncNow.ShowImage = true;
            this.btnSyncNow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSyncNow_Click);
            // 
            // btnSyncLog
            // 
            this.btnSyncLog.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSyncLog.Image = global::WFM_For_Outlook.Properties.Resources.log1;
            this.btnSyncLog.Label = "Open Sync Log";
            this.btnSyncLog.Name = "btnSyncLog";
            this.btnSyncLog.ShowImage = true;
            this.btnSyncLog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSyncLog_Click);
            // 
            // btnResetSettings
            // 
            this.btnResetSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnResetSettings.Label = "Reset Settings";
            this.btnResetSettings.Name = "btnResetSettings";
            this.btnResetSettings.OfficeImageId = "SyncSettingsMenu";
            this.btnResetSettings.ShowImage = true;
            this.btnResetSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetSettings_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHelp.Label = "Help";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.OfficeImageId = "Help";
            this.btnHelp.ShowImage = true;
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            // 
            // CalendarIntegrationRibbon
            // 
            this.Name = "CalendarIntegrationRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabWfmForOutlook);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CalendarIntegrationRibbon_Load);
            this.tabWfmForOutlook.ResumeLayout(false);
            this.tabWfmForOutlook.PerformLayout();
            this.grpMeetingOptions.ResumeLayout(false);
            this.grpMeetingOptions.PerformLayout();
            this.grpSyncOptions.ResumeLayout(false);
            this.grpSyncOptions.PerformLayout();
            this.grpSyncMisc.ResumeLayout(false);
            this.grpSyncMisc.PerformLayout();
            this.grpSyncStatus.ResumeLayout(false);
            this.grpSyncStatus.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabWfmForOutlook;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMeetingOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSyncOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSyncNow;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryReminder;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryAvailStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSubject;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryDaysToPull;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryPollingInterval;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSegmentName;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSyncStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelLastSyncTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelLastSyncStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelNextSyncTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSyncLog;
        public System.Windows.Forms.Timer syncTimer;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryCategory;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSyncMisc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetSettings;
        public System.ComponentModel.BackgroundWorker syncBackgroundWorker;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallerySyncMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
    }

    partial class ThisRibbonCollection
    {
        internal CalendarIntegrationRibbon CalendarIntegrationRibbon
        {
            get { return this.GetRibbon<CalendarIntegrationRibbon>(); }
        }
    }
}

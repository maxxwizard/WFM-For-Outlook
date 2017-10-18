using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO;
using System.ComponentModel;
using System.Threading;

namespace WFM_For_Outlook
{
    public partial class CalendarIntegrationRibbon
    {
        public const string LAST_SYNC_TIME = "Last sync time: ";
        public const string LAST_SYNC_STATUS = "Last sync status: ";
        public const string NEXT_SYNC_TIME = "Next sync time: ";
        
        private void CalendarIntegrationRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // attempt to retrieve saved options. if unsuccessful, create a new one and persist it to Exchange.
            //MessageBox.Show("loading user options now");
            Globals.ThisAddIn.userOptions = Options.LoadFromConfigItem();
            if (Globals.ThisAddIn.userOptions == null)
            {
                Globals.ThisAddIn.userOptions = new Options();
                Globals.ThisAddIn.userOptions.Save();
            }

            // initialize our controls to the stored config
            InitializeControls();

            // set up our background worker thread
            syncBackgroundWorker.WorkerSupportsCancellation = false;
            syncBackgroundWorker.WorkerReportsProgress = true;
            syncBackgroundWorker.ProgressChanged += syncBackgroundWorker_ProgressChanged;
            syncBackgroundWorker.RunWorkerCompleted += syncBackgroundWorker_RunWorkerCompleted;

            // start our timer
            syncTimer.Start();
        }

        void syncBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // update our labels
            RedrawSyncStatus();
        }

        void syncBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            labelLastSyncStatus.Label = LAST_SYNC_STATUS + e.ProgressPercentage + "%";
        }

        private void InitializeControls()
        {
            InitializeReminderGallery();
            InitializeAvailStatusGallery();
            InitializePollingIntervalGallery();
            InitializeDaysToPullGallery();
            InitializeCategoryGallery();
            InitializeSyncModeGallery();
            RedrawSyncStatus();
        }

        private void InitializeSyncModeGallery()
        {
            gallerySyncMode.Items.Clear();

            foreach (SyncMode mode in Enum.GetValues(typeof(SyncMode)))
            {
                var item = this.Factory.CreateRibbonDropDownItem();
                item.Label = mode.GetDescription();
                item.Tag = mode;
                gallerySyncMode.Items.Add(item);
                if (Globals.ThisAddIn.userOptions.syncMode == mode)
                {
                    gallerySyncMode.SelectedItem = item;
                }
            }
        }

        private void InitializeCategoryGallery()
        {
            galleryCategory.Items.Clear();

            Outlook.Categories categories = Globals.ThisAddIn.Application.ActiveExplorer().Session.Categories;

            var dropdownItem = this.Factory.CreateRibbonDropDownItem();
            dropdownItem.Label = "None";
            dropdownItem.Tag = string.Empty;
            galleryCategory.Items.Clear();
            galleryCategory.Items.Add(dropdownItem);
            galleryCategory.SelectedItem = dropdownItem;

            foreach (Outlook.Category cat in categories)
            {
                dropdownItem = this.Factory.CreateRibbonDropDownItem();
                dropdownItem.Label = cat.Name;
                dropdownItem.Tag = cat.CategoryID;
                galleryCategory.Items.Add(dropdownItem);
                if (Globals.ThisAddIn.userOptions.categoryId == cat.CategoryID)
                {
                    galleryCategory.SelectedItem = dropdownItem;
                }
            }
        }

        private void RedrawSyncStatus()
        {
            var pollingIntervalInMinutes = Globals.ThisAddIn.userOptions.pollingIntervalInMinutes;
            DateTime nextSyncTime = DateTime.Now.AddMinutes(pollingIntervalInMinutes);
            if (Globals.ThisAddIn.userOptions.lastSyncTime != DateTime.MinValue) // check for "null"
            {
                labelLastSyncTime.Label = LAST_SYNC_TIME + Globals.ThisAddIn.userOptions.lastSyncTime.ToString();
                labelLastSyncStatus.Label = LAST_SYNC_STATUS + (Globals.ThisAddIn.userOptions.lastSyncStatus == true ? "SUCCESS" : "FAILURE");
                nextSyncTime = Globals.ThisAddIn.userOptions.lastSyncTime.AddMinutes(pollingIntervalInMinutes);
                labelNextSyncTime.Label = NEXT_SYNC_TIME + nextSyncTime.ToString();
            }
            else
            {
                labelLastSyncTime.Label = LAST_SYNC_TIME + "n/a";
                labelLastSyncStatus.Label = LAST_SYNC_STATUS + "n/a";
                labelNextSyncTime.Label = NEXT_SYNC_TIME + nextSyncTime.ToString();
            }

            // store our next sync time value so sync timer has the update
            Globals.ThisAddIn.nextSyncTime = nextSyncTime;
        }

        private void InitializePollingIntervalGallery()
        {
            galleryPollingInterval.Items.Clear();

            // 1 hour, 2 hours, 4 hours, 8 hours
            var item1Hour = this.Factory.CreateRibbonDropDownItem();
            item1Hour.Label = "1 hour";
            item1Hour.Tag = 60;
            galleryPollingInterval.Items.Add(item1Hour);

            var item2Hours = this.Factory.CreateRibbonDropDownItem();
            item2Hours.Label = "2 hours";
            item2Hours.Tag = 120;
            galleryPollingInterval.Items.Add(item2Hours);

            var item4Hours = this.Factory.CreateRibbonDropDownItem();
            item4Hours.Label = "4 hours";
            item4Hours.Tag = 240;
            galleryPollingInterval.Items.Add(item4Hours);

            var item8Hours = this.Factory.CreateRibbonDropDownItem();
            item8Hours.Label = "8 hours";
            item8Hours.Tag = 480;
            galleryPollingInterval.Items.Add(item8Hours);

            switch (Globals.ThisAddIn.userOptions.pollingIntervalInMinutes)
            {
                case 60:
                    galleryPollingInterval.SelectedItem = item1Hour;
                    break;
                case 120:
                    galleryPollingInterval.SelectedItem = item2Hours;
                    break;
                case 240:
                    galleryPollingInterval.SelectedItem = item4Hours;
                    break;
                case 480:
                    galleryPollingInterval.SelectedItem = item8Hours;
                    break;
                default:
                    MessageBox.Show("Invalid value for pollingIntervalInMinutes retrieved.");
                    break;
            }
        }

        private void InitializeDaysToPullGallery()
        {
            galleryDaysToPull.Items.Clear();

            // 7 days, 14 days, 21 days, 28 days
            var item7Days = this.Factory.CreateRibbonDropDownItem();
            item7Days.Label = "7 days";
            item7Days.Tag = 7;
            galleryDaysToPull.Items.Add(item7Days);

            var item14Days = this.Factory.CreateRibbonDropDownItem();
            item14Days.Label = "14 days";
            item14Days.Tag = 14;
            galleryDaysToPull.Items.Add(item14Days);

            var item21Days = this.Factory.CreateRibbonDropDownItem();
            item21Days.Label = "21 days";
            item21Days.Tag = 21;
            galleryDaysToPull.Items.Add(item21Days);

            var item28Days = this.Factory.CreateRibbonDropDownItem();
            item28Days.Label = "28 days";
            item28Days.Tag = 28;
            galleryDaysToPull.Items.Add(item28Days);

            switch (Globals.ThisAddIn.userOptions.daysToPull)
            {
                case 7:
                    galleryDaysToPull.SelectedItem = item7Days;
                    break;
                case 14:
                    galleryDaysToPull.SelectedItem = item14Days;
                    break;
                case 21:
                    galleryDaysToPull.SelectedItem = item21Days;
                    break;
                case 28:
                    galleryDaysToPull.SelectedItem = item28Days;
                    break;
                default:
                    MessageBox.Show("Invalid value for daysToPull retrieved.");
                    break;
            }
        }

        private void InitializeReminderGallery()
        {
            galleryReminder.Items.Clear();

            // none, 0 minutes, 5 minutes, 15 minutes, 30 minutes, 1 hour
            var itemNone = this.Factory.CreateRibbonDropDownItem();
            itemNone.Label = "None";
            itemNone.Tag = -1;
            galleryReminder.Items.Add(itemNone);

            var item0Mins = this.Factory.CreateRibbonDropDownItem();
            item0Mins.Label = "0 minutes";
            item0Mins.Tag = 0;
            galleryReminder.Items.Add(item0Mins);

            var item5Mins = this.Factory.CreateRibbonDropDownItem();
            item5Mins.Label = "5 minutes";
            item5Mins.Tag = 5;
            galleryReminder.Items.Add(item5Mins);

            var item15Mins = this.Factory.CreateRibbonDropDownItem();
            item15Mins.Label = "15 minutes";
            item15Mins.Tag = 15;
            galleryReminder.Items.Add(item15Mins);

            var item30Mins = this.Factory.CreateRibbonDropDownItem();
            item30Mins.Label = "30 minutes";
            item30Mins.Tag = 30;
            galleryReminder.Items.Add(item30Mins);

            var item60Mins = this.Factory.CreateRibbonDropDownItem();
            item60Mins.Label = "1 hour";
            item60Mins.Tag = 60;
            galleryReminder.Items.Add(item60Mins);

            if (Globals.ThisAddIn.userOptions.reminderSet == false)
            {
                galleryReminder.SelectedItem = itemNone;
            }
            else
            {
                switch (Globals.ThisAddIn.userOptions.reminderMinutesBeforeStart)
                {
                    case 0:
                        galleryReminder.SelectedItem = item0Mins;
                        break;
                    case 5:
                        galleryReminder.SelectedItem = item5Mins;
                        break;
                    case 15:
                        galleryReminder.SelectedItem = item15Mins;
                        break;
                    case 30:
                        galleryReminder.SelectedItem = item30Mins;
                        break;
                    case 60:
                        galleryReminder.SelectedItem = item60Mins;
                        break;
                    default:
                        MessageBox.Show("Invalid value for reminderMinutesBeforeStart retrieved.");
                        break;
                }

            }
        }

        private void InitializeAvailStatusGallery()
        {
            galleryAvailStatus.Items.Clear();

            // Free, Working Elsewhere, Tentative, Busy, Out of Office
            var itemFree = this.Factory.CreateRibbonDropDownItem();
            itemFree.Label = "Free";
            itemFree.Tag = Outlook.OlBusyStatus.olFree;
            galleryAvailStatus.Items.Add(itemFree);

            var itemWorkingElsewhere = this.Factory.CreateRibbonDropDownItem();
            itemWorkingElsewhere.Label = "Working Elsewhere";
            itemWorkingElsewhere.Tag = Outlook.OlBusyStatus.olWorkingElsewhere;
            galleryAvailStatus.Items.Add(itemWorkingElsewhere);

            var itemTentative = this.Factory.CreateRibbonDropDownItem();
            itemTentative.Label = "Tentative";
            itemTentative.Tag = Outlook.OlBusyStatus.olTentative;
            galleryAvailStatus.Items.Add(itemTentative);

            var itemBusy = this.Factory.CreateRibbonDropDownItem();
            itemBusy.Label = "Busy";
            itemBusy.Tag = Outlook.OlBusyStatus.olBusy;
            galleryAvailStatus.Items.Add(itemBusy);

            var itemOOF = this.Factory.CreateRibbonDropDownItem();
            itemOOF.Label = "Out of Office";
            itemOOF.Tag = Outlook.OlBusyStatus.olOutOfOffice;
            galleryAvailStatus.Items.Add(itemOOF);

            switch (Globals.ThisAddIn.userOptions.availStatus)
            {
                case Outlook.OlBusyStatus.olFree:
                    galleryAvailStatus.SelectedItem = itemFree;
                    break;
                case Outlook.OlBusyStatus.olWorkingElsewhere:
                    galleryAvailStatus.SelectedItem = itemWorkingElsewhere;
                    break;
                case Outlook.OlBusyStatus.olTentative:
                    galleryAvailStatus.SelectedItem = itemTentative;
                    break;
                case Outlook.OlBusyStatus.olBusy:
                    galleryAvailStatus.SelectedItem = itemBusy;
                    break;
                case Outlook.OlBusyStatus.olOutOfOffice:
                    galleryAvailStatus.SelectedItem = itemOOF;
                    break;
            }
        }

        private void galleryReminder_Click(object sender, RibbonControlEventArgs e)
        {
            var gallery = sender as RibbonGallery;
            int newReminderMinutesBeforeStart = (int)gallery.SelectedItem.Tag;
            switch (newReminderMinutesBeforeStart)
            {
                case -1:
                    Globals.ThisAddIn.userOptions.reminderSet = false;
                    Globals.ThisAddIn.userOptions.reminderMinutesBeforeStart = -1;
                    break;
                default:
                    Globals.ThisAddIn.userOptions.reminderSet = true;
                    Globals.ThisAddIn.userOptions.reminderMinutesBeforeStart = newReminderMinutesBeforeStart;
                    break;
            }

            Globals.ThisAddIn.userOptions.Save();
        }

        private void galleryAvailStatus_Click(object sender, RibbonControlEventArgs e)
        {
            var gallery = sender as RibbonGallery;
            var newAvailStatus = (Outlook.OlBusyStatus)gallery.SelectedItem.Tag;
            Globals.ThisAddIn.userOptions.availStatus = newAvailStatus;

            Globals.ThisAddIn.userOptions.Save();
        }

        private void btnSubject_Click(object sender, RibbonControlEventArgs e)
        {
            var result = InputBox("Subject", ref Globals.ThisAddIn.userOptions.meetingPrefix);

            if (result == DialogResult.OK)
            {
                Globals.ThisAddIn.userOptions.Save();
            }
        }

        private void galleryDaysToPull_Click(object sender, RibbonControlEventArgs e)
        {
            var gallery = sender as RibbonGallery;
            int newDaysToPull = (int)gallery.SelectedItem.Tag;

            Globals.ThisAddIn.userOptions.daysToPull = newDaysToPull;

            Globals.ThisAddIn.userOptions.Save();
        }

        private void galleryPollingInterval_Click(object sender, RibbonControlEventArgs e)
        {
            var gallery = sender as RibbonGallery;
            int newPollingInterval = (int)gallery.SelectedItem.Tag;

            Globals.ThisAddIn.userOptions.pollingIntervalInMinutes = newPollingInterval;

            // update our next sync time as well
            RedrawSyncStatus();

            Globals.ThisAddIn.userOptions.Save();
        }

        private void btnSyncNow_Click(object sender, RibbonControlEventArgs e)
        {
            AttemptBackgroundSync();
        }

        private void btnSegmentName_Click(object sender, RibbonControlEventArgs e)
        {
            PromptUserForSegmentFilter();
        }

        private DialogResult PromptUserForSegmentFilter()
        {
            var result = InputBox(String.Format("Segment Filter ({0})", Globals.ThisAddIn.userOptions.syncMode), ref Globals.ThisAddIn.userOptions.segmentFilter);

            if (result == DialogResult.OK)
            {
                Globals.ThisAddIn.userOptions.Save();
            }

            return result;
        }

        /// <summary>
        /// If new string is valid, overwrites the previously stored string.
        /// </summary>
        /// <param name="formTitle"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        private DialogResult InputBox(string formTitle, ref string text)
        {
            System.Drawing.Size size = new System.Drawing.Size(200, 70);
            Form inputBox = new Form();
            inputBox.StartPosition = FormStartPosition.CenterParent;
            inputBox.ControlBox = false;

            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = formTitle;

            System.Windows.Forms.TextBox textBox = new TextBox();
            textBox.Size = new System.Drawing.Size(size.Width - 10, 23);
            textBox.Location = new System.Drawing.Point(5, 5);
            textBox.Text = text;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 80 - 95, 39);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 95, 39);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            DialogResult result = inputBox.ShowDialog();
            string newText = textBox.Text.Trim();

            if (result == DialogResult.OK)
            {
                if (newText.Length >= 512)
                {
                    MessageBox.Show("The string is too long. Reverting to last stored value.");
                    return DialogResult.Cancel;
                }
                else
                {
                    text = newText;
                }
            }

            return result;
        }

        private void btnSyncLog_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(Log.filePath);
            }
            catch (Exception exc)
            {
                MessageBox.Show(String.Format("Unable to open {0}.\r\n{1}", Log.filePath, exc.ToString()), "", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void AttemptBackgroundSync()
        {
            /*
            if (Globals.ThisAddIn.userOptions.segmentFilter == string.Empty)
            {
                MessageBox.Show("Before WFM for Outlook can sync, we need the name of your CritWatch segment. Please visit http://wfm and locate yours - it'll read like 'CritWatch' or 'GE-ECS XADM T3 SR'.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult result = PromptUserForCWSegmentName();
                if (result != DialogResult.OK)
                {
                    return;
                }
            }
            */

            if (syncBackgroundWorker.IsBusy)
            {
                MessageBox.Show("A sync cannot be started while one is already running.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // initiate the sync
                try
                {
                    syncBackgroundWorker.RunWorkerAsync();
                }
                catch (InvalidOperationException exc)
                {
                    Log.WriteEntry("Failed to start a sync.\r\n" + exc.ToString());
                }
            }
        }

        private void syncBackgroundWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;

                bool success = Globals.ThisAddIn.SyncNow();

                var dtNow = DateTime.Now;

                // store our last sync time and status in Exchange
                Globals.ThisAddIn.userOptions.lastSyncTime = dtNow;
                Globals.ThisAddIn.userOptions.lastSyncStatus = success;
                Globals.ThisAddIn.userOptions.Save();
            }
            catch (Exception exc)
            {
                Log.WriteEntry("Background worker thread crashed.\r\n" + exc.ToString());
            }
        }

        /// <summary>
        /// Fires every minute checking if we need to sync again.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void syncTimer_Tick(object sender, EventArgs e)
        {
            // check to see if we're supposed to be syncing
            if (Globals.ThisAddIn.nextSyncTime <= DateTime.Now)
            {
                // make sure we're not already running
                if (!syncBackgroundWorker.IsBusy)
                {
                    AttemptBackgroundSync();
                }
            }
        }

        private void galleryCategory_Click(object sender, RibbonControlEventArgs e)
        {
            var gallery = sender as RibbonGallery;

            Globals.ThisAddIn.userOptions.categoryId = gallery.SelectedItem.Tag as string;
            Globals.ThisAddIn.userOptions.categoryName = galleryCategory.SelectedItem.Label as string;

            Globals.ThisAddIn.userOptions.Save();
        }

        private void btnResetSettings_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult confirm = MessageBox.Show("Are you sure you wish to reset to default settings? This will delete all upcoming WFM meetings on your calendar.", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm == DialogResult.Yes)
            {
                // delete all future synced segments
                Globals.ThisAddIn.DeleteFutureMeetings();

                // reset
                Globals.ThisAddIn.userOptions = new Options();

                // save
                Globals.ThisAddIn.userOptions.Save();

                // redraw
                InitializeControls();
            }
        }
        
        private void gallerySyncMode_Click(object sender, RibbonControlEventArgs e)
        {
            var gallery = sender as RibbonGallery;
            SyncMode mode = (SyncMode)gallery.SelectedItem.Tag;

            Globals.ThisAddIn.userOptions.syncMode = mode;

            Globals.ThisAddIn.userOptions.Save();

            MessageBox.Show("As you've changed your Sync Mode, please be sure to update your Segment Filter as well.", "Sync Mode Changed", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            PromptUserForSegmentFilter();
        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start(@"https://www.github.com/maxxwizard/WFM-for-Outlook");
        }
    }
}

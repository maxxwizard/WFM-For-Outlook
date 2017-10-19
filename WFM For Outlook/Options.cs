using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO;
using System.ComponentModel;
using System.Reflection;

namespace WFM_For_Outlook
{
    public enum SyncMode
    {
        [Description("Inclusive")]
        Inclusive,

        [Description("Exclusive")]
        Exclusive
    }

    public class Options
    {
        public const string CONFIG_MESSAGE_SUBJECT = "WFM for Outlook";
        public const string DEFAULT_MEETING_PREFIX = "WFM: ";
        public const string DEFAULT_SEGMENT_FILTER = "Research,Shift,Meal";

        public bool reminderSet;
        public int reminderMinutesBeforeStart;
        public Outlook.OlBusyStatus availStatus;
        public string meetingPrefix;
        public int daysToPull;
        public int pollingIntervalInMinutes;
        public DateTime lastSyncTime;
        public bool lastSyncStatus;
        public string segmentFilter;
        public string employeeSK;
        public string categoryName;
        public SyncMode syncMode;

        /// <summary>
        /// Constructor with default values.
        /// </summary>
        public Options()
        {
            // initialize a blank new one with default values
            this.reminderSet = false;
            this.reminderMinutesBeforeStart = 0;
            this.availStatus = Outlook.OlBusyStatus.olFree;
            this.meetingPrefix = DEFAULT_MEETING_PREFIX;
            this.daysToPull = 28;
            this.pollingIntervalInMinutes = 480;
            this.segmentFilter = DEFAULT_SEGMENT_FILTER;
            this.syncMode = SyncMode.Exclusive;
            this.categoryName = CONFIG_MESSAGE_SUBJECT;
        }

        /// <summary>
        /// Persist this object to Exchange as a StorageItem.
        /// </summary>
        public void Save()
        {
            Outlook.MAPIFolder inboxFolder = Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.StorageItem configItem = inboxFolder.GetStorage(Options.CONFIG_MESSAGE_SUBJECT, Outlook.OlStorageIdentifierType.olIdentifyBySubject);
            configItem.Subject = Options.CONFIG_MESSAGE_SUBJECT;

            XmlSerializer x = new XmlSerializer(typeof(Options));
            using (StringWriter writer = new StringWriter())
            {
                // serialize this object into XML and store into the config item's body property
                x.Serialize(writer, this);
                configItem.Body = writer.ToString();

                // persist the item to Exchange
                configItem.Save();

                Log.WriteEntry("User options were saved to Exchange.");
            }
        }

        public static Options LoadFromConfigItem()
        {
            // grab the FAI message that houses user config
            Outlook.MAPIFolder inboxFolder = Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.StorageItem configItem = inboxFolder.GetStorage(Options.CONFIG_MESSAGE_SUBJECT, Outlook.OlStorageIdentifierType.olIdentifyBySubject);

            XmlSerializer x = new XmlSerializer(typeof(Options));
            using (StringWriter writer = new StringWriter())
            {
                // we found existing user config so deserialize the stored XML into an Options object for add-in to use
                if (configItem.EntryID != null && configItem.EntryID != "")
                {
                    Log.WriteEntry("User options were loaded from Exchange.");
                    // if we fail to deserialize, just return an empty one
                    Options opts;
                    try
                    {
                        opts = (Options)x.Deserialize(new StringReader(configItem.Body));
                    } catch
                    {
                        MessageBox.Show("There was an error reading your WFM for Outlook settings from Exchange so it has been reset. Please update your settings again accordingly.");
                        opts = new Options();
                    }
                    
                    return opts;
                }
            }

            return null;
        }
    }
}

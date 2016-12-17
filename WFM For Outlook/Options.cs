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

namespace WFM_For_Outlook
{
    public class Options
    {
        public const string CONFIG_MESSAGE_SUBJECT = "WFM for Outlook";
        public const string DEFAULT_CRITWATCH_SUBJECT = "CritWatch";

        public bool reminderSet;
        public int reminderMinutesBeforeStart;
        public Outlook.OlBusyStatus availStatus;
        public string critwatchSubject;
        public int daysToPull;
        public int pollingIntervalInMinutes;
        public DateTime lastSyncTime;
        public bool lastSyncStatus;
        public string segmentNameToMatch;
        public string employeeSK;
        public string categoryId;
        public string categoryName;

        /// <summary>
        /// Constructor with default values.
        /// </summary>
        public Options()
        {
            // initialize a blank new one with default values
            this.reminderSet = false;
            this.reminderMinutesBeforeStart = 0;
            this.availStatus = Outlook.OlBusyStatus.olTentative;
            this.critwatchSubject = DEFAULT_CRITWATCH_SUBJECT;
            this.daysToPull = 14;
            this.pollingIntervalInMinutes = 480;
            this.segmentNameToMatch = String.Empty;
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
                    return (Options)x.Deserialize(new StringReader(configItem.Body));
                }
            }

            return null;
        }
    }
}

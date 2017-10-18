using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Web;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using WFM_For_Outlook.WFM_API;
using Microsoft.Office.Tools.Ribbon;
using System.Threading;

namespace WFM_For_Outlook
{
    public partial class ThisAddIn
    {

        //Outlook.Inspectors inspectors;
        public Options userOptions;
        public DateTime nextSyncTime;

        public const string CUSTOM_MESSAGE_CLASS = "IPM.Appointment.WFM";
        public const string PROP_LAST_SYNC_TIME = "LastSyncTime";

        HttpClientHandler clientHandler;
        HttpClient client;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            /*
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
             */

            /* don't add anything here else Outlook will flag us as a slow-loading add-in */
        }

        private void InitializeHttpClient()
        {
            // http://haacked.com/archive/2004/05/15/http-web-request-expect-100-continue.aspx/
            System.Net.ServicePointManager.Expect100Continue = false;

            if (this.clientHandler == null)
            {
                clientHandler = new HttpClientHandler();
                clientHandler.UseDefaultCredentials = true;
                clientHandler.PreAuthenticate = true;
                clientHandler.ClientCertificateOptions = ClientCertificateOption.Automatic;

                client = new HttpClient(clientHandler);

                client.BaseAddress = new Uri("http://wfm");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("image/jpeg"));
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/x-ms-application"));
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("image/gif"));
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/xaml+xml"));
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("image/pjpeg"));
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/x-ms-xbap"));
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("*/*"));

                //client.DefaultRequestHeaders.Connection.Add("Keep-Alive");
                //client.DefaultRequestHeaders.CacheControl.NoCache = false;

            }
        }

        /// <summary>
        /// Queries WFM for employeeSK and saves it.
        /// </summary>
        /// <returns>True if success, false if not.</returns>
        public async Task<bool> QueryWfmForEmployeeSK()
        {
            Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(10);

            if (userOptions.employeeSK == null || userOptions.employeeSK == string.Empty)
            {
                HttpResponseMessage response = null;
                string targetUrl = string.Empty;
                string userXML = string.Empty;

                targetUrl = client.BaseAddress.ToString() + "/EAMWeb/WFMPRD/ENU/Common/servlet/AdminGetUserProfile.ewfm";
                Log.WriteEntry("HTTP GET from " + targetUrl);
                try
                {
                    //userXML = await client.GetStringAsync(targetUrl);
                    response = await client.GetAsync(targetUrl);
                }
                catch (Exception exc)
                {
                    Log.WriteEntry("Exception while querying WFM for employee SK.\r\n" + exc.ToString());
                    return false;
                }

                if (!response.IsSuccessStatusCode)
                {
                    Log.WriteEntry("Failed to query WFM for employee SK.\r\n" + response.ToString());
                    return false;
                }
                else
                {
                    userXML = await response.Content.ReadAsStringAsync();
                    Log.WriteEntry(response.ToString() + "\r\n" + userXML);
                    Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(30);
                    var xmlDoc = XDocument.Parse(userXML);

                    var employees = from c in xmlDoc.Root.Descendants("Employee")
                                    select c;

                    foreach (var e in employees)
                    {
                        userOptions.employeeSK = e.Attribute("SK").Value;
                    }

                    //userOptions.Save();
                }
            }

            Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(35);

            return true;
        }

        public async Task<string> QueryWfmForSegments(DateTime start, DateTime end)
        {
            HttpResponseMessage response = null;
            string targetUrl = string.Empty;
            string userXML = string.Empty;

            SegmentFilter filter = new SegmentFilter(userOptions.employeeSK, start, end);

            MediaTypeFormatter formatter = new SegmentFilterFormatter();

            targetUrl = client.BaseAddress.ToString() + "/EAMWeb/WFMPRD/ENU/ScheduleEditor/servlet/RequestScheduleView.ewfm";
            Log.WriteEntry("HTTP POST to " + targetUrl);
            Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(40);
            try
            {
                //client.CancelPendingRequests();
                //HttpContent hc;

                response = await client.PostAsync(targetUrl, filter, formatter);
            }
            catch (Exception exc)
            {
                Log.WriteEntry("Exception while querying WFM for employee schedule.\r\n" + exc.ToString());
                return string.Empty;
            }

            if (!response.IsSuccessStatusCode)
            {
                Log.WriteEntry("Failed to query WFM for employee schedule.\r\n" + response.ToString());
                return string.Empty;
            }
            else
            {
                //t.Wait();
                var scheduleXML = await response.Content.ReadAsStringAsync();

                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(50);

                Log.WriteEntry(response.ToString() + "\r\n" + scheduleXML);

                return scheduleXML;
            }
        }

        /// <summary>
        /// Immediately does a sync against WFM.
        /// </summary>
        /// <returns>True if sync was successful.</returns>
        public bool SyncNow()
        {
            // 1) retrieve XML data
            // 2) parse XML data for appropriate segments
            // 3) retrieve all meetings of class IPM.Appointment.WFM in the current sync range and delete them
            // 4) create the meetings again

            Log.WriteEntry(String.Format("Sync for {0} days initiated.", userOptions.daysToPull));

            try
            {
                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(5);

                InitializeHttpClient();

                bool success = QueryWfmForEmployeeSK().Result;
                if (!success)
                {
                    return false;
                }

                // we subtract one from DaysToPull because the current day counts
                string scheduleXml = QueryWfmForSegments(DateTime.Now, DateTime.Now.AddDays(userOptions.daysToPull-1)).Result;
                if (scheduleXml == string.Empty)
                {
                    return false;
                }

                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(55);

                InternalSyncBetter(scheduleXml);

                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(100);

            }
            catch (Exception e)
            {
                Log.WriteEntry("Internal sync issue.\r\n" + e.ToString());
                return false;
            }

            return true;
        }

        private void InternalSyncBetter(string scheduleXml)
        {
            // sync logic operates with the building block of a day. a day is actually defined by WFM's NominalDate field (start date is what counts).
            // for each day, we use this sync logic:
            /*
                grab segments by nominal date from WFM and calendar

                meetingsToBeCreated = 0

                for each WFM segment {
                   find the corresponding calendar meeting and mark it as processed
                   if not found
                      meetingsToBeCreated++
                }

                for (;meetingsToBeCreated > 0; meetingsToBeCreated--) {
                   if there is an unprocessed calendar meeting left
                      use it (by changing the start/end time) and set processed flag
                }

                for each calendar meeting that is still not processed {
                   delete the meeting
                }
             */

            WfmSchedule schedule = WfmSchedule.Parse(scheduleXml);

            string[] segmentNames = Globals.ThisAddIn.userOptions.segmentNameToMatch.ToLower().Split(new char[] { ';', ',' });
            var matchingSegments = schedule.GetMatchingSegments(segmentNames);

            Log.WriteEntry(String.Format("Found {0} segments from WFM with segment names: {1}", matchingSegments.Count, String.Join(", ", segmentNames)));

            if (Globals.ThisAddIn.userOptions.lastSyncTime == DateTime.MinValue && matchingSegments.Count == 0)
            {
                // this is the first run of the meeting and there were no segments found
                MessageBox.Show(String.Format("I noticed this is your first WFM for Outlook sync and we did not find any" +
                                                " segments titled '{0}' from WFM. Are you sure that" +
                                                " your CW segment name is set correctly?", userOptions.segmentNameToMatch),
                                                "First sync notification", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var syncTimeNow = DateTime.Now;

            Outlook.AppointmentItem unprocessedMeeting;

            int stats_totalSegmentsFromWfm = schedule.Segments.Count;
            int stats_meetingsCreated = 0;
            int stats_meetingsDeleted = 0;
            int stats_meetingsSynched = 0;
            int stats_meetingsUpdated = 0;

            int PercentPerSegment = (int)Math.Floor(25.0 / (double)userOptions.daysToPull);
            int progressPercent = 75;

            Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(progressPercent);

            DateTime current = DateTime.Now;
            for (int i = 0; i < userOptions.daysToPull; i++)
            {
                /*
                Outlook.Items cwMeetingsOnCalendar = GetCritWatchMeetingsNominalDate(current);
                IEnumerable<XElement> segmentsInWfm = from s in matchingSegments
                                                      where s.Element("NominalDate").Value == current.ToString("yyyy-MM-dd")
                                                      select s;
                int countSegmentsInWfm = matchingSegments.Count();

                List<XElement> meetingsToBeCreated = new List<XElement>();

                Log.WriteDebug("syncEngine : Starting initial sync loop, date " + current.ToShortDateString());
                foreach (var seg in segmentsInWfm)
                {
                    string subject = seg.Attribute("SK").Value;
                    DateTime start = DateTime.Parse(seg.Element("StartTime").Value);
                    DateTime end = DateTime.Parse(seg.Element("StopTime").Value);
                    Outlook.AppointmentItem foundItem = FindItemInMeetings(cwMeetingsOnCalendar, start, end);

                    if (foundItem == null)
                    {
                        Log.WriteDebug(String.Format("syncEngine : saving meeting {0} to be created/adjusted later", start));
                        meetingsToBeCreated.Add(seg);
                    }
                    else
                    {
                        var lastSyncTime = foundItem.UserProperties.Find(PROP_LAST_SYNC_TIME);
                        if (lastSyncTime == null)
                        {
                            lastSyncTime = foundItem.UserProperties.Add(PROP_LAST_SYNC_TIME, Outlook.OlUserPropertyType.olDateTime);
                        }
                        lastSyncTime.Value = syncTimeNow;
                        foundItem.Save();
                        stats_meetingsSynched++;
                        Log.WriteDebug(String.Format("syncEngine : meeting {0} found and synched", foundItem.Start.ToString()));
                    }
                }

                Log.WriteDebug(String.Format("syncEngine : meetingsToBeCreated.Count = {0}", meetingsToBeCreated.Count));

                Log.WriteDebug("syncEngine : Starting meeting creation/update loop");
                foreach (var seg in meetingsToBeCreated)
                {
                    DateTime start = DateTime.Parse(seg.Element("StartTime").Value);
                    DateTime end = DateTime.Parse(seg.Element("StopTime").Value);

                    // find the corresponding calendar meeting to use
                    unprocessedMeeting = FindUnprocessedMeetingToUse(syncTimeNow, cwMeetingsOnCalendar);

                    // if we didn't find one, we need to create a new meeting otherwise we just need to adjust it
                    if (unprocessedMeeting == null)
                    {
                        CreateCritWatchSegment(syncTimeNow, start, end);
                        stats_meetingsCreated++;
                    }
                    else
                    {
                        unprocessedMeeting.UserProperties.Find(PROP_LAST_SYNC_TIME).Value = syncTimeNow;
                        unprocessedMeeting.Start = start;
                        unprocessedMeeting.End = end;
                        unprocessedMeeting.Save();
                        stats_meetingsUpdated++;
                    }
                }

                Log.WriteDebug("syncEngine : Starting deletion loop");
                // https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.outlook._appointmentitem.delete(v=office.15).aspx
                // https://social.msdn.microsoft.com/Forums/office/en-US/53a07e5d-ab16-4930-90bf-52215f084d59/appointmentitemrecipientsremoveindex-question?forum=outlookdev
                // Seriously Outlook Object Model? So the index starts at 1 and we have to always decrement instead of increment.
                for (int j = cwMeetingsOnCalendar.Count; j >= 1; j--)
                {
                    
                    Outlook.AppointmentItem item = cwMeetingsOnCalendar[j];
                    var lastSyncTimeProp = item.UserProperties.Find(PROP_LAST_SYNC_TIME);

                    Log.WriteDebug(String.Format("syncEngine : deletion loop : (lastSyncTime={0}, syncTimeNow={1})", (lastSyncTimeProp != null ? lastSyncTimeProp.Value.ToString() : "null"), syncTimeNow.ToString()));

                    // if the LastSyncTime doesn't match syncTimeNow, we delete it
                    // we compare their ToStrings because for some unknown reason, the ticks are different from when we save them to the object and retrieve them
                    if (lastSyncTimeProp != null && !syncTimeNow.ToString().Equals(lastSyncTimeProp.Value.ToString()))
                    {
                        Log.WriteDebug("syncEngine : Deleting meeting " + item.Start.ToString());
                        item.Delete();
                        stats_meetingsDeleted++;
                    }
                }

                // move to next day
                current = current.AddDays(1);
                */

                // update progress
                progressPercent += PercentPerSegment;
                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(progressPercent);
            }

            // output stats to our sync log
            Log.WriteEntry(String.Format("Segments from WFM: {0}\r\nMeetings matched/synched: {1}\r\nMeetings updated: {2}\r\nMeetings created: {3}\r\nMeetings deleted: {4}",
                stats_totalSegmentsFromWfm, stats_meetingsSynched, stats_meetingsUpdated, stats_meetingsCreated, stats_meetingsDeleted));
        }

        /// <summary>
        /// Returns a meeting whose LastSyncTime isn't equal to the current sync attempt (<paramref name="syncTime"/>). Returns null if none available.
        /// </summary>
        /// <param name="syncTime"></param>
        /// <param name="meetings"></param>
        /// <returns></returns>
        private Outlook.AppointmentItem FindUnprocessedMeetingToUse(DateTime syncTime, Outlook.Items meetings)
        {
            Log.WriteEntry(String.Format("FindUnprocessedMeetingToUse() with LastSyncTime older than {0}", syncTime.ToString()));
            for (int j = meetings.Count; j >= 1; j--)
            {
                Outlook.AppointmentItem item = meetings[j];
                var lastSyncTime = item.UserProperties.Find(PROP_LAST_SYNC_TIME);

                if (lastSyncTime != null && !syncTime.ToString().Equals(lastSyncTime.Value.ToString()))
                {
                    Log.WriteEntry(String.Format("FindUnprocessedMeetingToUse() found meeting with LastSyncTime of {0}", lastSyncTime.Value.ToString()));
                    return item;
                }
            }

            return null;
        }

        /// <summary>
        /// The meeting if found, null if not.
        /// </summary>
        /// <param name="meetings"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        private Outlook.AppointmentItem FindItemInMeetings(Outlook.Items meetings, DateTime start, DateTime end)
        {
            
            foreach (Outlook.AppointmentItem m in meetings)
            {
                Log.WriteEntry(String.Format("syncEngine : comparing meeting {0} against param {1}", m.Start.ToString(), start.ToString()));
                if (m.Start.Equals(start) && m.End.Equals(end))
                {
                    return m;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets our custom created meetings from the user's calendar.
        /// </summary>
        /// <param name="startDate">Meeting start time must fall on this date.</param>
        /// <returns></returns>
        private Outlook.Items GetCritWatchMeetingsNominalDate(DateTime startDate)
        {
            Outlook.MAPIFolder calendar = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = calendar.Items;

            //DateTime nominalDate = new DateTime(startDate.Year, startDate.Month, startDate.Day);

            // filter to just our custom items in our date range
            string filter = String.Format("[Start] >= '{0}' and [Start] < '{1}' and [MessageClass] = '{2}'",
                startDate.ToShortDateString(), startDate.AddDays(1).ToShortDateString(), CUSTOM_MESSAGE_CLASS);
            items = items.Restrict(filter);
            Log.WriteEntry(String.Format("GetCritWatchMeetingsNominalDate() : {0} : count = {1}", filter, items.Count));

            return items;
        }

        private void DeleteFutureMeetings()
        {
            Log.WriteDebug("Deleting all future meetings");

            Outlook.MAPIFolder calendar = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = calendar.Items;

            DateTime startDate = DateTime.Now;

            string filter = String.Format("[Start] >= '{0}' and [MessageClass] = '{2}'",
                startDate.ToShortDateString(), CUSTOM_MESSAGE_CLASS);
            items = items.Restrict(filter);

            foreach (Outlook.AppointmentItem item in items)
            {
                item.Delete();
            }

            Log.WriteDebug("Deleted all future meetings");
        }

        /// <summary>
        /// Finds all meetings created by our add-in inside Deleted Items and deletes them.
        /// </summary>
        private void PurgeCritWatchMeetingsFromDeletedItems()
        {
            Outlook.MAPIFolder deletedItems = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
            Outlook.Items itemsToPurge = deletedItems.Items;
            itemsToPurge = itemsToPurge.Restrict(String.Format("[MessageClass] = '{0}'", CUSTOM_MESSAGE_CLASS));

            Outlook.AppointmentItem item = itemsToPurge.GetLast();
            while (item != null)
            {
                item.Delete();
                item = itemsToPurge.GetLast();
            }
        }

        /// <summary>
        /// Get the CritWatch meetings in a date range.
        /// </summary>
        /// <returns></returns>
        private Outlook.Items GetCritWatchMeetings(DateTime startDate, DateTime endDate)
        {
            Outlook.MAPIFolder calendar = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = calendar.Items;

            // filter to just our custom items in our date range
            string filter = String.Format("[Start] > '{0}' and [End] < '{1}' and [MessageClass] = '{2}'",
                startDate.ToShortDateString(), endDate.ToShortDateString(), CUSTOM_MESSAGE_CLASS);
            items = items.Restrict(filter);

            return items;
        }

        private void CreateCritWatchSegment(DateTime lastSyncTime, DateTime startTime, DateTime endTime)
        {
            Outlook.AppointmentItem newMeeting = Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
            if (newMeeting != null)
            {
                newMeeting.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olNonMeeting;
                //newMeeting.Location = "Conference Room";
                newMeeting.Subject = userOptions.critwatchSubject;
                newMeeting.Body = "Created by the WFM for Outlook add-in";
                newMeeting.Start = startTime;
                newMeeting.End = endTime;
                newMeeting.Categories = userOptions.categoryName;
                newMeeting.BusyStatus = userOptions.availStatus;
                newMeeting.ReminderSet = userOptions.reminderSet;
                if (userOptions.reminderSet == true)
                {
                    newMeeting.ReminderMinutesBeforeStart = userOptions.reminderMinutesBeforeStart;
                }
                newMeeting.MessageClass = CUSTOM_MESSAGE_CLASS;
                var lastSyncTimeProp = newMeeting.UserProperties.Add(PROP_LAST_SYNC_TIME, Outlook.OlUserPropertyType.olDateTime);
                lastSyncTimeProp.Value = lastSyncTime;
                newMeeting.Save();
                //MessageBox.Show("CritWatch segment for " + startTime.ToString() + " created");
            }

            Log.WriteEntry(String.Format("CritWatch meeting created (Start={0}, End={1}, LastSyncTime={2})", startTime.ToString(), endTime.ToString(), lastSyncTime.ToString()));
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                // check that the mail item is new
                if (mailItem.EntryID == null)
                {
                    /*
                    mailItem.Subject = "This text was added using code";
                    mailItem.Body = "This text was added using code";
                     */
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

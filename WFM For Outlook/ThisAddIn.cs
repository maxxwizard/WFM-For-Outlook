using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using WFM_For_Outlook.WFM_API;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace WFM_For_Outlook
{
    public partial class ThisAddIn
    {

        public Options userOptions;
        public DateTime nextSyncTime;

        public const string CUSTOM_MESSAGE_CLASS = "IPM.Appointment.WFM";
        public const string PROP_LAST_SYNC_TIME = "LastSyncTime";

        HttpClientHandler clientHandler;
        HttpClient client;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            /* don't add anything here else Outlook will flag us as a slow-loading add-in */
        }

        private void InitializeHttpClient()
        {
            // http://haacked.com/archive/2004/05/15/http-web-request-expect-100-continue.aspx/
            System.Net.ServicePointManager.Expect100Continue = false;

            if (clientHandler == null)
            {
                clientHandler = new HttpClientHandler();
                clientHandler.UseDefaultCredentials = true;
                clientHandler.PreAuthenticate = true;
                clientHandler.ClientCertificateOptions = ClientCertificateOption.Automatic;

                client = new HttpClient(clientHandler);

                client.BaseAddress = new Uri(Globals.ThisAddIn.userOptions.wfmUrl);
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
                catch (System.Exception exc)
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
            catch (System.Exception exc)
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
        /// <remarks>
        /// 1) retrieve XML data
        /// 2) parse XML data for appropriate segments
        /// 3) delete all future synced meetings
        /// 4) create the meetings based on WFM schedule
        /// </remarks>
        /// <returns>True if sync was successful.</returns>
        public bool SyncNow()
        {
            var dict = new Dictionary<string, string>();
            dict.Add("DaysToPull", userOptions.daysToPull.ToString());
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
                string scheduleXml = QueryWfmForSegments(DateTime.Now, DateTime.Now.AddDays(userOptions.daysToPull - 1)).Result;
                if (scheduleXml == string.Empty)
                {
                    return false;
                }

                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(55);

                InternalSync(scheduleXml);

                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(100);

            }
            catch (System.Exception e)
            {
                Log.WriteEntry("Internal sync issue.\r\n" + e.ToString());
                return false;
            }

            return true;
        }

        /// <summary>
        /// Syncs WFM schedule XML to Outlook calendar
        /// </summary>
        /// <remarks>
        /// Sync logic is very basic. We delete all future meetings and create new ones based on WFM schedule.
        /// </remarks>
        /// <param name="scheduleXml">Full XML response from WFM</param>
        private void InternalSync(string scheduleXml)
        {
            WfmSchedule schedule = WfmSchedule.Parse(scheduleXml);

            string[] segmentFilter = Globals.ThisAddIn.userOptions.segmentFilter.ToLower().Split(new char[] { ';', ',' }).Select(s => s.Trim()).ToArray();
            var matchingSegments = schedule.GetMatchingSegments(segmentFilter);

            Log.WriteEntry(String.Format("Found {0} segments from WFM with {1} segment names: {2}", matchingSegments.Count, userOptions.syncMode, String.Join(", ", segmentFilter)));

            if (Globals.ThisAddIn.userOptions.lastSyncTime == DateTime.MinValue && matchingSegments.Count == 0)
            {
                // this is the first run of the meeting and there were no segments found
                MessageBox.Show(String.Format("I noticed this is your first WFM for Outlook sync and we did not find any segments from WFM. Please ensure that your segment filter is configured correctly.", userOptions.segmentFilter),
                                                "First sync notification", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int stats_meetingsDeleted = DeleteFutureMeetings();

            var syncTimeNow = DateTime.Now;

            int stats_totalSegmentsFromWfm = schedule.Segments.Count;
            int stats_meetingsCreated = 0;

            int PercentPerSegment = (int)Math.Floor(40.0 / (double)stats_totalSegmentsFromWfm);
            int progressPercent = 60;

            Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(progressPercent);

            Log.WriteDebug("SyncEngine : Starting meeting creation loop");
            foreach (var seg in matchingSegments)
            {
                CreateMeetingOnCalendar(syncTimeNow, seg);
                stats_meetingsCreated++;

                Log.WriteDebug(String.Format("SyncEngine : meeting created \r\n{0}", seg.ToString()));

                // update progress
                progressPercent += PercentPerSegment;
                Globals.Ribbons.CalendarIntegrationRibbon.syncBackgroundWorker.ReportProgress(progressPercent);
            }

            // output stats to our sync log
            Log.WriteEntry(String.Format("Segments from WFM: {0}\r\nMeetings created: {1}\r\nMeetings deleted: {2}",
                stats_totalSegmentsFromWfm, stats_meetingsCreated, stats_meetingsDeleted));
        }

        /// <summary>
        /// Gets our custom created meetings from the user's calendar.
        /// </summary>
        /// <param name="startDate">Meeting start time must fall on this date.</param>
        /// <returns></returns>
        private Outlook.Items GetMeetingsNominalDate(DateTime startDate)
        {
            Outlook.MAPIFolder calendar = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = calendar.Items;

            //DateTime nominalDate = new DateTime(startDate.Year, startDate.Month, startDate.Day);

            // filter to just our custom items in our date range
            string filter = String.Format("[Start] >= '{0}' and [Start] < '{1}' and [MessageClass] = '{2}'",
                startDate.ToShortDateString(), startDate.AddDays(1).ToShortDateString(), CUSTOM_MESSAGE_CLASS);
            items = items.Restrict(filter);
            Log.WriteEntry(String.Format("GetCritWatchMeetingsNominalDate() : {0} : count = {1}", filter, items.Count));

            return items;
        }

        /// <summary>
        /// Hard-deletes all meetings newer than today that the add-in created.
        /// </summary>
        public int DeleteFutureMeetings()
        {
            Log.WriteDebug("Deleting all future meetings");

            Outlook.MAPIFolder calendar = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = calendar.Items;

            // start deleting from 11:59pm the previous night to capture the all-day events
            DateTime startDate = DateTime.Now.Date.AddMinutes(-1);

            string filter = String.Format("[Start] >= '{0}' and [MessageClass] = '{1}'",
                String.Format("{0} {1}", startDate.ToShortDateString(), startDate.ToShortTimeString()), CUSTOM_MESSAGE_CLASS);
            items = items.Restrict(filter);

            // https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.outlook._appointmentitem.delete(v=office.15).aspx
            // https://social.msdn.microsoft.com/Forums/office/en-US/53a07e5d-ab16-4930-90bf-52215f084d59/appointmentitemrecipientsremoveindex-question?forum=outlookdev
            // Seriously Outlook Object Model? So the index starts at 1 and we have to always decrement instead of increment.
            for (int j = items.Count; j >= 1; j--)
            {
                AppointmentItem item = items[j] as AppointmentItem;
                item.Delete();
            }

            PurgeMeetingsFromDeletedItems();

            Log.WriteEntry("Deleted all future meetings");

            return items.Count;
        }

        /// <summary>
        /// Finds all meetings created by our add-in inside Deleted Items and deletes them.
        /// </summary>
        private void PurgeMeetingsFromDeletedItems()
        {
            Outlook.MAPIFolder deletedItems = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
            Outlook.Items itemsToPurge = deletedItems.Items;
            itemsToPurge = itemsToPurge.Restrict(String.Format("[MessageClass] = '{0}'", CUSTOM_MESSAGE_CLASS));

            for (int j = itemsToPurge.Count; j >= 1; j--)
            {
                AppointmentItem item = itemsToPurge[j] as AppointmentItem;
                item.Delete();
            }
        }

        /// <summary>
        /// Get the CritWatch meetings in a date range.
        /// </summary>
        /// <returns></returns>
        private Outlook.Items GetCritWatchMeetings(DateTime startDate, DateTime endDate)
        {
            Outlook.MAPIFolder calendar = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items items = calendar.Items;

            // filter to just our custom items in our date range
            string filter = String.Format("[Start] > '{0}' and [End] < '{1}' and [MessageClass] = '{2}'",
                startDate.ToShortDateString(), endDate.ToShortDateString(), CUSTOM_MESSAGE_CLASS);
            items = items.Restrict(filter);

            return items;
        }

        private void CreateMeetingOnCalendar(DateTime lastSyncTime, Segment segment)
        {
            Outlook.AppointmentItem newMeeting = Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
            if (newMeeting != null)
            {
                newMeeting.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olNonMeeting;

                newMeeting.Subject = userOptions.meetingPrefix + segment.Name;

                newMeeting.Body = "Created by the WFM for Outlook add-in";
                if (!String.IsNullOrEmpty(segment.Memo))
                {
                    newMeeting.Body = segment.Memo + "\r\n\r\n" + newMeeting.Body;
                }

                if (segment.IsAllDay)
                {
                    newMeeting.AllDayEvent = true;
                    newMeeting.Start = segment.NominalDate;
                    // https://msdn.microsoft.com/en-us/library/office/ff184629.aspx states the end date needs to be midnight of the next day
                    newMeeting.End = segment.NominalDate.AddDays(1);
                }
                else
                {
                    newMeeting.Start = segment.StartTime;
                    newMeeting.End = segment.EndTime;
                }

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
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

using Microsoft.ApplicationInsights;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WFM_For_Outlook
{
    class Log
    {
        public static string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"WFM For Outlook");
        public static string filename = "sync.log";
        public static string filePath = Path.Combine(folderPath, filename);
        public static StreamWriter writer;
        private static TelemetryClient tc;

        public static TelemetryClient TelemetryClient
        {
            get
            {
                if (tc == null)
                {
                    tc = new TelemetryClient();

                    tc.Context.User.Id = Environment.UserName;
                    tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString();
                    tc.Context.Device.Id = Environment.MachineName;
                    tc.Context.Session.Id = Guid.NewGuid().ToString();
                }
                return tc;
            }
        }

        public static void TrackEvent(string eventName, IDictionary<string, string> properties = null)
        {
            TelemetryClient.TrackEvent(eventName, properties);
        }

        public static void WriteEntry(string logMessage)
        {
            try
            {
                if (writer == null)
                {
                    // every time the add-in starts, we're going to create a blank new log
                    Directory.CreateDirectory(folderPath);
                    writer = File.CreateText(filePath);
                }

                writer.WriteLine("{0} {1}\r\n{2}\r\n", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString(), logMessage);
                writer.Flush();
                TelemetryClient.TrackTrace(logMessage);
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Format("WFM for Outlook failed to write to {0}.\r\n{1}", filePath, e.ToString()));
                TelemetryClient.TrackException(e);
            }
        }

        public static void WriteDebug(string logMessage)
        {
#if (DEBUG)
            WriteEntry(logMessage);
#endif
        }
    }
}

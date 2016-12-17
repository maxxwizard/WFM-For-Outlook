using System;
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
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Format("WFM for Outlook failed to write to {0}.\r\n{1}", filePath, e.ToString()));
            }
        }

        public static void WriteDebug(string logMessage)
        {
#if (DEBUG)
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
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Format("WFM for Outlook failed to write to {0}.\r\n{1}", filePath, e.ToString()));
            }
#endif
        }
    }
}

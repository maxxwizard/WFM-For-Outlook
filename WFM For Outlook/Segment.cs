using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WFM_For_Outlook
{
    class Segment
    {
        public DateTime NominalDate
        {
            get; set;
        }

        public string Code
        {
            get; set;
        }

        public string Name
        {
            get; set;
        }

        public string Memo
        {
            get; set;
        }

        public bool IsAllDay
        {
            get; set;
        }

        public DateTime StartTime
        {
            get; set;
        }

        public DateTime EndTime
        {
            get; set;
        }

        public override string ToString()
        {
            string s = String.Format("{0}\r\n{1}", this.Name, this.NominalDate.ToShortDateString());

            if (this.IsAllDay)
            {
                return s;
            }
            else
            {
                s += String.Format("\r\n{0}\r\n{1}", this.StartTime.ToLocalTime().ToString(), this.EndTime.ToLocalTime().ToString());
            }

            return s;
        }
    }
}

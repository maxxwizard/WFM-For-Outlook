using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WFM_For_Outlook.WFM_API
{
    public class DateRange
    {
        public string Start;
        public string Stop;

        public DateRange()
        {

        }

        public DateRange(DateTime start, DateTime stop)
        {
            this.Start = start.ToString("yyyy-MM-dd");
            this.Stop = stop.ToString("yyyy-MM-dd");
        }
    }
}

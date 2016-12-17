using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace WFM_For_Outlook.WFM_API
{
    public class SegmentFilter
    {
        [XmlElement("EmployeeSelector")]
        public EmployeeSelector empSelector;

        [XmlElement("DateRange")]
        public DateRange dateRange;

        public SegmentFilter()
        {
        }

        public SegmentFilter(string employeeSK, DateTime start, DateTime stop)
        {
            this.empSelector = new EmployeeSelector(employeeSK);
            this.dateRange = new DateRange(start, stop);
        }
    }
}

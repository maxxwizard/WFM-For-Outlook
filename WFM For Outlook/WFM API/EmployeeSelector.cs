using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace WFM_For_Outlook.WFM_API
{
    public class EmployeeSelector
    {
        [XmlElement("Employee")]
        public Employee e;

        public EmployeeSelector()
        {
        }

        public EmployeeSelector(string SK)
        {
            this.e = new Employee(SK);
        }
    }
}

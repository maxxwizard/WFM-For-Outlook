using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace WFM_For_Outlook.WFM_API
{
    public class Employee
    {
        [XmlAttribute("SK")]
        public string SK;

        public Employee()
        {
        }

        public Employee(string SK)
        {
            this.SK = SK;
        }
    }
}

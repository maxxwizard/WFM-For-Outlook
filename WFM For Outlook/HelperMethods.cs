using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WFM_For_Outlook
{
    static class HelperMethods
    {
        public static string GetDescription(this SyncMode mode)
        {
            FieldInfo fi = mode.GetType().GetField(mode.ToString());
            var desc = fi.GetCustomAttributes(typeof(DescriptionAttribute), false) as DescriptionAttribute[];
            if (null != desc && desc.Length > 0)
            {
                return desc[0].Description;
            }
            else
            {
                return mode.ToString();
            }
        }
    }
}

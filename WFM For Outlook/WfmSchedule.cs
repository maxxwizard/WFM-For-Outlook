using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WFM_For_Outlook
{
    class WfmSchedule
    {
        public Dictionary<string, string> SegmentCodeToName
        {
            get; set;
        } = new Dictionary<string, string>();

        public List<Segment> Segments
        {
            get; set;
        } = new List<Segment>();

        public List<Segment> GetMatchingSegments(string[] segmentNames)
        {
            List<Segment> subset = new List<Segment>();

            var matches = this.Segments.Where(s =>
            {
                if (Globals.ThisAddIn.userOptions.syncMode == SyncMode.Inclusive)
                {
                    return segmentNames.Contains(s.Name.ToLower());
                }
                
                return !segmentNames.Contains(s.Name.ToLower());
            });
            subset.AddRange(matches);

            return subset;
        }       

        public static WfmSchedule Parse(string xml)
        {
            WfmSchedule schedule = new WfmSchedule();

            //var xmlDoc = XDocument.Load(@"C:\sources\WFM-For-Outlook\WFM For Outlook\eeSchedule.xml");
            var xmlDoc = XDocument.Parse(xml);

            // extract segment code-to-name dictionary
            var segmentCodes = from c in xmlDoc.Root.Descendants("SegmentCodes").Descendants("SegmentCode")
                               select c;
            foreach (var segmentCode in segmentCodes)
            {
                var code = segmentCode.Attribute("SK").Value;
                var name = segmentCode.Element("Code").Value;
                schedule.SegmentCodeToName.Add(code, name);
            }

            // extract segments
            var segments = from s in xmlDoc.Root.Descendants("Segments").Descendants()
                           where s.Name == "DetailSegment" || s.Name == "GeneralSegment"
                           select s;
            foreach (var segment in segments)
            {
                string code = segment.Element("SegmentCode").Attribute("SK").Value;
                string name; schedule.SegmentCodeToName.TryGetValue(code, out name);
                Segment s = new Segment() {
                    Code = code,
                    Name = name,
                    Memo = segment.Element("Memo").Value,
                    IsAllDay = segment.Name.ToString().Equals("GeneralSegment", StringComparison.InvariantCultureIgnoreCase),
                    StartTime = segment.Element("StartTime") == null ? DateTime.MinValue : DateTime.Parse(segment.Element("StartTime").Value),
                    EndTime = segment.Element("StopTime") == null ? DateTime.MinValue : DateTime.Parse(segment.Element("StopTime").Value),
                    NominalDate = DateTime.Parse(segment.Element("NominalDate").Value),
                };
                schedule.Segments.Add(s);
            }            

            return schedule;
        }
    }
}

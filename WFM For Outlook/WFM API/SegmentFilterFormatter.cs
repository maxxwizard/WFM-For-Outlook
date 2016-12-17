using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Xml;
using System.Xml.Serialization;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.IO;

namespace WFM_For_Outlook.WFM_API
{
    class SegmentFilterFormatter : BufferedMediaTypeFormatter
    {
        public SegmentFilterFormatter()
        {
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/x-www-form-urlencoded"));
        }

        public override bool CanWriteType(Type type)
        {
            if (type == typeof(SegmentFilter))
            {
                return true;
            }

            return false;
        }

        public override bool CanReadType(Type type)
        {
            return false;
        }

        public override void WriteToStream(Type type, object value, System.IO.Stream writeStream, HttpContent content)
        {
            using (var writer = new StreamWriter(writeStream))
            {
                SegmentFilter segFilter = value as SegmentFilter;
                if (segFilter != null)
                {
                    string postRequestUrl = BuildPostRequestUrl(segFilter);
                    try
                    {
                        //content.Headers.ContentLength = postRequestUrl.Length;
                        writer.WriteLine(postRequestUrl);
                    }
                    catch (Exception e)
                    {
                        Log.WriteEntry("SegmentFilterFormatter::WriteToStream() exception.\r\n" + e.ToString());
                    }
                }
            }
        }

        static string BuildPostRequestUrl(SegmentFilter filter)
        {
            var xns = new XmlSerializerNamespaces();
            xns.Add(string.Empty, string.Empty);

            StringBuilder sb = new StringBuilder();

            var xmlSettings = new XmlWriterSettings();
            xmlSettings.OmitXmlDeclaration = true;

            string postRequest = "Stylesheet=../../ScheduleEditor/Styles/ScheduleLoadData.xsl&data_in=";

            XmlSerializer x = new XmlSerializer(typeof(SegmentFilter));
            using (XmlWriter writer = XmlWriter.Create(sb, xmlSettings))
            {
                x.Serialize(writer, filter, xns);
                postRequest += WebUtility.UrlEncode(sb.ToString());
            }

            return postRequest;
        }
    }
}

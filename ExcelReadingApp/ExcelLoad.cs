using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelReadingApp
{
    class ExcelLoad
    {
        XmlDocument doc = new XmlDocument();

        public void xmlLoadData()
        {
            doc.LoadXml(@"F:\XMLFilesDemo\_MetersDoneInPhillippines\ZonaLibre_20050908.xml");
            foreach(XmlNode node in doc.DocumentElement)
            {
                string name = node.Attributes[0].Name;
            }
        }
    }
}

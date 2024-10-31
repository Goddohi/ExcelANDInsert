using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace excelChange.Entity
{
    [XmlRoot("TableEntity")]

    public class TableEntity
    {
        [XmlElement("Name")]
        public string Name { get; set; }
    }
}

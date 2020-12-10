using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace Test {
    class Program {
        static void Main(string[] args) {
			StreamWriter sw = new StreamWriter(Environment.CurrentDirectory+"test.xml");
			sw.AutoFlush = true;
			ProductType yo = new ProductType("Buns", FilingPriority.High, "DECBUNS", true);
			XmlSerializer xser = new XmlSerializer(typeof(ProductType), new XmlRootAttribute("xml"));
			xser.Serialize(sw, yo);
			sw.Close();
        }
    }


	[Serializable]
	[XmlRoot("productType")]
	public  sealed class ProductType {
		[XmlAttribute("name")]
		public string Name { get; set; }
		[XmlAttribute("filingPriority")]
		public FilingPriority Priority { get; set; }
		[XmlAttribute("skuTag")]
		public string SkuTag { get; set; }
		[XmlAttribute("categorise")]
		public bool Categorise { get; set; }
		public ProductType() { }
		public ProductType(string name,
						FilingPriority priority,
						string skuTag = null,
						bool categorise = false) {
			Name = name;
			Priority = priority;
			SkuTag = skuTag;
			Categorise = categorise;
		}
	}

	[Serializable]
	public enum FilingPriority {
		[XmlEnum(Name ="0")]
		None = 0,
		[XmlEnum(Name = "1")]
		Low = 1,
		[XmlEnum(Name = "2")]
		Mid = 2,
		[XmlEnum(Name = "4")]
		High = 4,
		[XmlEnum(Name = "8")]
		Immediate =8
    }
}

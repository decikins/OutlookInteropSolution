using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml;

namespace FPBInterop {
    internal class XmlHandling {
		private static readonly TraceSource Tracer = new TraceSource("FPBInterop.XmlHandling");

		private const string ColourXmlPath = "./ColourChart.xml";
		private const string FranchiseeXmlPath = "./Franchises.xml";
		private const string OrderTypeXmlPath = "./OrderTypeInfo.xml";
		private const string SKUTypeTable = "./SKUTypeTable.xml";

		private static XmlDocument xml = new XmlDocument();

		public static Dictionary<string,Colour> LoadColours() {
			Tracer.TraceEvent(TraceEventType.Information, 0, $"Begin loading ColourChart.xml from {ColourXmlPath}");
			if(!xml.HasChildNodes)
				xml.Load(ColourXmlPath);

			Dictionary<string, Colour> colours = new Dictionary<string, Colour>();
			XmlNode sideColours = xml.SelectSingleNode("//sideColours");

			Tracer.TraceEvent(TraceEventType.Verbose, 0, 
				$"ColourChart.xml contains {sideColours.ChildNodes.Count} entries");

			foreach (XmlNode node in sideColours) {
				if (node.Name != "colour" || node.Attributes.Count == 0)
					continue;

				string name;
				bool fondant;
				bool sprinkle;
				bool coconut;
				try {
					name = node.Attributes.GetNamedItem("name").Value;
					fondant = bool.Parse(node.Attributes.GetNamedItem("fondant").Value);
					sprinkle = bool.Parse(node.Attributes.GetNamedItem("sprinkle").Value);
					coconut = bool.Parse(node.Attributes.GetNamedItem("coconut").Value);
				} catch(FormatException fe) {
					Tracer.TraceEvent(TraceEventType.Error, 0, fe.Message);
					return null;
                }

				SideType sides = fondant ? SideType.Fondant : 0;
				sides |= sprinkle ? SideType.Sprinkles : 0;
				sides |= coconut ? SideType.Coconut : 0;

				colours.Add(name, new Colour(name, sides));
			}
			xml = null;
			Tracer.TraceEvent(TraceEventType.Information, 0,
				$"Loading ColourChart.xml complete");
			return colours;
		}
		public static void AddColour(string name, bool fondant, bool sprinkle, bool coconut) {
			if (!xml.HasChildNodes)
				xml.Load(ColourXmlPath);

			if (xml.SelectSingleNode($"//colour/name[text()='{name}']") == null) {
				Tracer.TraceEvent(TraceEventType.Information,0,"Selected name already exists");
				return;
			}

			XmlElement colour = xml.CreateElement("colour");

			XmlElement n = xml.CreateElement("name");
			XmlElement f = xml.CreateElement("fondant");
			XmlElement s = xml.CreateElement("sprinkle");
			XmlElement c = xml.CreateElement("coconut");

			n.InnerText = name;
			f.InnerText = fondant.ToString().ToUpper();
			s.InnerText = sprinkle.ToString().ToUpper();
			c.InnerText = coconut.ToString().ToUpper();

			colour.AppendChild(n);
			colour.AppendChild(f);
			colour.AppendChild(s);
			colour.AppendChild(c);

			xml.DocumentElement.AppendChild(colour);
			xml.Save(ColourXmlPath);
		}
		public static void RemoveColour(string name) {
			if (!xml.HasChildNodes)
				xml.Load(ColourXmlPath);
			try {
				XmlNode child = xml.SelectSingleNode($"//colour/name[text()='{name}']").ParentNode;
				child.ParentNode.RemoveChild(child);
			}
			catch (Exception e) {
				Tracer.TraceEvent(TraceEventType.Error,0,e.Message);
				return;
			}
			xml.Save(ColourXmlPath);
		}

		public static Dictionary<string,Franchise> LoadFranchises() {
			xml.Load(FranchiseeXmlPath);

			Dictionary<string, Franchise> franchises = new Dictionary<string, Franchise>();
			XmlNode franchiseXml = xml.SelectSingleNode("//franchises");
			foreach (XmlNode node in franchiseXml) {
				if (node.Name != "franchise" || node.Attributes.Count == 0)
					continue;

				string name;
				string email;
				List<string> alias = new List<string>();

				string aliasString = node.Attributes.GetNamedItem("alias").Value;
				name = node.Attributes.GetNamedItem("name").Value;
				email = node.Attributes.GetNamedItem("email").Value;
				if (!String.IsNullOrEmpty(aliasString)) {
					if (aliasString.Contains(",")) {
						string[] aliases = aliasString.Split(',');
						foreach (string a in aliases) {
							alias.Add(a.Trim(' '));
						}
					} else {
						alias.Add(aliasString);
                    }
				}
                
				franchises.Add(name, new Franchise(name, email, alias));
			}
			xml = null;
			return franchises;
		}

		public static Dictionary<string,OrderType> LoadOrderTypes(string orderTypeTag) {
			xml.Load(OrderTypeXmlPath);

			Dictionary<string, OrderType> types = new Dictionary<string, OrderType>();
			XmlNode stdTypeXml = xml.SelectSingleNode(orderTypeTag);
			foreach (XmlNode node in stdTypeXml) {
				if (node.Name != "type" || node.Attributes.Count == 0)
					continue;

				string name;
				TimeSpan cutoffSpan;
				DayOfWeekFlag daysNotAvailable = DayOfWeekFlag.None;

				name = node.Attributes.GetNamedItem("name").Value;
				cutoffSpan = new TimeSpan(int.Parse(node.Attributes.GetNamedItem("cutoffSpan").Value),0,0,0);
				string daysString = node.Attributes.GetNamedItem("notAvailableDays").Value;
				if (!String.IsNullOrEmpty(daysString)) {
					if (daysString.Contains(",")) {
						string[] days = daysString.Split(',');
						foreach (string day in days) {
							daysNotAvailable |= (DayOfWeekFlag)Enum.Parse(typeof(DayOfWeekFlag), day.Trim(' '));
						}
                    }
                    else {
						daysNotAvailable = (DayOfWeekFlag)Enum.Parse(typeof(DayOfWeekFlag), daysString);
					}
				}
					
				int filingPriority = int.Parse(node.Attributes.GetNamedItem("filingPriority").Value);
				XmlNode skuTagAttr = node.Attributes.GetNamedItem("skuTag");
				string skuTag = (skuTagAttr == null) ? skuTagAttr.Value : null;
				types.Add(name, new OrderType(name, cutoffSpan, daysNotAvailable, (FilingPriority)filingPriority, skuTag));
			}
			xml = null;
			return types;
		}
	}

	[Flags]
	internal enum DayOfWeekFlag {
		None = 0,
		Monday = 1,
		Tuesday = 2,
		Wednesday = 4,
		Thursday = 8,
		Friday = 16,
		Saturday = 32,
		Sunday = 64
    }

	internal sealed class Franchise {
		public string StoreName { get; private set; }
		public string Email { get; private set; }
		public List<string> Aliases { get; private set; }

		public Franchise(string name, string email, List<string> alias = null) {
			StoreName = name;
			Email = email;
			Aliases = alias;
		}
	}

	internal sealed class OrderType {
		public string Type { get; private set; }
		public TimeSpan CutoffPeriod { get; private set; }
		public DayOfWeekFlag DaysNotAvailable { get; private set; }
		public FilingPriority Priority { get; private set; }
		public string SkuTag { get; private set; }
		public OrderType(string type, 
						TimeSpan cutoff, 
						DayOfWeekFlag daysUnavailable, 
						FilingPriority priority,
						string skuTag = null) {
			Type = type;
			CutoffPeriod = cutoff;
			DaysNotAvailable = daysUnavailable;
			Priority = priority;
			SkuTag = skuTag;
        }
    }
	internal static class OrderTypeXmlTags {
		public static string Standard { get { return "//standard"; } }
		public static string Event { get { return "//event"; } }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml;
using System.Xml.Serialization;

namespace FPBInterop {
    internal static class XmlHandler {
		private static readonly TraceSource Tracer = new TraceSource("FPBInterop.XmlHandling");

		internal static Dictionary<string, Franchise> Franchises;
		internal static Dictionary<string, Colour> Colours;
		internal static Dictionary<string, ProductType> ProductTypesStandard;

		private const string ConfigXmlPath = "./Config.xml";

		private static XmlDocument xml = new XmlDocument();

		internal static void LoadConfig() {
			xml.Load(ConfigXmlPath);

			if (xml.SelectSingleNode("//sideColours") == null)
				throw new XmlException(
					$"Config.xml does not contain 'sideColours' node, no colour info loaded");
			else
				Colours = LoadColours();

			if (xml.SelectSingleNode("//franchises") == null)
				throw new XmlException(
					$"Config.xml does not contain 'franchises' node, no franchise info loaded");
			else
				Franchises = LoadFranchises();

			if (xml.SelectSingleNode("//flavours") == null)
				throw new XmlException(
					$"Config.xml does not contain 'flavours' node, no flavour info loaded");
			else;
				//hasFlavourInfo = true;

			if (xml.SelectSingleNode("//productTypeInfo") == null)
				throw new XmlException(
					$"Config.xml does not contain 'productTypeInfo' node, no product type info loaded");
			else
				ProductTypesStandard = LoadStandardProductTypes();
		}

		internal static Dictionary<string,Colour> LoadColours() {
			Tracer.TraceEvent(TraceEventType.Information, 0, $"Begin loading colour table");

			Dictionary<string, Colour> colours = new Dictionary<string, Colour>();
			XmlNode sideColours = xml.SelectSingleNode("//sideColours");

			if (sideColours == null) 
				throw new XmlException(
					$"Config.xml does not contain 'sideColours' node, no colour info loaded");
			

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
			Tracer.TraceEvent(TraceEventType.Information, 0,
				$"Loading ColourChart.xml complete");
			return colours;
		}
		internal static Dictionary<string,Franchise> LoadFranchises() {
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
			return franchises;
		}
		internal static Dictionary<string,ProductType> LoadStandardProductTypes() {
			xml.Load(ConfigXmlPath);

			Dictionary<string, ProductType> types = new Dictionary<string, ProductType>();
			XmlNode stdTypeXml = xml.SelectSingleNode("//productTypeInfo/standard");
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
				types.Add(skuTagAttr.Value, new ProductType(name, cutoffSpan, daysNotAvailable, (FilingPriority)filingPriority, skuTagAttr.Value));
			}
			xml = null;
			return types;
		}

		internal static void AddColour(string name, bool fondant, bool sprinkle, bool coconut) {
			if (xml.SelectSingleNode($"//colour/name[text()='{name}']") != null) {
				Tracer.TraceEvent(TraceEventType.Information, 0, "Selected name already exists");
				return;
			}

			XmlElement colour = xml.CreateElement("colour");

			XmlElement n = xml.CreateElement("name");
			XmlElement f = xml.CreateElement("fondant");
			XmlElement s = xml.CreateElement("sprinkle");
			XmlElement c = xml.CreateElement("coconut");

			n.InnerText = name;
			f.InnerText = fondant.ToString().ToLower();
			s.InnerText = sprinkle.ToString().ToLower();
			c.InnerText = coconut.ToString().ToLower();

			colour.AppendChild(n);
			colour.AppendChild(f);
			colour.AppendChild(s);
			colour.AppendChild(c);

			xml.DocumentElement.SelectSingleNode("//sideColours").AppendChild(colour);
			xml.Save(ConfigXmlPath);
		}
		internal static void RemoveColour(string name) {
			try {
				XmlNode colour = xml.SelectSingleNode($"//colour/name[text()='{name}']").ParentNode;
				if (colour == null) {
					Tracer.TraceEvent(TraceEventType.Information, 0, $"No colour info exists with name {name}");
					return;
				}
				colour.ParentNode.RemoveChild(colour);
			}
			catch (Exception e) {
				Tracer.TraceEvent(TraceEventType.Error, 0, e.Message);
				return;
			}
			xml.Save(ConfigXmlPath);
		}

		internal static ProductType GetProductType(string sku) {
			foreach(KeyValuePair<string,ProductType> entry in ProductTypesStandard) {
				if (sku.StartsWith(entry.Key));
					return entry.Value;
            }
			throw new Exception($"No Product type found that matches SKU {sku}");
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

	internal sealed class ProductType {
		public string Type { get; private set; }
		public TimeSpan CutoffPeriod { get; private set; }
		public DayOfWeekFlag DaysNotAvailable { get; private set; }
		public FilingPriority Priority { get; private set; }
		public string SkuTag { get; private set; }
		public ProductType(string type, 
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
}

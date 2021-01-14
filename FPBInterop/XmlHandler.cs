using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace FPBInterop {
    internal class XmlHandler {
		private const string OrderConfigXmlPath = "./OrderConfig.xml";

		private readonly XmlDocument xml = new XmlDocument();
		private Dictionary<string, Franchise> Franchises;
		private Dictionary<string, ProductType> ProductTypesStandard;
		internal Dictionary<string, ProductType> GetProductTypesStandard() {
			return ProductTypesStandard;
        }
		internal Dictionary<string, Franchise> GetFranchises() {
			return Franchises;
		}

		internal static readonly ProductType DefaultProductType = new ProductType(
			"Default", FilingPriority.GENERAL);

		public XmlHandler() {
			xml.Load(OrderConfigXmlPath);
			Franchises = LoadFranchises();
			ProductTypesStandard = LoadStandardProductTypes();
		}

		/*internal static Dictionary<string,Colour> LoadColours() {
			Logger.TraceEvent(TraceEventType.Information, $"Loading colour table");

			Dictionary<string, Colour> colours = new Dictionary<string, Colour>();
			XmlNode colourNode = xml.SelectSingleNode("//sideColours");

			if (colourNode == null) 
				throw new XmlException(
					$"OrderConfig.xml does not contain 'sideColours' node, no colour info loaded");
			

			Logger.TraceEvent(TraceEventType.Verbose,
				$"OrderConfig.xml contains {colourNode.ChildNodes.Count} entries");

			foreach (XmlNode node in colourNode) {
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
					Logger.TraceEvent(TraceEventType.Error, 0, fe.Message);
					return null;
                }

				SideType sides = fondant ? SideType.Fondant : 0;
				sides |= sprinkle ? SideType.Sprinkles : 0;
				sides |= coconut ? SideType.Coconut : 0;

				colours.Add(name, new Colour(name, sides));
			}
			Logger.TraceEvent(TraceEventType.Information, 0,
				$"Loading ColourChart.xml complete");
			return colours;
		}*/
		internal Dictionary<string,Franchise> LoadFranchises() {
			Dictionary<string, Franchise> franchises = new Dictionary<string, Franchise>();
			XmlNode franchiseNode = xml.SelectSingleNode("//franchises");

			if (franchiseNode == null)
				throw new XmlException(
					$"OrderConfig does not contain 'franchise' node");

			Logger.TraceEvent(TraceEventType.Verbose,
				$"OrderConfig contains {franchiseNode.ChildNodes.Count} franchise entries");

			foreach (XmlNode node in franchiseNode) {
				if (node.Name != "franchise" || node.Attributes.Count == 0)
					continue;

				string name = node.Attributes.GetNamedItem("name").Value;
				string email = node.Attributes.GetNamedItem("email").Value;
				string aliasString = node.Attributes.GetNamedItem("alias").Value;

				List<string> alias = new List<string>();
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
		internal Dictionary<string,ProductType> LoadStandardProductTypes() {
			Dictionary<string, ProductType> types = new Dictionary<string, ProductType>();
			XmlNode standardTypeNode = xml.SelectSingleNode("//productTypeInfo/standard");
			if (standardTypeNode == null)
				throw new XmlException(
					$"OrderConfig does not contain standard 'productTypeInfo' node");
            
			Logger.TraceEvent(TraceEventType.Verbose,
				$"OrderConfig contains {standardTypeNode.ChildNodes.Count} standard productTypeInfo entries");

			foreach (XmlNode node in standardTypeNode) {
				if (node.Name != "type" || node.Attributes.Count == 0)
					continue;

				string name = node.Attributes.GetNamedItem("name").Value;
				int filingPriority = int.Parse(node.Attributes.GetNamedItem("filingPriority").Value);
				XmlNode skuTagAttr = node.Attributes.GetNamedItem("skuTag");
				if (String.IsNullOrEmpty(skuTagAttr.Value))
					skuTagAttr.Value = "Not Applicable";

				bool categorise = bool.Parse(node.Attributes.GetNamedItem("categorise").Value);

				types.Add(skuTagAttr.Value, new ProductType(name, (FilingPriority)filingPriority, skuTagAttr.Value,categorise));
			}
			return types;
		}

		internal Franchise GetFranchiseInfo(string nameOrAlias) {
			if (Franchises.ContainsKey(nameOrAlias))
				return Franchises[nameOrAlias];
			else {
				foreach (Franchise f in Franchises.Values) {
					foreach (string alias in f.Aliases) {
						if (alias == nameOrAlias)
							return f;
					}
				}
			}
			throw new ArgumentException($"Store name {nameOrAlias} not found");
		}
		internal ProductType GetProductType(string sku) {
			foreach(KeyValuePair<string,ProductType> entry in ProductTypesStandard) {
				if (sku.StartsWith(entry.Key))
					return entry.Value;
            }
			throw new ArgumentException($"No Product type found that matches SKU {sku}");
		}

		private static DayOfWeekFlag GetDaysFromStringList(string list) {
			DayOfWeekFlag daysFlag = DayOfWeekFlag.None;
			if (!String.IsNullOrEmpty(list)) {
				if (list.Contains(",")) {
					string[] days = list.Split(',');
					foreach (string day in days) {
						daysFlag |= (DayOfWeekFlag)Enum.Parse(typeof(DayOfWeekFlag), day.Trim(' '));
					}
				}
				else {
					daysFlag = (DayOfWeekFlag)Enum.Parse(typeof(DayOfWeekFlag), list);
				}
			}
			return daysFlag;
		}
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
		public string Name { get; private set; }
		public FilingPriority Priority { get; private set; }
		public string SkuTag { get; private set; }
		public bool Categorise { get; private set; }
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
}

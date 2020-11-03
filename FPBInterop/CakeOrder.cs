using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Diagnostics;

namespace FPBInterop {
	internal abstract class BaseOrder {
		public int Store { get; private set; }
		public string StoreName { get; private set; }
		public string Email { get; private set; }
		public DateTime DeliveryDate { get; private set; }

		public BaseOrder() { }
	}

	internal abstract class Tier {
		public CakeFlavour flavour { get; private set; }
		public SideDecoration side { get; private set; }
		public SideColourEnum sideColour { get; private set; }

		public Tier(CakeFlavour flavour, SideDecoration side) {
			this.flavour = flavour;
			this.side = side;
        }
    }

	internal class DecoratedCakeOrder : BaseOrder
	{
		public int orderNumber;
		public string designCode;
		public string customerName;
		public string contactNumber;
		public string cakeDetails;
		public Tier topTier;
		public Tier flavourMiddleTier;
		public Tier flavourBottomTier;
		public DecoratedCakeOrder() 
		{
		}
	}

	public static class SideColour {
		public static Dictionary<int, Colour> Colours = new Dictionary<int, Colour>();

		private const string ColourChartXml = "./ColourChart.xml";
		private static XmlDocument ColourChart = new XmlDocument();

		public static void GetColours() {
			if (!ColourChart.HasChildNodes)
				ColourChart.Load(ColourChartXml);

			XmlNode sideColours = ColourChart.SelectSingleNode("//sideColours");
			int index = 0;
			foreach (XmlNode node in sideColours) {
				if (node.Name != "colour" || node.Attributes.Count==0)
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
				}
				catch (Exception e) {
					Trace.WriteLine($"Could not parse colour xml: {e.Message}");
					return;
				}

				Colour colour = new Colour(name, fondant, sprinkle, coconut);
				Colours.Add(index++, colour);
			}
			Trace.WriteLine(Colours.Count);
		}

		public static void AddColour(string name, bool fondant, bool sprinkle, bool coconut) {
			if (!ColourChart.HasChildNodes)
				ColourChart.Load(ColourChartXml);

			if (ColourChart.SelectSingleNode($"//colour/name[text()='{name}']") == null) { 
				Trace.WriteLine("Selected name already exists");
				return; 
			}

			XmlElement colour = ColourChart.CreateElement("colour");

			XmlElement n = ColourChart.CreateElement("name");
			XmlElement f = ColourChart.CreateElement("fondant");
			XmlElement s = ColourChart.CreateElement("sprinkle");
			XmlElement c = ColourChart.CreateElement("coconut");

			n.InnerText = name;
			f.InnerText = fondant.ToString().ToUpper();
			s.InnerText = sprinkle.ToString().ToUpper();
			c.InnerText = coconut.ToString().ToUpper();

			colour.AppendChild(n);
			colour.AppendChild(f);
			colour.AppendChild(s);
			colour.AppendChild(c);

			ColourChart.DocumentElement.AppendChild(colour);
			ColourChart.Save(ColourChartXml);
		}
		public static void RemoveColour(string name) {
			if (!ColourChart.HasChildNodes)
				ColourChart.Load(ColourChartXml);
			try {
				XmlNode child = ColourChart.SelectSingleNode($"//colour/name[text()='{name}']").ParentNode;
				Trace.WriteLine(child.Name);
				child.ParentNode.RemoveChild(child);
			}
			catch (Exception e) {
				Trace.WriteLine(e.Message);
				return;
			}
			ColourChart.Save(ColourChartXml);
		}

		public sealed class Colour {
			public string Name { get; private set; }
			public bool Fondant { get; private set; }
			public bool Sprinkle { get; private set; }
			public bool Coconut { get; private set; }

			public Colour(string name, bool fondant, bool sprinkle, bool coconut) {
				Name = name;
				Fondant = fondant;
				Sprinkle = sprinkle;
				Coconut = coconut;
			}
		}
	}

	internal enum CakeFlavour
	{
		NONE,
		ANGELMUD,
		CHOCOLATEMUD,
		MARBLEMUD,
		ORANGEMUD,
		REDVELVET,
		RAINBOWMUD,
		CARAMELMUD,
		RASPBERRYROUGH,
		VANBUTTERSPONGE,
		CHOCBUTTERSPONGE,
		VANFRESHSPONGE,
		CHOCFRESHSPONGE
	}
	internal enum SideDecoration
	{
		NONE,
		FULLYICED,
		SPRINKLES,
		COCONUT,
		BUTTERCREAM
	}
	internal enum SideColourEnum
	{
		None,
		Red,
		HotPink,
		PalePink,
		Purple,
		PalePurple,
		NavyBlue,
		RoyalBlue,
		SatinBlue,
		PaleBlue,
		AquaBlue,
		EmeraldGreen,
		GrassGreen,
		PaleGreen,
		KellyGreen,
		Yellow,
		PaleYellow,
		Orange,
		PaleOrange,
		Chocolate,
		Black,
		Grey,
		White
	}
}

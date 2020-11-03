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
		public int store;
		public string storename;
		public string email;
		public DateTime deliveryDate;
		public DateTime deliveryDay;
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
		private static XmlDocument ColourChart = new XmlDocument();
		public sealed class TierColour {
			public string Name { get; private set; }
			public bool Fondant { get; private set; }
			public bool Sprinkle { get; private set; }
			public bool Coconut { get; private set; }

			public TierColour(string name, bool fondant, bool sprinkle, bool coconut) {
				Name = name;
				Fondant = fondant;
				Sprinkle = sprinkle;
				Coconut = coconut;
			}
		}

		public static void AddColour(string name, bool fondant, bool sprinkle, bool coconut) {
			if (!ColourChart.HasChildNodes)
				ColourChart.Load("./ColourChart.xml");

			XmlNode root = ColourChart.DocumentElement;
			XmlElement colour = ColourChart.CreateElement("colour");

			XmlElement n = ColourChart.CreateElement(name);
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
		}

		public static Dictionary<int,TierColour> Colours = new Dictionary<int, TierColour>();
		public static void GetColours() {
			if(!ColourChart.HasChildNodes)
			ColourChart.Load("./ColourChart.xml");

			int index = 0;
			foreach(XmlNode node in ColourChart.DocumentElement) {
				if (node.Name != "colour" || !node.HasChildNodes)
					return;
				string name;
				bool fondant;
				bool sprinkle;
				bool coconut;
				try {
					name = node.SelectSingleNode("./name").InnerText;
					fondant = bool.Parse(node.SelectSingleNode("./fondant").InnerText);
					sprinkle = bool.Parse(node.SelectSingleNode("./sprinkle").InnerText);
					coconut = bool.Parse(node.SelectSingleNode("./coconut").InnerText);
				}
				catch (Exception e) {
					Trace.WriteLine($"Could not parse colour xml: {e.Message}");
					return;
                }

				TierColour colour = new TierColour(name, fondant, sprinkle, coconut);
				Colours.Add(index, colour);
				index++;
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
		VANILLABUTTERSPONGE,
		CHOCBUTTERSPONGE,
		VANILLAFRESHSPONGE,
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

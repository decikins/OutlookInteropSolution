using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace FPBInterop {

	internal class BaseOrder {
		public Franchise Location { get; private set; }
		public DateTime DeliveryDate { get { return DeliveryDate.Date(); } private set { DeliveryDate = value; } }
		public OrderMetadata Meta { get; private set; }

		public BaseOrder(Franchise location, DateTime deliveryDate, OrderMetadata meta) {
			Location = location;
			DeliveryDate = deliveryDate.Date;
			Meta = meta;
		}
	}

	internal abstract class Tier {
		public CakeFlavour flavour { get; private set; }
		public SideType side { get; private set; }
		public Colour sideColour { get; private set; }

		public Tier(CakeFlavour flavour, SideType side, Colour colour) {
			this.flavour = flavour;
			this.side = side;
			this.sideColour = colour;
        }
    }

	internal class DecoratedCakeOrder : BaseOrder
	{
		public int OrderNumber { get; private set; }
		public string SKU { get; private set; }
		public string CustomerName { get; private set; }
		public string ContactNumber { get; private set; }
		public string Details { get; private set; }
		public List<Tier> Tiers { get; private set; }
		public int Price { get; private set; }

		public DecoratedCakeOrder(
			Franchise location,
			DateTime deliveryDate,
			OrderMetadata meta) : base(location,deliveryDate,meta)
		{
		}
	}

	internal sealed class Colour {
		public string Name { get; private set; }
		public SideType SideType { get; private set; }
		public Colour(string name, SideType sideTypes) {
			Name = name;
			SideType = sideTypes;
		}
	}

	public static class DateExtension {
		public static DateTime Date(this DateTime date) {
			return date.Date;
        }
    }

	public enum CakeFlavour
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

	[Flags]
	public enum SideType
	{
		None,
		Fondant,
		Sprinkles,
		Coconut,
		Buttercream
	}

	public enum SideColour
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

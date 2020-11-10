using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace FPBInterop {

	internal class Product {
		public string Name { get; private set; }
		public string SKU { get; private set; }
		public OrderType ProductType { get; private set; }
		public int Price { get; private set; }
		public Product(string name, string skuCode, OrderType type, int price) {
			Name = name;
			SKU = skuCode;
			ProductType = type;
			Price = price;
		}
	}

	internal class BaseOrder {
		public Franchise Location { get; private set; }
		public DateTime DeliveryDate { get; private set; }
		public OrderMetadata Meta { get; private set; }

		public BaseOrder(Franchise location, DateTime deliveryDate, OrderMetadata meta) {
			Location = location;
			DeliveryDate = deliveryDate;
			Meta = meta;
		}
	}

	[Flags]
	internal enum FilingPriority : byte {
		NONE = 0,
		GENERAL = 1,
		COOKIE = 2,
		CUSTOM = 4
    }

	internal class MagentoOrder : BaseOrder { 
		public List<Product> Products { get; private set; }
		public FilingPriority OrderPriority { get; private set; }
		public MagentoOrder(Franchise location, 
							DateTime deliveryDate, 
							OrderMetadata meta,
							List<Product> products) 
			: base(location,deliveryDate,meta) {
			Products = products;
			OrderPriority = FilingPriority.NONE;
			foreach (Product p in products) {
				OrderPriority |= p.ProductType.Priority;
            }
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

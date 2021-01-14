using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace FPBInterop {

	internal sealed class MagentoProduct {
		public string Name { get; private set; }
		public string SKU { get; private set; }
		public ProductType ProductType { get; private set; }
		public MagentoProduct(string name, string skuCode, ProductType type) {
			Name = name;
			SKU = skuCode;
			ProductType = type;
		}
	}

	internal interface IBaseOrder {
		Franchise Location { get; }
		DateTime DeliveryDate { get; }
		OrderMetadata Meta { get; }
    }

	internal class BaseOrder : IBaseOrder {
		public Franchise Location { get; private set; }
		public DateTime DeliveryDate { get; private set; }
		public OrderMetadata Meta { get; private set; }
		public FilingPriority Priority { get; }

		public BaseOrder(Franchise location, DateTime deliveryDate, OrderMetadata meta) {
			Location = location;
			DeliveryDate = deliveryDate;
			Meta = meta;
		}
	}

	internal class WufooOrder : BaseOrder {
		public ProductType Type { get; private set; }
		public new FilingPriority Priority { get { return Type.Priority; } }
		public WufooOrder(Franchise location, DateTime deliveryDate, OrderMetadata meta, ProductType type) 
			: base(location, deliveryDate, meta) {
			Type = type;
        }
    }

	internal class MagentoOrder : BaseOrder { 
		public List<MagentoProduct> Products { get; private set; }
		public new FilingPriority Priority { get {
				FilingPriority priority = FilingPriority.NONE;
				foreach (MagentoProduct p in Products) {
					priority |= p.ProductType.Priority;
				}
				return priority;
			} 
		}

		public MagentoOrder(Franchise location, 
							DateTime deliveryDate, 
							OrderMetadata meta,
							List<MagentoProduct> products) 
			: base(location,deliveryDate,meta) {
			Products = products;
		}
	}

	/*internal class DecoratedOrder : WufooOrder
	{
		public int OrderNumber { get; private set; }
		public string SKU { get; private set; }
		public string CustomerName { get; private set; }
		public string ContactNumber { get; private set; }
		public string Details { get; private set; }
		public List<Tier> Tiers { get; private set; }
		public int Price { get; private set; }

		public DecoratedOrder(
			Franchise location,
			DateTime deliveryDate,
			OrderMetadata meta,
			ProductType type) : base(location,deliveryDate,meta,type)
		{
		}
	}

	internal sealed class Tier {
		public CakeFlavour flavour { get; private set; }
		public SideType side { get; private set; }
		public Colour sideColour { get; private set; }

		public Tier(CakeFlavour flavour, SideType side, Colour colour) {
			this.flavour = flavour;
			this.side = side;
			this.sideColour = colour;
		}
	}*/

	internal sealed class Colour {
		public string Name { get; private set; }
		public SideType SideType { get; private set; }
		public Colour(string name, SideType sideTypes) {
			Name = name;
			SideType = sideTypes;
		}
	}

	[Flags]
	internal enum FilingPriority : byte {
		NONE = 0,
		GENERAL = 1,
		COOKIE = 2,
		CUSTOM = 4
	}

	[Flags]
	internal enum SideType
	{
		None,
		Fondant,
		Sprinkles,
		Coconut,
		Buttercream
	}

	internal enum SideColour
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

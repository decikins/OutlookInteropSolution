using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FPBInteropConsole
{
	
	public class DecoratedCakeOrder
	{
		public int store;
		public int orderNumber;
		public string storename;
		public string email;
		public string customerName;
		public string contactNumber;
		public DateTime deliveryDate;
		public DateTime deliveryDay;
		public string designCode;
		public string cakeDetails;
		public string flavourTopTier;
		public string flavourMiddleTier;
		public string flavourBottomTier;
		public string sidesTopTier;
		public string sidesMiddleTier;
		public string sidesBottomTier;
		public string sideColourTopTier;
		public string sideColourMiddleTier;
		public string sideColourBottomTier;
		public DecoratedCakeOrder() 
		{

		}
	}

	public enum CakeFlavour
	{
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

	public enum SideDecoration
	{
		FULLYICED,
		SPRINKLES,
		COCONUT,
		BUTTERCREAM,
		CHOCOLATEDIPPED
	}

	public enum FondantColours
	{
		RED,
		HOTPINK,
		PALEPINK,
		PURPLE,
		MAUVE,
		NAVYBLUE,
		ROYALBLUE,
		SATINBLUE,
		PALEBLUE,
		AQUABLUE,
		EMERALDGREEN,
		PALEGREEN,
		KELLYGREEN,
		YELLOW,
		PALEYELLOW,
		ORANGE,
		CHOCOLATE,
		BLACK,
		GREY,
		WHITE
	}

	public enum SprinkleColours 
	{
		RED,
		HOTPINK,
		PALEPINK,
		MAUVE,
		ROYALBLUE,
		PALEBLUE,
		GREEN,
		YELLOW,
		ORANGE,
		CHOCOLATE,
		RAINBOW,
		BLACK
	}

	public enum CoconutColours 
	{
		PINK,
		MAUVE,
		BLUE,
		GREEN,
		YELLOW,
		APRICOT,
		WHITE
	}
}

using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace FPBInteropConsole
{
	public class Franchise
	{
		private readonly string storename;
		private readonly string emailAddress;
		private readonly bool isCurrentlyOpen;

		public string StoreName 
		{
			get { return storename; }
		}
		public string EmailAddress
		{
			get { return emailAddress; }
		}
		public bool IsCurrentlyOpen 
		{
			get { return isCurrentlyOpen; }
		}

		public Franchise(string name, string email, bool isOpen)
		{
			this.storename = name;
			this.emailAddress = email;
			this.isCurrentlyOpen = isOpen;
		}

		public void ModifyStoreList() 
		{
			while (true) 
			{
				switch (Console.ReadLine()) 
				{
					default:
						return;
				}
			}
		}
	}
}

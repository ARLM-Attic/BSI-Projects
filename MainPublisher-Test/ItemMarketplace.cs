//
// AMGD
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// Marketplaces:
// Amazon
//  A1 = ShopUsLast
//  A2 = 
//  A3 = 
//  A4 = 
// eBay
//  E1 = Mecalzo
//  E2 = OneMillionShoes
//  E3 = 
//  E4 = 
//  E5 = 
// Websites
//  W1 =
//  W2 =
//  W3

namespace Publisher_Test
{
    public class ItemMarketplace
    {
        public const int MARKETPLACE_TYPE_AMAZON = 1;
        public const int MARKETPLACE_TYPE_EBAY = 2;

        public const int MARKETPLACE_MAXMARKETS = 12;
        public const uint MARKETPLACE_AMAZON = 1;
        public const uint MARKETPLACE_EBAY_MECALZO = 16;
        public const uint MARKETPLACE_EBAY_1MS = 32;

        public int item { get; set; }
        public String ItemLookupCode { get; set; }
        public uint market { get; set; }
        public String Name { get; set; }
        public int Quantity { get; set; }
        public Decimal Price { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        public ItemMarketplace()
        {
            item = 0;
            ItemLookupCode = "";
            market = 0;
            Name = "";
            Quantity = 0;
            Price = 0;
            StartDate = EndDate = new DateTime(1903, 1, 1);
        } // ItemMarketplace

        public ItemMarketplace(int pitem, string pilc, uint pmarket, string pname, int pq, Decimal pprice, DateTime pstart, DateTime pend)
        {
            item = pitem;
            ItemLookupCode = pilc;
            market = pmarket;
            Name = pname;
            Quantity = pq;
            Price = pprice;
            StartDate = pstart;
            EndDate = pend;
        } // ItemMarketplace(int pitem, string pilc, uint pmarket, string pname, int pq, Decimal pprice, DateTime pstart, DateTime pend)

        public ItemMarketplace(ItemMarketplace pm)
        {
            item = pm.item;
            ItemLookupCode = String.Copy(pm.ItemLookupCode);
            market = pm.market;
            Name = String.Copy(pm.Name);
            Quantity = pm.Quantity;
            Price = pm.Price;
            StartDate = pm.StartDate;
            EndDate = pm.EndDate;
        } // ItemMarketplace

    } // ItemMarketplace
} // Publisher_Test

//
// AMGD
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
//

using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;

namespace Publisher_Test
{
    public class ItemExcel
    {
        public static int ITEM_TYPE_SINGLE = 0;
        public static int ITEM_TYPE_PARENT = 10;

        public bool Ok2Publish { get; set; }
        public int Type { get; set; }
        public String Gender { get; set; }
        public String Brand { get; set; }
        public String SKU { get; set; }
        public String ItemLookupCode { get; set; }
        public String UPC { get; set; }
        public String Alias { get; set; }
        public String Size { get; set; }
        public String Width { get; set; }
        public String Condition { get; set; }
        public String Category { get; set; }
        public String Style { get; set; }
        public String RMS_Description { get; set; }
        public String FullDescription { get; set; }
        public String Keywords { get; set; }
        public String Material { get; set; }
        public String Color { get; set; }
        public String Shade { get; set; }
        public String HeelHeight { get; set; }
        public Decimal MSRP { get; set; }

        public int Variation { get; set; }
        public String Widths { get; set; }

        public int Received { get; set; }
        public Decimal Cost { get; set; }
        
        public String Title { get; set; }

        // This item will be associated with a post. Let's save the post ID so we can update the post ifo like
        // eBay posting ID and end date back in the DB after publishing
        public int post_id { get; set; }

        // Qty & price for single item. When this item referes to a single product it will have the q & p here
        public int SingleQuantity { get; set; }
        public Decimal SinglePrice { get; set; }

        // Qtys & prices for multiple listings 
        public int QtyA1 { get; set; }
        public int QtyA2 { get; set; }
        public int QtyA3 { get; set; }
        public int QtyA4 { get; set; }
        public int QtyE1 { get; set; }
        public int QtyE2 { get; set; }
        public int QtyE3 { get; set; }
        public int QtyE4 { get; set; }
        public int QtyE5 { get; set; }
        public int QtyW1 { get; set; }
        public int QtyW2 { get; set; }
        public int QtyW3 { get; set; }

        public Decimal PriceA1 { get; set; }
        public Decimal PriceA2 { get; set; }
        public Decimal PriceA3 { get; set; }
        public Decimal PriceA4 { get; set; }
        public Decimal PriceE1 { get; set; }
        public Decimal PriceE2 { get; set; }
        public Decimal PriceE3 { get; set; }
        public Decimal PriceE4 { get; set; }
        public Decimal PriceE5 { get; set; }

        public Decimal PriceW1 { get; set; }
        public Decimal PriceW2 { get; set; }
        public Decimal PriceW3 { get; set; }

        public String SellingFormat { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        public uint MarketPlaces { get; set; } // BIT Flag for marketplaces

        public List<ItemExcel> Items;
        public List<ItemMarketplace> Markets;

        public List<string> Pictures;
        public List<string> URLPictures;

        public ItemExcel()
        {
            Type = ITEM_TYPE_SINGLE;
            Ok2Publish = false;
            Brand = ItemLookupCode = Alias = RMS_Description = Gender = Title = FullDescription = Keywords = "";
            SKU = Size = Width = Condition = Category = Style = Material = Color = Shade = HeelHeight = SellingFormat = "";
            Widths = UPC = "";
            Variation = 0;
            Received = 0;

            SingleQuantity = QtyA1 = QtyA2 = QtyA3 = QtyA4 = 0;
            SinglePrice = PriceA1 = PriceA2 = PriceA3 = PriceA4 = 0;

            QtyE1 = QtyE2 = QtyE3 = QtyE4 = QtyE5 = 0;
            PriceE1 = PriceE2 = PriceE3 = PriceE4 = PriceE5 = 0;
            
            QtyW1 = QtyW2 = QtyW3 = 0;
            PriceW1 = PriceW2 = PriceW3 = 0;

            post_id = 0;

            Cost = MSRP = 0; 
            StartDate = new DateTime(1903, 1, 1); // Very old date
            EndDate = StartDate.AddDays(1);
            MarketPlaces = 0;
            Items = new List<ItemExcel>();
            Markets = new List<ItemMarketplace>();
            Pictures = new List<string>();
            URLPictures = new List<string>();
        } // ItemExcel

        public ItemExcel(ItemExcel pc)
        {
            Ok2Publish = pc.Ok2Publish;
            Brand = String.Copy(pc.Brand);
            SKU = String.Copy(pc.SKU);
            ItemLookupCode = String.Copy(pc.ItemLookupCode);
            Alias = String.Copy(pc.Alias);
            RMS_Description = String.Copy(pc.RMS_Description);
            Gender = String.Copy(pc.Gender);
            Condition = String.Copy(pc.Condition);
            Category = String.Copy(pc.Category);
            Style = String.Copy(pc.Style);
            Variation = pc.Variation;
            Widths = String.Copy(pc.Widths);
            UPC = String.Copy(pc.UPC);
            Size = String.Copy(pc.Size);
            Width = String.Copy(pc.Width);

            Title = String.Copy(pc.Title);
            FullDescription = String.Copy(pc.FullDescription);
            Keywords = String.Copy(pc.Keywords);

            Material = String.Copy(pc.Material);
            Color = String.Copy(pc.Color);
            Shade = String.Copy(pc.Shade);
            HeelHeight = String.Copy(pc.HeelHeight);

            Received = pc.Received;
            Cost = pc.Cost;
            MSRP = pc.MSRP;

            post_id = pc.post_id;

            SinglePrice = pc.SinglePrice;
            SingleQuantity = pc.SingleQuantity;

            QtyA1 = pc.QtyA1; PriceA1 = pc.PriceA1;
            QtyA2 = pc.QtyA2; PriceA2 = pc.PriceA2;
            QtyA3 = pc.QtyA3; PriceA3 = pc.PriceA3;
            QtyA4 = pc.QtyA4; PriceA4 = pc.PriceA4;

            QtyE1 = pc.QtyE1; PriceE1 = pc.PriceE1;
            QtyE2 = pc.QtyE2; PriceE2 = pc.PriceE2;
            QtyE3 = pc.QtyE3; PriceE3 = pc.PriceE3;
            QtyE4 = pc.QtyE4; PriceE4 = pc.PriceE4;
            QtyE5 = pc.QtyE5; PriceE5 = pc.PriceE5;

            QtyW1 = pc.QtyW1; PriceW1 = pc.PriceW1;
            QtyW2 = pc.QtyW2; PriceW2 = pc.PriceW2;
            QtyW3 = pc.QtyW3; PriceW3 = pc.PriceW3;

            SellingFormat = String.Copy(pc.SellingFormat);
            StartDate = pc.StartDate;
            EndDate = pc.EndDate;

            Items = new List<ItemExcel>();
            foreach (ItemExcel pi in pc.Items)
                Items.Add(new ItemExcel(pi));

            Markets = new List<ItemMarketplace>();
            foreach (ItemMarketplace pm in pc.Markets)
                Markets.Add(new ItemMarketplace(pm));

            Pictures = new List<string>();
            foreach (String ls in pc.Pictures)
                Pictures.Add(ls);

            URLPictures = new List<string>();
            foreach (String ls in pc.URLPictures)
                URLPictures.Add(ls);
        } // ItemExcel(ItemExcel pc)
    } // ExcelItem
} // Publisher_Test

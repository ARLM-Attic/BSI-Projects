//
// AMGD
//

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;

using eBay.Service.Call;
using eBay.Service.Core.Sdk;
using eBay.Service.Core.Soap;
using eBay.Service.Util;

namespace Publisher_Test
{
    public partial class Form1 : Form
    {

        public bool DEBUG_MODE = false;
        public const int POSTING_STATUS_ACTIVE = 0;
        public const int POSTING_STATUS_READY2PUBLISH = 10;
        public const int POSTING_STATUS_BLOCKED = 100;

        public const int QUANTITY_RECORD_TYPE_POSTING = 0;
        public const int QUANTITY_RECORD_TYPE_VARIATION = 10;

        public const int VARIATIONS_NONE = 0;
        public const int VARIATIONS_SIZE = 1;
        public const int VARIATIONS_WIDTH = 2;
        public const int VARIATIONS_COLOR = 4;

        public static string EXCEL_COLUMN_INITIAL = "A";
        public static string EXCEL_COLUMN_FINAL = "AD";

        public static int EXCEL_INTCOLUMN_BRAND = 1;
        public static int EXCEL_INTCOLUMN_SKU = 2;
        public static int EXCEL_INTCOLUMN_LOOKUPCODE = 3;
        public static int EXCEL_INTCOLUMN_SIZE = 4;
        public static int EXCEL_INTCOLUMN_WIDTH = 5;
        public static int EXCEL_INTCOLUMN_CONDITION = 6;
        public static int EXCEL_INTCOLUMN_CATEGORY = 7;
        public static int EXCEL_INTCOLUMN_STYLE = 8;
        public static int EXCEL_INTCOLUMN_TITLE = 9;
        public static int EXCEL_INTCOLUMN_COUNT = 10;
        public static int EXCEL_INTCOLUMN_FULLD = 11;
        public static int EXCEL_INTCOLUMN_KEYWORDS = 12;
        public static int EXCEL_INTCOLUMN_MATERIAL = 13;
        public static int EXCEL_INTCOLUMN_COLOR = 14;
        public static int EXCEL_INTCOLUMN_SHADE = 15;
        public static int EXCEL_INTCOLUMN_HEEL = 16;
        public static int EXCEL_INTCOLUMN_MSRP = 17;
        public static int EXCEL_INTCOLUMN_RMSDESCRIPTION = 18;
        public static int EXCEL_INTCOLUMN_GENDER = 19;
        public static int EXCEL_INTCOLUMN_RECEIVED = 20;
        public static int EXCEL_INTCOLUMN_COST = 21;
        public static int EXCEL_INTCOLUMN_UPC = 22;

        public static int EXCEL_INTCOLUMN_QTY_AMAZON = 23;
        public static int EXCEL_INTCOLUMN_QTY_MECALZO = 24;
        public static int EXCEL_INTCOLUMN_QTY_1MS = 25;

        public static int EXCEL_INTCOLUMN_PRICE_AMAZON = 26;
        public static int EXCEL_INTCOLUMN_PRICE_MECALZO = 27;
        public static int EXCEL_INTCOLUMN_PRICE_1MS = 28;

        public static int EXCEL_INTCOLUMN_SELLINGFORMAT = 29;
        public static int EXCEL_INTCOLUMN_STARTDATE = 30;

        public static int EXCEL_INTCOLUMN_PRICE = 5;

        public const double SHIPPING_SURCHARGE = 8.50;

        private static ApiContext apiContext = null;
        private string _descriptionHeader, _descriptionFooter, lorginalpathfile, lpicturespath;

        private uint[] MarketPlaces = { ItemMarketplace.MARKETPLACE_AMAZON, 
                                        ItemMarketplace.MARKETPLACE_EBAY_MECALZO,
                                        ItemMarketplace.MARKETPLACE_EBAY_1MS };

        Boolean lstop;

        List<ItemExcel> theProducts;
        List<ItemType> itemsOnline;

        berkeleyDataSet.bsi_marketplacesDataTable ldsMarkets = new berkeleyDataSet.bsi_marketplacesDataTable();
        berkeleyDataSet.bsi_marketplacesRow currentMarketPlace = null;



        public Form1()
        {
            InitializeComponent();
        } // Form1

        private void btnStop_Click(object sender, EventArgs e)
        {
            lstop = true;
        } // btnStop_Click

        private void btnStart_Click(object sender, EventArgs e)
        {
            lstop = false;
            publishProducts();
        } // btnStart_Click

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtItemID.Text))
            {
                int lid = (int)cmbMarketplace.SelectedValue;

                try
                {
                    String lresponse = "\r\n" + txtStatus.Text;
                    currentMarketPlace = ldsMarkets[cmbMarketplace.SelectedIndex];
                    GetApiContext();

                    /*
                    GetItemCall lgic = new GetItemCall(apiContext);
                    lgic.ItemID = txtItemID.Text.Trim();
                    lgic.Execute();
                    ItemType litem = lgic.Item;                    
                    if (litem != null)
                    {
                        lresponse += "Item.SellingStatus.ListingStatus = " + litem.SellingStatus.ListingStatus + "\r\n" +
                                     "End time:" + litem.ListingDetails.EndTime.ToLongDateString();
                    }
                    else
                    {
                        lresponse += "--- NO ITEM FOUND ---";
                    };
                    */

                    GetMyeBaySellingCall lgmes = new GetMyeBaySellingCall(apiContext);
                    /*
                    lgmes.ActiveList.Include = false;
                    lgmes.BidList.Include = false;
                    lgmes.DeletedFromSoldList.Include = false;
                    lgmes.DeletedFromUnsoldList.Include = false;
                    lgmes.ScheduledList.Include = false;
                    lgmes.SellingSummary.Include = false;
                    lgmes.SoldList.Include = false;
                    */
                    //lgmes.UnsoldList = new ItemListCustomizationType();

                    lgmes.DetailLevelList = new DetailLevelCodeTypeCollection(new DetailLevelCodeType[] { DetailLevelCodeType.ReturnAll });
                    int lpage = 0, lcount = 0;
                    List<ItemType> TheUnsoldProducts = new List<ItemType>();
                    try
                    {
                        do
                        {
                            ++lpage;
                            txtStatus.Text += "\r\nReading page " + lpage.ToString() + "\r\n";
                            txtStatus.Update();
                            Application.DoEvents();

                            ItemListCustomizationType lunsolds = new ItemListCustomizationType();
                            lunsolds.Pagination = new PaginationType();
                            lunsolds.Pagination.PageNumber = lpage;
                            lunsolds.Pagination.EntriesPerPage = 20;
                            lunsolds.Include = true;
                            lunsolds.DurationInDays = 7;
                            lunsolds.IncludeNotes = false;
                            SellingSummaryType lst = lgmes.GetMyeBaySelling(null, null, null, lunsolds, null, null, null, null, true);
                            lcount += lgmes.UnsoldListReturn.ItemArray.Count;
                            TheUnsoldProducts.AddRange(lgmes.UnsoldListReturn.ItemArray.ToArray());
                            foreach (ItemType li in lgmes.UnsoldListReturn.ItemArray)
                            {
                                txtStatus.Text += "\r\n " + li.ItemID + "\t " + li.Title;
                            } // foreach

                        } while (lgmes.UnsoldListReturn.PaginationResult.TotalNumberOfPages > lpage);
                    }
                    catch (Exception pe)
                    {
                        MessageBox.Show("Error: " + pe.ToString());
                    }

                    //txtStatus.Text = lresponse;
                    MessageBox.Show("The process ended!");
                }
                catch (Exception pe)
                {
                    MessageBox.Show(pe.ToString());
                }
            }
            else
            {
                MessageBox.Show(" PLEASE ENTER AN EBAY PRODUCT ID IN THE BOX ");
                txtItemID.Focus();
            } // if
        } // btnSearch_Click

        private void btnReadMarket_Click(object sender, EventArgs e)
        {
            currentMarketPlace = ldsMarkets[cmbMarketplace.SelectedIndex];
            readMarketplace();
        } // btnReadMarket_Click

        // ----------------------------- Service methods --------------------------------

        // Sort method for 2 items: by brand
        private int sortItems(ItemExcel p1, ItemExcel p2)
        {
            int lres = 0;
            
            lres = p1.Brand.CompareTo(p2.Brand);
            return lres;
        } // sortItems

        int sortBySize(ItemExcel p1, ItemExcel p2)
        {
            float lsize1, lsize2;
            int lres = 0;

            lsize1 = float.Parse(p1.Size);
            lsize2 = float.Parse(p2.Size);

            if ((lsize1 - lsize2) < 0)
                lres = -1;
            else
                if ((lsize1 - lsize2) > 0)
                    lres = 1;
                else
                    lres = 0;

            if (lres == 0) lres = string.Compare(p1.Width, p2.Width);

            return lres;
        } // sortBySize

        String convertWidth(string pgender, string pwidth)
        {
            string lwidth = "";

            switch (pgender)
            {
                case "MENS":
                    switch (pwidth)
                    {
                        case "XN": lwidth = "Extra Narrow (A+)"; break;
                        case "N": lwidth = "Narrow (C, B)"; break;
                        case "D":
                        case "M": lwidth = "Medium (D, M)"; break;
                        case "E":
                        case "W": lwidth = "Wide (E,W)"; break;
                        case "XW":
                        case "2E":
                        case "3E":
                        case "EEE":
                        case "EE":
                        case "WW": lwidth = "Extra Wide (EE+)"; break;
                    } // swtich
                    break;

                case "WOMENS":
                    switch (pwidth)
                    {
                        case "XN": lwidth = "Extra Narrow (AAA+)"; break;
                        case "N": lwidth = "Narrow (AA, N)"; break;
                        case "M":
                        case "B": lwidth = "Medium (B, M)"; break;
                        case "W":
                        case "C":
                        case "D": lwidth = "Wide (C, D, W)"; break;
                        case "XW":
                        case "WW": lwidth = "Extra Wide (E+)"; break;
                    }
                    break;

                case "YOUTH":
                    switch (pwidth)
                    {
                        case "XN": lwidth = "X Narrow"; break;
                        case "N": lwidth = "Narrow"; break;
                        case "M":
                        case "B": lwidth = "Medium"; break;
                        case "W":
                        case "C":
                        case "D": lwidth = "Wide"; break;
                        case "XW":
                        case "WW": lwidth = "X Wide"; break;
                    }
                    break;
            } // switch(pgender)

            return lwidth;
        } // convertWidth

        String getEbaySizeName(String pgender)
        {
            String lname = "";

            switch (pgender)
            {
                case "MENS": lname = "US Shoe Size (Men's)"; break;
                case "WOMENS": lname = "US Shoe Size (Women's)"; break;
                case "JUNIOR": lname = "US Shoe Size (Youth)"; break;
            } // switch

            return lname;
        } // getEbaySizeName

        private String properNameString(String ps)
        {
            StringBuilder ls;

            ls = new StringBuilder((!String.IsNullOrEmpty(ps)) ? ps.ToLower() : "");
            if (ls.Length > 0) ls[0] = Char.ToUpper(ls[0]);

            return ls.ToString();
        } // properNameString

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        } // releaseObject

        private ApiContext GetApiContext()
        {
            apiContext = new ApiContext();

            //set Api Server Url

            apiContext.SoapApiServerUrl = Properties.Settings.Default.ApiServerUrl;
            //set Api Token to access eBay Api Server
            ApiCredential apiCredential = new ApiCredential();
            apiCredential.eBayToken = currentMarketPlace.eBayToken;

            apiContext.ApiCredential = apiCredential;
            //set eBay Site target to US
            apiContext.Site = SiteCodeType.US;

            //set Api logging
            apiContext.ApiLogManager = new ApiLogManager();
            apiContext.ApiLogManager.ApiLoggerList.Add(
                new FileLogger("listing_log.txt", true, true, true)
                );
            apiContext.ApiLogManager.EnableLogging = true;

            return apiContext;
        } // GetApiContext

        ItemExcel isTheProductOnEbay(ItemExcel lix)
        {
            ItemExcel lnewItem = null;

            // Look for the item
            ItemType litem = null;

            litem = itemsOnline.Find(
                                      delegate(ItemType pi)
                                      {
                                       return pi.SKU == lix.SKU;
                                      }
                                    );

            if (litem != null)
            {
                switch (lix.SellingFormat)
                {
                    case "A":
                    case "A3":
                    case "A5":
                        // TO-DO Simply add qty to current item on our DB, nothing else.

                        break;
                    default:
                            bool lfoundFlag = false; // We'll use this to look for items
                            if (litem.ListingType != ListingTypeCodeType.Chinese)
                            {
                                // TO-DO Cancel the eBay product

                                // Combine both products into one
                                // Ours is a parent with children, let's add the kid(s) of the found one
                                if (litem.Variations != null)
                                {
                                    // Let's see if we already have this item
                                    lfoundFlag = false;
                                    foreach (ItemExcel lax in lix.Items)
                                    {
                                        if (lax.ItemLookupCode == litem.SKU.Trim())
                                        {
                                            lfoundFlag = true;
                                            lax.SingleQuantity += litem.Quantity - litem.SellingStatus.QuantitySold;
                                            break; // Get out of the loop
                                        }
                                    } // foreach

                                    if (!lfoundFlag) // Create and add the size
                                    {
                                        ItemExcel ltempxl = new ItemExcel(lix.Items[0]);

                                        // Set the size and width
                                        ltempxl.ItemLookupCode = litem.SKU;
                                        String[] lprodinfo = litem.SKU.Split(new char[] { '-' });
                                        if (lprodinfo.Length > 2)
                                        {
                                            ltempxl.Size = lprodinfo[1];
                                            ltempxl.Width = lprodinfo[2];
                                            ltempxl.SingleQuantity = litem.Quantity - litem.SellingStatus.QuantitySold;
                                        }
                                    }
                                }
                                else
                                {
                                    // We have a parent with kids and we'll add more kids. We need to check one by one
                                    foreach (VariationType lnewKidOneBay in litem.Variations.Variation)
                                    {
                                        lfoundFlag = false;
                                        foreach (ItemExcel lax in lix.Items)
                                        {
                                            if (lax.ItemLookupCode == lnewKidOneBay.SKU.Trim())
                                            {
                                                lfoundFlag = true;
                                                lax.SingleQuantity += lnewKidOneBay.Quantity - lnewKidOneBay.SellingStatus.QuantitySold;
                                                break; // Get out of the loop
                                            }
                                        } // foreach

                                        if (!lfoundFlag) // Create the size
                                        {
                                            ItemExcel ltempxl = new ItemExcel(lix.Items[0]);

                                            // Set the size and width
                                            String[] lprodinfo = lnewKidOneBay.SKU.Split(new char[] { '-' });
                                            if (lprodinfo.Length > 2)
                                            {
                                                ltempxl.Size = lprodinfo[1];
                                                ltempxl.Width = lprodinfo[2];
                                                ltempxl.SingleQuantity = lnewKidOneBay.Quantity - lnewKidOneBay.SellingStatus.QuantitySold;
                                            }
                                        }
                                    } // for each
                                } // if (litem.Variations != null)
                            } // if (litem.ListingType != ListingTypeCodeType.Chinese)
                        break;
                } // switch
            } // if (litem != null)

            return lnewItem;
        } // isTheProductOnEbay

        private void publishProducts()
        {
            // Let's publish products by marketplace
            uint[] mktPlaces = { ItemMarketplace.MARKETPLACE_EBAY_MECALZO, ItemMarketplace.MARKETPLACE_EBAY_1MS };
            int lgrandTotalProducts = 0, lgrandTotalItems = 0;

            lstop = false;
            // foreach (uint lmarket in mktPlaces)

            SqlConnection lconn = null;
            berkeleyDataSetTableAdapters.bsi_postingTableAdapter lda;
            berkeleyDataSetTableAdapters.bsi_postsTableAdapter lposts_da;
            berkeleyDataSetTableAdapters.bsi_quantitiesTableAdapter lqtys_da;

            try
            {
                lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString());
                lconn.Open();

                lda = new berkeleyDataSetTableAdapters.bsi_postingTableAdapter();
                lposts_da = new berkeleyDataSetTableAdapters.bsi_postsTableAdapter();
                lqtys_da = new berkeleyDataSetTableAdapters.bsi_quantitiesTableAdapter();

                lda.Connection = lconn;
                lposts_da.Connection = lconn;
                lqtys_da.Connection = lconn;

                theProducts = new List<ItemExcel>();

                foreach (berkeleyDataSet.bsi_marketplacesRow lmarketPlace in ldsMarkets.Rows)
                {
                    currentMarketPlace = lmarketPlace;

                    if (lstop) break;
                    if (currentMarketPlace.type == ItemMarketplace.MARKETPLACE_TYPE_AMAZON) continue;

                    _descriptionHeader = currentMarketPlace.template_header;
                    _descriptionFooter = currentMarketPlace.template_footer;

                    GetApiContext();

                    txtStatus.Text = "Publishing products for " + currentMarketPlace.name + "\r\n" + txtStatus.Text;
                    txtStatus.Update();

                    // Let's get all the posts from this marketplace 
                    theProducts = new List<ItemExcel>();
                    berkeleyDataSet.bsi_postsDataTable lposts = new berkeleyDataSet.bsi_postsDataTable();

                    string po = cbSelectPO.SelectedValue.ToString();

                    lposts_da.FillByUnpublished(lposts, dtPickerStart.Value, po);

                    foreach (berkeleyDataSet.bsi_postsRow lpost in lposts.Rows)
                    {
                        if ( (lpost.marketplace & lmarketPlace.maskId) == 0 ) continue;

                        ItemExcel lix = new ItemExcel();

                        // Get the information of the posting
                        berkeleyDataSet.bsi_postingDataTable lpostingT = new berkeleyDataSet.bsi_postingDataTable();
                        lda.FillById(lpostingT,lpost.postingID);
                        if (lpostingT.Rows.Count > 0)
                        {
                            berkeleyDataSet.bsi_postingRow lposting = lpostingT[0];

                            lix.Brand = lposting.brand;
                            lix.Category = lposting.category;
                            lix.Color = lposting.color;
                            lix.Condition = lposting.condition;
                            // lix.Cost = lposting.cost;
                            lix.StartDate = lpost.startDate;
                            lix.EndDate = lpost.endDate;
                            lix.FullDescription = lposting.fullDescription;
                            lix.Gender = lposting.gender;
                            lix.HeelHeight = lposting.heelHeight;

                            //lix.ItemLookupCode =  ; <================ DEPENDS On QUANTIY!!!
                            lix.Keywords = lposting.keywords;
                            lix.MarketPlaces = (uint)lposting.marketplaces;
                            lix.Material = lposting.material;
                            // lix.MSRP = decimal.Parse(lposting.MSRP);
                            lix.Ok2Publish = true;
                            
                            String[] lpictures = lposting.pictures.Split(new string[] { " | " }, StringSplitOptions.RemoveEmptyEntries);
                            // Let's sort the pictures according 
                            if (currentMarketPlace.pictureNo <= lpictures.Length & currentMarketPlace.pictureNo != 1)
                            {
                                // Interchange
                                String laux = lpictures[currentMarketPlace.pictureNo - 1];
                                lpictures[currentMarketPlace.pictureNo - 1] = lpictures[0];
                                lpictures[0] = laux;
                            }
                            lix.Pictures.AddRange(lpictures);
                            lix.URLPictures.AddRange(lpictures);

                            lix.SellingFormat = lpost.sellingFormat;

                            lix.Title = lpost.title;

                            lix.Shade = lposting.shade;
                            //lix.Size = lposting.size;
                            lix.SKU = lposting.sku;
                            lix.Style = lposting.style;
                            lix.Width = lposting.width;

                            // lix.Type = lposting.type
                            // Let's see if this is a variation item or a simple one
                            lix.Variation = lposting.variationType;
                            lix.Widths = lposting.variationDimensions;

                            lix.post_id = lpost.id;

                            berkeleyDataSet.bsi_quantitiesDataTable lqtysT = new berkeleyDataSet.bsi_quantitiesDataTable();
                            berkeleyDataSet.bsi_quantitiesRow lqtsrow = null;

                            lqtys_da.FillByPostId(lqtysT, lpost.id);

                            /*
                            berkeleyDataSet.bsi_qsandpricesDataTable lqsT = new berkeleyDataSet.bsi_qsandpricesDataTable();
                            berkeleyDataSet.bsi_qsandpricesRow lqs = null;
                            
                            String lmkt2diminish = "";
                            Decimal lprice = 0;
                            */

                            //if (lix.Variation == VARIATIONS_NONE)
                            if ( lqtysT.Rows.Count == 1 )
                            {
                                // Set the quantity, size, price and ILCto publish by reading the qs based on posting
                                if (lqtysT.Rows.Count > 0)
                                {
                                    lqtsrow = lqtysT[0];
                                    lix.Size = lqtsrow.size;
                                    lix.ItemLookupCode = lqtsrow.itemLookupCode;
                                    lix.SinglePrice = lqtsrow.price;
                                }
                                else
                                    lix.SinglePrice = 0;

                                if (lix.SellingFormat.Contains('A'))
                                    lix.SingleQuantity = 1; // All auctions will be set to 1
                                else
                                    lix.SingleQuantity = lqtsrow.quantity;

                                lix.Title = lqtsrow.title;
                            }
                            else
                            {
                                // Check all the variations and their quantities
                                foreach (berkeleyDataSet.bsi_quantitiesRow lqr in lqtysT)
                                {
                                    ItemExcel lvari = new ItemExcel();

                                    lvari.SKU = lposting.sku;
                                    lvari.ItemLookupCode = lqr.itemLookupCode;
                                    lvari.Title = lqr.title;
                                    lvari.Size = lqr.size;
                                    lvari.Width = lqr.width;
                                    lvari.SinglePrice = lqr.price;
                                    lvari.SingleQuantity = lqr.quantity;

                                    if ( lqr.quantity > 0 ) lix.Items.Add(lvari);
                                } // foreach (berkeleyDataSet.bsi_quantitiesRow lqr in lqtysT)

                                lix.Items.Sort(sortBySize);
                            } // if (lix.Variation == VARIATIONS_NONE)

                            theProducts.Add(lix);
                        } // if (lpostingT.Rows.Count > 0)

                    } // foreach (berkeleyDataSet.bsi_postsRow lpost in lposts.Rows)

                    int ltotalItems = 0, ltotalProducts = 0;
                    foreach (ItemExcel lix in theProducts)
                    {
                        // if ((lix.MarketPlaces & currentMarketPlace.maskId) == 0) continue; // Skip items that do not belong to this marketplace
                        if (!lix.Ok2Publish) continue;

                        ltotalProducts++;

                        // txtStatus.Text = "Publishing " + lix.Title + " [" + lix.ItemLookupCode + ( (isTheProductOnEbay(lix))? " >>REPEATED<< " : "" ) + "]\r\n" + txtStatus.Text  ;
                        txtStatus.Text = "Publishing " + lix.Title + " [" + lix.ItemLookupCode +  "]\r\n" + txtStatus.Text;
                        txtStatus.Update();

                        ItemType lproduct;

                        lproduct = BuildItem(lix);

                        if (lix.Items.Count == 0) // ONLY set price and QTY for individual products, not for Parents with children
                        {
                            lproduct.SKU = lix.ItemLookupCode; // We need the ILC as ID of the product
                            lproduct.Quantity = lix.SingleQuantity;
                            lproduct.StartPrice.Value = (double)lix.SinglePrice;
                            ltotalItems++;
                        }
                        else ltotalItems += lix.Items.Count;

                        try
                        {
                            FeeTypeCollection fees=null;
                            txtStatus.Text = " ...API call started..." + txtStatus.Text;
                            txtStatus.Update();

                            // Choose the correct API call. AddItemCall works for auctions and for single items with best offer
                            if (!DEBUG_MODE)
                            {
                                if (lix.SellingFormat == "A" || lix.Items.Count == 0)
                                {
                                    AddItemCall api_AUCTION_Call = new AddItemCall(apiContext);
                                    fees = api_AUCTION_Call.AddItem(lproduct);
                                }
                                else
                                {
                                    try
                                    {
                                        lproduct.BestOfferEnabled = false;
                                        lproduct.BestOfferDetails = null;
                                        AddFixedPriceItemCall api_FP_Call = new AddFixedPriceItemCall(apiContext);
                                        fees = api_FP_Call.AddFixedPriceItem(lproduct);
                                        int lfeesN = fees.Count;
                                    }
                                    catch (Exception pe)
                                    {
                                        MessageBox.Show(pe.ToString());
                                    }
                                }

                                double listingFee = 0.0;
                                foreach (FeeType fee in fees)
                                {
                                    if (fee.Name == "ListingFee")
                                    {
                                        listingFee = fee.Fee.Value;
                                    }
                                }
                            }
                            
                            // txtStatus.Text = "Listing fee is: " + listingFee + "\r\n" + txtStatus.Text;
                            int lauxInt = (lix.Items.Count > 0) ? lix.Items.Count : 1;
                            txtStatus.Text = "\r\nThe product was listed successfully! (" + lauxInt.ToString() + 
                                             " item(s)) Item ID:" + lproduct.ItemID + " | " + txtStatus.Text + " ";
                            txtStatus.Update();
                            
                            // We need to update the post information AND diminish the quantities
                            String lcommandStr = "UPDATE bsi_posts SET status=" + POSTING_STATUS_ACTIVE.ToString() +
                                                 ",markerplaceItemID=" + lproduct.ItemID +
                                                 ",endDate='" + lproduct.ListingDetails.EndTime +
                                                 "' WHERE id=" + lix.post_id;
                            SqlCommand lc = new SqlCommand(lcommandStr , lconn);
                            lc.ExecuteNonQuery();

                            /*
                            * WE NO LONGER DECREASE THE QUANTITIES OF THE ITEMS WE'LL DO THIS ONCE THE ITEM SELLS!!
                            lcommandStr = "UPDATE bsi_quantities SET quantity=quantity-1 WHERE postId=" + lix.post_id;
                            lc = new SqlCommand(lcommandStr, lconn);
                            lc.ExecuteNonQuery();
                            */
                        }
                        catch (Exception pe)
                        {
                            txtStatus.Text = pe.ToString() + "\r\n" + txtStatus.Text;
                            if (MessageBox.Show("ERROR WHILE PUBLISHING ITEM:\r\n\r\n\r\n" + pe.ToString() +
                                                "\r\n\r\nDO YOU WANT TO STOP THE PROCESS?\r\n\r\n",
                                                "Error", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                lstop = true;
                                break;
                            };
                        };

                    } // foreach (ItemExcel in lproducts)

                    txtStatus.Text = "Total products: " + ltotalProducts + " with " +
                                                             ltotalItems + " items \r\n" + txtStatus.Text;

                    lgrandTotalProducts += ltotalProducts;
                    lgrandTotalItems += ltotalItems;

                } // foreach (uint lmarket in mktPlaces)

            }
            catch (Exception pe)
            {
                MessageBox.Show(pe.ToString(), "Error while publishing products");
            }
            finally
            {
                if (lconn != null) lconn.Close();
            }

            txtStatus.Text = "GRAND TOTAL PRODUCTS: " + lgrandTotalProducts + " WITH " +
                              lgrandTotalItems + " ITEMS \r\n\r\n" + txtStatus.Text;

            MessageBox.Show("PROCESS ENDED");

        } // publishProducts

        VariationsType createVariations(ItemExcel pi, out int outVariations, out string outWidths)
        {
            VariationsType lvt = new VariationsType();
            lvt.Variation = new VariationTypeCollection();
            String lwidths = "";

            pi.Items.Sort(sortBySize);

            outVariations = 0;

            // Create the content of specifics for variations
            NameValueListTypeCollection TheVariations = new NameValueListTypeCollection(); // At least one variation will be by size. Maybe by width
            NameValueListType lvarBySize = new NameValueListType();
            NameValueListType lvarByWidth = new NameValueListType();

            lvarByWidth.Name = "Width";
            lvarByWidth.Value = new StringCollection();

            lvarBySize.Name = getEbaySizeName(pi.Gender);
            lvarBySize.Value = new StringCollection();

            // First: Let's see what sizes and widths are available and create their respective lists
            // Situations: Available sizes: 7M, 7W, 7XW, 7.5M, 8M, 8XW... same sizes, diff widths, avoid repeat sizes and widths
            String lcurrWidth = convertWidth(pi.Gender, pi.Items[0].Width); // Let's assume all the items are the same width than the first one
            lwidths = lcurrWidth;
            foreach (ItemExcel li in pi.Items)
            {
                String lnewWidth = convertWidth(pi.Gender, li.Width);
                if (!lvarBySize.Value.Contains(li.Size)) // Add the size only if it is not already here
                    lvarBySize.Value.Add(li.Size);
                if (lcurrWidth.CompareTo(lnewWidth) != 0) // New width
                {
                    lwidths += "," + lnewWidth;
                    // Create a new width if necessary
                    if (lvarByWidth.Value.Count == 0) // If this is the first change in width, then we need to add the original width
                        lvarByWidth.Value.Add(lcurrWidth);
                    if (!lvarByWidth.Value.Contains(lnewWidth))
                        lvarByWidth.Value.Add(lnewWidth);
                    lcurrWidth = lnewWidth;
                };

                // We cannot add here the individual product variation because we'll only know if we need
                // to specify width for each item until we look for width-changes with all of the items.
                // Unfortunately we need to create a separate loop to create the individual items.
            } // foreach

            TheVariations.Add(lvarBySize);
            outVariations |= VARIATIONS_SIZE;
            if (lvarByWidth.Value.Count > 0)
            {
                TheVariations.Add(lvarByWidth);
                outVariations |= VARIATIONS_WIDTH;
            }

            lvt.VariationSpecificsSet = TheVariations;

            // Now create the particular variation of each item
            foreach (ItemExcel li in pi.Items)
            {
                VariationType TheVariation = new VariationType(); // Wrapper for all the variations of this item

                NameValueListTypeCollection itemVariation = new NameValueListTypeCollection(); // All the ways this individual item varies

                // These are going to be how the item varies
                NameValueListType variationBySize = new NameValueListType(); // Size is obligatory.
                variationBySize.Name = getEbaySizeName(pi.Gender);
                variationBySize.Value = new StringCollection();
                variationBySize.Value.Add(li.Size);
                itemVariation.Add(variationBySize);

                if (lvarByWidth.Value.Count > 0)
                {
                    NameValueListType variationByWidth = new NameValueListType();
                    variationByWidth.Name = "Width";
                    variationByWidth.Value = new StringCollection();
                    variationByWidth.Value.Add(convertWidth(pi.Gender, li.Width));
                    itemVariation.Add(variationByWidth);
                }

                AmountType price = new AmountType();
                price.currencyID = CurrencyCodeType.USD;

                // TO PUBLISH JUST ONE JUST UNCOMMENT THE FOLLOWING LINE, OTHERWISE IT WILL POST EVERYTHING
                // TheVariation.Quantity = 1;
                TheVariation.Quantity = li.SingleQuantity;
                price.Value = (double)li.SinglePrice;
                /*
                switch ((uint)currentMarketPlace.maskId)
                {
                    case ItemMarketplace.MARKETPLACE_EBAY_MECALZO:
                        TheVariation.Quantity = li.QtyE1;
                        price.Value = (double)li.PriceE1;
                        break;

                    case ItemMarketplace.MARKETPLACE_EBAY_1MS:
                        TheVariation.Quantity = li.QtyE2;
                        price.Value = (double)li.PriceE2;
                        break;
                } // switch
                */
                TheVariation.StartPrice = price;

                // Set variation title and SKU
                TheVariation.VariationTitle = li.Title;
                TheVariation.SKU = li.ItemLookupCode;

                TheVariation.VariationSpecifics = itemVariation;

                lvt.Variation.Add(TheVariation);
            } // foreach

            outWidths = lwidths;
            return lvt;
        } // createVariations

        private ItemType BuildItem(ItemExcel excelItem)
        {
            ItemType item = new ItemType();

            // item title
            item.Title = excelItem.Title;
            // item description
            item.Description = _descriptionHeader + excelItem.Title + " " + excelItem.FullDescription + _descriptionFooter;
            item.SKU = excelItem.SKU;

            // Create the picture, save the URL and then pass it to the item
            item.PictureDetails = new PictureDetailsType();
            item.PictureDetails.PhotoDisplay = PhotoDisplayCodeType.PicturePack;
            item.PictureDetails.GalleryType = GalleryTypeCodeType.Gallery;
            item.PictureDetails.PictureURL = new StringCollection();
            foreach (String lpic in excelItem.URLPictures)
                item.PictureDetails.PictureURL.Add(lpic);

            // listing type
            BestOfferDetailsType lbo = null;
            switch (excelItem.SellingFormat)
            {
                case "A":
                    item.ListingType = ListingTypeCodeType.Chinese;
                    item.ListingDuration = "Days_7";
                    break;
                case "A1":
                    item.ListingType = ListingTypeCodeType.Chinese;
                    item.ListingDuration = "Days_1";
                    break;

                case "A3":
                    item.ListingType = ListingTypeCodeType.Chinese;
                    item.ListingDuration = "Days_3";
                    break;
                case "A5":
                    item.ListingType = ListingTypeCodeType.Chinese;
                    item.ListingDuration = "Days_5";
                    break;
                case "BIN":
                    item.ListingType = ListingTypeCodeType.FixedPriceItem;
                    item.ListingDuration = "Days_30";
                    lbo = new BestOfferDetailsType();
                    lbo.BestOfferEnabled = true;
                    item.BestOfferDetails = lbo;
                    item.BestOfferEnabled = true;
                    break;
                case "GTC":
                    item.ListingType = ListingTypeCodeType.FixedPriceItem;
                    item.ListingDuration = "GTC";
                    lbo = new BestOfferDetailsType();
                    lbo.BestOfferEnabled = true;
                    item.BestOfferDetails = lbo;
                    item.BestOfferEnabled = true;
                    break;
            }; // switch

            // Start time if specified
            if (lix.StartDate > DateTime.Now) item.ScheduleTime = lix.StartDate;

            item.HitCounter = HitCounterCodeType.BasicStyle;

            // item condition, New=1000, New without box=1500, New with defects=1750, Pre-owned=3000
            switch (excelItem.Condition)
            {
                case "NEW": item.ConditionID = 1000; break;
                case "NWB": item.ConditionID = 1500; break;
                case "NWD": item.ConditionID = 1750; break;
                case "PRE": item.ConditionID = 3000; break;
            }; // switch 

            // Item specifics
            item.ItemSpecifics = new NameValueListTypeCollection();
            NameValueListType litemspec = null;

            // Do not specify size nor width for products with variation. Each variation has its own specifics
            // Also, do not state size/width for watches
            if (excelItem.Items.Count == 0 && excelItem.Category != "31387" && excelItem.Category != "63852")
            {
                litemspec = new NameValueListType();
                litemspec.Name = getEbaySizeName(excelItem.Gender);
                litemspec.Value = new StringCollection(new String[] { excelItem.Size });
                item.ItemSpecifics.Add(litemspec);

                litemspec = new NameValueListType();
                litemspec.Name = "Width";
                String lwidth = convertWidth(excelItem.Gender, excelItem.Width);
                litemspec.Value = new StringCollection(new String[] { lwidth });
                item.ItemSpecifics.Add(litemspec);
            }

            int ebayCategory = int.Parse(excelItem.Category);

            litemspec = new NameValueListType();
            litemspec.Name = "Brand";
            litemspec.Value = new StringCollection(new String[] { excelItem.Brand });
            item.ItemSpecifics.Add(litemspec);

            litemspec = new NameValueListType();
            litemspec.Name = "Style";
            litemspec.Value = new StringCollection(new String[] { excelItem.Style });
            item.ItemSpecifics.Add(litemspec);

            if (!String.IsNullOrEmpty(excelItem.Color))
            {
                litemspec = new NameValueListType();
                litemspec.Name = "Color";
                litemspec.Value = new StringCollection(new String[] { excelItem.Color });
                item.ItemSpecifics.Add(litemspec);
            }

            if (!String.IsNullOrEmpty(excelItem.Material))
            {
                litemspec = new NameValueListType();

                if (ebayCategory != 31387)
                {
                    litemspec.Name = "Material";
                }
                else
                {
                    litemspec.Name = "Band Material";
                }

                litemspec.Value = new StringCollection(new String[] { excelItem.Material });
                item.ItemSpecifics.Add(litemspec);
            }

            if (!String.IsNullOrEmpty(excelItem.Shade) && ebayCategory != 63852)
            {
                litemspec = new NameValueListType();
                litemspec.Name = "Shade";
                litemspec.Value = new StringCollection(new String[] { excelItem.Shade });
                item.ItemSpecifics.Add(litemspec);
            }

            // listing price
            item.Currency = CurrencyCodeType.USD;

            if (excelItem.Items.Count == 0) // Do not set price or quantity for products with children
            {
                item.StartPrice = new AmountType();
                item.StartPrice.currencyID = CurrencyCodeType.USD;

                // item quantity
                item.Quantity = 1; // It will be overriden later, after the product creation
            }


            // item location and country
            item.Location = "Very near to you!";
            item.Country = CountryCodeType.US;

            // listing category
            CategoryType category = new CategoryType();
            category.CategoryID = ebayCategory.ToString(); // Primary Category
            item.PrimaryCategory = category;

            // Payment methods
            item.PaymentMethods = new BuyerPaymentMethodCodeTypeCollection();
            item.PaymentMethods.AddRange(
                new BuyerPaymentMethodCodeType[] { BuyerPaymentMethodCodeType.PayPal }
                );
            // email is required if paypal is used as payment method
            item.PayPalEmailAddress = currentMarketPlace.eBayPayPalAccount;

            // item specifics
            // item.ItemSpecifics = buildItemSpecifics();

            // handling time is required
            item.DispatchTimeMax = 1;

            // return policy
            item.ReturnPolicy = new ReturnPolicyType();
            item.ReturnPolicy.ReturnsAcceptedOption = "ReturnsAccepted";
            item.ReturnPolicy.ReturnsWithinOption = "Days_30";
            item.ReturnPolicy.ShippingCostPaidByOption = "Buyer";
            item.ReturnPolicy.Description = currentMarketPlace.ReturnsPolicies;

            // Create item variations if necessary
            if (excelItem.Items.Count > 0)
            {
                int pvariations = 0;
                String lwidthsList = "";
                item.Variations = createVariations(excelItem, out pvariations, out lwidthsList);
                excelItem.Variation = pvariations;
                excelItem.Widths = lwidthsList;

                // Let's see what variations were not set to set them in default
                if ((pvariations & VARIATIONS_WIDTH) == 0)
                {
                    // There were sizes but not widths, then set the general width
                    litemspec = new NameValueListType();
                    litemspec.Name = "Width";
                    String lwidth = convertWidth(excelItem.Gender, excelItem.Width);
                    litemspec.Value = new StringCollection(new String[] { lwidth });
                    item.ItemSpecifics.Add(litemspec);
                }
            }

            // shipping details
            item.ShippingDetails = BuildShippingDetails();

            return item;
        } 

        ShippingDetailsType BuildShippingDetails()
        {
            AmountType amount;

            // Shipping details
            ShippingDetailsType sd = new ShippingDetailsType();
            sd.ShippingServiceOptions = new ShippingServiceOptionsTypeCollection();

            sd.PaymentInstructions = "";

            sd.ShippingType = ShippingTypeCodeType.Flat; // All options will be flat

            // Let's create the domestic ground 
            ShippingServiceOptionsType shippingOptions = new ShippingServiceOptionsType();
            shippingOptions.ShippingServicePriority = 1; // First one to be displayed
            // shippingOptions.ShippingService = ShippingServiceCodeType.ShippingMethodStandard.ToString();
            shippingOptions.ShippingService = ShippingServiceCodeType.UPSGround.ToString();
            // shippingOptions.ShippingService = ShippingServiceCodeType.FedExHomeDelivery.ToString();
            if (double.Parse(currentMarketPlace.shippingDomesticStandard) == 0)
                shippingOptions.FreeShipping = true; // Each additional will be 0 so shippingOptions.ShippingServiceAdditionalCost is default 0
            else
            {
                amount = new AmountType();
                amount.currencyID = CurrencyCodeType.USD;
                amount.Value = double.Parse(currentMarketPlace.shippingDomesticStandard);
                shippingOptions.ShippingServiceCost = amount;
                amount = new AmountType();
                amount.currencyID = CurrencyCodeType.USD;
                amount.Value = double.Parse(currentMarketPlace.shippingDomesticStandardAdd);
                shippingOptions.ShippingServiceAdditionalCost = amount;

                // Add surcharge
                amount = new AmountType();
                amount.currencyID = CurrencyCodeType.USD;
                amount.Value = SHIPPING_SURCHARGE;
                shippingOptions.ShippingSurcharge = amount;
            }
            sd.ShippingServiceOptions.Add(shippingOptions); // Add to the list of shipping options

            // Now create the domestic next day
            shippingOptions = new ShippingServiceOptionsType();
            shippingOptions.ShippingServicePriority = 2; // Second to be displayed
            shippingOptions.ShippingService = ShippingServiceCodeType.ShippingMethodOvernight.ToString();
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = double.Parse(currentMarketPlace.shippingDomesticNextDay);
            shippingOptions.ShippingServiceCost = amount;
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = double.Parse(currentMarketPlace.shippingDomesticNextDayAdd);
            shippingOptions.ShippingServiceAdditionalCost = amount;
            sd.ShippingServiceOptions.Add(shippingOptions); // Add to the list of shipping options


            // Time to add the international shipping options
            InternationalShippingServiceOptionsType internationalShippingOptions;
            sd.InternationalShippingServiceOption = new InternationalShippingServiceOptionsTypeCollection();

            // First to Canada
            internationalShippingOptions = new InternationalShippingServiceOptionsType();
            internationalShippingOptions.ShippingServicePriority = 1; // First to be shown
            internationalShippingOptions.ShippingService = ShippingServiceCodeType.USPSPriorityMailInternational.ToString();
            internationalShippingOptions.ShipToLocation = new StringCollection();
            internationalShippingOptions.ShipToLocation.Add(CountryCodeType.CA.ToString()); // An specific country
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = double.Parse(currentMarketPlace.shippingCanadaPriority);
            internationalShippingOptions.ShippingServiceCost = amount;
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = double.Parse(currentMarketPlace.shippingCanadaPriorityAdd);
            internationalShippingOptions.ShippingServiceAdditionalCost = amount;

            sd.InternationalShippingServiceOption.Add(internationalShippingOptions);

            // Second Worldwide
            internationalShippingOptions = new InternationalShippingServiceOptionsType();
            internationalShippingOptions.ShippingServicePriority = 2; // Second to be shown
            internationalShippingOptions.ShippingService = ShippingServiceCodeType.USPSPriorityMailInternational.ToString();
            internationalShippingOptions.ShipToLocation = new StringCollection();
            internationalShippingOptions.ShipToLocation.Add(ShippingRegionCodeType.Worldwide.ToString()); // A region
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = double.Parse(currentMarketPlace.shippingInternationalPriority);
            internationalShippingOptions.ShippingServiceCost = amount;
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = double.Parse(currentMarketPlace.shippingInternationalPriorityAdd);
            internationalShippingOptions.ShippingServiceAdditionalCost = amount;
            sd.InternationalShippingServiceOption.Add(internationalShippingOptions);

            // UPS Worldwide Express option
            internationalShippingOptions = new InternationalShippingServiceOptionsType();
            internationalShippingOptions.ShippingServicePriority = 3; // Third to be shown
            internationalShippingOptions.ShippingService = ShippingServiceCodeType.UPSWorldWideExpress.ToString();
            internationalShippingOptions.ShipToLocation = new StringCollection();
            internationalShippingOptions.ShipToLocation.Add(ShippingRegionCodeType.Worldwide.ToString()); // A region
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = 60.0;
            internationalShippingOptions.ShippingServiceCost = amount;
            amount = new AmountType();
            amount.currencyID = CurrencyCodeType.USD;
            amount.Value = 10.00;
            internationalShippingOptions.ShippingServiceAdditionalCost = amount;
            sd.InternationalShippingServiceOption.Add(internationalShippingOptions);

            return sd;
        } // BuildShippingDetails

        private void Form1_Load(object sender, EventArgs e)
        {
            dtPickerStart.Value = DateTime.Now;
            cmbMarketplace.SelectedIndex = 0;

            // Read and store all the marketplaces
            SqlConnection lconn = null;
            berkeleyDataSetTableAdapters.bsi_marketplacesTableAdapter lda;

            try
            {
                lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString());
                lconn.Open();
                lda = new berkeleyDataSetTableAdapters.bsi_marketplacesTableAdapter();
                lda.Connection = lconn;
                lda.Fill(ldsMarkets);

                // Fill the cmbMarketplace
                cmbMarketplace.DataSource = ldsMarkets;
                cmbMarketplace.DisplayMember = "name";
                cmbMarketplace.ValueMember = "id";


                SqlCommand command = new SqlCommand("SELECT purchaseOrder FROM bsi_posts WHERE status = 10 AND marketplace <> 512");
                command.Connection = lconn;
                SqlDataReader dr = command.ExecuteReader();


                List<string> pos = new List<string>();

                while (dr.Read())
                {
                    pos.Add(dr.GetString(0));
                }

                cbSelectPO.DataSource = pos.Distinct().ToList();

                dr.Close();
            }
            catch (Exception pe)
            {
                MessageBox.Show("\nERROR WHILE READING MARKETPLACES: " + pe.ToString() + "\n", " Error on Load ");
            }
            finally
            {
                if (lconn != null)
                {
                    lconn.Close();
                };
            }
        } // Form1_Load

        private void readMarketplace()
        {
            // int lid = (int)cmbMarketplace.SelectedValue;

            try
            {
                String lresponse = "\r\n" + txtStatus.Text;
                // currentMarketPlace = ldsMarkets[cmbMarketplace.SelectedIndex];
                // currentMarketPlace = ldsMarkets[pmarket];
                GetApiContext();

                GetSellerListRequestType request = new GetSellerListRequestType();

                request.EndTimeFromSpecified = true;
                request.EndTimeFrom = DateTime.Now;
                request.EndTimeTo = DateTime.Now.AddDays(30);
                request.GranularityLevel = GranularityLevelCodeType.Fine;

                request.Pagination = new PaginationType();
                request.Pagination.EntriesPerPage = 200;

                request.IncludeVariationsSpecified = true;
                request.IncludeVariations = true;

                /*
                StringCollection lskus = new StringCollection();
                lskus.AddRange(txtItemID.Text.Split(new char[] { ',' }));
                request.SKUArray = lskus;
                */

                GetSellerListCall call = new GetSellerListCall(apiContext);
                itemsOnline = new List<ItemType>();
                int lpage = 1;

                try
                {
                    int totalPages = 0;
                    do
                    {
                        request.Pagination.PageNumber = lpage;
                        GetSellerListResponseType response = call.ExecuteRequest(request) as GetSellerListResponseType;
                        totalPages = response.PaginationResult.TotalNumberOfPages;
                        itemsOnline.AddRange(response.ItemArray.ToArray());
                        txtStatus.Text = "Reading page: " + lpage + "\r\n" + txtStatus.Text;
                        txtStatus.Update();
                        Application.DoEvents();
                        ++lpage;
                    } while (lpage <= totalPages);
                }
                catch (Exception pe)
                {
                    MessageBox.Show("Error: " + pe.ToString());
                }

                //txtStatus.Text = lresponse;
                MessageBox.Show("\r\nThe process ended reading marketpalce contents!\r\n");
            }
            catch (Exception pe)
            {
                MessageBox.Show(pe.ToString());
            }
        } // readMarketplace()

    } // class Form1
} // namespace Publisher_Test

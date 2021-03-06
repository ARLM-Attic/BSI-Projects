﻿//
// A la Mayor Gloria a Dios
//


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Configuration;

using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;

using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

using eBay.Service.Call;
using eBay.Service.Core.Sdk;
using eBay.Service.Core.Soap;
using eBay.Service.Util;


namespace BSI_InventoryPreProcessor
{
    public partial class Form1 : Form
    {
        public bool DEBUG_MODE = false;
        public const int POSTING_STATUS_ACTIVE   = 0;
        public const int POSTING_STATUS_READY2PUBLISH = 10;
        public const int POSTING_STATUS_BLOCKED  = 100;

        public const int QUANTITY_RECORD_TYPE_POSTING = 0;
        public const int QUANTITY_RECORD_TYPE_VARIATION = 10;

        public const int VARIATIONS_NONE = 0;
        public const int VARIATIONS_SIZE = 1;
        public const int VARIATIONS_WIDTH = 2;
        public const int VARIATIONS_COLOR = 4;

        public static string EXCEL_COLUMN_INITIAL = "A";
        public static string EXCEL_COLUMN_FINAL = "AQ";

        public static int EXCEL_INTCOLUMN_PO = 1;
        public static int EXCEL_INTCOLUMN_LISTUSER = 2;

        public static int EXCEL_INTCOLUMN_BRAND = 3;
        public static int EXCEL_INTCOLUMN_SKU = 4;
        public static int EXCEL_INTCOLUMN_LOOKUPCODE = 5;
        public static int EXCEL_INTCOLUMN_SIZE = 6;
        public static int EXCEL_INTCOLUMN_WIDTH = 7;
        public static int EXCEL_INTCOLUMN_CONDITION = 8;
        public static int EXCEL_INTCOLUMN_CATEGORY = 9;
        public static int EXCEL_INTCOLUMN_STYLE = 10;
        public static int EXCEL_INTCOLUMN_TITLE = 11;
        public static int EXCEL_INTCOLUMN_COUNT = 12;
        public static int EXCEL_INTCOLUMN_FULLD = 13;
        public static int EXCEL_INTCOLUMN_KEYWORDS = 14;
        public static int EXCEL_INTCOLUMN_MATERIAL = 15;
        public static int EXCEL_INTCOLUMN_COLOR = 16;
        public static int EXCEL_INTCOLUMN_SHADE = 17;
        public static int EXCEL_INTCOLUMN_HEEL = 18;
        public static int EXCEL_INTCOLUMN_RMSDESCRIPTION = 19;
        public static int EXCEL_INTCOLUMN_GENDER = 20;
        public static int EXCEL_INTCOLUMN_RECEIVED = 21;
        public static int EXCEL_INTCOLUMN_COST = 22;
        public static int EXCEL_INTCOLUMN_UPC = 23;

        // 2013-Jan-02: New posting sheet format with a single qty/price
        public static int EXCEL_INTCOLUMN_QUANTITY = 24;
        public static int EXCEL_INTCOLUMN_PRICE = 25;

        public static int EXCEL_INTCOLUMN_SELLINGFORMAT = 26; // 37;
        public static int EXCEL_INTCOLUMN_STARTDATE = 27; // 38;

        // Previous format with store info per store 2013-01-02
        public static int EXCEL_INTCOLUMN_MSRP = 19; 
        public static int EXCEL_INTCOLUMN_QTY_AMAZON = 25;
        public static int EXCEL_INTCOLUMN_QTY_HARVARD = 26;
        public static int EXCEL_INTCOLUMN_QTY_MECALZO = 27;
        public static int EXCEL_INTCOLUMN_QTY_1MS = 28;
        public static int EXCEL_INTCOLUMN_QTY_PAS = 29;
        public static int EXCEL_INTCOLUMN_QTY_SA = 30;

        public static int EXCEL_INTCOLUMN_PRICE_AMAZON = 31;
        public static int EXCEL_INTCOLUMN_PRICE_HARVARD = 32;
        public static int EXCEL_INTCOLUMN_PRICE_MECALZO = 33;
        public static int EXCEL_INTCOLUMN_PRICE_1MS = 34;
        public static int EXCEL_INTCOLUMN_PRICE_PAS = 35;
        public static int EXCEL_INTCOLUMN_PRICE_SA = 36;

        private static ApiContext apiContext = null;
        private string _descriptionHeader, _descriptionFooter, lorginalpathfile, lpicturespath;

        public static int EBAY_STARTINGINDEX = 2;
        public static int WEB_STARTINGINDEX = 6;

        private uint[] MarketPlaces = { ItemMarketplace.MARKETPLACE_AMAZON, 
                                        ItemMarketplace.MARKETPLACE_AMAZON_HARVARD, 
                                        
                                        ItemMarketplace.MARKETPLACE_EBAY_MECALZO,
                                        ItemMarketplace.MARKETPLACE_EBAY_1MS,
                                        ItemMarketplace.MARKETPLACE_EBAY_PAS,
                                        ItemMarketplace.MARKETPLACE_EBAY_SA,

                                        ItemMarketplace.MARKETPLACE_WEB_SF
                                      };

        Boolean lstop;

        int gCurrentMarketplace = 0;
        List<ItemExcel> theProducts;

        // Products on the marketplaces. Each element is a marketplace
        List<ItemType>[] itemsOnline = new List<ItemType>[12];

        berkeleyDataSet.bsi_marketplacesDataTable ldsMarkets = new berkeleyDataSet.bsi_marketplacesDataTable();
        berkeleyDataSet.bsi_marketplacesRow currentMarketPlace = null;

        public Form1()
        {
            InitializeComponent();            
        }  // public Form1()

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                lorginalpathfile = openFileDialog1.FileName;
            txtOriginalFile.Text = lorginalpathfile;
        } // btnSearch_Click

        private void btnSearchPictures_Click(object sender, EventArgs e)
        {
        } // btnSearchPictures_Click

        private string removeSize(string ltext)
        {
            StringBuilder lsb = new StringBuilder();

            if (ltext.ToLower().Contains(" size") || ltext.ToLower().Contains(" sz"))
            {
                int lszTextpos = -1;

                if (ltext.ToLower().Contains(" size"))
                {
                    lszTextpos = ltext.ToLower().IndexOf(" size");
                }
                else
                {
                    lszTextpos = ltext.ToLower().IndexOf(" sz");
                }

                if (lszTextpos >= 0)
                {
                    lsb.Append(ltext.Substring(0,lszTextpos));
                    
                    // Let's find the size or white space
                    int li = lszTextpos, lconsecutiveWidthChars=0;
                    bool lflag = false, lcheckingSizeW = false, lwidthfound = false;
                    while ( !lflag && li < ltext.Length )
                    {
                        if ( char.IsNumber(ltext[li]) ) lcheckingSizeW = true;
                        if ( char.IsWhiteSpace(ltext[li]) )
                        {
                            if ( lwidthfound )
                                lflag = true;
                        }
                        if ( char.IsLetter(ltext[li]) )
                        {
                            if ( lcheckingSizeW )
                            {
                                lwidthfound = true;
                                lconsecutiveWidthChars++;
                                if ( lconsecutiveWidthChars > 3 ) // Get out! This is now width!!
                                {
                                    li -= lconsecutiveWidthChars;
                                    lflag = true;
                                }
                            }
                        }
                        li++;
                    } // while
                    lsb.Append(' ');
                    lsb.Append(ltext.Substring(li));
                }
                else
                    lsb.Append(ltext);
            }
            else
            {
                lsb = new StringBuilder(ltext);
            }
            return lsb.ToString();
        } // removeSize

        private void updateMarketplaces()
        {
            // lmktindex starts in 1 cause 0=Amazon, 1=AMZ-HarvardStation, 2=Mecalzo, 3=1MS, 4=PickAShoe, 5=Shoefestival.com
            if (cmbMarkets.SelectedIndex >= EBAY_STARTINGINDEX && cmbMarkets.SelectedIndex < WEB_STARTINGINDEX )
            {
                itemsOnline[cmbMarkets.SelectedIndex] = new List<ItemType>();
                currentMarketPlace = ldsMarkets[cmbMarkets.SelectedIndex];
                GetApiContext();
                txtStatus.Text = "UPDATING " + currentMarketPlace.name + " CATALOG... THIS MIGHT TAKE A FEW MINUTES...\r\n" +
                                 "-------------------------------------------------------------------------------------------\r\n" +
                                 txtStatus.Text;
                try
                {
                    String lresponse = "\r\n" + txtStatus.Text;

                    GetSellerListRequestType request = new GetSellerListRequestType();

                    request.EndTimeFromSpecified = true;
                    request.EndTimeFrom = DateTime.Now;
                    request.EndTimeTo = DateTime.Now.AddDays(30);
                    request.GranularityLevel = GranularityLevelCodeType.Medium;
                    request.Pagination = new PaginationType();
                    request.Pagination.EntriesPerPage = Properties.Settings.Default.eBayPageSize;

                    request.IncludeVariationsSpecified = true;
                    request.IncludeVariations = true;

                    /*
                    StringCollection lskus = new StringCollection();
                    lskus.AddRange(txtItemID.Text.Split(new char[] { ',' }));
                    request.SKUArray = lskus;
                    */

                    GetSellerListCall call = new GetSellerListCall(apiContext);
                    int lpage = 1;

                    try
                    {
                        int totalPages = 0;
                        do
                        {
                            request.Pagination.PageNumber = lpage;
                            GetSellerListResponseType response = (GetSellerListResponseType)call.ExecuteRequest(request);
                            totalPages = response.PaginationResult.TotalNumberOfPages;
                            itemsOnline[cmbMarkets.SelectedIndex].AddRange(response.ItemArray.ToArray());
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
                    txtStatus.Update();
                    Application.DoEvents();
                }
                catch (Exception pe)
                {
                    MessageBox.Show(pe.ToString());
                }
                txtStatus.Update();
                Application.DoEvents();
            }
            else
            {
                MessageBox.Show("\r\nPLEASE SELECT A VALID eBay MARKETPLACE\r\n");
            } // if

            txtStatus.Text = "FINISHED READING... RESUMING CHECKING OF PRODUCTS...\r\n" +
                             txtStatus.Text;
        } // updateMarketplaces

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtOriginalFile.Text.Trim()))
            {
                MessageBox.Show("Please select the original Excel file with the inventory to add");
                txtOriginalFile.Focus();
                return;
            }
            
            if (String.IsNullOrEmpty(txtPicturesPath.Text.Trim()))
            {
                MessageBox.Show("Please specify the path where the pictures are stored");
                txtPicturesPath.Focus();
                return;
            }

            if (MessageBox.Show("ABOUT TO PUBLISH PRODUCTS FOR \r\n\r\n" + cmbMarkets.Text + "\r\n\r\nREADY TO PROCEED?", "Confirmation", MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
                return;

            // Everything is okay to go
            lorginalpathfile = txtOriginalFile.Text.Trim();
            lpicturespath = txtPicturesPath.Text.Trim();

            var excelApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook theWorkbook = excelApp.Workbooks.Open(lorginalpathfile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
            /*
            excelApp.Workbooks.Open(lorginalpathfile, 0, true, 5,
            "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
            "\t", false, false,0, true);
            */
            Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)theWorkbook.ActiveSheet; // (Excel.Worksheet)excelApp.ActiveSheet;

            btnStart.Enabled = false;
            ItemExcel lcurrentItem = null;
            int lcurrRow = 2, lmarketplace = 0;
            lstop = false;
            String lsku = "", lcurrentSku = "", lbrand = "";
            theProducts = new List<ItemExcel>();
            currentMarketPlace = ldsMarkets[cmbMarkets.SelectedIndex];

            using (SqlConnection lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString()))
            {
                try
                {
                    lconn.Open();
                    while (!lstop)
                    {
                        ItemExcel lix = new ItemExcel();
                        Decimal laux = 0; int lauxi = 0, lqtyreceived = 0;
                        String lauxS = "";

                        Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range(EXCEL_COLUMN_INITIAL + lcurrRow.ToString(),
                                                                EXCEL_COLUMN_FINAL + lcurrRow.ToString());
                        System.Array myvalues = (System.Array)range.Cells.Value;
                        lbrand = Convert.ToString(myvalues.GetValue(1, 1)); // "myvalues" is a 2-D array (no matter if the range was from one single row)

                        if (!String.IsNullOrEmpty(lbrand))
                        {
                            // Read the rest of the line

                            lsku = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_SKU));

                            // Read the individual information. This will be different for each item no matter it belongs to a father-children
                            lix.SKU = lsku;
                            lix.purchaseOrder = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PO));
                            lix.listUser = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_LISTUSER));

                            lix.ItemLookupCode = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_LOOKUPCODE));
                            lix.Title = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_TITLE));
                            lix.Size = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_SIZE));
                            lix.Width = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_WIDTH));

                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_RECEIVED)), out lauxi);
                            lqtyreceived = lauxi;
                    
                            int ltotal2publish = 0;
                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_QUANTITY)), out lauxi);
                            lix.Quantity = lauxi; ltotal2publish += lauxi;
                            Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PRICE)), out laux);
                            lix.Price = laux;

                            /* 2013-01-02
                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_QTY_AMAZON)), out lauxi);
                            lix.QtyA1 = lauxi; ltotal2publish += lauxi;
                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_QTY_HARVARD)), out lauxi);
                            lix.QtyA2 = lauxi; ltotal2publish += lauxi;
                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_QTY_MECALZO)), out lauxi);
                            lix.QtyE1 = lauxi; ltotal2publish += lauxi;
                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_QTY_1MS)), out lauxi);
                            lix.QtyE2 = lauxi; ltotal2publish += lauxi;
                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_QTY_PAS)), out lauxi);
                            lix.QtyE3 = lauxi; ltotal2publish += lauxi;
                            Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_QTY_SA)), out lauxi);
                            lix.QtyE4 = lauxi; ltotal2publish += lauxi;

                            Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PRICE_AMAZON)), out laux);
                            lix.PriceA1 = laux; if (laux > 0) lmarketplace = 1;
                            Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PRICE_HARVARD)), out laux);
                            lix.PriceA2 = laux; if (laux > 0) lmarketplace = 2;
                            Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PRICE_MECALZO)), out laux);
                            lix.PriceE1 = laux; if (laux > 0) lmarketplace = 16;
                            Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PRICE_1MS)), out laux);
                            lix.PriceE2 = laux; if (laux > 0) lmarketplace = 32;
                            Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PRICE_PAS)), out laux);
                            lix.PriceE3 = laux; if (laux > 0) lmarketplace = 64;
                            Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_PRICE_SA)), out laux);
                            lix.PriceE4 = laux; if (laux > 0) lmarketplace = 128;
                            */

                            lix.MarketPlaces = (uint)currentMarketPlace.maskId;

                            bool loverpublished = false;
                            if (ltotal2publish > lqtyreceived)
                            {
                                loverpublished = true;
                            }
                            else
                            {
                                /*
                                String lqs = "SELECT SUM(quantity) AS totalQ FROM bsi_quantities where purchaseorder='" + 
                                             lix.purchaseOrder + "' AND itemlookupcode='" + lix.ItemLookupCode + "'";
                                SqlCommand lc = new SqlCommand(lqs,lconn);
                                SqlDataReader lr = lc.ExecuteReader();
                                if ( lr.Read() )
                                {
                                    int lq = 0;
                                    String lqS = lr["totalQ"].ToString();
                                    Int32.TryParse(lqS, out lq);
                                    loverpublished = ( (lq + ltotal2publish) > lqtyreceived );
                                }
                                lr.Close();
                                */
                            }

                            if ( loverpublished )
                            {
                                MessageBox.Show("\r\n\r\n THE SKU:" + 
                                                lix.ItemLookupCode + 
                                                " IS TRYING TO PUBLISH MORE PRODUCTS THAN IT RECEIVED.\r\n" +
                                                "   CHECK THE LISTING, FIX THE PROBLEM AND TRY AGAIN. \r\n\r\n", 
                                                "Error in quantities");
                                return;
                            }
                        }
                        else
                            lsku = "";

                        if (lsku.CompareTo(lcurrentSku) != 0)
                        {
                            if (!String.IsNullOrEmpty(lbrand))
                            {
                                // Change of SKU
                                // if (lcurrentItem != null) lcurrentItem.checkMarketPlaces();
                                // if (lcurrentItem != null) lcurrentItem.MarketPlaces = (uint)currentMarketPlace.maskId; // Cambio 2013-01-02

                                lcurrentSku = lsku;
                                lcurrentItem = new ItemExcel(lix); // The new current item is the the one that's being read. Get its current attributes

                                lcurrentItem.SKU = lsku;
                                lcurrentItem.Alias = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_UPC)).Trim();
                                lcurrentItem.Brand = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_BRAND)).Trim();
                                lcurrentItem.Condition = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_CONDITION)).ToUpper().Trim();
                                if (String.IsNullOrEmpty(lcurrentItem.Condition))
                                    lcurrentItem.Condition = "NEW";

                                lcurrentItem.Category = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_CATEGORY)).Trim();
                                Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_COST)).Trim(), out laux); lcurrentItem.Cost = laux;

                                lcurrentItem.FullDescription = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_FULLD)).Trim();
                                lauxS = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_GENDER)).ToUpper().Trim();
                                if (lauxS == "MEN") lauxS = "MENS";
                                if (lauxS == "WOMEN") lauxS = "WOMENS";
                                if (lauxS == "JUNIORS") lauxS = "JUNIOR";
                                lcurrentItem.Gender = lauxS;

                                lcurrentItem.Keywords = properNameString(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_KEYWORDS))).Trim();

                                lcurrentItem.Material = properNameString(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_MATERIAL))).Trim();
                                lcurrentItem.Color = properNameString(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_COLOR))).Trim();
                                lcurrentItem.Shade = properNameString(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_SHADE))).Trim();
                                lcurrentItem.HeelHeight = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_HEEL)).Trim();

                                Decimal.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_MSRP)).Trim(), out laux);
                                lcurrentItem.MSRP = laux;

                                Int32.TryParse(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_RECEIVED)), out lauxi);
                                lcurrentItem.Received = lauxi;

                                lcurrentItem.RMS_Description = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_RMSDESCRIPTION)).Trim();
                                lcurrentItem.Size = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_SIZE));
                                lcurrentItem.Width = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_WIDTH));
                                lcurrentItem.Style = properNameString(Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_STYLE))).Trim();

                                lcurrentItem.SellingFormat = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_SELLINGFORMAT)).ToUpper().Trim();
                                if (String.IsNullOrEmpty(lcurrentItem.SellingFormat))
                                    lcurrentItem.SellingFormat = "GTC";

                                String ldateS = Convert.ToString(myvalues.GetValue(1, EXCEL_INTCOLUMN_STARTDATE));
                                lcurrentItem.StartDate = new DateTime(1900, 1, 1);
                                if (!String.IsNullOrEmpty(ldateS))
                                {
                                    lcurrentItem.StartDate = DateTime.Parse(ldateS);
                                    lcurrentItem.StartDate = lcurrentItem.StartDate.AddHours(7); // Add 7 hours to convert from PDT to GMT
                                };

                                DateTime lendDate = new DateTime();
                                switch (lcurrentItem.SellingFormat)
                                {
                                    case "A": lendDate = lcurrentItem.StartDate.AddDays(7); break;
                                    case "A1": lendDate = lcurrentItem.StartDate.AddDays(1); break;
                                    case "A3": lendDate = lcurrentItem.StartDate.AddDays(3); break;
                                    case "A5": lendDate = lcurrentItem.StartDate.AddDays(5); break;
                                    case "BIN": lendDate = lcurrentItem.StartDate.AddDays(30); break;
                                    case "GTC": lendDate = new DateTime(2020, 12, 25); break;
                                } // switch
                                lcurrentItem.EndDate = lendDate;
                                theProducts.Add(lcurrentItem);
                            }
                            else
                            {
                                lstop = true;
                                // if ( lcurrentItem.Items.Count > 0 ) theProducts.Add(lcurrentItem);
                            }
                        }
                        else
                        {
                            if (lcurrentItem != null)
                            {
                                lcurrentItem.Type = ItemExcel.ITEM_TYPE_PARENT;
                                if (lcurrentItem.Items.Count == 0) lcurrentItem.Items.Add(lcurrentItem); // Add itself as first item 
                                lcurrentItem.Items.Add(lix);
                            };
                        };
                        ++lcurrRow;
                    } // while
                }
                catch(Exception pe)
                {
                    MessageBox.Show(pe.ToString());
                }
            } // using


            //if (theProducts.Last<ItemExcel>().Items.Count > 0)
            //    theProducts.Last<ItemExcel>().checkMarketPlaces();

            theWorkbook.Close();
            excelApp.Quit();

            releaseObject(theWorkbook);
            releaseObject(excelApp);

            // Sort theProducts by brand
            theProducts.Sort(sortItems);

            // Now, it is time to look for the pictures
            DirectoryInfo ldi;
            FileInfo[] ldirEntries = new FileInfo[] { }; ;

            bool lflag = true;
            foreach (ItemExcel lxi in theProducts)
            {
                // Let's see how many pictures are available
                String lpath = this.txtPicturesPath.Text + '\\' + lxi.Brand;

                ldi = new DirectoryInfo(lpath);
                ldirEntries = ldi.GetFiles(lxi.SKU + "*.jpg");
                foreach (FileInfo lfi in ldirEntries)
                {
                    byte[] lfilecontents = File.ReadAllBytes(lfi.FullName);
                    if (lfilecontents.Length > 0) lxi.Pictures.Add(lfi.FullName);
                } // foreach

                if (lxi.Pictures.Count > 0 )
                {
                    lxi.Ok2Publish = true;
                    lxi.Pictures.Sort();
                }
                else
                {
                    lxi.Ok2Publish = false;
                    lflag = false;
                }

                txtStatus.Text = "Style " + lxi.ItemLookupCode + " has " + lxi.Pictures.Count + " pictures\r\n" + txtStatus.Text;
                txtStatus.Update();
            } // foreach

            if (!lflag && (cmbMarkets.SelectedIndex >= EBAY_STARTINGINDEX && cmbMarkets.SelectedIndex < WEB_STARTINGINDEX) )
            {
                // We cannot publish on eBay if there's one item without pics
                MessageBox.Show("\r\nAT LEAST ONE ITEM DOES NOT HAVE PICTURES. PLEASE CHECK THE LIST AND TRY AGAIN.\r\n" +
                                "You won't be able to publish on any eBay market without pictures.\r\n");
            }
            else
            {
                btnStartPublishing.Enabled = true;
                MessageBox.Show("\n\nPLEASE VERIFY THE STATUS OF THE INITIAL VERIFICATION AND PROCEED\nTO SAVE PRODUCTS IN OUR DATABASES IF EVERYTHING IS CORRECT\n\n", "PROCESSING FINISHED");
            }
        } // btnStart_Click

        private void btnStartPublishing_Click(object sender, EventArgs e)
        {
            btnStartPublishing.Enabled = false;
            if (cmbMarkets.SelectedIndex >= EBAY_STARTINGINDEX && cmbMarkets.SelectedIndex < WEB_STARTINGINDEX)
            {
                updateMarketplaces();
            }
            publishProducts();
            MessageBox.Show("Process ended with " + theProducts.Count + " products");
        } // btnStartPublishing_Click

        // ------------------------------------------ Service methods

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

        private void publishProducts()
        {
            // Let's publish products by marketplace
            uint[] mktPlaces = { 
                                   ItemMarketplace.MARKETPLACE_AMAZON, 
                                   ItemMarketplace.MARKETPLACE_AMAZON_HARVARD, 
                                   ItemMarketplace.MARKETPLACE_EBAY_MECALZO, 
                                   ItemMarketplace.MARKETPLACE_EBAY_1MS,
                                   ItemMarketplace.MARKETPLACE_EBAY_PAS,
                                   ItemMarketplace.MARKETPLACE_EBAY_SA,
                                   ItemMarketplace.MARKETPLACE_WEB_SF
                               };

            // SqlCeDataAdapter lda = new SqlCeDataAdapter("select * from marketplace", ConfigurationManager.ConnectionStrings["ProcessInventory.Properties.Settings.marketsConnectionString"].ToString());
            // marketsDataSet.marketplaceDataTable lmarketsTable = new marketsDataSet.marketplaceDataTable();
            // lda.Fill(lmarketsTable);
            // lda.Dispose();

            lstop = false;
            // foreach (uint lmarket in mktPlaces)

            SqlConnection lconn = null;
            berkeleyDataSetTableAdapters.bsi_postingTableAdapter lda;
            berkeleyDataSetTableAdapters.bsi_postsTableAdapter lposts_da;
            berkeleyDataSetTableAdapters.bsi_quantitiesTableAdapter lqtys_da;
            /*
            berkeleyDataSetTableAdapters.bsi_variationsTableAdapter lvar_da;
            berkeleyDataSetTableAdapters.bsi_qsandpricesTableAdapter lqs_da;
            */
            try
            {
                lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString());
                lconn.Open();

                lda = new berkeleyDataSetTableAdapters.bsi_postingTableAdapter();
                lposts_da = new berkeleyDataSetTableAdapters.bsi_postsTableAdapter();
                lqtys_da = new berkeleyDataSetTableAdapters.bsi_quantitiesTableAdapter();

                // lvar_da = new berkeleyDataSetTableAdapters.bsi_variationsTableAdapter();
                // lqs_da = new berkeleyDataSetTableAdapters.bsi_qsandpricesTableAdapter();
                lda.Connection = lconn;
                lposts_da.Connection = lconn;
                lqtys_da.Connection = lconn;

                // lvar_da.Connection = lconn;
                // lqs_da.Connection = lconn;

                // foreach (berkeleyDataSet.bsi_marketplacesRow lmarketPlace in ldsMarkets.Rows)
                {
                    // currentMarketPlace = (berkeleyDataSet.bsi_marketplacesRow)ldsMarkets.Rows[0]; // lmarketPlace;
                    currentMarketPlace = (berkeleyDataSet.bsi_marketplacesRow)ldsMarkets.Rows[cmbMarkets.SelectedIndex]; // lmarketPlace;
                    GetApiContext();

                    //if (lstop) break;
                    /*
                    // Get the marketplace information
                    //foreach (marketsDataSet.marketplaceRow lmarketPlace in lmarketsTable.Rows)
                
                    foreach (berkeleyDataSet.bsi_marketplacesRow lmarketPlace in ldsMarkets.Rows)
                        if (lmarketPlace.id == lmarket)
                            currentMarketPlace = lmarketPlace;
                
                    _descriptionHeader = System.IO.File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + "\\" + currentMarketPlace.templateHeader);
                    _descriptionFooter = System.IO.File.ReadAllText(System.IO.Directory.GetCurrentDirectory() + "\\" + currentMarketPlace.tempateFooter);
                     * 
                    String llocation = Assembly.GetExecutingAssembly().Location.Replace(Assembly.GetExecutingAssembly().ManifestModule.Name, "");
                    _descriptionHeader = System.IO.File.ReadAllText(llocation + currentMarketPlace.templateHeader);
                    _descriptionFooter = System.IO.File.ReadAllText(llocation + currentMarketPlace.tempateFooter);                
                    */

                    _descriptionHeader = currentMarketPlace.template_header;
                    _descriptionFooter = currentMarketPlace.template_footer;

                    //txtStatus.Text = "Publishing products for " + currentMarketPlace.name + "\r\n" + txtStatus.Text;
                    txtStatus.Update();
                    foreach (ItemExcel xlProduct in theProducts)
                    {
                        //if ((xlProduct.MarketPlaces & currentMarketPlace.maskId) == 0) continue; // Skip items that do not belong to this marketplace
                        if (!xlProduct.Ok2Publish && !this.chkPublishWOPics.Checked ) continue;

                        if (xlProduct.Items.Count > 1)
                            xlProduct.Title = removeSize(xlProduct.Title);

                        // Let's see if we have already posted this item
                        SqlCommand lc = new SqlCommand();
                        lc.Connection = lconn;
                        lc.Parameters.Add(new SqlParameter("ppo", xlProduct.purchaseOrder));
                        lc.Parameters.Add(new SqlParameter("psku", xlProduct.SKU));
                        lc.Parameters.Add(new SqlParameter("psf", xlProduct.SellingFormat));
                        lc.Parameters.Add(new SqlParameter("pmkt", (int)xlProduct.MarketPlaces));
                        
                        String lqs = null;
                        if (xlProduct.SellingFormat.Contains("A")) // If it is an auction we need to check the quantities table too
                        {
                            lqs = "SELECT theQ.postid, theQ.quantity, theQ.itemlookupcode, " +
                                  "thePost.id, thePost.purchaseorder, thePost.sku, thePost.sellingformat, thePost.marketplace " +
                                  "from bsi_quantities as theQ inner join bsi_posts as thePost " + 
                                  "on theQ.postid = thePost.id " +
                                  "where thePost.purchaseorder=@ppo AND thePost.sku=@psku AND " +
                                  "thePost.sellingFormat=@psf AND thePost.marketplace=@pmkt AND theQ.itemlookupcode='" + xlProduct.ItemLookupCode + "' ";
                            /*
                            lqs = "SELECT thePost.id, thePost.purchaseorder, thePost.sku, thePost.sellingformat, " + 
                                  "thePost.marketplace from bsi_posts AS thePost " + 
                                  "inner join bsi_quantities AS theQ on theQ.postid=thePost.id " + 
                                  "WHERE thePost.purchaseorder=@ppo AND thePost.sku=@psku AND " + 
                                  "thePost.sellingFormat=@psf AND thePost.marketplace=@pmkt";
                            */
                        }
                        else
                        {
                            lqs = "SELECT * FROM bsi_posts WHERE purchaseorder=@ppo AND " +
                                  "sku=@psku AND sellingFormat=@psf AND marketplace=@pmkt";
                        }

                        lc.CommandText = lqs;
                        SqlDataReader lr = lc.ExecuteReader();
                        bool lexists = lr.Read();
                        lr.Close();
                        lc.Cancel();
                        if ( lexists )
                        {
                            MessageBox.Show("\r\n\r\nYOU ARE TRYING TO RE-PUBLISH AN ITEM ("+ xlProduct.ItemLookupCode + " | " +
                                            xlProduct.SellingFormat +
                                            ") FROM THE SAME PURCHASE ORDER! THIS ITEM WAS PREVIOUSLY LISTED BY " + 
                                            xlProduct.listUser + ". THE PROGRAM WILL STOP AND END\r\n\r\n",
                                            "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }

                        // First, let's process the product to clean it and see if it is duplicated
                        if (cmbMarkets.SelectedIndex >= EBAY_STARTINGINDEX && cmbMarkets.SelectedIndex < WEB_STARTINGINDEX)
                            isTheProductOnEbay(xlProduct);
                        else
                        {
                            if (cmbMarkets.SelectedIndex < EBAY_STARTINGINDEX)
                            {
                                //isTheProductOnAmazon(xlProduct);
                            }
                            else
                                isTheProductOnWebsite(xlProduct);
                        }

                        txtStatus.Text = "Publishing " + xlProduct.Title + " [" + 
                                         xlProduct.ItemLookupCode + " | " + xlProduct.SellingFormat + 
                                         "]\r\n" + txtStatus.Text;
                        txtStatus.Update();

                        ItemType lproduct = lproduct = BuildItem(xlProduct);
                        
                        if (xlProduct.Items.Count == 0) // ONLY set price and QTY for individual products, not for Parents with children
                        {
                            // Set a price and Q temporal
                            lproduct.Quantity = 1;
                            lproduct.StartPrice.Value = 99.99;
                         
                            /*
                            lproduct.SKU = xlProduct.ItemLookupCode; // We need the ILC as ID of the product
                            switch ((uint)currentMarketPlace.maskId)
                            {
                                case ItemMarketplace.MARKETPLACE_AMAZON:
                                    lproduct.Quantity = xlProduct.QtyA1;
                                    lproduct.StartPrice.Value = (double)xlProduct.PriceA1; break;
                                case ItemMarketplace.MARKETPLACE_EBAY_1MS:
                                    lproduct.Quantity = xlProduct.QtyE2;
                                    lproduct.StartPrice.Value = (double)xlProduct.PriceE2; break;
                                case ItemMarketplace.MARKETPLACE_EBAY_MECALZO:
                                    lproduct.Quantity = xlProduct.QtyE1;
                                    lproduct.StartPrice.Value = (double)xlProduct.PriceE1; break;
                            } // swtich
                            */
                        }

                        try
                        {
                            FeeTypeCollection fees;
                            txtStatus.Text = " ...API call started..." + txtStatus.Text;
                            txtStatus.Update();

                            // Set one picture to the eBay product
                            lproduct.PictureDetails = new PictureDetailsType();
                            lproduct.PictureDetails.PictureURL = new StringCollection(new string[] { "http://www.tools4inet.com/0/products/tim/10061.jpg" });

                            // Choose the correct API call. AddItemCall works for auctions and for single items with best offer
                            if (!DEBUG_MODE && (currentMarketPlace.maskId > 8 && currentMarketPlace.maskId < 512)  ) // Publish only those who '8 < mask id < 512' (Not Amazons, nor websites)
                            {
                                if (xlProduct.SellingFormat == "A" || xlProduct.Items.Count == 0)
                                {
                                    VerifyAddItemCall api_AUCTION_Call = new VerifyAddItemCall(apiContext);
                                    fees = api_AUCTION_Call.VerifyAddItem(lproduct);
                                }
                                else
                                {
                                    VerifyAddFixedPriceItemCall api_FP_Call = new VerifyAddFixedPriceItemCall(apiContext);
                                    fees = api_FP_Call.VerifyAddFixedPriceItem(lproduct);
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
                            txtStatus.Text = "\r\nThe item was listed successfully! " + txtStatus.Text + " ";
                            txtStatus.Update();

                            // Let's see if the posting already exists, if not save this posting. Later we'll save the pictures
                            String lpostingID = null;
                            SqlCommand lcmd = new SqlCommand("SELECT * FROM bsi_posting WHERE sku='" + xlProduct.SKU + "'", lconn);
                            SqlDataReader ldr = lcmd.ExecuteReader();
                            if (ldr.Read())
                            {
                                lpostingID = ldr["id"].ToString();
                            }
                            ldr.Close();
                            lcmd.Cancel();

                            if (chkOverridePosting.Checked && lpostingID != null)
                            {
                                // Override and overwrite the product info of the product
                                lcmd = new SqlCommand("UPDATE bsi_posting SET gender=@gender,brand=@brand,style=@style,fullDescription=@fullDescription,keywords=@keywords,material=@material,color=@color,shade=@shade,heelHeight=@heelHeight WHERE id=" + lpostingID, lconn);
                                lcmd.Parameters.Add("@brand", SqlDbType.NVarChar).Value = xlProduct.Brand;
                                lcmd.Parameters.Add("@gender", SqlDbType.NVarChar).Value = xlProduct.Gender;
                                lcmd.Parameters.Add("@style", SqlDbType.NVarChar).Value = xlProduct.Style;
                                lcmd.Parameters.Add("@fullDescription", SqlDbType.NVarChar).Value = xlProduct.FullDescription;
                                lcmd.Parameters.Add("@keywords", SqlDbType.NVarChar).Value = xlProduct.Keywords;
                                lcmd.Parameters.Add("@material", SqlDbType.NVarChar).Value = xlProduct.Material;
                                lcmd.Parameters.Add("@color", SqlDbType.NVarChar).Value = xlProduct.Color;
                                lcmd.Parameters.Add("@shade", SqlDbType.NVarChar).Value = xlProduct.Shade;
                                lcmd.Parameters.Add("@heelHeight", SqlDbType.NVarChar).Value = xlProduct.HeelHeight;

                                lcmd.ExecuteNonQuery();
                                lpostingID = null;
                            }

                            if ( lpostingID == null )
                               lpostingID = lda.InsertQuery(POSTING_STATUS_ACTIVE, xlProduct.SKU, DateTime.Now,
                                       "", xlProduct.Gender, xlProduct.Brand,
                                       xlProduct.Size, xlProduct.Width, xlProduct.Condition, xlProduct.Category, xlProduct.Style, xlProduct.FullDescription,
                                       xlProduct.Keywords, xlProduct.Material, xlProduct.Color, xlProduct.Shade, xlProduct.HeelHeight, 
                                       xlProduct.Title, (int)xlProduct.MarketPlaces, xlProduct.Variation,xlProduct.Widths,"","","","").ToString();

                            // Create posts
                            int lx = -1, litemID = -1;

                            // Do we need to create variations or do we need to make an individual post? 
                            if (xlProduct.Variation != VARIATIONS_NONE) 
                            {
                                // Create a post for each marketplace
                                lx = 0;                                
                                for (int li = 1; lx < ItemMarketplace.MARKETPLACE_MAXMARKETS; lx++, li <<= 1)
                                {
                                    if ((xlProduct.MarketPlaces & li) != 0)
                                    {
                                            // We'll create a single post for this item
                                            // Decimal lprice = xlProduct.getPriceForMarketplace(lx);
                                            Decimal lprice = xlProduct.Price; // 2013-01-02
                                            if (lprice == 0)
                                            {
                                                // The item was created without price, use the price of the first sibling
                                                foreach (ItemExcel lixl in xlProduct.Items)
                                                {
                                                    //if (lixl.getPriceForMarketplace(lx) > 0) 2013-01-02
                                                    if (lixl.Price > 0)
                                                    {
                                                        //lprice = lixl.getPriceForMarketplace(lx);
                                                        lprice = lixl.Price; // 2013-01-02
                                                        break;
                                                    }
                                                } // foreach
                                            }

                                            String lpostID = lposts_da.InsertQuery(Int32.Parse(lpostingID), li,
                                                             "", POSTING_STATUS_READY2PUBLISH, xlProduct.SKU,
                                                             xlProduct.Title,
                                                             lprice.ToString(),
                                                             xlProduct.StartDate, xlProduct.EndDate,
                                                             xlProduct.SellingFormat, 
                                                             "", "", "", "", "", "", "",
                                                             xlProduct.purchaseOrder,
                                                             xlProduct.listUser).ToString();

                                            // Now, let's create the quantities
                                            foreach (ItemExcel liex in xlProduct.Items)
                                            {
                                                // We need the item ID of the product
                                                // if (liex.getQuantityForMarketplace(lx) > 0) 2013-01-02
                                                if (liex.Quantity > 0)
                                                {
                                                    litemID = -1;
                                                    lcmd = new SqlCommand("SELECT ID,ItemLookupCode FROM item WHERE ItemLookupCode='" + liex.ItemLookupCode + "'", lconn);
                                                    ldr = lcmd.ExecuteReader();
                                                    if (ldr.Read())
                                                    {
                                                        litemID = int.Parse(ldr["ID"].ToString());
                                                    }
                                                    ldr.Close();
                                                    lcmd.Cancel();

                                                    // Create the quantity
                                                    lqtys_da.Insert(int.Parse(lpostID), litemID,
                                                                    liex.ItemLookupCode, liex.Title,
                                                                    liex.Size, liex.Width, liex.Color,
                                                                    /*liex.getQuantityForMarketplace(lx),
                                                                    liex.getPriceForMarketplace(lx), 2013-01-02*/
                                                                    liex.Quantity,
                                                                    liex.Price,
                                                                    xlProduct.purchaseOrder,
                                                                    xlProduct.listUser);
                                                }
                                            } // foreach (ItemExcel liex in xlProduct.Items)

                                    } // if ((xlProduct.MarketPlaces & li) != 0)
                                } // for (int li = 1; lx < ItemMarketplace.MARKETPLACE_MAXMARKETS; lx++, li <<= 1)

                            }
                            else
                            {
                                // Create an individual post for each market this item goes to
                                lx = 0;
                                for (int li = 1; lx < ItemMarketplace.MARKETPLACE_MAXMARKETS; lx++, li <<= 1)
                                {
                                    if ((xlProduct.MarketPlaces & li) != 0)
                                    {
                                        // Create the post in the marketplace
                                        String lpostID = lposts_da.InsertQuery(Int32.Parse(lpostingID), li,
                                                                               "", POSTING_STATUS_READY2PUBLISH, 
                                                                               xlProduct.SKU, xlProduct.Title,
                                                                               /*xlProduct.getPriceForMarketplace(lx).ToString(), 2013-01-02*/
                                                                               xlProduct.Price.ToString(),
                                                                               xlProduct.StartDate, xlProduct.EndDate,
                                                                               xlProduct.SellingFormat,
                                                                               "", "", "", "", "", "", "",
                                                                               xlProduct.purchaseOrder,
                                                                               xlProduct.listUser).ToString();

                                        // We need the item ID of the product
                                        litemID = -1;
                                        lcmd = new SqlCommand("SELECT ID,ItemLookupCode FROM item WHERE ItemLookupCode='" + xlProduct.ItemLookupCode + "'", lconn);
                                        ldr = lcmd.ExecuteReader();
                                        if (ldr.Read())
                                        {
                                            litemID = int.Parse(ldr["ID"].ToString());
                                        }
                                        ldr.Close();
                                        lcmd.Cancel();

                                        // Create the quantity
                                        lqtys_da.Insert(int.Parse(lpostID), litemID, 
                                                        xlProduct.ItemLookupCode, xlProduct.Title,
                                                        xlProduct.Size, xlProduct.Width, xlProduct.Color,
                                                        /*xlProduct.getQuantityForMarketplace(lx),
                                                        xlProduct.getPriceForMarketplace(lx),*/
                                                        xlProduct.Quantity,
                                                        xlProduct.Price,
                                                        xlProduct.purchaseOrder,
                                                        xlProduct.listUser);
                                    }
                                } // for

                            } // if (xlProduct.Variation != VARIATIONS_NONE)

                            // Finally, if everything went OK we'll upload the pictures to eBay. Then
                            // we'll update the posting information with the urls of the pictures

                            // Use this code to upload pictures to the eBay server
                            // first let's put all the pix URL in a single, comma delimited, string
                            if (currentMarketPlace.maskId > 8 && currentMarketPlace.maskId < 512)
                            {
                                String lpix = "pictures";
                                foreach (String lpic in xlProduct.Pictures)
                                {
                                    String lpicURL = uploadPicture(lpic);
                                    if (!String.IsNullOrEmpty(lpicURL))
                                    {
                                        lpix += " | " + lpicURL;
                                    }
                                } // foreach (String lpic in xlProduct.Pictures)
                                lpix = lpix.Replace("pictures | ", "");

                                SqlCommand lc2 = new SqlCommand("UPDATE BSI_POSTING SET PICTURES='" + lpix + "' WHERE ID=" + lpostingID, lconn);
                                txtStatus.Text = "\r\nPictures upload status:" +
                                                 lc2.ExecuteNonQuery().ToString() + " " +
                                                 txtStatus.Text;
                            }
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

                        txtStatus.Text = "\r\n" + txtStatus.Text;

                    } // foreach (ItemExcel in lproducts)
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

            btnStart.Enabled = true;
            btnStartPublishing.Enabled = false;
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

                /* For this process we do not need the real qts & prices
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
                TheVariation.StartPrice = price;
                */

                TheVariation.Quantity = 1;
                price.Value = 9.99;
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

        ItemType BuildItem(ItemExcel lix)
        {
            ItemType item = new ItemType();

            // item title
            item.Title = lix.Title;
            // item description
            item.Description = _descriptionHeader + lix.Title + " " + lix.FullDescription + _descriptionFooter;
            item.SKU = lix.SKU;

            // Create the picture, save the URL and then pass it to the item
            item.PictureDetails = new PictureDetailsType();
            item.PictureDetails.PhotoDisplay = PhotoDisplayCodeType.PicturePack;
            item.PictureDetails.GalleryType = GalleryTypeCodeType.Gallery;
            item.PictureDetails.PictureURL = new StringCollection();
            foreach (String lpic in lix.URLPictures)
                item.PictureDetails.PictureURL.Add(lpic);

            // listing type
            BestOfferDetailsType lbo = null;
            switch (lix.SellingFormat)
            {
                case "A":
                    item.ListingType = ListingTypeCodeType.Chinese;
                    item.ListingDuration = "Days_7";
                    break;
                case "A1" :
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

            // Start time if specified. We cannot use "lix.StartDate" because some items will be posted to be published
            // for later times accepted by eBay
            // if (lix.StartDate > DateTime.Now)
            // item.ScheduleTime = (DateTime.Now).AddHours(3); // lix.StartDate;

            item.HitCounter = HitCounterCodeType.BasicStyle;

            // item condition, New=1000, New without box=1500, New with defects=1750, Pre-owned=3000
            switch (lix.Condition)
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
            if (lix.Items.Count == 0 && lix.Category != "31387")
            {
                litemspec = new NameValueListType();
                litemspec.Name = getEbaySizeName(lix.Gender);
                litemspec.Value = new StringCollection(new String[] { lix.Size });
                item.ItemSpecifics.Add(litemspec);

                litemspec = new NameValueListType();
                litemspec.Name = "Width";
                String lwidth = convertWidth(lix.Gender, lix.Width);
                litemspec.Value = new StringCollection(new String[] { lwidth });
                item.ItemSpecifics.Add(litemspec);
            }

            int lcategory = 0;
            if (lix.Gender == "MENS")
            {
                // Let's see if the category was entered as a number or as text
                if (!Int32.TryParse(lix.Category, out lcategory))
                {
                    switch (lix.Category.ToUpper())
                    {
                        case "CASUAL": lcategory = 24087; break;
                        case "ATHLETIC": lcategory = 15709; break;
                        case "BOOTS": lcategory = 11498; break;
                        case "DRESS/FORMAL": lcategory = 53120; break;
                        case "SANDALS": lcategory = 11504; break;
                        case "SLIPPERS": lcategory = 11505; break;
                        default: lcategory = 63850; break; // Mixed items & lots
                    }
                }
            }
            else
            {
                // Let's see if the category was entered as a number or as text
                if (lix.Gender == "WOMENS")
                {
                    if (!Int32.TryParse(lix.Category, out lcategory))
                    {
                        switch (lix.Category.ToUpper())
                        {
                            case "HEELS": lcategory = 55793; break;
                            case "CASUAL": lcategory = 45333; break;
                            case "ATHLETIC": lcategory = 95672; break;
                            case "BOOTS": lcategory = 53557; break;
                            case "SANDALS": lcategory = 62107; break;
                            case "SLIPPERS": lcategory = 11632; break;
                            case "OCCUPATIONAL": lcategory = 53548; break;
                            default: lcategory = 63889; break; // Mixed items & lots
                        }
                    }
                }
                else
                {
                    if (lix.Gender == "JUNIOR")
                        lcategory = 155202;
                }
            }

            litemspec = new NameValueListType();
            litemspec.Name = "Brand";
            litemspec.Value = new StringCollection(new String[] { lix.Brand });
            item.ItemSpecifics.Add(litemspec);

            litemspec = new NameValueListType();
            litemspec.Name = "Style";
            litemspec.Value = new StringCollection(new String[] { lix.Style });
            item.ItemSpecifics.Add(litemspec);

            if (!String.IsNullOrEmpty(lix.Color))
            {
                litemspec = new NameValueListType();
                litemspec.Name = "Color";
                litemspec.Value = new StringCollection(new String[] { lix.Color });
                item.ItemSpecifics.Add(litemspec);
            }

            if (!String.IsNullOrEmpty(lix.Material))
            {
                litemspec = new NameValueListType();
                if (lcategory != 31387)
                    litemspec.Name = "Material";
                else
                    litemspec.Name = "Band Material";
                litemspec.Value = new StringCollection(new String[] { lix.Material });
                item.ItemSpecifics.Add(litemspec);
            }

            if (!String.IsNullOrEmpty(lix.Shade))
            {
                litemspec = new NameValueListType();
                litemspec.Name = "Shade";
                litemspec.Value = new StringCollection(new String[] { lix.Shade });
                item.ItemSpecifics.Add(litemspec);
            }

            // end-of-item-specifics

            // listing price
            item.Currency = CurrencyCodeType.USD;

            if (lix.Items.Count == 0) // Do not set price or quantity for products with children
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
            category.CategoryID = lcategory.ToString(); // Primary Category
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
            if (lix.Items.Count > 0)
            {
                int pvariations = 0;
                String lwidthsList = "";
                item.Variations = createVariations(lix, out pvariations,out lwidthsList);
                lix.Variation = pvariations;
                lix.Widths = lwidthsList;

                // Let's see what variations were not set to set them in default
                if ((pvariations & VARIATIONS_WIDTH) == 0)
                {
                    // There were sizes but not widths, then set the general width
                    litemspec = new NameValueListType();
                    litemspec.Name = "Width";
                    String lwidth = convertWidth(lix.Gender, lix.Width);
                    litemspec.Value = new StringCollection(new String[] { lwidth });
                    item.ItemSpecifics.Add(litemspec);
                }
            }

            // shipping details
            item.ShippingDetails = BuildShippingDetails();

            return item;
        } // BuildItem

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
            shippingOptions.ShippingService = ShippingServiceCodeType.ShippingMethodStandard.ToString();
            if ( double.Parse(currentMarketPlace.shippingDomesticStandard) == 0)
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

            return sd;
        } // BuildShippingDetails

        private void Form1_Load(object sender, EventArgs e)
        {
            btnStart.Enabled = true;
            btnStartPublishing.Enabled = false;

            // Read all the marketplaces
            SqlConnection lconn = null;
            berkeleyDataSetTableAdapters.bsi_marketplacesTableAdapter lda;
            
            try
            {
                lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString());
                lconn.Open();
                lda = new berkeleyDataSetTableAdapters.bsi_marketplacesTableAdapter();
                lda.Connection = lconn;
                lda.Fill(ldsMarkets); // This will fill sort by maskid
                foreach (berkeleyDataSet.bsi_marketplacesRow lmarketPlace in ldsMarkets.Rows)
                {
                    cmbMarkets.Items.Add(lmarketPlace);
                } // foreach
                cmbMarkets.DisplayMember = "name";
                cmbMarkets.SelectedIndex = EBAY_STARTINGINDEX; // Select the first one from ebay
            }
            catch (Exception pe)
            {
                MessageBox.Show("\nERROR WHILE READING MARKETPLACES: " + pe.ToString() + "\n", " Error on Load ");
            }
            finally
            {
                if ( lconn != null ) 
                {
                    lconn.Close();
                };
            }
        } // Form1_Load

        private string uploadPicture(string pfname)
        {
            string lurlpic = null;

            //read the image file as a byte array
            if (DEBUG_MODE) return null;

            System.IO.FileStream fs = new System.IO.FileStream(pfname, FileMode.Open, FileAccess.Read);
            fs.Seek(0, SeekOrigin.Begin);
            System.IO.BinaryReader br = new System.IO.BinaryReader(fs);

            byte[] image = br.ReadBytes((int)fs.Length);
            br.Close();
            fs.Close();

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create("https://api.ebay.com/ws/api.dll");
            HttpWebResponse resp = null;

            string boundary = "MIME_boundary";
            string CRLF = "\r\n";

            //Add the request headers
            req.Headers.Add("X-EBAY-API-COMPATIBILITY-LEVEL", "515");
            req.Headers.Add("X-EBAY-API-DEV-NAME", Properties.Settings.Default.eBayDevID); //use your devid
            req.Headers.Add("X-EBAY-API-APP-NAME", Properties.Settings.Default.eBayAppID); //use your appid
            req.Headers.Add("X-EBAY-API-CERT-NAME", Properties.Settings.Default.eBayCertID); //use your certid
            req.Headers.Add("X-EBAY-API-SITEID", "0");
            req.Headers.Add("X-EBAY-API-DETAIL-LEVEL", "0");
            req.Headers.Add("X-EBAY-API-CALL-NAME", "UploadSiteHostedPictures");
            req.ContentType = "multipart/form-data; boundary=" + boundary;

            //set the method to POST
            req.Method = "POST";

            //set the HTTP version to 1.0
            req.ProtocolVersion = HttpVersion.Version10;

            //replace token with your own token
            string token = currentMarketPlace.eBayToken;

            //Construct the request
            string strReq1 = "--" + boundary + CRLF
                             + "Content-Disposition: form-data; name=document" + CRLF
                             + "Content-Type: text/xml; charset=\"UTF-8\"" + CRLF + CRLF
                             + "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                             + "<UploadSiteHostedPicturesRequest xmlns=\"urn:ebay:apis:eBLBaseComponents\">"
                             + "<RequesterCredentials>"
                             + "<eBayAuthToken>" + token + "</eBayAuthToken>"
                             + "</RequesterCredentials>"
                             + "<PictureSet>Supersize</PictureSet>"
                             + "</UploadSiteHostedPicturesRequest>"
                             + CRLF + "--" + boundary + CRLF
                             + "Content-Disposition: form-data; name=image; filename=image" + CRLF
                             + "Content-Type: application/octet-stream" + CRLF
                             + "Content-Transfer-Encoding: binary" + CRLF + CRLF;

            string strReq2 = CRLF + "--" + boundary + "--" + CRLF;

            //Convert the string to a byte array
            byte[] postDataBytes1 = System.Text.Encoding.ASCII.GetBytes(strReq1);
            byte[] postDataBytes2 = System.Text.Encoding.ASCII.GetBytes(strReq2);

            int len = postDataBytes1.Length + postDataBytes2.Length + image.Length;
            req.ContentLength = len;

            //Post the request to eBay
            System.IO.Stream requestStream = req.GetRequestStream();
            requestStream.Write(postDataBytes1, 0, postDataBytes1.Length);
            requestStream.Write(image, 0, image.Length);
            requestStream.Write(postDataBytes2, 0, postDataBytes2.Length);
            requestStream.Close();

            string response;
            try
            {
                // get response and write to console
                resp = (HttpWebResponse)req.GetResponse();

                StreamReader responseReader = new StreamReader(resp.GetResponseStream(), Encoding.UTF8);
                response = responseReader.ReadToEnd();
                resp.Close();

                //response contains our pictures url
                System.Xml.XmlDocument xml = new System.Xml.XmlDocument();
                xml.LoadXml(response);

                //Extract the FullURL from the response
                System.Xml.XmlNodeList list = xml.GetElementsByTagName("FullURL", "urn:ebay:apis:eBLBaseComponents");
                lurlpic = list[0].InnerText;

                /*
                 * Get the other elements from the response if required
                list = xml.GetElementsByTagName("PictureSet", "urn:ebay:apis:eBLBaseComponents");
                if ( list != null )
                    txtStatus.Text = "Result of PictureSet:" + list[0].InnerText + "\r\n" + txtStatus.Text; 
                */
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return lurlpic;
        } // uploadPicture

        bool deleteItemOnEbay(string pitemid)
        {
            bool lreturn = false;

            try
            {
                GetApiContext();
                EndItemCall lendItemCall = new EndItemCall(apiContext);
                lendItemCall.EndingReason = EndReasonCodeType.Incorrect;
                lendItemCall.ItemID = pitemid;
                lendItemCall.Execute();
                lreturn = true;
            }
            catch (Exception pe)
            {
                txtStatus.Text = "Error while deleting the product: " + pe.ToString() + "\r\n" + txtStatus.Text;
            }

            return lreturn;
        } // deleteItemOnEbay

        bool isTheProductOnWebsite(ItemExcel lix)
        {
            bool lreturn = true;

            using (SqlConnection lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString()))
            {
                lconn.Open();
                try
                {
                    String lcmdS = "SELECT thePost.id, thePost.marketplace,thePost.markerplaceitemid,thePost.status," +
                                   "theQ.postid,theQ.itemlookupcode,theQ.Title,theQ.size,theQ.width,theQ.quantity " +
                                   "FROM bsi_posts as thePost, " +
                                   "bsi_quantities as theQ " +
                                   "where theQ.postid=thePost.id and thePost.sku='" + lix.SKU +
                                   "' AND thePost.marketplace=512 AND (thePost.status=0 OR thePost.status=10)";
                    SqlCommand lc = new SqlCommand(lcmdS, lconn);
                    lc.Connection = lconn;

                    SqlDataReader lr = lc.ExecuteReader();
                    if (lr.Read())
                    {
                        // We need to create at least one item
                        String lpostid = lr["id"].ToString().Trim();
                        ItemExcel laux = new ItemExcel();
                        laux.copyNewItem(lix);
                        do
                        {
                            ItemExcel ltempxl = new ItemExcel(laux);
                            ltempxl.ItemLookupCode = lr["itemlookupcode"].ToString().Trim();
                            ltempxl.Title = lr["Title"].ToString().Trim();
                            ltempxl.Size = lr["Size"].ToString().Trim();
                            ltempxl.Width = lr["Width"].ToString().Trim();
                            // ltempxl.setPriceForMarketplace(lix.getPriceForMarketplace(lmarketplace), lmarketplace); 2013-01-02
                            ltempxl.Price = lix.Price;
                            int lq = 0;
                            int.TryParse(lr["Quantity"].ToString().Trim(), out lq);
                            // ltempxl.setQuantityForMarketplace(lq, lmarketplace); 2013-01-02
                            ltempxl.Quantity = lq;
                            laux.Items.Add(ltempxl);
                        } while (lr.Read());
                        lr.Close();

                        // Cancel the actual product and update its status
                        lc.CommandText = "UPDATE bsi_posts SET status=110 WHERE id=" + lpostid;
                        lc.ExecuteNonQuery();

                        // TO-DO: Cancel on Amazon...?

                        // Now combine both items... but first check if this item is single...
                        if ((lix.Items == null || lix.Items.Count < 1)) // If so, then we need to make it father w/1 child
                        {
                            ItemExcel ltempxl = new ItemExcel();
                            if (lix.Items.Count > 0)
                                ltempxl.copyNewItem(lix.Items[0]); // Copy from the first item
                            else
                                ltempxl.copyNewItem(lix);
                            lix.Title = removeSize(lix.Title);
                            //ltempxl.setPriceForMarketplace(lix.getPriceForMarketplace(lmarketplace), lmarketplace); 2013-01-02
                            ltempxl.Price = lix.Price;
                            lix.Items.Add(ltempxl);
                        }

                        foreach (ItemExcel lax in laux.Items)
                        {
                            ItemExcel lu = lix.Items.Find(delegate(ItemExcel pi)
                            {
                                return pi.ItemLookupCode == lax.ItemLookupCode;
                            });
                            if (lu != null)
                            {   /* 2013-01-02
                                int lnewQty = lu.getQuantityForMarketplace(lmarketplace) + 
                                               lax.getQuantityForMarketplace(lmarketplace);
                                lu.setQuantityForMarketplace(lnewQty, lmarketplace);
                                */
                                lu.Quantity = lu.Quantity + lax.Quantity;
                            }
                            else
                            {
                                lix.Items.Add(lax);
                            }
                        } // foreach
                        lix.Items.Sort(sortBySize);
                    }
                    lr.Close();
                }
                catch (Exception pe)
                {
                    txtStatus.Text = "Error while checking on Website: " + pe.ToString() + "\r\n" + txtStatus.Text;
                }
            } // using
            return lreturn;
        } // isTheProductOnWebsite

        bool isTheProductOnAmazon(ItemExcel lix)
        {
            bool lreturn = true;
            int lmarketplace = cmbMarkets.SelectedIndex;

            using (SqlConnection lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString()))
            {
                lconn.Open();
                try
                {
                    String lcmdS = "SELECT thePost.id, thePost.marketplace,thePost.markerplaceitemid,thePost.status," +
                                   "theQ.postid,theQ.itemlookupcode,theQ.Title,theQ.size,theQ.width,theQ.quantity " + 
                                   "FROM bsi_posts as thePost, " +
                                   "bsi_quantities as theQ " +
                                   "where theQ.postid=thePost.id and thePost.sku='" + lix.SKU +
                                   "' AND thePost.marketplace=1 AND (thePost.status=0 OR thePost.status=10)";
                    SqlCommand lc = new SqlCommand(lcmdS, lconn);
                    lc.Connection = lconn;

                    SqlDataReader lr = lc.ExecuteReader();
                    if (lr.Read())
                    {
                        // We need to create at least one item
                        String lpostid = lr["id"].ToString().Trim();
                        ItemExcel laux = new ItemExcel();
                        laux.copyNewItem(lix);
                        do
                        {
                            ItemExcel ltempxl = new ItemExcel(laux);
                            ltempxl.ItemLookupCode = lr["itemlookupcode"].ToString().Trim();
                            ltempxl.Title = lr["Title"].ToString().Trim();
                            ltempxl.Size = lr["Size"].ToString().Trim();
                            ltempxl.Width = lr["Width"].ToString().Trim();
                            // ltempxl.setPriceForMarketplace(lix.getPriceForMarketplace(lmarketplace), lmarketplace); 2013-01-02
                            ltempxl.Price = lix.Price;
                            int lq = 0;
                            int.TryParse(lr["Quantity"].ToString().Trim(), out lq);
                            // ltempxl.setQuantityForMarketplace(lq, lmarketplace); 2013-01-02
                            ltempxl.Quantity = lq;
                            laux.Items.Add(ltempxl);
                        } while (lr.Read());
                        lr.Close();

                        // Cancel the actual product and update its status
                        lc.CommandText = "UPDATE bsi_posts SET status=110 WHERE id=" + lpostid;
                        lc.ExecuteNonQuery();

                        // TO-DO: Cancel on Amazon...?

                        // Now combine both items... but first check if this item is single...
                        if ((lix.Items == null || lix.Items.Count < 1)) // If so, then we need to make it father w/1 child
                        {
                            ItemExcel ltempxl = new ItemExcel();
                            if (lix.Items.Count > 0)
                                ltempxl.copyNewItem(lix.Items[0]); // Copy from the first item
                            else
                                ltempxl.copyNewItem(lix);
                            lix.Title = removeSize(lix.Title);
                            //ltempxl.setPriceForMarketplace(lix.getPriceForMarketplace(lmarketplace), lmarketplace); 2013-01-02
                            ltempxl.Price = lix.Price;
                            lix.Items.Add(ltempxl);
                        }

                        foreach (ItemExcel lax in laux.Items)
                        {
                            ItemExcel lu = lix.Items.Find(delegate(ItemExcel pi)
                                               {
                                                   return pi.ItemLookupCode == lax.ItemLookupCode;
                                               });
                            if (lu != null)
                            {   /* 2013-01-02
                                int lnewQty = lu.getQuantityForMarketplace(lmarketplace) + 
                                               lax.getQuantityForMarketplace(lmarketplace);
                                lu.setQuantityForMarketplace(lnewQty, lmarketplace);
                                */
                                lu.Quantity = lu.Quantity + lax.Quantity;
                            }
                            else
                            {
                                lix.Items.Add(lax);
                            }
                        } // foreach
                        lix.Items.Sort(sortBySize);
                    }
                    lr.Close();
                }
                catch (Exception pe)
                {
                    txtStatus.Text = "Error while checking on Amazon: " + pe.ToString() + "\r\n" + txtStatus.Text;
                }
            } // using

            return lreturn;
        } // isTheProductOnAmazon

        bool isTheProductOnEbay(ItemExcel lix)
        {
            bool lreturn = false;
            SqlConnection lconn = null;
            SqlCommand lcmd = null;

            // Look for the item
            ItemType litem = null;
            int lmarketplace = EBAY_STARTINGINDEX + cmbMarkets.SelectedIndex;

            litem = itemsOnline[cmbMarkets.SelectedIndex].
                                Find(
                                      delegate(ItemType pi)
                                      {
                                          bool lf = pi.SKU == lix.SKU;
                                          if (!lf)
                                          {
                                              // Maybe we are a parent (SKU w/o size) and this one is a child (SKU w/size)
                                              try
                                              {
                                                  String[] lsplittedsku = pi.SKU.Split(new char[] { '-' });
                                                  lf = (lix.SKU == lsplittedsku[0]);
                                              }
                                              catch (Exception pe)
                                              {
                                                  String lsku = (pi.SKU != null) ? pi.SKU : pi.ItemID;
                                                  MessageBox.Show("Error searching: " + lix.SKU + " caused by: " + lsku + " - " + pe.ToString());
                                              }
                                          }
                                          return lf;
                                      }
                                    );

            if (litem != null)
            {
                txtStatus.Text = "Product is already listed... " + txtStatus.Text;
                lreturn = true;

                // If we found the item on eBay then we need to make our item a father if it is a single product
                if ((lix.Items == null || lix.Items.Count < 1) && !lix.SellingFormat.Contains('A'))
                {
                    ItemExcel ltempxl = new ItemExcel();
                    if (lix.Items.Count > 0)
                        ltempxl.copyNewItem(lix.Items[0]); // Copy from the first item
                    else
                        ltempxl.copyNewItem(lix);
                    lix.Title = removeSize(lix.Title);
                    //ltempxl.setPriceForMarketplace(lix.getPriceForMarketplace(lmarketplace), lmarketplace); 2013-01-02
                    ltempxl.Price = lix.Price;
                    lix.Items.Add(ltempxl);
                }

                switch (lix.SellingFormat)
                {
                    case "A":
                    case "A3":
                    case "A5":
                        // Simply add qty to current item on our DB, nothing else.

                        if (lix.ItemLookupCode == litem.SKU && litem.ListingType == ListingTypeCodeType.Chinese)
                        {
                            txtStatus.Text = "\r\n\r\nWARNING! PLEASE NOTE: eBAY ITEM [" + lix.ItemLookupCode + "] IS ALREADY IN AUCTION. REVIEW AND TRY TO PUBLISH IT AGAIN.\r\n\r\n" + txtStatus.Text;
                        }

                        break;
                    default:
                        bool lfoundFlag = false; // We'll use this to look for items
                        if (litem.ListingType != ListingTypeCodeType.Chinese)
                        {
                            // Cancel the eBay product
                            deleteItemOnEbay(litem.ItemID);

                            // Update the current item published with the eBay item id
                            try
                            {
                                string lscmd = "UPDATE bsi_posts SET status=110 WHERE markerplaceItemID='" + litem.ItemID + "'";
                                lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString());
                                lconn.Open();

                                lcmd = new SqlCommand(lscmd, lconn);
                                lcmd.ExecuteNonQuery();
                                lcmd.Cancel();
                            }
                            catch (Exception pe)
                            {
                                txtStatus.Text = "Error while trying to update our database: " + pe.ToString() + "\r\n" + txtStatus.Text;
                            }
                            finally
                            {
                                if ( lcmd != null ) lcmd.Cancel();
                                if ( lconn != null ) lconn.Close();
                            }

                            // Combine both products into one
                            // Ours is a parent with children, let's add the kid(s) of the found one
                            if (litem.Variations == null)
                            {
                                // Let's see if we already have this item
                                lfoundFlag = false;
                                foreach (ItemExcel lax in lix.Items)
                                {
                                    if (lax.ItemLookupCode == litem.SKU.Trim())
                                    {
                                        lfoundFlag = true;
                                        /*int lqx = lax.getQuantityForMarketplace(lmarketplace); 2013-01-02
                                        lax.setQuantityForMarketplace(lqx+litem.Quantity - litem.SellingStatus.QuantitySold,lmarketplace);*/
                                        lax.Quantity = lax.Quantity+litem.Quantity - litem.SellingStatus.QuantitySold;
                                        break; // Get out of the loop
                                    }
                                } // foreach

                                if (!lfoundFlag) // Create and add the size
                                {
                                    ItemExcel ltempxl = new ItemExcel();
                                    if (lix.Items.Count > 0)
                                        ltempxl.copyNewItem(lix.Items[0]); // Copy from the first item
                                    else
                                        ltempxl.copyNewItem(lix);

                                    // Set the size and width
                                    ltempxl.ItemLookupCode = litem.SKU;
                                    String[] lprodinfo = litem.SKU.Split(new char[] { '-' });
                                    if (lprodinfo.Length > 2)
                                    {
                                        ltempxl.Title = litem.Title.Trim();
                                        ltempxl.Size = lprodinfo[1];
                                        ltempxl.Width = lprodinfo[2];

                                        /* 2013-01-02
                                        ltempxl.setPriceForMarketplace(lix.getPriceForMarketplace(lmarketplace), lmarketplace);
                                        ltempxl.setQuantityForMarketplace(litem.Quantity - litem.SellingStatus.QuantitySold,lmarketplace);
                                        */
                                        ltempxl.Price = lix.Price;
                                        ltempxl.Quantity = litem.Quantity - litem.SellingStatus.QuantitySold;

                                        lix.Items.Add(ltempxl);
                                        lix.Items.Sort(sortBySize);
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
                                            /* 2013-01-02
                                            int lqx = lax.getQuantityForMarketplace(lmarketplace);
                                            lax.setQuantityForMarketplace(lqx+lnewKidOneBay.Quantity - lnewKidOneBay.SellingStatus.QuantitySold,lmarketplace);
                                            */
                                            lax.Quantity = lax.Quantity + (lnewKidOneBay.Quantity - lnewKidOneBay.SellingStatus.QuantitySold);
                                            break; // Get out of the loop
                                        }
                                    } // foreach

                                    if (!lfoundFlag) // Create the size
                                    {
                                        ItemExcel ltempxl = new ItemExcel();
                                        if ( lix.Items != null && lix.Items.Count > 0 )
                                           ltempxl.copyNewItem(lix.Items[0]); // lix.Items[0].clone(); 
                                        else
                                           ltempxl.copyNewItem(lix);

                                        // Set the size and width
                                        String[] lprodinfo = lnewKidOneBay.SKU.Split(new char[] { '-' });
                                        if (lprodinfo.Length > 2)
                                        {
                                            ltempxl.ItemLookupCode = lnewKidOneBay.SKU;
                                            ltempxl.Title = (lnewKidOneBay.VariationTitle != null) ? lnewKidOneBay.VariationTitle.Trim() : "";
                                            ltempxl.Size = lprodinfo[1];
                                            ltempxl.Width = lprodinfo[2];
                                            /*
                                            ltempxl.setPriceForMarketplace(lix.getPriceForMarketplace(lmarketplace),lmarketplace);                                            
                                            ltempxl.setQuantityForMarketplace(lnewKidOneBay.Quantity - lnewKidOneBay.SellingStatus.QuantitySold,lmarketplace);
                                            */
                                            ltempxl.Price = lix.Price;
                                            ltempxl.Quantity = lnewKidOneBay.Quantity - lnewKidOneBay.SellingStatus.QuantitySold;
                                            lix.Items.Add(ltempxl);
                                            lix.Items.Sort(sortBySize);
                                        }
                                    }
                                } // for each
                            } // if (litem.Variations != null)
                        } // if (litem.ListingType != ListingTypeCodeType.Chinese)
                        break;
                } // switch
            } // if (litem != null)

            return lreturn;
        } // isTheProductOnEbay

        private void btnUpdateMarketplaces_Click(object sender, EventArgs e)
        {
            updateMarketplaces();
        } // btnUpdateMarketplaces_Click


        private void cmbMarkets_SelectedIndexChanged(object sender, EventArgs e)
        {
            chkPublishWOPics.Enabled = (cmbMarkets.SelectedIndex < EBAY_STARTINGINDEX || cmbMarkets.SelectedIndex >= WEB_STARTINGINDEX);
            if (cmbMarkets.SelectedIndex >= EBAY_STARTINGINDEX && cmbMarkets.SelectedIndex < WEB_STARTINGINDEX) chkPublishWOPics.Checked = false;
        } // cmbMarkets_SelectedIndexChanged

        private void eBayPageSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmEbayPageSize lf = new frmEbayPageSize();

            lf.ShowDialog();
        } // eBayPageSizeToolStripMenuItem_Click

    } // partial class Form1
} // namespace BSI_InventoryPreProcessor

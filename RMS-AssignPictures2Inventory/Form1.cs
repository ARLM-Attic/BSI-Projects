using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Net;

namespace RMS_AssignPictures2Inventory
{
    public partial class Form1 : Form
    {
        public const int EXCEL_COLS_BRAND = 4; // D
        public const int EXCEL_COLS_SKU = 5; // E
        
        public const int EXCEL_COLS_CATEGORY = 10; // J
        public const int EXCEL_COLS_TYPE = 11; // K
        public const int EXCEL_COLS_CLASS = 12; // L

        public const int EXCEL_COLS_TITLE = 13; // M
        public const int EXCEL_COLS_FULLDESCRIPTION = 15; // O

        public const int EXCEL_COLS_MATERIAL = 17; // Q
        public const int EXCEL_COLS_COLOR = 18; // R
        public const int EXCEL_COLS_SHADE = 19; // S

        public const int EXCEL_COLS_GENDER = 22; // V

        public const int IMAGE_WRAP_HEIGHT = 100;
        public const int IMAGE_HEIGHT = 90;

        bool lstop = false;
        string lFileName = "", lFullPath = "";

        public Form1()
        {
            InitializeComponent();
        } // Form1

        private void btnStart_Click(object sender, EventArgs e)
        {

            if (String.IsNullOrEmpty(txtInventoryFile.Text.Trim()))
            {
                MessageBox.Show("\n\nPLEASE SELECT AN EXCEL FILE\n\n", "Error: not file");
                return;
            }

            if (String.IsNullOrEmpty(txtPicturesPath.Text.Trim()))
            {
                MessageBox.Show("\n\nPLEASE SELECT THE PATH WHERE THE PRODUCT PICTURES FOR RMS ARE\n\n", "Error: not file");
                return;
            }

            SqlConnection lconn = null;
            SqlCommand lcmd = null;
            SqlDataReader lr = null;
            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString());
                lconn.Open();

                Microsoft.Office.Interop.Excel.Workbook theWorkbook = excelApp.Workbooks.Open(txtInventoryFile.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)theWorkbook.ActiveSheet; // (Excel.Worksheet)excelApp.ActiveSheet;
                int lcurrRow = 2;
                lstop = false;
                String lbrand = "", lsku = "", lpicture = "";
                Double ltop = 0;
                string lwithpic = "";
                while (!lstop)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A" + lcurrRow.ToString(),
                                                            "AZ" + lcurrRow.ToString());
                    range.RowHeight = IMAGE_WRAP_HEIGHT;
                    System.Array myvalues = (System.Array)range.Cells.Value;
                    ltop = Convert.ToDouble(range.Top);
                    lbrand = Convert.ToString(myvalues.GetValue(1, EXCEL_COLS_BRAND)); // "myvalues" is a 2-D array (no matter if the range was from one single row)

                    if (!String.IsNullOrEmpty(lbrand))
                    {
                        lwithpic = "";
                        lsku = Convert.ToString(myvalues.GetValue(1, EXCEL_COLS_SKU)).Trim();
                        if (!String.IsNullOrEmpty(lsku))
                        {
                            lpicture = lsku + ".jpg"; // Convert.ToString(myvalues.GetValue(1, EXCEL_COLS_PICTURENAME));
                            if (!String.IsNullOrEmpty(lpicture))
                            {
                                lwithpic = "NO LOCAL PIC | ";
                                if (File.Exists(txtPicturesPath.Text + "\\" + lbrand + "\\" + lpicture))
                                {
                                    try
                                    {
                                        byte[] lfilecontents = File.ReadAllBytes(txtPicturesPath.Text + "\\" + lbrand + "\\" + lpicture);

                                        if (lfilecontents.Count<byte>() > 0)
                                        {
                                            txtStatus.Text += lsku + " has a local picture!\r\n";
                                            workSheet.Shapes.AddPicture(txtPicturesPath.Text + "\\" + lbrand + "\\" + lpicture,
                                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                                                        1, (float)ltop + 5,
                                                                        IMAGE_HEIGHT, IMAGE_HEIGHT);
                                            lwithpic = "LOCAL PIC | ";
                                        }
                                    }
                                    catch (Exception pe)
                                    {
                                        txtStatus.Text += "Local picture error " + pe.ToString() + "\r\n";
                                    }
                                }
                            }

                            // Let's see if we have it in our databases
                            lcmd = new SqlCommand("SELECT * FROM bsi_posting WHERE sku='" + lsku + "'",lconn);
                            lr = lcmd.ExecuteReader();
                            if ( lr.Read() )
                            {
                                lwithpic += "DESCR | ";

                                if ( chkDescription.Checked ) workSheet.Cells[lcurrRow,EXCEL_COLS_FULLDESCRIPTION] = removeSize(lr["fullDescription"].ToString());
                                if ( chkCategory.Checked ) workSheet.Cells[lcurrRow,EXCEL_COLS_CATEGORY] = lr["category"].ToString();
                                if ( chkStyle.Checked ) workSheet.Cells[lcurrRow,EXCEL_COLS_TYPE] = lr["style"].ToString();
                                
                                if ( chkMaterial.Checked ) workSheet.Cells[lcurrRow,EXCEL_COLS_MATERIAL] = lr["material"].ToString();
                                if ( chkColor.Checked ) workSheet.Cells[lcurrRow,EXCEL_COLS_COLOR] = lr["color"].ToString();
                                if ( chkShade.Checked ) workSheet.Cells[lcurrRow,EXCEL_COLS_SHADE] = lr["shade"].ToString();

                                if ( chkGender.Checked ) workSheet.Cells[lcurrRow, EXCEL_COLS_GENDER] = lr["gender"].ToString();

                                // Let's see if we have a remote picture and if that is still available
                                /*
                                string lpictures = lr["pictures"].ToString();
                                if ( !String.IsNullOrEmpty(lpictures) )
                                {
                                    string[] lremotepix = lpictures.Split( new char[] { '|' });
                                    try {
                                        WebClient client = new WebClient();
                                        Stream stream = client.OpenRead(lremotepix[0]);
                                        if ( stream != null )
                                        {
                                            lwithpic += "REMOTE PIX";
                                            stream.Close();
                                        }
                                    }
                                    catch(Exception pe)
                                    {
                                        txtStatus.Text += "Error while getting the remote pic: " + pe.ToString() + "\r\n";
                                    }
                                }
                                */
                                
                            }
                            lr.Close();
                            txtStatus.Text += lsku + " = " + lwithpic + "\r\n";
                        }

                        txtStatus.Update();
                        Application.DoEvents();
                    }
                    else lstop = true;
                    lcurrRow++;
                } // while

                object misValue = System.Reflection.Missing.Value;


                //theWorkbook.SaveAs(txtInventoryFile.Text,Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                theWorkbook.SaveCopyAs(lFullPath + "\\PIX_" + lFileName);
                theWorkbook.Save();
                theWorkbook.Close(true, misValue, misValue);
                excelApp.Quit();

                releaseObject(theWorkbook);
                releaseObject(excelApp);

                MessageBox.Show("PROCESS FINISHED WITH " + lcurrRow + " ROWS PROCESSED");
            }
            catch (Exception pe)
            {
                MessageBox.Show(pe.ToString(), "Error while processing");
            }
            finally
            {
                if (lconn != null) lconn.Close();
            }
        } // btnStart_Click

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

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbMarketplaces.SelectedIndex = 0;
        } // Form1_Load

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtInventoryFile.Text = openFileDialog1.FileName;
                lFileName = openFileDialog1.SafeFileName;
                lFullPath = System.IO.Path.GetDirectoryName(txtInventoryFile.Text);
            }
        } // btnSearch_Click

        private void btnSetPath_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPicturesPath.Text = folderBrowserDialog1.SelectedPath;

                if (!txtPicturesPath.Text.EndsWith("\\"))
                    txtPicturesPath.Text += "\\";
            };
        } // btnSetPath_Click

        private void btnStop_Click(object sender, EventArgs e)
        {
            lstop = true;
        } // btnStop_Click

        private string getConnectionString(string pn)
        {
            string ls = null;

            foreach (ConnectionStringSettings lcs in ConfigurationManager.ConnectionStrings)
            {
                if (lcs.Name.IndexOf(pn) >= 0)
                {
                    ls = lcs.ConnectionString;
                    break;
                }
            }; // foreach

            return ls;
        } // getConnectionString

        private void btnChkSKUs_Click(object sender, EventArgs e)
        {
            SqlCommand lc = null;
            SqlDataReader lr = null;
            SqlConnection lconn = null;

            if (String.IsNullOrEmpty(txtInventoryFile.Text.Trim()))
            {
                MessageBox.Show("\n\nPLEASE SELECT AN EXCEL FILE\n\n", "Error: not file");
                return;
            }

            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook theWorkbook = excelApp.Workbooks.Open(txtInventoryFile.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)theWorkbook.ActiveSheet; // (Excel.Worksheet)excelApp.ActiveSheet;
                int lcurrRow = 2;
                lstop = false;
                String lsku = "", lresult ="", lebayid ="" ;
                Double ltop = 0;


                string lcs = getConnectionString("berkeleyConnectionString");
                lconn = new SqlConnection(lcs);
                lconn.Open();

                while (!lstop)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A" + lcurrRow.ToString(),
                                                            "K" + lcurrRow.ToString());

                    System.Array myvalues = (System.Array)range.Cells.Value2;
                    ltop = Convert.ToDouble(range.Top);

                    // "myvalues" is a 2-D array (non-zero-indexed) (no matter if the range was from one single row)
                    lebayid = Convert.ToString(myvalues.GetValue(1, 1)).Trim(); 
                    lsku = Convert.ToString(myvalues.GetValue(1, 2)).Trim(); 

                    if (!String.IsNullOrEmpty(lebayid))
                    {
                        try
                        {

                            // Is this a repeated order?
                            lc = new SqlCommand("SELECT * FROM item where itemlookupcode='" + lsku.Trim() + "'", lconn);

                            lr = lc.ExecuteReader();
                            lresult = (lr.Read()) ? "OK" : "NOT FOUND";

                            workSheet.Cells[lcurrRow, 3] = lresult;

                            txtStatus.Text = (lcurrRow-1).ToString() + ".- " + lsku + " in RMS " + lresult + "\r\n" + txtStatus.Text;

                            lr.Close();
                        }
                        catch (Exception pe)
                        {
                            MessageBox.Show("SEVERE ERROR: " + pe.ToString(), "Error while processing orders");
                        }
                        finally
                        {

                        }
                    }
                    else lstop = true;

                    txtStatus.Update();
                    Application.DoEvents();
                    lcurrRow++;
                } // while

                
                if (lconn != null) lconn.Close();

                object misValue = System.Reflection.Missing.Value;


                //theWorkbook.SaveAs(txtInventoryFile.Text,Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                theWorkbook.SaveCopyAs(lFullPath + "\\RES_" + lFileName);
                theWorkbook.Save();
                theWorkbook.Close(true, misValue, misValue);
                excelApp.Quit();

                releaseObject(theWorkbook);
                releaseObject(excelApp);

                MessageBox.Show("PROCESS FINISHED WITH " + lcurrRow + " ROWS PROCESSED");
            }
            catch (Exception pe)
            {
                MessageBox.Show(pe.ToString(), "Error while processing");
            }
        } // btnChkSKUs_Click

        private void btnRegisterQtys_Click(object sender, EventArgs e)
        {
            SqlCommand lc = null;
            SqlDataReader lr = null;
            SqlConnection lconn = null;
            String[] lmkts = { "AmzQty", "MecalzoQty", "OMSQty" };

            int SKU_COLUMN = 1,
                QTY_COLUMN = 2;

            if (String.IsNullOrEmpty(txtInventoryFile.Text.Trim()))
            {
                MessageBox.Show("\n\nPLEASE SELECT AN EXCEL FILE\n\n", "Error: not file");
                return;
            }

            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook theWorkbook = excelApp.Workbooks.Open(txtInventoryFile.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)theWorkbook.ActiveSheet; // (Excel.Worksheet)excelApp.ActiveSheet;
                
                lstop = false;
                String lsku = "", lqtyS = "";
                int lcurrRow = 2, lqty = 0, lnotFounds = 0;
                Double ltop = 0;

                string lcs = getConnectionString("berkeleyConnectionString");
                lconn = new SqlConnection(lcs);
                lconn.Open();

                String lMktColumn = lmkts[cmbMarketplaces.SelectedIndex];

                while (!lstop)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A" + lcurrRow.ToString(),
                                                            "K" + lcurrRow.ToString());

                    System.Array myvalues = (System.Array)range.Cells.Value2;
                    ltop = Convert.ToDouble(range.Top);

                    // "myvalues" is a 2-D array (non-zero-indexed) (no matter if the range was from one single row)
                    lsku = Convert.ToString(myvalues.GetValue(1, SKU_COLUMN)).Trim();
                    lqtyS = Convert.ToString(myvalues.GetValue(1, QTY_COLUMN)).Trim();

                    if (!int.TryParse(lqtyS, out lqty)) lqty = 0;

                    if (!String.IsNullOrEmpty(lsku))
                    {
                        try
                        {
                            // Check if the item is in RMS
                            lc = new SqlCommand("SELECT * FROM item where itemlookupcode='" + lsku.Trim() + "'", lconn);

                            lr = lc.ExecuteReader();
                            if (!lr.Read())
                            {
                                workSheet.Cells[lcurrRow, 3] = "NF";
                                lnotFounds++;
                            }
                            lr.Close();                            

                            // Check to see if the item is already registered
                            lc = new SqlCommand("SELECT * FROM bsi_qtys where sku='" + lsku.Trim() + "'", lconn);

                            lr = lc.ExecuteReader();
                            if (lr.Read())
                            {
                                lr.Close();
                                lc.Cancel();

                                // Update qtys
                                lc = new SqlCommand("UPDATE bsi_qtys SET " + lMktColumn + "=" + lMktColumn + "+" + lqty + " where sku='" + lsku.Trim() + "'", lconn);
                                lc.ExecuteNonQuery();
                            }
                            else
                            {
                                lr.Close();
                                lc.Cancel();

                                // Create qtys
                                lc = new SqlCommand("INSERT INTO bsi_qtys (sku," + lMktColumn + ") values('" + lsku + "'," + lqty + ")", lconn);
                                lc.ExecuteNonQuery();
                            }

                            txtStatus.Text = (lcurrRow - 1).ToString() + ".- " + lsku + " in RMS " + lqtyS + "\r\n" + txtStatus.Text;
                            txtStatus.Update();
                            Application.DoEvents();

                            lr.Close();
                        }
                        catch (Exception pe)
                        {
                            MessageBox.Show("SEVERE ERROR: " + pe.ToString(), "Error while processing orders");
                        }
                    }
                    else lstop = true;

                    lcurrRow++;
                } // while


                if (lconn != null) lconn.Close();

                object misValue = System.Reflection.Missing.Value;

                //theWorkbook.SaveAs(txtInventoryFile.Text,Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                theWorkbook.SaveCopyAs(lFullPath + "\\INVENTORY_" + lFileName);
                theWorkbook.Save();
                theWorkbook.Close(true, misValue, misValue);
                excelApp.Quit();

                releaseObject(theWorkbook);
                releaseObject(excelApp);

                MessageBox.Show("PROCESS FINISHED WITH " + lcurrRow + " ROWS PROCESSED AND " + lnotFounds + " NOT FOUND IN RMS");
            }
            catch (Exception pe)
            {
                MessageBox.Show("Error while reading file: " + pe.ToString(), "Error for quantities");
            }
        } // btnRegisterQtys_Click

        private void btnVerifyOnEbay_Click(object sender, EventArgs e)
        {
            txtSKU.Text = txtSKU.Text.Trim();
            if (String.IsNullOrEmpty(txtSKU.Text))
            {
                MessageBox.Show("\r\nPLEASE SPECIFY A SKU\r\n");
                return;
            }

        }

        private void txtPicturesPath_TextChanged(object sender, EventArgs e)
        {

        } // btnVerifyOnEbay_Click

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
                    lsb.Append(ltext.Substring(0, lszTextpos));

                    // Let's find the size or white space
                    int li = lszTextpos, lconsecutiveWidthChars = 0;
                    bool lflag = false, lcheckingSizeW = false, lwidthfound = false;
                    while (!lflag && li < ltext.Length)
                    {
                        if (char.IsNumber(ltext[li])) lcheckingSizeW = true;
                        if (char.IsWhiteSpace(ltext[li]))
                        {
                            if (lwidthfound)
                                lflag = true;
                        }
                        if (char.IsLetter(ltext[li]))
                        {
                            if (lcheckingSizeW)
                            {
                                lwidthfound = true;
                                lconsecutiveWidthChars++;
                                if (lconsecutiveWidthChars > 3) // Get out! This is now width!!
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

        private void btnCheckMarketplaces_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtInventoryFile.Text.Trim()))
            {
                MessageBox.Show("\n\nPLEASE SELECT AN EXCEL FILE\n\n", "Error: not file");
                return;
            }

            SqlConnection lconn = null;
            SqlCommand lcmd = null;
            SqlDataReader lr = null;
            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                lconn = new SqlConnection(Properties.Settings.Default.berkeleyConnectionString.ToString());
                lconn.Open();

                Microsoft.Office.Interop.Excel.Workbook theWorkbook = excelApp.Workbooks.Open(txtInventoryFile.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)theWorkbook.ActiveSheet; // (Excel.Worksheet)excelApp.ActiveSheet;
                int lcurrRow = 2;
                lstop = false;

                String lbrand = "", lsku = "", lpicture = "";
                Double ltop = 0;
                string lwithpic = "";
                while (!lstop)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A" + lcurrRow.ToString(),
                                                            "AZ" + lcurrRow.ToString());
                    range.RowHeight = IMAGE_WRAP_HEIGHT;
                    System.Array myvalues = (System.Array)range.Cells.Value;
                    ltop = Convert.ToDouble(range.Top);
                    lbrand = Convert.ToString(myvalues.GetValue(1, EXCEL_COLS_BRAND)); // "myvalues" is a 2-D array (no matter if the range was from one single row)

                    if (!String.IsNullOrEmpty(lbrand))
                    {
                        lwithpic = "";
                        lsku = Convert.ToString(myvalues.GetValue(1, EXCEL_COLS_SKU)).Trim();
                        if (!String.IsNullOrEmpty(lsku))
                        {
                            lpicture = lsku + ".jpg"; // Convert.ToString(myvalues.GetValue(1, EXCEL_COLS_PICTURENAME));
                            if (!String.IsNullOrEmpty(lpicture))
                            {
                                lwithpic = "NO LOCAL PIC | ";
                                if (File.Exists(txtPicturesPath.Text + "\\" + lbrand + "\\" + lpicture))
                                {
                                    try
                                    {
                                        byte[] lfilecontents = File.ReadAllBytes(txtPicturesPath.Text + "\\" + lbrand + "\\" + lpicture);

                                        if (lfilecontents.Count<byte>() > 0)
                                        {
                                            txtStatus.Text += lsku + " has a local picture!\r\n";
                                            workSheet.Shapes.AddPicture(txtPicturesPath.Text + "\\" + lbrand + "\\" + lpicture,
                                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                                                        1, (float)ltop + 5,
                                                                        IMAGE_HEIGHT, IMAGE_HEIGHT);
                                            lwithpic = "LOCAL PIC | ";
                                        }
                                    }
                                    catch (Exception pe)
                                    {
                                        txtStatus.Text += "Local picture error " + pe.ToString() + "\r\n";
                                    }
                                }
                            }

                            // Let's see if we have it in our databases
                            lcmd = new SqlCommand("SELECT * FROM bsi_posting WHERE sku='" + lsku + "'", lconn);
                            lr = lcmd.ExecuteReader();
                            if (lr.Read())
                            {
                                lwithpic += "DESCR | ";

                                workSheet.Cells[lcurrRow, EXCEL_COLS_FULLDESCRIPTION] = removeSize(lr["fullDescription"].ToString());
                                workSheet.Cells[lcurrRow, EXCEL_COLS_CATEGORY] = lr["category"].ToString();

                                workSheet.Cells[lcurrRow, EXCEL_COLS_TYPE] = lr["style"].ToString();

                                workSheet.Cells[lcurrRow, EXCEL_COLS_MATERIAL] = lr["material"].ToString();
                                workSheet.Cells[lcurrRow, EXCEL_COLS_COLOR] = lr["color"].ToString();
                                workSheet.Cells[lcurrRow, EXCEL_COLS_SHADE] = lr["shade"].ToString();

                                // Let's see if we have a remote picture and if that is still available
                                /*
                                string lpictures = lr["pictures"].ToString();
                                if ( !String.IsNullOrEmpty(lpictures) )
                                {
                                    string[] lremotepix = lpictures.Split( new char[] { '|' });
                                    try {
                                        WebClient client = new WebClient();
                                        Stream stream = client.OpenRead(lremotepix[0]);
                                        if ( stream != null )
                                        {
                                            lwithpic += "REMOTE PIX";
                                            stream.Close();
                                        }
                                    }
                                    catch(Exception pe)
                                    {
                                        txtStatus.Text += "Error while getting the remote pic: " + pe.ToString() + "\r\n";
                                    }
                                }
                                */

                            }
                            lr.Close();
                            txtStatus.Text += lsku + " = " + lwithpic + "\r\n";
                        }

                        txtStatus.Update();
                        Application.DoEvents();
                    }
                    else lstop = true;
                    lcurrRow++;
                } // while

                object misValue = System.Reflection.Missing.Value;


                //theWorkbook.SaveAs(txtInventoryFile.Text,Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                theWorkbook.SaveCopyAs(lFullPath + "\\PIX_" + lFileName);
                theWorkbook.Save();
                theWorkbook.Close(true, misValue, misValue);
                excelApp.Quit();

                releaseObject(theWorkbook);
                releaseObject(excelApp);

                MessageBox.Show("PROCESS FINISHED WITH " + lcurrRow + " ROWS PROCESSED");
            }
            catch (Exception pe)
            {
                MessageBox.Show(pe.ToString(), "Error while processing");
            }
            finally
            {
                if (lconn != null) lconn.Close();
            }
        } // btnCheckMarketplaces_Click

    } // public partial class Form1 : Form
} // namespace RMS_AssignPictures2Inventory

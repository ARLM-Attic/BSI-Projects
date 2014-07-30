namespace RMS_AssignPictures2Inventory
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtInventoryFile = new System.Windows.Forms.TextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.btnSetPath = new System.Windows.Forms.Button();
            this.txtPicturesPath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnStop = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnChkSKUs = new System.Windows.Forms.Button();
            this.cmbMarketplaces = new System.Windows.Forms.ComboBox();
            this.btnRegisterQtys = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.label4 = new System.Windows.Forms.Label();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.btnVerifyOnEbay = new System.Windows.Forms.Button();
            this.txtSKU = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.btnCheckMarketplaces = new System.Windows.Forms.Button();
            this.chkCategory = new System.Windows.Forms.CheckBox();
            this.chkStyle = new System.Windows.Forms.CheckBox();
            this.chkMaterial = new System.Windows.Forms.CheckBox();
            this.chkColor = new System.Windows.Forms.CheckBox();
            this.chkShade = new System.Windows.Forms.CheckBox();
            this.chkGender = new System.Windows.Forms.CheckBox();
            this.chkDescription = new System.Windows.Forms.CheckBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Inventory file:";
            // 
            // txtInventoryFile
            // 
            this.txtInventoryFile.Location = new System.Drawing.Point(142, 22);
            this.txtInventoryFile.Name = "txtInventoryFile";
            this.txtInventoryFile.Size = new System.Drawing.Size(368, 26);
            this.txtInventoryFile.TabIndex = 1;
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(516, 22);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(90, 26);
            this.btnSearch.TabIndex = 2;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(318, 92);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(128, 27);
            this.btnStart.TabIndex = 3;
            this.btnStart.Text = "Start Process";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // txtStatus
            // 
            this.txtStatus.Location = new System.Drawing.Point(16, 263);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtStatus.Size = new System.Drawing.Size(586, 402);
            this.txtStatus.TabIndex = 4;
            // 
            // btnSetPath
            // 
            this.btnSetPath.Location = new System.Drawing.Point(356, 39);
            this.btnSetPath.Name = "btnSetPath";
            this.btnSetPath.Size = new System.Drawing.Size(90, 26);
            this.btnSetPath.TabIndex = 7;
            this.btnSetPath.Text = "Search";
            this.btnSetPath.UseVisualStyleBackColor = true;
            this.btnSetPath.Click += new System.EventHandler(this.btnSetPath_Click);
            // 
            // txtPicturesPath
            // 
            this.txtPicturesPath.Location = new System.Drawing.Point(10, 39);
            this.txtPicturesPath.Name = "txtPicturesPath";
            this.txtPicturesPath.Size = new System.Drawing.Size(340, 26);
            this.txtPicturesPath.TabIndex = 6;
            this.txtPicturesPath.Text = "P:\\products\\";
            this.txtPicturesPath.TextChanged += new System.EventHandler(this.txtPicturesPath_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 20);
            this.label2.TabIndex = 5;
            this.label2.Text = "Path to pictures:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnStop
            // 
            this.btnStop.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStop.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnStop.Location = new System.Drawing.Point(16, 77);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(120, 65);
            this.btnStop.TabIndex = 8;
            this.btnStop.Text = "Stop All Processes";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 240);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 20);
            this.label3.TabIndex = 9;
            this.label3.Text = "Status:";
            // 
            // btnChkSKUs
            // 
            this.btnChkSKUs.Location = new System.Drawing.Point(125, 39);
            this.btnChkSKUs.Name = "btnChkSKUs";
            this.btnChkSKUs.Size = new System.Drawing.Size(169, 32);
            this.btnChkSKUs.TabIndex = 10;
            this.btnChkSKUs.Text = "Start Checking SKUs";
            this.btnChkSKUs.UseVisualStyleBackColor = true;
            this.btnChkSKUs.Click += new System.EventHandler(this.btnChkSKUs_Click);
            // 
            // cmbMarketplaces
            // 
            this.cmbMarketplaces.FormattingEnabled = true;
            this.cmbMarketplaces.Items.AddRange(new object[] {
            "Amazon",
            "Mecalzo",
            "One Million Shoes"});
            this.cmbMarketplaces.Location = new System.Drawing.Point(115, 31);
            this.cmbMarketplaces.Name = "cmbMarketplaces";
            this.cmbMarketplaces.Size = new System.Drawing.Size(148, 28);
            this.cmbMarketplaces.TabIndex = 11;
            // 
            // btnRegisterQtys
            // 
            this.btnRegisterQtys.Location = new System.Drawing.Point(13, 74);
            this.btnRegisterQtys.Name = "btnRegisterQtys";
            this.btnRegisterQtys.Size = new System.Drawing.Size(214, 28);
            this.btnRegisterQtys.TabIndex = 12;
            this.btnRegisterQtys.Text = "Start reading quantities";
            this.btnRegisterQtys.UseVisualStyleBackColor = true;
            this.btnRegisterQtys.Click += new System.EventHandler(this.btnRegisterQtys_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Location = new System.Drawing.Point(142, 77);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(460, 158);
            this.tabControl1.TabIndex = 13;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.chkDescription);
            this.tabPage1.Controls.Add(this.chkGender);
            this.tabPage1.Controls.Add(this.chkShade);
            this.tabPage1.Controls.Add(this.chkColor);
            this.tabPage1.Controls.Add(this.chkMaterial);
            this.tabPage1.Controls.Add(this.chkStyle);
            this.tabPage1.Controls.Add(this.chkCategory);
            this.tabPage1.Controls.Add(this.btnStart);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txtPicturesPath);
            this.tabPage1.Controls.Add(this.btnSetPath);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(452, 125);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Assign pictures";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnChkSKUs);
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(452, 125);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Check SKUs in RMS";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.label4);
            this.tabPage3.Controls.Add(this.cmbMarketplaces);
            this.tabPage3.Controls.Add(this.btnRegisterQtys);
            this.tabPage3.Location = new System.Drawing.Point(4, 29);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(452, 125);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Read marketplace qtys";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 34);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 20);
            this.label4.TabIndex = 13;
            this.label4.Text = "Marketplace:";
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.btnVerifyOnEbay);
            this.tabPage4.Controls.Add(this.txtSKU);
            this.tabPage4.Controls.Add(this.label5);
            this.tabPage4.Location = new System.Drawing.Point(4, 29);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(452, 125);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Verify item on Mecalzo";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // btnVerifyOnEbay
            // 
            this.btnVerifyOnEbay.Location = new System.Drawing.Point(18, 57);
            this.btnVerifyOnEbay.Name = "btnVerifyOnEbay";
            this.btnVerifyOnEbay.Size = new System.Drawing.Size(148, 31);
            this.btnVerifyOnEbay.TabIndex = 2;
            this.btnVerifyOnEbay.Text = "Verify on eBay";
            this.btnVerifyOnEbay.UseVisualStyleBackColor = true;
            this.btnVerifyOnEbay.Click += new System.EventHandler(this.btnVerifyOnEbay_Click);
            // 
            // txtSKU
            // 
            this.txtSKU.Location = new System.Drawing.Point(66, 13);
            this.txtSKU.Name = "txtSKU";
            this.txtSKU.Size = new System.Drawing.Size(100, 26);
            this.txtSKU.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 16);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(46, 20);
            this.label5.TabIndex = 0;
            this.label5.Text = "SKU:";
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.btnCheckMarketplaces);
            this.tabPage5.Location = new System.Drawing.Point(4, 29);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(452, 125);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "Check marketplaces";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // btnCheckMarketplaces
            // 
            this.btnCheckMarketplaces.Location = new System.Drawing.Point(22, 13);
            this.btnCheckMarketplaces.Name = "btnCheckMarketplaces";
            this.btnCheckMarketplaces.Size = new System.Drawing.Size(145, 38);
            this.btnCheckMarketplaces.TabIndex = 0;
            this.btnCheckMarketplaces.Text = "Start Checking";
            this.btnCheckMarketplaces.UseVisualStyleBackColor = true;
            this.btnCheckMarketplaces.Click += new System.EventHandler(this.btnCheckMarketplaces_Click);
            // 
            // chkCategory
            // 
            this.chkCategory.AutoSize = true;
            this.chkCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkCategory.Location = new System.Drawing.Point(10, 72);
            this.chkCategory.Name = "chkCategory";
            this.chkCategory.Size = new System.Drawing.Size(74, 19);
            this.chkCategory.TabIndex = 8;
            this.chkCategory.Text = "Category";
            this.chkCategory.UseVisualStyleBackColor = true;
            // 
            // chkStyle
            // 
            this.chkStyle.AutoSize = true;
            this.chkStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkStyle.Location = new System.Drawing.Point(10, 92);
            this.chkStyle.Name = "chkStyle";
            this.chkStyle.Size = new System.Drawing.Size(52, 19);
            this.chkStyle.TabIndex = 9;
            this.chkStyle.Text = "Style";
            this.chkStyle.UseVisualStyleBackColor = true;
            // 
            // chkMaterial
            // 
            this.chkMaterial.AutoSize = true;
            this.chkMaterial.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkMaterial.Location = new System.Drawing.Point(90, 72);
            this.chkMaterial.Name = "chkMaterial";
            this.chkMaterial.Size = new System.Drawing.Size(71, 19);
            this.chkMaterial.TabIndex = 10;
            this.chkMaterial.Text = "Material";
            this.chkMaterial.UseVisualStyleBackColor = true;
            // 
            // chkColor
            // 
            this.chkColor.AutoSize = true;
            this.chkColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkColor.Location = new System.Drawing.Point(90, 92);
            this.chkColor.Name = "chkColor";
            this.chkColor.Size = new System.Drawing.Size(55, 19);
            this.chkColor.TabIndex = 11;
            this.chkColor.Text = "Color";
            this.chkColor.UseVisualStyleBackColor = true;
            // 
            // chkShade
            // 
            this.chkShade.AutoSize = true;
            this.chkShade.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkShade.Location = new System.Drawing.Point(167, 71);
            this.chkShade.Name = "chkShade";
            this.chkShade.Size = new System.Drawing.Size(62, 19);
            this.chkShade.TabIndex = 12;
            this.chkShade.Text = "Shade";
            this.chkShade.UseVisualStyleBackColor = true;
            // 
            // chkGender
            // 
            this.chkGender.AutoSize = true;
            this.chkGender.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkGender.Location = new System.Drawing.Point(167, 92);
            this.chkGender.Name = "chkGender";
            this.chkGender.Size = new System.Drawing.Size(67, 19);
            this.chkGender.TabIndex = 13;
            this.chkGender.Text = "Gender";
            this.chkGender.UseVisualStyleBackColor = true;
            // 
            // chkDescription
            // 
            this.chkDescription.AutoSize = true;
            this.chkDescription.Checked = true;
            this.chkDescription.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDescription.Location = new System.Drawing.Point(244, 71);
            this.chkDescription.Name = "chkDescription";
            this.chkDescription.Size = new System.Drawing.Size(88, 19);
            this.chkDescription.TabIndex = 14;
            this.chkDescription.Text = "Description";
            this.chkDescription.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(624, 677);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.txtStatus);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.txtInventoryFile);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.Text = "RMS - Toolbox";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.tabPage5.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtInventoryFile;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.Button btnSetPath;
        private System.Windows.Forms.TextBox txtPicturesPath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnChkSKUs;
        private System.Windows.Forms.ComboBox cmbMarketplaces;
        private System.Windows.Forms.Button btnRegisterQtys;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button btnVerifyOnEbay;
        private System.Windows.Forms.TextBox txtSKU;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.Button btnCheckMarketplaces;
        private System.Windows.Forms.CheckBox chkDescription;
        private System.Windows.Forms.CheckBox chkGender;
        private System.Windows.Forms.CheckBox chkShade;
        private System.Windows.Forms.CheckBox chkColor;
        private System.Windows.Forms.CheckBox chkMaterial;
        private System.Windows.Forms.CheckBox chkStyle;
        private System.Windows.Forms.CheckBox chkCategory;
    }
}


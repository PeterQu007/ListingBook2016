namespace ListingBook2016
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.button4 = this.Factory.CreateRibbonButton();
            this.buttonLoadSubject = this.Factory.CreateRibbonButton();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.editBoxClient = this.Factory.CreateRibbonEditBox();
            this.group9 = this.Factory.CreateRibbonGroup();
            this.chkBoxNewHomes = this.Factory.CreateRibbonCheckBox();
            this.chkBoxLanguage = this.Factory.CreateRibbonCheckBox();
            this.checkBoxSumTable = this.Factory.CreateRibbonCheckBox();
            this.Options2 = this.Factory.CreateRibbonGroup();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnDetachedSoldCMA = this.Factory.CreateRibbonButton();
            this.btnDetachedActiveCMA = this.Factory.CreateRibbonButton();
            this.btnAllCMA = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.btmCondoSold = this.Factory.CreateRibbonButton();
            this.btnCondoActive = this.Factory.CreateRibbonButton();
            this.btnCondoCMA = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.btnBuyerDetachedReport = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnTownhouseAllCities = this.Factory.CreateRibbonButton();
            this.btnTownhouseAllCommunities = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnDetachedAllCities = this.Factory.CreateRibbonButton();
            this.btnDetachedMonthlyCommunities = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group8.SuspendLayout();
            this.box1.SuspendLayout();
            this.group9.SuspendLayout();
            this.Options2.SuspendLayout();
            this.group4.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group8);
            this.tab1.Groups.Add(this.group9);
            this.tab1.Groups.Add(this.Options2);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group8
            // 
            this.group8.Items.Add(this.box1);
            this.group8.Items.Add(this.comboBox1);
            this.group8.Items.Add(this.editBoxClient);
            this.group8.Label = ".               Subject Property               .";
            this.group8.Name = "group8";
            // 
            // box1
            // 
            this.box1.Items.Add(this.button4);
            this.box1.Items.Add(this.buttonLoadSubject);
            this.box1.Name = "box1";
            // 
            // button4
            // 
            this.button4.Label = "ReConnect  To MySQL Database";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // buttonLoadSubject
            // 
            this.buttonLoadSubject.Label = "|    Load Subjects";
            this.buttonLoadSubject.Name = "buttonLoadSubject";
            this.buttonLoadSubject.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLoadSubject_Click);
            // 
            // comboBox1
            // 
            ribbonDropDownItemImpl1.Label = "1346 Napier Place";
            ribbonDropDownItemImpl2.Label = "1785 137A ST";
            this.comboBox1.Items.Add(ribbonDropDownItemImpl1);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl2);
            this.comboBox1.Label = "Addr";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.SizeString = "700000000000000000000000000000000";
            this.comboBox1.Text = null;
            this.comboBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox1_TextChanged);
            // 
            // editBoxClient
            // 
            this.editBoxClient.Label = "Client:";
            this.editBoxClient.Name = "editBoxClient";
            this.editBoxClient.SizeString = "700000000000000000000000000000000000";
            this.editBoxClient.Text = null;
            // 
            // group9
            // 
            this.group9.Items.Add(this.chkBoxNewHomes);
            this.group9.Items.Add(this.chkBoxLanguage);
            this.group9.Items.Add(this.checkBoxSumTable);
            this.group9.Label = "Options";
            this.group9.Name = "group9";
            // 
            // chkBoxNewHomes
            // 
            this.chkBoxNewHomes.Label = "New Homes";
            this.chkBoxNewHomes.Name = "chkBoxNewHomes";
            // 
            // chkBoxLanguage
            // 
            this.chkBoxLanguage.Label = "Chinese";
            this.chkBoxLanguage.Name = "chkBoxLanguage";
            // 
            // checkBoxSumTable
            // 
            this.checkBoxSumTable.Label = "Sum Table";
            this.checkBoxSumTable.Name = "checkBoxSumTable";
            // 
            // Options2
            // 
            this.Options2.Items.Add(this.checkBox1);
            this.Options2.Label = "Options";
            this.Options2.Name = "Options2";
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "CorCoe";
            this.checkBox1.Name = "checkBox1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnDetachedSoldCMA);
            this.group4.Items.Add(this.btnDetachedActiveCMA);
            this.group4.Items.Add(this.btnAllCMA);
            this.group4.Label = "Detached CMA";
            this.group4.Name = "group4";
            // 
            // btnDetachedSoldCMA
            // 
            this.btnDetachedSoldCMA.Label = "Sold";
            this.btnDetachedSoldCMA.Name = "btnDetachedSoldCMA";
            this.btnDetachedSoldCMA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDetachedSoldCMA_Click);
            // 
            // btnDetachedActiveCMA
            // 
            this.btnDetachedActiveCMA.Label = "Active";
            this.btnDetachedActiveCMA.Name = "btnDetachedActiveCMA";
            this.btnDetachedActiveCMA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDetachedActiveCMA_Click);
            // 
            // btnAllCMA
            // 
            this.btnAllCMA.Label = "CMA";
            this.btnAllCMA.Name = "btnAllCMA";
            this.btnAllCMA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDetachedAllCMA_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.btmCondoSold);
            this.group6.Items.Add(this.btnCondoActive);
            this.group6.Items.Add(this.btnCondoCMA);
            this.group6.Label = "Attached CMA";
            this.group6.Name = "group6";
            // 
            // btmCondoSold
            // 
            this.btmCondoSold.Label = "Sold";
            this.btmCondoSold.Name = "btmCondoSold";
            this.btmCondoSold.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btmCondoSold_Click);
            // 
            // btnCondoActive
            // 
            this.btnCondoActive.Label = "Active";
            this.btnCondoActive.Name = "btnCondoActive";
            this.btnCondoActive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCondoActive_Click);
            // 
            // btnCondoCMA
            // 
            this.btnCondoCMA.Label = "CMA";
            this.btnCondoCMA.Name = "btnCondoCMA";
            this.btnCondoCMA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCondoCMA_Click);
            // 
            // group7
            // 
            this.group7.Items.Add(this.btnBuyerDetachedReport);
            this.group7.Items.Add(this.button8);
            this.group7.Label = "Deals for the Buyer";
            this.group7.Name = "group7";
            // 
            // btnBuyerDetachedReport
            // 
            this.btnBuyerDetachedReport.Label = "Buyer Detached";
            this.btnBuyerDetachedReport.Name = "btnBuyerDetachedReport";
            this.btnBuyerDetachedReport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBuyerDetachedReport_Click);
            // 
            // button8
            // 
            this.button8.Label = "Buyer Attached";
            this.button8.Name = "button8";
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBuyerAttachedReport_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button1);
            this.group3.Items.Add(this.button2);
            this.group3.Label = "Condo Monthly";
            this.group3.Name = "group3";
            // 
            // button1
            // 
            this.button1.Label = "All Cities";
            this.button1.Name = "button1";
            // 
            // button2
            // 
            this.button2.Label = "All Communities";
            this.button2.Name = "button2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnTownhouseAllCities);
            this.group2.Items.Add(this.btnTownhouseAllCommunities);
            this.group2.Label = "Townhouse Monthly";
            this.group2.Name = "group2";
            // 
            // btnTownhouseAllCities
            // 
            this.btnTownhouseAllCities.Label = "All Cities";
            this.btnTownhouseAllCities.Name = "btnTownhouseAllCities";
            // 
            // btnTownhouseAllCommunities
            // 
            this.btnTownhouseAllCommunities.Label = "All Commnities";
            this.btnTownhouseAllCommunities.Name = "btnTownhouseAllCommunities";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDetachedAllCities);
            this.group1.Items.Add(this.btnDetachedMonthlyCommunities);
            this.group1.Label = "Detached Monthly";
            this.group1.Name = "group1";
            // 
            // btnDetachedAllCities
            // 
            this.btnDetachedAllCities.Label = "All Cities";
            this.btnDetachedAllCities.Name = "btnDetachedAllCities";
            this.btnDetachedAllCities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDetachedMonthlyCities_Click);
            // 
            // btnDetachedMonthlyCommunities
            // 
            this.btnDetachedMonthlyCommunities.Label = "All Communities";
            this.btnDetachedMonthlyCommunities.Name = "btnDetachedMonthlyCommunities";
            this.btnDetachedMonthlyCommunities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDetachedMonthlyCommunities_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.button3);
            this.group5.Items.Add(this.button6);
            this.group5.Items.Add(this.button7);
            this.group5.Label = "Specialty Report";
            this.group5.Name = "group5";
            // 
            // button3
            // 
            this.button3.Label = "Price Change Top10";
            this.button3.Name = "button3";
            // 
            // button6
            // 
            this.button6.Label = "Price Active Top10";
            this.button6.Name = "button6";
            // 
            // button7
            // 
            this.button7.Label = "Price Sold Top10";
            this.button7.Name = "button7";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group8.ResumeLayout(false);
            this.group8.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group9.ResumeLayout(false);
            this.group9.PerformLayout();
            this.Options2.ResumeLayout(false);
            this.Options2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDetachedAllCities;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDetachedMonthlyCommunities;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTownhouseAllCities;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTownhouseAllCommunities;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDetachedSoldCMA;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDetachedActiveCMA;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAllCMA;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btmCondoSold;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCondoActive;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCondoCMA;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBuyerDetachedReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group8;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkBoxLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkBoxNewHomes;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group9;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxClient;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSumTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Options2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLoadSubject;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}

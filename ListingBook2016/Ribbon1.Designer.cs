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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnDetachedAllCities = this.Factory.CreateRibbonButton();
            this.btnDetachedMonthlyCommunities = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnTownhouseAllCities = this.Factory.CreateRibbonButton();
            this.btnTownhouseAllCommunities = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.btmCondoSold = this.Factory.CreateRibbonButton();
            this.btnCondoActive = this.Factory.CreateRibbonButton();
            this.btnCondoCMA = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group6.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
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
            // group4
            // 
            this.group4.Items.Add(this.button3);
            this.group4.Items.Add(this.button4);
            this.group4.Items.Add(this.button5);
            this.group4.Label = "Detached CMA";
            this.group4.Name = "group4";
            // 
            // group5
            // 
            this.group5.Label = "Townhouse CMA";
            this.group5.Name = "group5";
            // 
            // group6
            // 
            this.group6.Items.Add(this.btmCondoSold);
            this.group6.Items.Add(this.btnCondoActive);
            this.group6.Items.Add(this.btnCondoCMA);
            this.group6.Label = "Condo CMA";
            this.group6.Name = "group6";
            // 
            // button3
            // 
            this.button3.Label = "Sold";
            this.button3.Name = "button3";
            // 
            // button4
            // 
            this.button4.Label = "Active";
            this.button4.Name = "button4";
            // 
            // button5
            // 
            this.button5.Label = "CMA";
            this.button5.Name = "button5";
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
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btmCondoSold;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCondoActive;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCondoCMA;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}

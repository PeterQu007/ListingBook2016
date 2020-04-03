using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace ListingBook2016
{
    public partial class Ribbon1
    {
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
        public string CMADataSheet = string.Empty;
        public string PublicReportDataSheet = string.Empty;
        public string ReportDataSheet = string.Empty;
        private string SubjectPropertyAdress = "1385 137A ST";
        private string SubjectUnitNo = "805";
        private int SubjectPropertyAge = 35;
        private int LandSize = 14113;
        private int FloorArea = 2200;
        private decimal BCAssessLand = 1421000;
        private decimal BCAssessImprove = 271000;
        private decimal BCAssessTotal = 1692000;
        private string PropertyType;
        private decimal MaintenanceFee;
        private string City;
        private string Neighborhood;
        private int CMAAction;
        private int SubjectID;

        private DBConnection dbCon;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            CMADataSheet = ListingDataSheet.MLSHelperExport;
            PublicReportDataSheet = ListingDataSheet.ParagonExport;
            dbCon = DBConnection.Instance();
            dbCon.DatabaseName = "local";
            if (dbCon.IsConnect())
            {
                //suppose col0 and col1 are defined as VARCHAR in the DB
                comboBox1.Items.Clear();

                string query = "select * From pid_cma_subjects Where CMA_Action = 1";
                var cmd = new MySqlCommand(query, dbCon.Connection);
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    SubjectPropertyAdress = reader.IsDBNull(1) ? "" : reader.GetString("Subject_Address");
                    SubjectUnitNo = reader.IsDBNull(2) ? "" : reader.GetString("Unit_No");
                    SubjectPropertyAge = reader.IsDBNull(3) ? 0 : reader.GetInt16("Age");
                    RibbonDropDownItem ddItem1 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    ddItem1.Label = SubjectPropertyAdress;
                    comboBox1.Items.Add(ddItem1);
                }
                reader.Close();
                dbCon.Close();
            }

        }

        #region MARKETING.REPORTS
        //PUBLIC.MARKETING.REPORTS
        private void btnDetachedMonthlyCities_Click(object sender, RibbonControlEventArgs e)
        {
            ReportMonthlyDetached report = new ReportMonthlyDetached(Globals.ThisAddIn.Application.Worksheets[PublicReportDataSheet]);
            report.ListingSheet.Activate();
            report.ListingSheet.AutoFilterMode = false;

            report.AllCities();
        }

        private void btnDetachedMonthlyCommunities_Click(object sender, RibbonControlEventArgs e)
        {
            ReportMonthlyDetached report = new ReportMonthlyDetached(Globals.ThisAddIn.Application.Worksheets[PublicReportDataSheet]);
            report.ListingSheet.Activate();
            report.ListingSheet.AutoFilterMode = false;

            report.AllCommunities();
        }
        #endregion

        #region RESIDENTIAL.ATTCHED.PROPERTY.CMA.REPORTS
        //RESIDENTIAL.ATTACHED.PROPERTY
        private void btmCondoSold_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMAAttached);
            cma.Residential(ListingStatus.Sold);
        }

        private void btnCondoActive_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMAAttached);
            cma.Residential(ListingStatus.Active);
        }

        private void btnCondoCMA_Click(object sender, RibbonControlEventArgs e)
        {
            if( string.IsNullOrEmpty(comboBox1.Text))
            {
                string caption = "Error Detected in Input";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                MessageBox.Show("Please Select Subject Property", caption, buttons);
                return;
            }
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMAAttached, chkBoxLanguage.Checked ? "Chinese" : "English");
            if (cma.ListingDataValidated)
            {
                try { cma.Residential(ListingStatus.Sold); } catch (Exception ex) { Debug.Write(ex); };
                try { cma.Residential(ListingStatus.Active); } catch (Exception ex) { Debug.Write(ex); };
                try { cma.Residential(ListingStatus.OffMarket, true); } catch (Exception ex) { Debug.Write(ex); };
            }
        }
        #endregion

        #region RESIDENTIAL.DETACHED.PROPERTY.CMA.REPORTS
        //RESIDENTIAL.DETACHED
        //SOLD
        private void btnDetachedSoldCMA_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMADetached);
            cma.Residential(ListingStatus.Sold);
        }
        //ACTIVE
        private void btnDetachedActiveCMA_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMADetached);
            cma.Residential(ListingStatus.Active);
        }
        //SOLD, ACTIVE. OFFMARKET
        private void btnDetachedAllCMA_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox1.Text))
            {
                string caption = "Error Detected in Input";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                MessageBox.Show("Please Select Subject Property", caption, buttons);
                return;
            }
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMADetached);
            if (cma.ListingDataValidated)
            {
                try { cma.Residential(ListingStatus.Sold); } catch (Exception ex) { Debug.Write(ex); };
                try { cma.Residential(ListingStatus.Active); } catch (Exception ex) { Debug.Write(ex); };
                try { cma.Residential(ListingStatus.OffMarket, true); } catch (Exception ex) { Debug.Write(ex); };
            }
        }


        #endregion

        #region BUYER.REPORTS
        //BUYER.REPORTS
        private void btnBuyerDetachedReport_Click(object sender, RibbonControlEventArgs e)
        {
            ReportBuyer buyerReport = new ReportBuyer(Globals.ThisAddIn.Application.Worksheets[ReportDataSheet], ReportType.CMADetached);
            try { buyerReport.Residential(ListingStatus.Active); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.Sold); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.OffMarket); } catch (Exception ex) { throw ex; };
        }

        private void btnBuyerAttachedReport_Click(object sender, RibbonControlEventArgs e)
        {
            ReportBuyer buyerReport = new ReportBuyer(Globals.ThisAddIn.Application.Worksheets[ReportDataSheet], ReportType.CMAAttached);
            try { buyerReport.Residential(ListingStatus.Active); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.Sold); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.OffMarket); } catch (Exception ex) { throw ex; };
        }
        #endregion

        private void buttonLoadSubject_Click(object sender, RibbonControlEventArgs e)
        {
            //dbCon
            if (dbCon.IsConnect())
            {
                //suppose col0 and col1 are defined as VARCHAR in the DB
                Excel.Worksheet SubjectsWorkSheet = null;
                string query = "select * From pid_cma_subjects Where CMA_Action = 1";
                var cmd = new MySqlCommand(query, dbCon.Connection);
                var reader = cmd.ExecuteReader();
                if (Library.SheetExist("Subjects"))
                {
                    SubjectsWorkSheet = Globals.ThisAddIn.Application.Worksheets["Subjects"];
                    SubjectsWorkSheet.Cells.Clear();
                }
                else
                {
                    SubjectsWorkSheet = Globals.ThisAddIn.Application.Worksheets.Add();
                    SubjectsWorkSheet.Name = "Subjects";
                    SubjectsWorkSheet.Activate();
                }
                SubjectsWorkSheet.Cells[1, 1] = "ID";
                SubjectsWorkSheet.Cells[1, 2] = "Address";
                SubjectsWorkSheet.Cells[1, 3] = "Unit NO";
                SubjectsWorkSheet.Cells[1, 4] = "Age";
                SubjectsWorkSheet.Cells[1, 5] = "Land Size";
                SubjectsWorkSheet.Cells[1, 6] = "Floor Area";
                SubjectsWorkSheet.Cells[1, 7] = "BC Assess Land";
                SubjectsWorkSheet.Cells[1, 8] = "BC Assess Improve";
                SubjectsWorkSheet.Cells[1, 9] = "BC_Assess_Total";
                SubjectsWorkSheet.Cells[1, 10] = "Property Type";
                SubjectsWorkSheet.Cells[1, 11] = "Maintenance Fee";
                SubjectsWorkSheet.Cells[1, 12] = "City";
                SubjectsWorkSheet.Cells[1, 13] = "Neighborhood";
                SubjectsWorkSheet.Cells[1, 14] = "CMA Action";
                var i = 2;

                while (reader.Read())
                {
                    SubjectID= reader.IsDBNull(1) ? 0 : reader.GetInt16("ID");
                    SubjectPropertyAdress = reader.IsDBNull(1) ? "" : reader.GetString("Subject_Address");
                    SubjectUnitNo = reader.IsDBNull(2) ? "" : reader.GetString("Unit_No");
                    SubjectPropertyAge = reader.IsDBNull(3) ? 0 : reader.GetInt16("Age");
                    LandSize = reader.IsDBNull(4) ? 0 : reader.GetInt16("Land_Size");
                    FloorArea = reader.IsDBNull(5) ? 0 : reader.GetInt16("Floor_Area");
                    BCAssessLand = reader.IsDBNull(6) ? 0 : reader.GetDecimal("BC_Assess_Land");
                    BCAssessImprove = reader.IsDBNull(7) ? 0 : reader.GetDecimal("BC_Assess_Improve");
                    BCAssessTotal = reader.IsDBNull(8) ? 0 : reader.GetDecimal("BC_Assess_Total");
                    PropertyType = reader.IsDBNull(15) ? "" : reader.GetString("Subject_Property_Type");
                    MaintenanceFee = reader.IsDBNull(16) ? 0 : reader.GetInt16("Maintenance_Fee");
                    City = reader.IsDBNull(17) ? "" : reader.GetString("City");
                    Neighborhood = reader.IsDBNull(18) ? "" : reader.GetString("Neighborhood");
                    CMAAction = reader.IsDBNull(19) ? 0 : reader.GetInt16("CMA_Action");

                    SubjectsWorkSheet.Cells[i, 1] = SubjectID;
                    SubjectsWorkSheet.Cells[i, 2] = SubjectPropertyAdress;
                    SubjectsWorkSheet.Cells[i, 3] = SubjectUnitNo;
                    SubjectsWorkSheet.Cells[i, 4] = SubjectPropertyAge;
                    SubjectsWorkSheet.Cells[i, 5] = LandSize;
                    SubjectsWorkSheet.Cells[i, 6] = FloorArea;
                    SubjectsWorkSheet.Cells[i, 7] = BCAssessLand;
                    SubjectsWorkSheet.Cells[i, 8] = BCAssessImprove;
                    SubjectsWorkSheet.Cells[i, 9] = BCAssessTotal;
                    SubjectsWorkSheet.Cells[i, 10] = PropertyType;
                    SubjectsWorkSheet.Cells[i, 11] = MaintenanceFee;
                    SubjectsWorkSheet.Cells[i, 12] = City;
                    SubjectsWorkSheet.Cells[i, 13] = Neighborhood;
                    SubjectsWorkSheet.Cells[i, 14] = CMAAction;

                    i++;
                }
                reader.Close();
                dbCon.Close();
            }

            //sql query

            //read records

            //ad worksheet Subjects


        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Debug.Write("TEST");
        }
    }
}

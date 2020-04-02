using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace ListingBook2016
{
    public class ReportCMA
    {
        public Excel.Worksheet ListingSheet;
        public Excel.Worksheet PivotSheet;
        private bool PivotSheetAdded;
        private DataProcessing dp;
        public bool ListingDataValidated = true;
        private ReportType CMAReportType;
        private PropertyType CMAPropertyType;
        private string CMAReportLanguage;
        private bool TitleAdded;
        private bool SubTitleAdded;
        private bool SubjectEvaluationAdded;
        private string SubjectPropertyAdress = "1385 137A ST";
        private string SubjectUnitNo = "805";
        private int SubjectPropertyAge = 35;
        private int LandSize = 14113;
        private int FloorArea = 2200;
        private decimal BCAssessLand = 1421000;
        private decimal BCAssessImprove = 271000;
        private decimal BCAssessTotal = 1692000;
        private decimal AvgLandSoldPricePerSF = 105;
        private decimal AvgImproveSoldPricePerSF = 93;
        private decimal AvgTotalSoldPricePerSF = 93;
        private decimal HiLandSoldPricePerSF = 115;
        private decimal HiImproveSoldPricePerSF = 122;
        private decimal HiTotalSoldPricePerSF = 560;
        //private decimal LoLandPricePerSF;
        //private decimal LoImprovePricePerSF;
        //private decimal MedLandPricePerSF;
        //private decimal MedImprovePricePerSF;
        private DBConnection dbCon;

        public ReportCMA(Excel.Worksheet ws, ReportType cmaType, string cmaLang = "English")
        {
            dbCon = DBConnection.Instance();
            dbCon.DatabaseName = "local";
            if (dbCon.IsConnect())
            {
                //suppose col0 and col1 are defined as VARCHAR in the DB
                string query = "select * From pid_cma_subjects Where CMA_Action = 1";
                var cmd = new MySqlCommand(query, dbCon.Connection);
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    SubjectPropertyAdress = reader.IsDBNull(1) ? "" : reader.GetString("Subject_Address");
                    SubjectUnitNo = reader.IsDBNull(2) ? "" : reader.GetString("Unit_No");
                    SubjectPropertyAge = reader.IsDBNull(3) ? 0 :  reader.GetInt16("Age");
                    LandSize = reader.IsDBNull(4) ? 0 : reader.GetInt16("Land_Size");
                    FloorArea = reader.IsDBNull(5) ? 0 : reader.GetInt16("Floor_Area");
                    BCAssessLand = reader.IsDBNull(6) ? 0 : reader.GetDecimal("BC_Assess_Land");
                    BCAssessImprove = reader.IsDBNull(7) ? 0 : reader.GetDecimal("BC_Assess_Improve");
                    BCAssessTotal = reader.IsDBNull(8) ? 0 : reader.GetDecimal("BC_Assess_Total");
                }
                reader.Close();
                dbCon.Close();
            }
            this.ListingSheet = ws;
            dp = new DataProcessing(ws);
            this.CMAReportType = cmaType;
            switch (cmaType)
            {
                case ReportType.CMADetached:
                    this.CMAPropertyType = PropertyType.Detached;
                    break;
                case ReportType.CMAAttached:
                    this.CMAPropertyType = PropertyType.Attached;
                    break;
                default:
                    this.CMAPropertyType = PropertyType.Detached;
                    break;
            }
            this.CMAReportLanguage = cmaLang;
            this.TitleAdded = false;
            this.SubTitleAdded = false;
            this.SubjectEvaluationAdded = false;
            this.PivotSheetAdded = false;
            string PivotSheetName = "PivotSheet";

            //DATA VALIDATE
            ListingSheet.Activate();
            if (dp.ValidatePropertyType(CMAReportType, CMAPropertyType))
            {
                Debug.Write("The property type in the listing sheet is wrong!");
                ListingDataValidated = false;
                return;
            }
            if (dp.ValidateData(CMAReportType))
            {
                Debug.Write("Listing Data Needs To Be Reviewed");
                ListingDataValidated = false;
                return;
            }
            //DATA ADD LOT AND IMPROVE UNIT PRICE AS PER BCA PERCENTAGE%
            dp.AddLotAndImproveUnitPrice();

            if (Library.SheetExist(PivotSheetName))
            {
                Globals.ThisAddIn.Application.Worksheets[PivotSheetName].Application.DisplayAlerts = false;
                Globals.ThisAddIn.Application.Worksheets[PivotSheetName].Delete();
            }
            Excel.Worksheet NewSheet = Globals.ThisAddIn.Application.Worksheets.Add();
            NewSheet.Name = PivotSheetName;
        }

        public void Residential( ListingStatus Status, bool AddSumTable = false)
        {
            Residential(Status);
            if (AddSumTable)
            {
                switch (CMAPropertyType)
                {
                    case PropertyType.Attached:
                        this.AddCMASubjectEvaluation_Attached(this.PivotSheet);
                        break;
                    case PropertyType.Detached:
                        this.AddCMASubjectEvaluation_Detached(this.PivotSheet);
                        break;
                }
            }
        }
        public void Residential(ListingStatus Status)
        {
            int PivotTableTopPaddingRows = 5;
            string PivotSheetName = "PivotSheet";
            string PivotTableName = "";
            PivotTableCMA ptCMA = null;

            /////////////////////
            //PIVOT TABLE
            PivotTableName = "PivotTable_" + Status;
            Globals.ThisAddIn.Application.ScreenUpdating = true;

            ptCMA = new PivotTableCMA(PivotSheetName, PivotTableTopPaddingRows, PivotTableName, Status, CMAReportType);
            if (ptCMA.ListingDataRows <= 0)
            {
                try
                {
                    this.PivotSheet.Select();
                }
                catch (Exception ex)
                {
                    Debug.Write(ex);
                    throw;
                }
                return;
            }

            ptCMA.Create();
            this.PivotSheet = ptCMA.PivotSheet;
            ptCMA.PivotSheet.Select();
            if (!ptCMA.FormatColumnsWidthDone)
            {
                ptCMA.Format(ptCMA.PivotSheet, PivotTableName, Status, "");
                if(Status == ListingStatus.Sold)
                {
                    this.HiImproveSoldPricePerSF = ptCMA.MaxSoldPricePerSF_Improve;
                    this.HiLandSoldPricePerSF = ptCMA.MaxSoldPricePerSF_Land;
                    this.HiTotalSoldPricePerSF = ptCMA.MaxSoldPricePerSF_Total;
                    this.AvgImproveSoldPricePerSF = ptCMA.AverageSoldPricePerSF_Improve;
                    this.AvgLandSoldPricePerSF = ptCMA.AverageSoldPricePerSF_Land;
                    this.AvgTotalSoldPricePerSF = ptCMA.AverageSoldPricePerSF_Total;
                }
            }
            ptCMA.AddMedianSummary(ptCMA.PivotSheet, PivotTableName, Status);
            //Globals.ThisAddIn.Application.ScreenUpdating = true;
            ptCMA.AddCorCoeSummary_Attached(ptCMA.PivotSheet, ListingSheet);
            ptCMA.AddDisclaimer(ptCMA.PivotSheet);
            this.AddCMATitle(PivotSheet, this.CMAReportLanguage == "English" ? "CMA REPORT" : "CMA 物业评估报告");
            this.AddCMASubTitle(PivotSheet, "Peter Qu");
            //this.AddCMASubjectEvaluation(PivotSheet);
            //Excel.Range line = (Excel.Range)ptCMA.PivotSheet.Rows[3];
            //line.Select();
            //line.Insert();
            //line = (Excel.Range)ptCMA.PivotSheet.Rows[4,5];
            //line.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            //line.Select();
        }
        public void AddCMATitle(Excel.Worksheet WS, string Title)
        {
            if (this.TitleAdded) return;
            long LastCol = Library.GetLastCol(this.PivotSheet);
            Excel.Range cell1 = WS.Cells[1, 1];
            Excel.Range cell2 = WS.Cells[1, LastCol - 1];
            WS.Range[cell1, cell2].Merge();

            if(this.CMAPropertyType == PropertyType.Attached)
            {
                cell1.Value = this.SubjectUnitNo + "-" + this.SubjectPropertyAdress + " " + Title;
            }
            else
            {
                cell1.Value = this.SubjectPropertyAdress + " " + Title;
            }
            cell1.Font.Size = 28;
            cell1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); 
            cell1.Font.Bold = true;
            cell1.Font.Italic = false;
            cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell1.EntireRow.RowHeight = 58;
            this.TitleAdded = true;
        }

        public void AddCMASubTitle(Excel.Worksheet WS, string SubTitle)
        {
            if (this.SubTitleAdded) return;
            long LastCol = Library.GetLastCol(this.PivotSheet);
            Excel.Range cell = WS.Cells[2, 1];
            Excel.Range cell2 = WS.Cells[2, LastCol - 1];
            WS.Range[cell, cell2].Merge();

            cell.Value = SubTitle + " " + System.DateTime.Now.Date.ToShortDateString();
            cell.Font.Size = 14;
            cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            cell.Font.Bold = false;
            cell.Font.Italic = false;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell.EntireRow.RowHeight = 42;
            this.SubTitleAdded = true;
        }

        #region Add CMA Subject Evaluation Table
        public void AddCMASubjectEvaluation_Detached(Excel.Worksheet pivotSheet)
        {
            if (this.SubjectEvaluationAdded) return;
            this.SubjectEvaluationAdded = true;
            //Insert new rows for the sum table
            Excel.Range line = (Excel.Range)pivotSheet.Rows[3];
            line.Select();
            line.Insert();
            Excel.Range line2 = (Excel.Range)pivotSheet.Rows["3:5"];
            line2.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            line2.Select();
            line2 = (Excel.Range)pivotSheet.Rows["3:6"];
            line2.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            line2.Select();
            //Creat tabel heading
            Excel.Range cellBox1 = pivotSheet.Range["A3", "B3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Subject Property Address";
            cellBox1 = pivotSheet.Range["A4", "B4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.SubjectPropertyAdress + "(" + this.SubjectPropertyAge.ToString() + " years)" ;
            //Land Size
            cellBox1 = pivotSheet.Range["C3", "D3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Land Size";
            cellBox1 = pivotSheet.Range["C4", "D4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.LandSize;
            //Floor Area
            cellBox1 = pivotSheet.Range["E3", "F3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Floor Area";
            cellBox1 = pivotSheet.Range["E4", "F4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.FloorArea;
            //BC Assess Land
            cellBox1 = pivotSheet.Range["G3", "H3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "BC Assess Land";
            cellBox1 = pivotSheet.Range["G4", "H4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.BCAssessLand;
            //BC Assess Improve.
            cellBox1 = pivotSheet.Range["I3", "J3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "BC Assess Improve.";
            cellBox1 = pivotSheet.Range["I4", "J4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.BCAssessImprove;
            //Total Value
            cellBox1 = pivotSheet.Range["K3", "L3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Total Value";
            cellBox1 = pivotSheet.Range["K4", "L4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=G4+I4";

            //Change% to BCA
            cellBox1 = pivotSheet.Range["M3", "N3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Change% to BCA";
            cellBox1 = pivotSheet.Range["M4", "N4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //empty cell
            //Price Per SF
            cellBox1 = pivotSheet.Range["O3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Price Per SF";
            cellBox1 = pivotSheet.Range["O4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=K4/E4";
            //Average and Highest Box
            cellBox1 = pivotSheet.Range["A5", "A6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Average:";
            cellBox1 = pivotSheet.Range["B5"];
            cellBox1.Select();
            cellBox1.Value = "Pirce /SF";
            cellBox1 = pivotSheet.Range["B6"];
            cellBox1.Select();
            cellBox1.Value = "Evaluation";
            cellBox1 = pivotSheet.Range["C5", "D5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.AvgLandSoldPricePerSF;
            cellBox1 = pivotSheet.Range["C6", "D6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=C4*C5";
            cellBox1 = pivotSheet.Range["E5", "F5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.AvgImproveSoldPricePerSF;
            cellBox1 = pivotSheet.Range["E6", "F6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=E4*E5";
            cellBox1 = pivotSheet.Range["G5", "H5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=G4/C4"; //Average Price per SF of BCA Land
            cellBox1 = pivotSheet.Range["G6", "L6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=C6+E6"; //Average Price per SF of BCA Land
            cellBox1 = pivotSheet.Range["I5", "J5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=I4/E4"; //Average Price per SF of BCA Improve.
            cellBox1 = pivotSheet.Range["I6", "J6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=I4"; //Average Price per SF of BCA Improve.
            cellBox1 = pivotSheet.Range["K5", "L5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Total Value
            cellBox1 = pivotSheet.Range["M5", "N5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Change % to BCA
            cellBox1 = pivotSheet.Range["M6", "N6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=(G6-K4)/K4"; //Change % to BCA
            cellBox1 = pivotSheet.Range["O6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=G6/E4"; //Price Per SF

            cellBox1 = pivotSheet.Range["A7", "A8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Highest:";
            cellBox1 = pivotSheet.Range["B7"];
            cellBox1.Select();
            cellBox1.Value = "Pirce /SF";
            cellBox1 = pivotSheet.Range["B8"];
            cellBox1.Select();
            cellBox1.Value = "Evaluation";
            cellBox1 = pivotSheet.Range["C7", "D7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.HiLandSoldPricePerSF;
            cellBox1 = pivotSheet.Range["C8", "D8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=C4*C7";
            cellBox1 = pivotSheet.Range["E7", "F7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.HiImproveSoldPricePerSF;
            cellBox1 = pivotSheet.Range["E8", "F8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=E4*E7";
            cellBox1 = pivotSheet.Range["G7", "H7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=G4/C4"; //Average Price per SF of BCA Land
            cellBox1 = pivotSheet.Range["G8", "L8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=C8+E8"; //Average Price per SF of BCA Land
            cellBox1 = pivotSheet.Range["I7", "J7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=I4/E4"; //Average Price per SF of BCA Improve.
            cellBox1 = pivotSheet.Range["I8", "J8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=I4"; //Average Price per SF of BCA Improve.
            cellBox1 = pivotSheet.Range["K7", "L7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Total Value
            cellBox1 = pivotSheet.Range["M7", "N7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Change % to BCA
            cellBox1 = pivotSheet.Range["M8", "N8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=(G8-K4)/K4"; //Change % to BCA
            cellBox1 = pivotSheet.Range["O8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=G8/E4"; //Price Per SF
            //Sub Table Footer
            cellBox1 = pivotSheet.Range["A9", "O9"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Market Valuation Range:";
            cellBox1 = pivotSheet.Range["A10", "O10"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            cellBox1.EntireRow.RowHeight = 18;
            cellBox1 = pivotSheet.Range["A11", "O11"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Comparable Criteria:";

            //format the Header cells:
            cellBox1 = pivotSheet.Range["A3", "O3"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.EntireRow.RowHeight = 38;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format the Subject Info Row:
            cellBox1 = pivotSheet.Range["A4", "O4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.EntireRow.RowHeight = 32;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //Address
            cellBox1 = pivotSheet.Range["A4", "B4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.NumberFormat = "$#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //Land Size:
            cellBox1 = pivotSheet.Range["C4", "D4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.NumberFormat = "#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format Values Cells
            cellBox1 = pivotSheet.Range["C5", "L8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.NumberFormat = "$#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format Floor Area
            cellBox1 = pivotSheet.Range["E4", "F4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.NumberFormat = "#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format BC Assess Land
            cellBox1 = pivotSheet.Range["G4", "H8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format BC Assess Improve.
            cellBox1 = pivotSheet.Range["I4", "J4"];
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue);
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format total Value:
            cellBox1 = pivotSheet.Range["K4", "L4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Change % to BCA:
            cellBox1 = pivotSheet.Range["M4", "N8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.NumberFormat = "##.00%;[RED](##.00%)";
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            //cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Price Per SF:
            cellBox1 = pivotSheet.Range["O4", "O8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format footer
            cellBox1 = pivotSheet.Range["A9"];
            cellBox1.Font.Size = 26;
            cellBox1.EntireRow.RowHeight = 42;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.Font.Bold = false;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            cellBox1.WrapText = true;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Comparable Criteria:
            cellBox1 = pivotSheet.Range["A11", "O11"];
            cellBox1.Font.Size = 24;
            cellBox1.EntireRow.RowHeight = 42;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Chocolate); ;
            cellBox1.Font.Bold = false;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            cellBox1.WrapText = true;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Average / Highests Boxes:
            cellBox1 = pivotSheet.Range["A5", "B8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            cellBox1.EntireRow.RowHeight = 22;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            //add border double line
            cellBox1 = pivotSheet.Range["A3", "O8"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            cellBox1 = pivotSheet.Range["A3", "O3"];
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); 
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            cellBox1 = pivotSheet.Range["A4", "O4"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;

            cellBox1 = pivotSheet.Range["A6", "O6"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1 = pivotSheet.Range["B6", "O6"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1 = pivotSheet.Range["B8", "O8"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
        }

        public void AddCMASubjectEvaluation_Attached(Excel.Worksheet pivotSheet)
        {
            if (this.SubjectEvaluationAdded) return;
            this.SubjectEvaluationAdded = true;
            //Insert new rows for the sum table
            Excel.Range line = (Excel.Range)pivotSheet.Rows[3];
            line.Select();
            line.Insert();
            Excel.Range line2 = (Excel.Range)pivotSheet.Rows["3:5"];
            line2.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            line2.Select();
            line2 = (Excel.Range)pivotSheet.Rows["3:6"];
            line2.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            line2.Select();
            //Creat tabel heading
            Excel.Range cellBox1 = pivotSheet.Range["A3", "B3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Subject Property Address";
            cellBox1 = pivotSheet.Range["A4", "B4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.SubjectUnitNo + "-" + this.SubjectPropertyAdress + "(" + this.SubjectPropertyAge.ToString() + " years)";
            //Land Size
            cellBox1 = pivotSheet.Range["C3", "D3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Land Size";
            cellBox1 = pivotSheet.Range["C4", "D4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.LandSize;
            //Floor Area
            cellBox1 = pivotSheet.Range["E3", "F3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Floor Area";
            cellBox1 = pivotSheet.Range["E4", "F4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.FloorArea;
            //BC Assess Land
            cellBox1 = pivotSheet.Range["G3", "H3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "BC Assess Land";
            cellBox1 = pivotSheet.Range["G4", "H4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            //BC Assess Improve.
            cellBox1 = pivotSheet.Range["I3", "J3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "BC Assess Improve.";
            cellBox1 = pivotSheet.Range["I4", "J4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            //Total Value
            cellBox1 = pivotSheet.Range["K3", "L3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Total Value";
            cellBox1 = pivotSheet.Range["K4", "L4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.BCAssessTotal;

            //Change% to BCA
            cellBox1 = pivotSheet.Range["M3", "N3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Change% to BCA";
            cellBox1 = pivotSheet.Range["M4", "N4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //empty cell
            //Price Per SF
            cellBox1 = pivotSheet.Range["O3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Price Per SF";
            cellBox1 = pivotSheet.Range["O4"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=K4/E4";
            //Average and Highest Box
            cellBox1 = pivotSheet.Range["A5", "A6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Average:";
            cellBox1 = pivotSheet.Range["B5"];
            cellBox1.Select();
            cellBox1.Value = "Pirce /SF";
            cellBox1 = pivotSheet.Range["B6"];
            cellBox1.Select();
            cellBox1.Value = "Evaluation";
            cellBox1 = pivotSheet.Range["C5", "D5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            cellBox1 = pivotSheet.Range["C6", "D6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            cellBox1 = pivotSheet.Range["E5", "F5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.AvgTotalSoldPricePerSF;
            cellBox1 = pivotSheet.Range["E6", "F6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=E4*E5";
            cellBox1 = pivotSheet.Range["G5", "H5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Average Price per SF of BCA Land
            cellBox1 = pivotSheet.Range["G6", "L6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=E6"; //Total Value
            cellBox1 = pivotSheet.Range["I5", "J5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //
            cellBox1 = pivotSheet.Range["I6", "J6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=I4"; //Average Price per SF of BCA Improve.
            cellBox1 = pivotSheet.Range["K5", "L5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Total Value
            cellBox1 = pivotSheet.Range["M5", "N5"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Change % to BCA
            cellBox1 = pivotSheet.Range["M6", "N6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=(G6-K4)/K4"; //Change % to BCA
            cellBox1 = pivotSheet.Range["O6"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=G6/E4"; //Price Per SF

            cellBox1 = pivotSheet.Range["A7", "A8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Highest:";
            cellBox1 = pivotSheet.Range["B7"];
            cellBox1.Select();
            cellBox1.Value = "Pirce /SF";
            cellBox1 = pivotSheet.Range["B8"];
            cellBox1.Select();
            cellBox1.Value = "Evaluation";
            cellBox1 = pivotSheet.Range["C7", "D7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            cellBox1 = pivotSheet.Range["C8", "D8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            cellBox1 = pivotSheet.Range["E7", "F7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = this.HiTotalSoldPricePerSF;
            cellBox1 = pivotSheet.Range["E8", "F8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=E4*E7";
            cellBox1 = pivotSheet.Range["G7", "H7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Average Price per SF of BCA Land
            cellBox1 = pivotSheet.Range["G8", "L8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=E8"; //Total Value
            cellBox1 = pivotSheet.Range["I7", "J7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Average Price per SF of BCA Improve.
            //cellBox1 = pivotSheet.Range["I8", "J8"];
            //cellBox1.Select();
            //cellBox1.Merge();
            //cellBox1.Value = "=I4"; //.
            cellBox1 = pivotSheet.Range["K7", "L7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Total Value
            cellBox1 = pivotSheet.Range["M7", "N7"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = ""; //Change % to BCA
            cellBox1 = pivotSheet.Range["M8", "N8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=(G8-K4)/K4"; //Change % to BCA
            cellBox1 = pivotSheet.Range["O8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "=G8/E4"; //Price Per SF
            //Sub Table Footer
            cellBox1 = pivotSheet.Range["A9", "O9"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Market Valuation Range:";
            cellBox1 = pivotSheet.Range["A10", "O10"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "";
            cellBox1.EntireRow.RowHeight = 18;
            cellBox1 = pivotSheet.Range["A11", "O11"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Comparable Criteria:";

            //format the Header cells:
            cellBox1 = pivotSheet.Range["A3", "O3"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.EntireRow.RowHeight = 38;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format the Subject Info Row:
            cellBox1 = pivotSheet.Range["A4", "O4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.EntireRow.RowHeight = 32;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //Address
            cellBox1 = pivotSheet.Range["A4", "B4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.NumberFormat = "$#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlDouble;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //Land Size:
            cellBox1 = pivotSheet.Range["C4", "D4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.NumberFormat = "#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format Values Cells
            cellBox1 = pivotSheet.Range["C5", "L8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.NumberFormat = "$#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format Floor Area
            cellBox1 = pivotSheet.Range["E4", "F4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.NumberFormat = "#,###";
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format BC Assess Land
            cellBox1 = pivotSheet.Range["G4", "H8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //format BC Assess Improve.
            cellBox1 = pivotSheet.Range["I4", "J4"];
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue);
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format total Value:
            cellBox1 = pivotSheet.Range["K4", "L4"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Change % to BCA:
            cellBox1 = pivotSheet.Range["M4", "N8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.NumberFormat = "##.00%;[RED](##.00%)";
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            //cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Price Per SF:
            cellBox1 = pivotSheet.Range["O4", "O8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            cellBox1.NumberFormat = "$#,###";
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format footer
            cellBox1 = pivotSheet.Range["A9"];
            cellBox1.Font.Size = 26;
            cellBox1.EntireRow.RowHeight = 42;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.Font.Bold = false;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            cellBox1.WrapText = true;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Comparable Criteria:
            cellBox1 = pivotSheet.Range["A11", "O11"];
            cellBox1.Font.Size = 24;
            cellBox1.EntireRow.RowHeight = 42;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Chocolate); ;
            cellBox1.Font.Bold = false;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            cellBox1.WrapText = true;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //format Average / Highests Boxes:
            cellBox1 = pivotSheet.Range["A5", "B8"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            cellBox1.EntireRow.RowHeight = 22;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); ;
            cellBox1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cellBox1.WrapText = true;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            //add border double line
            cellBox1 = pivotSheet.Range["A3", "O8"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            cellBox1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            cellBox1 = pivotSheet.Range["A3", "O3"];
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            cellBox1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            cellBox1 = pivotSheet.Range["A4", "O4"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;

            cellBox1 = pivotSheet.Range["A6", "O6"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1 = pivotSheet.Range["B6", "O6"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
            cellBox1 = pivotSheet.Range["B8", "O8"];
            cellBox1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue); ;
        }
        #endregion
    }
}

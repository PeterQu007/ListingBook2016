using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListingBook2016
{
    public class ReportCMA
    {
        public Excel.Worksheet ListingSheet;
        public Excel.Worksheet PivotSheet;
        private DataProcessing dp;
        private ReportType CMAReportType;
        private string CMAReportLanguage;
        private bool TitleAdded;
        private bool SubTitleAdded;
        private bool SubjectEvaluationAdded;

        public ReportCMA(Excel.Worksheet ws, ReportType cmaType, string cmaLang = "English")
        {
            this.ListingSheet = ws;
            dp = new DataProcessing(ws);
            this.CMAReportType = cmaType;
            this.CMAReportLanguage = cmaLang;
            this.TitleAdded = false;
            this.SubTitleAdded = false;
            this.SubjectEvaluationAdded = false;
        }
        public void Residential(ListingStatus Status)
        {
            int PivotTableTopPaddingRows = 5;
            string PivotSheetName = "PivotSheet";
            string PivotTableName = "";
            PivotTableCMA ptCMA = null;

            //DATA VALIDATE
            ListingSheet.Activate();
            if (dp.ValidateData(CMAReportType))
            {
                Console.Write("Listing Data Needs To Be Reviewed");
                return;
            }
            //DATA ADD LOT AND IMPROVE UNIT PRICE AS PER BCA PERCENTAGE%
            dp.AddLotAndImproveUnitPrice();

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
            ptCMA.Format(ptCMA.PivotSheet, PivotTableName, Status, "");
            ptCMA.AddMedianSummary(ptCMA.PivotSheet, PivotTableName, Status);
            //Globals.ThisAddIn.Application.ScreenUpdating = true;
            ptCMA.AddCorCoeSummary_Attached(ptCMA.PivotSheet, ListingSheet);
            ptCMA.AddDisclaimer(ptCMA.PivotSheet);
            this.AddCMATitle(PivotSheet, this.CMAReportLanguage == "English" ? "CMA REPORT" : "CMA 物业评估报告");
            this.AddCMASubTitle(PivotSheet, "Peter Qu");
            this.AddCMASubjectEvaluation(PivotSheet);
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
            Excel.Range cell2 = WS.Cells[1, LastCol-1];
            WS.Range[cell1, cell2].Merge();

            cell1.Value = Title;
            cell1.Font.Size = 28;
            cell1.Font.Color = System.Drawing.Color.Brown.ToArgb();
            cell1.Font.Bold = true;
            cell1.Font.Italic = false;
            cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell1.EntireRow.RowHeight = 38;
            this.TitleAdded = true;
        }

        public void AddCMASubTitle(Excel.Worksheet WS, string SubTitle)
        {
            if (this.SubTitleAdded) return;
            long LastCol = Library.GetLastCol(this.PivotSheet);
            Excel.Range cell = WS.Cells[2, 1];
            Excel.Range cell2 = WS.Cells[2, LastCol-1];
            WS.Range[cell, cell2].Merge();

            cell.Value = SubTitle + " " + System.DateTime.Now.Date.ToShortDateString();
            cell.Font.Size = 14;
            cell.Font.Color = System.Drawing.Color.Black.ToArgb();
            cell.Font.Bold = false;
            cell.Font.Italic = false;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cell.EntireRow.RowHeight = 20;
            this.SubTitleAdded = true;
        }

        public void AddCMASubjectEvaluation(Excel.Worksheet pivotSheet)
        {
            if (this.SubjectEvaluationAdded) return;
            this.SubjectEvaluationAdded = true;
            Excel.Range line = (Excel.Range)pivotSheet.Rows[3];
            line.Select();
            line.Insert();
            Excel.Range line2 = (Excel.Range)pivotSheet.Rows["3:5"];
            line2.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            line2.Select();
            line2 = (Excel.Range)pivotSheet.Rows["3:6"];
            line2.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            line2.Select();
            Excel.Range cellBox1 = pivotSheet.Range["A3", "A8"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Subject Property";
            cellBox1 = pivotSheet.Range["B3", "C3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Address";
            cellBox1 = pivotSheet.Range["D3", "E3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Floor Area";
            cellBox1 = pivotSheet.Range["F3", "G3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "LOT Area";
            cellBox1 = pivotSheet.Range["H3", "J3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Total Market Value";
            cellBox1 = pivotSheet.Range["K3", "L3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "BC Assess";
            cellBox1 = pivotSheet.Range["M3", "N3"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Change % to BCA";
            cellBox1 = pivotSheet.Range["O3"];
            cellBox1.Select();
            cellBox1.Value = "Remarks";
            cellBox1 = pivotSheet.Range["A9", "O9"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Market Value Evaluation:";
            cellBox1 = pivotSheet.Range["A11", "O11"];
            cellBox1.Select();
            cellBox1.Merge();
            cellBox1.Value = "Comparable Creteria:";
            //format the cells:
            cellBox1 = pivotSheet.Range["B3", "O3"];
            cellBox1.Select();
            cellBox1.Font.Size = 16;
            cellBox1.Font.Color = System.Drawing.Color.Brown.ToArgb();
            cellBox1.Font.Bold = true;
            cellBox1.Font.Italic = false;
            cellBox1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cellBox1.EntireRow.RowHeight = 38;
            cellBox1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            cellBox1.Interior.Color = System.Drawing.Color.AliceBlue.ToArgb();

        }
    }
}

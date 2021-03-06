﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListingBook2016
{
    public class ReportBuyer
    {
        public Excel.Worksheet ListingSheet;
        public Excel.Worksheet PivotSheet;
        private DataProcessing dp;
        private ReportType BuyerReportType;
        public ReportBuyer(Excel.Worksheet ws, ReportType buyerReportType)
        {
            this.ListingSheet = ws;
            dp = new DataProcessing(ws);
            this.BuyerReportType = buyerReportType;
        }

        public void Residential(ListingStatus Status)
        {
            int PivotTableTopPaddingRows = 5;
            string PivotSheetName = "PivotSheet";
            string PivotTableName = "";
            PivotTableListingBuyer ptCMA = null;

            //DATA VALICATE
            ListingSheet.Activate();
            //if (dp.ValidateData(BuyerReportType))
            //{
            //    Console.Write("Listing Data Needs To Be Reviewed");
            //    return;
            //}

            /////////////////////
            //PIVOT TABLE
            PivotTableName = "PivotTable_" + Status;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            ptCMA = new PivotTableListingBuyer(PivotSheetName, PivotTableTopPaddingRows, PivotTableName, Status, BuyerReportType);
            if (ptCMA.ListingDataRows <= 0) return;

            ptCMA.Create();
            this.PivotSheet = ptCMA.PivotSheet;
            ptCMA.Format(ptCMA.PivotSheet, PivotTableName, Status, "");
            ptCMA.AddMedianSummary(ptCMA.PivotSheet, PivotTableName, Status);
            //Globals.ThisAddIn.Application.ScreenUpdating = true;
            ptCMA.AddCorCoeSummary_Attached(ptCMA.PivotSheet, ListingSheet);
            ptCMA.AddDisclaimer(ptCMA.PivotSheet);
            this.AddBuyerReportTitle(PivotSheet, "Buyer's REPORT");
            this.AddBuyerReportSubTitle(PivotSheet, "Peter Qu");
            ptCMA.PivotSheet.Select();
        }

        public void AddBuyerReportTitle(Excel.Worksheet WS, string Title)
        {
            long LastCol = Library.GetLastCol(this.PivotSheet);
            Excel.Range cell1 = WS.Cells[1, 1];
            Excel.Range cell2 = WS.Cells[1, LastCol];
            WS.Range[cell1, cell2].Merge();

            cell1.Value = Title;
            cell1.Font.Size = 28;
            cell1.Font.Color = System.Drawing.Color.Brown.ToArgb();
            cell1.Font.Bold = true;
            cell1.Font.Italic = false;
            cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell1.EntireRow.RowHeight = 38;
        }

        public void AddBuyerReportSubTitle(Excel.Worksheet WS, string SubTitle)
        {
            long LastCol = Library.GetLastCol(this.PivotSheet);
            Excel.Range cell = WS.Cells[2, 1];
            Excel.Range cell2 = WS.Cells[2, LastCol];
            WS.Range[cell, cell2].Merge();

            cell.Value = SubTitle + " " + System.DateTime.Now.Date.ToShortDateString();
            cell.Font.Size = 14;
            cell.Font.Color = System.Drawing.Color.Black.ToArgb();
            cell.Font.Bold = false;
            cell.Font.Italic = false;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell.EntireRow.RowHeight = 20;
        }
    }
}

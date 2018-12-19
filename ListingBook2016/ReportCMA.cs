﻿using System;
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

        public ReportCMA(Excel.Worksheet ws)
        {
            this.ListingSheet = ws;
            dp = new DataProcessing(ws);
        }
        public void Condo(ListingStatus Status, bool bCMA)
        {
            int PivotTableTopPaddingRows = 5;
            string PivotSheetName = "PivotSheet";
            string PivotTableName = "";
            bool bShowUnitNo = bCMA;
            PivotTableListingStatus lpt = null;

            //DATA VALICATE
            ListingSheet.Activate();
            if (dp.ValidateData_Attached())
            {
                Console.Write("Listing Data Needs To Be Reviewed");
                return;
            }

            /////////////////////
            //PIVOT TABLE
            PivotTableName = "PivotTable_" + Status;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            lpt = new PivotTableListingStatus(PivotSheetName, PivotTableTopPaddingRows, PivotTableName, Status, bShowUnitNo);
            this.PivotSheet = lpt.PivotSheet;
            lpt.Format(lpt.PivotSheet, PivotTableName, Library.GetStatus(Status), "");
            lpt.AddMedianSummary(lpt.PivotSheet, PivotTableName, Library.GetStatus(Status));
            //Globals.ThisAddIn.Application.ScreenUpdating = true;
            lpt.AddCorCoeSummary_Attached(lpt.PivotSheet, ListingSheet);
            lpt.AddDisclaimer(lpt.PivotSheet);

            lpt.PivotSheet.Select();
        }
        public void AddCMATitle(Excel.Worksheet WS, string Title)
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

        public void AddCMASubTitle(Excel.Worksheet WS, string SubTitle)
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

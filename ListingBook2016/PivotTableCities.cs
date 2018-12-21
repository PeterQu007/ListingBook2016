using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListingBook2016
{
    public class PivotTableCities
    {
        public Excel.Worksheet PivotSheet;
        public Excel.Worksheet ListingSheet;
        public Excel.Workbook ListingBook;
        private string PivotSheetName;
        private string PivotTableName;
        private int PivotTableTopPaddingRows;
        private string PivotTableLocation;
        private char Status;
        private bool bShowUnitNo;

        public PivotTableCities(string pvSheetName, int TopPadding, string TableName, ReportType TableType)
        {
            this.PivotSheetName = pvSheetName;
            this.PivotTableName = TableName;
            this.PivotTableTopPaddingRows = TopPadding;
            this.ListingSheet = Globals.ThisAddIn.Application.Worksheets["Sheet1"];
            this.ListingBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (!Library.SheetExist(PivotSheetName))
            {
                Excel.Worksheet NewSheet = Globals.ThisAddIn.Application.Worksheets.Add();
                NewSheet.Name = PivotSheetName;
            }
            PivotSheet = Globals.ThisAddIn.Application.Worksheets[PivotSheetName];
            PivotSheet.Activate();
            int PivotTableFirstRow = Library.GetLastRow(PivotSheet) + PivotTableTopPaddingRows;
            this.PivotTableLocation = "A" + PivotTableFirstRow;
            this.CreateCityPivotTable(PivotSheet, PivotTableLocation, PivotTableName, TableType);
        }

        public void CreateCityPivotTable(Excel.Worksheet PivotSheet, string Location, string TableName, ReportType TableType)
        {
            ListingSheet.Select();
            string LastRow = "";
            string LastCol = "";
            string LastCell = "";
            long lRow = 0;
            long lCol = 0;
            string RankBaseField = "";
            ////////////
            //FIND THE LAST NON-BLANK CELL IN COLUMN A
            lRow = ListingSheet.Cells[ListingSheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            LastRow = "R" + lRow;
            lCol = ListingSheet.Cells[1, ListingSheet.Columns.Count].End(Excel.XlDirection.xlToLeft).Column;
            LastCol = "C" + lCol;
            LastCell = ListingSheet.Cells[lRow, lCol].Address;

            Excel.Range PivotData = ListingSheet.Range["A1", LastCell];
            PivotData.Select();
            Excel.PivotCaches pch = ListingBook.PivotCaches();
            Excel.PivotCache pc = pch.Create(Excel.XlPivotTableSourceType.xlDatabase, PivotData);
            Excel.PivotTable pvt = pc.CreatePivotTable(PivotSheet.Range[Location], TableName);
            //pvt.MergeLabels = true; // The only thing I noticed this doing was centering the heading labels

            PivotSheet.Select();

            //Excel.PivotField pvf = pvt.PivotFields("Status");
            //pvf.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            //pvf.CurrentPage = Status;

            //Group 1 S/A
            switch (TableType)
            {
                case ReportType.MonthlyDetachedAllCities:
                    pvt.PivotFields("City").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                    RankBaseField = "City";
                    break;
                case ReportType.MonthlyDetachedAllCommunities:
                    pvt.PivotFields("Neighborhood").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                    RankBaseField = "Neighborhood";
                    break;
                default:
                    break;
            }

            //pvt.PivotFields("S/A").Name = "Neighborhood";

            //Sales Total Amount
            pvt.AddDataField(pvt.PivotFields("Sold Price"), "Rank", Excel.XlConsolidationFunction.xlSum);
            pvt.PivotFields("Rank").Calculation = Excel.XlPivotFieldCalculation.xlRankDecending;
            pvt.PivotFields("Rank").BaseField = RankBaseField;
            //Sort By Rank
            pvt.PivotFields(RankBaseField).AutoSort(2, "Rank");
            //Total Amount
            pvt.AddDataField(pvt.PivotFields("Sold Price"), "Total Sales Amount", Excel.XlConsolidationFunction.xlSum);
            pvt.PivotFields("Total Sales Amount").NumberFormat = "$#,##0";
            pvt.AddDataField(pvt.PivotFields("Sold Price"), "Market Share", Excel.XlConsolidationFunction.xlSum);
            pvt.PivotFields("Market Share").Calculation = Excel.XlPivotFieldCalculation.xlPercentOfTotal;
            //Sales Count
            pvt.AddDataField(pvt.PivotFields("Status"), "Sales", Excel.XlConsolidationFunction.xlCount);
            pvt.PivotFields("Sales").NumberFormat = "0";
            //Ave Sold Price
            pvt.AddDataField(pvt.PivotFields("Sold Price"), "Avg. Sold Price", Excel.XlConsolidationFunction.xlAverage);
            pvt.PivotFields("Avg. Sold Price").NumberFormat = "$#,##0";
            pvt.AddDataField(pvt.PivotFields("Sold Price"), "Avg. S.Price Rank", Excel.XlConsolidationFunction.xlAverage);
            pvt.PivotFields("Avg. S.Price Rank").Calculation = Excel.XlPivotFieldCalculation.xlRankDecending;
            pvt.PivotFields("Avg. S.Price Rank").BaseField = RankBaseField;
            //Price Per SqFt
            pvt.AddDataField(pvt.PivotFields("SP Sqft"), "Avg. $PerSQFT", Excel.XlConsolidationFunction.xlAverage);
            pvt.PivotFields("Avg. $PerSQFT").NumberFormat = "$#,##0";
            pvt.AddDataField(pvt.PivotFields("SP Sqft"), "Avg. $PSF Rank", Excel.XlConsolidationFunction.xlAverage);
            pvt.PivotFields("Avg. $PSF Rank").Calculation = Excel.XlPivotFieldCalculation.xlRankDecending;
            pvt.PivotFields("Avg. $PSF Rank").BaseField = RankBaseField;
            //Days On Market
            pvt.AddDataField(pvt.PivotFields("CDOM"), "Avg. Days OnMkt", Excel.XlConsolidationFunction.xlAverage);
            pvt.PivotFields("Avg. Days OnMkt").NumberFormat = "0";
            pvt.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);
        }

        public void FormatPivotTable(Excel.Worksheet PivotSheet, string TableName)
        {
            Excel.PivotTable pvt = PivotSheet.PivotTables(TableName);
            int FirstRow = 0;
            int LastRow = 0;
            int LastCol = 0;
            int TitleRow = 0;

            PivotSheet.Select();
            PivotSheet.Cells[1, 1].Select();

            FirstRow = pvt.TableRange1.Row + 1;
            LastRow = FirstRow + pvt.TableRange1.Rows.Count - 2;
            LastCol = pvt.ColumnRange.Columns.Count + pvt.ColumnRange.Column - 1;
            TitleRow = pvt.TableRange2.Row - 1;
            //Todo: Format Title

            //Hide Values Row
            Excel.Range rng0 = PivotSheet.Range["A" + (FirstRow - 1)];
            rng0.EntireRow.Hidden = true;
            //Title Row
            Excel.Range c1 = PivotSheet.Cells[FirstRow, 1];
            Excel.Range c2 = PivotSheet.Cells[FirstRow, LastCol];
            Excel.Range rng = PivotSheet.Range[c1, c2];
            rng.Select();
            rng.Style.Font.Name = "Roboto";
            rng.Style.Font.Size = 14;
            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
            rng.RowHeight = 38;
            rng.WrapText = true;
            rng.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rng.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            //Grand Total Row
            rng = PivotSheet.Range["A" + LastRow];
            rng.Select();
            rng.RowHeight = 36;
            rng.Style.Font.Name = "Roboto";
            rng.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //All Rows
            pvt.TableRange1.RowHeight = 36;
            //All DataRange
            pvt.DataBodyRange.Font.Size = 12;

            pvt.TableRange1.Select();
            pvt.DataBodyRange.Select();
            pvt.ColumnRange.Select();
            pvt.RowRange.Select();
            Console.WriteLine(pvt.ColumnRange.Columns.Count);
            Console.WriteLine(pvt.DataBodyRange.Columns.Count);
            pvt.DataLabelRange.Select();
            pvt.ColumnRange.Columns[2].Select();
            pvt.ColumnRange.Columns[2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.ColumnRange.Columns[2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            pvt.ColumnRange.Columns[2].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            pvt.DataBodyRange.Columns[2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.DataBodyRange.Columns[2].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            pvt.DataBodyRange.Columns[2].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;

            pvt.ColumnRange.Columns[3].Borders[Excel.XlBordersIndex.xlEdgeRight].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.ColumnRange.Columns[3].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            pvt.ColumnRange.Columns[3].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            pvt.DataBodyRange.Columns[3].Borders[Excel.XlBordersIndex.xlEdgeRight].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.DataBodyRange.Columns[3].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            pvt.DataBodyRange.Columns[3].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            pvt.ColumnRange.Columns[5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.ColumnRange.Columns[5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            pvt.ColumnRange.Columns[5].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            pvt.DataBodyRange.Columns[5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.DataBodyRange.Columns[5].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            pvt.DataBodyRange.Columns[5].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;

            pvt.ColumnRange.Columns[6].Borders[Excel.XlBordersIndex.xlEdgeRight].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.ColumnRange.Columns[6].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            pvt.ColumnRange.Columns[6].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            pvt.DataBodyRange.Columns[6].Borders[Excel.XlBordersIndex.xlEdgeRight].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
            pvt.DataBodyRange.Columns[6].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            pvt.DataBodyRange.Columns[6].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;


            //ListingBook.TableStyles.Add("Attached Report Style");
            Excel.TableStyle ptStyle = ListingBook.TableStyles["PivotStyleLight16"];


            // Set Table Style
            pvt.TableStyle2 = ptStyle;

            pvt.TableRange1.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexNone, System.Drawing.Color.DarkBlue);
            FormatTop3(pvt, FirstRow, LastRow, 4);
            FormatBottom3(pvt, FirstRow, LastRow, 4);
            FormatTop3(pvt, FirstRow, LastRow, 5);
            FormatBottom3(pvt, FirstRow, LastRow, 5);
            FormatTop3(pvt, FirstRow, LastRow, 7);
            FormatBottom3(pvt, FirstRow, LastRow, 7);
            FormatTop3(pvt, FirstRow, LastRow, 9);
            FormatBottom3(pvt, FirstRow, LastRow, 9);
            FormatColumnWidth();
        }

        private void FormatTop3(Excel.PivotTable pvt, int FirstRow, int LastRow, int iCol)
        {
            Excel.Range c1 = PivotSheet.Cells[FirstRow + 1, pvt.ColumnRange.Columns[iCol].Column];
            Excel.Range c2 = PivotSheet.Cells[LastRow - 1, pvt.ColumnRange.Columns[iCol].Column];
            PivotSheet.Range[c1, c2].Select();
            //pvt.PivotSelect("Sales", Excel.XlPTSelectionMode.xlDataOnly, true);
            Globals.ThisAddIn.Application.Selection.FormatConditions.AddTop10();
            Globals.ThisAddIn.Application.Selection.FormatConditions(Globals.ThisAddIn.Application.Selection.FormatConditions.Count).SetFirstPriority();
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).TopBottom = Excel.XlTopBottom.xlTop10Top;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Rank = 3;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Percent = false;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Font.Color = System.Drawing.Color.IndianRed;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Font.Bold = true;
        }

        private void FormatBottom3(Excel.PivotTable pvt, int FirstRow, int LastRow, int iCol)
        {

            Excel.Range c1 = PivotSheet.Cells[FirstRow + 1, pvt.ColumnRange.Columns[iCol].Column];
            Excel.Range c2 = PivotSheet.Cells[LastRow - 1, pvt.ColumnRange.Columns[iCol].Column];
            PivotSheet.Range[c1, c2].Select();
            //pvt.PivotSelect("Sales", Excel.XlPTSelectionMode.xlDataOnly, true);
            Globals.ThisAddIn.Application.Selection.FormatConditions.AddTop10();
            Globals.ThisAddIn.Application.Selection.FormatConditions(Globals.ThisAddIn.Application.Selection.FormatConditions.Count).SetFirstPriority();
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).TopBottom = Excel.XlTopBottom.xlTop10Bottom;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Rank = 3;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Percent = false;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Font.Color = System.Drawing.Color.DarkGreen;
            Globals.ThisAddIn.Application.Selection.FormatConditions(1).Font.Bold = true;

        }

        private void FormatColumnWidth()
        {
            if (PivotSheet.PivotTables(1).RowRange.Columns.Count == 3)
            {
                PivotSheet.Columns["A"].ColumnWidth = 19.5;
                PivotSheet.Columns["B"].ColumnWidth = 17;
                PivotSheet.Columns["C"].ColumnWidth = 18.5;
            }
            else
            {
                PivotSheet.Columns["A"].ColumnWidth = 19.5;
                PivotSheet.Columns["B"].ColumnWidth = 17;
                PivotSheet.Columns["C"].ColumnWidth = 18.5;
                PivotSheet.Columns["D"].ColumnWidth = 5;
            }

            int FirstCol = PivotSheet.PivotTables(1).ColumnRange.Column;
            int TotalCols = PivotSheet.PivotTables(1).ColumnRange.Columns.Count;
            PivotSheet.Columns[FirstCol].ColumnWidth = 6;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 15.5;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 9.5;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 8.5;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 10.5;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 8.5;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 8.5;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 8.5;
            PivotSheet.Columns[++FirstCol].ColumnWidth = 8.5;

        }

        public void AddSectionTitle(Excel.Worksheet WS, string PTName, string Title)
        {

            Excel.PivotTable PT = WS.PivotTables(PTName);
            int FirstRow = PT.TableRange2.Row;
            //ADD SECTION TITLE
            Excel.Range cell = WS.Range["A" + (FirstRow + 1)];
            cell.Value = Title;
            cell.Font.Size = 18;
            cell.Font.Color = System.Drawing.Color.Red.ToArgb();
            cell.Font.Bold = true;
            cell.Font.Italic = true;
            cell.EntireRow.RowHeight = 24;
            //HIDE PAGE GROUP FILTER
            cell = WS.Range["A" + FirstRow];
            cell.EntireRow.Hidden = true;
        }

        public void AddMedianSummary(Excel.Worksheet TableSheet, string TableName, char Status)
        {
            int lastRow = 0;
            int firstRow = 0;
            int avgRow = 0;
            int medianRow = 0;
            int lastCol = 0;
            double avgRowHeight = 38;
            double medianRowHeight = avgRowHeight;
            Excel.PivotTable pvt = null;

            TableSheet.Select();
            pvt = TableSheet.PivotTables(TableName);
            firstRow = pvt.TableRange1.Row + 2;
            lastRow = pvt.TableRange1.Row + pvt.TableRange1.Rows.Count - 2;
            avgRow = lastRow + 1;
            medianRow = avgRow + 1;

            int rw = 0;

            foreach (Excel.PivotField pvf in pvt.RowFields)
            {
                rw++;
            }
            Excel.Range Cell = TableSheet.Range["A" + avgRow];
            Cell.Value2 = "Average Values";
            Cell = TableSheet.Range["A" + medianRow];
            Cell.Value2 = "Median Values";

            TableSheet.Cells[medianRow, rw + 1].Value = Library.GetCount(ListingSheet, "B", Status, "", "");
            TableSheet.Cells[medianRow, rw + 2].Value = Library.GetMedianValue(ListingSheet, "G", Status, "", "");
            TableSheet.Cells[medianRow, rw + 3].Value = Library.GetMedianValue(ListingSheet, "I", Status, "", "");
            TableSheet.Cells[medianRow, rw + 4].Value = Library.GetMedianValue(ListingSheet, "L", Status, "", "");
            TableSheet.Cells[medianRow, rw + 5].Value = Library.GetMedianValue(ListingSheet, "M", Status, "", "");
            TableSheet.Cells[medianRow, rw + 6].Value = Library.GetMedianValue(ListingSheet, "O", Status, "", "");
            TableSheet.Cells[medianRow, rw + 7].Value = Library.GetMedianValue(ListingSheet, "P", Status, "", "");
            TableSheet.Cells[medianRow, rw + 8].Value = Library.GetMedianValue(ListingSheet, "R", Status, "", "");
            TableSheet.Cells[medianRow, rw + 9].Value = Library.GetMedianValue(ListingSheet, "S", Status, "", "");

            TableSheet.Select();
            lastCol = pvt.TableRange1.Columns.Count;
            Excel.Range rng = TableSheet.Range[TableSheet.Cells[medianRow, 1], TableSheet.Cells[medianRow, lastCol]];
            rng.RowHeight = medianRowHeight;
            Excel.Range rngSource = TableSheet.Range[TableSheet.Cells[avgRow, 1], TableSheet.Cells[avgRow, lastCol]];
            rngSource.Select();
            rngSource.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rngSource.Copy();
            rng.Select();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue);
            rng.Font.Bold = true;
            rngSource.Select();
            Globals.ThisAddIn.Application.CutCopyMode = 0;
        }

        public void AddCorCoeSummary_Attached(Excel.Worksheet DestSheet, Excel.Worksheet SourceSheet)
        {
            DestSheet.Select();
            int LastRow = Library.GetLastRow(DestSheet);
            Excel.Range Cell1 = null;
            Excel.Range Cell2 = null;

            //1) BCA - Price
            double BCA_Price_CorCoe = Library.GetCorCoeValue(SourceSheet, "G", "R");
            Cell1 = DestSheet.Range["A" + (LastRow + 1)];
            Cell1.Value = "BCA - Price: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 1)];
            Cell2.Value = BCA_Price_CorCoe;
            //2) BCA - Change
            double BCA_Change_CorCoe = Library.GetCorCoeValue(SourceSheet, "S", "R");
            Cell1 = DestSheet.Range["A" + (LastRow + 2)];
            Cell1.Value = "BCA - Change: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 2)];
            Cell2.Value = BCA_Change_CorCoe;
            //3) FloorArea - Price Per Square Feet
            double FloorArea_PricePSF_CorCoe = Library.GetCorCoeValue(SourceSheet, "L", "M");
            Cell1 = DestSheet.Range["A" + (LastRow + 3)];
            Cell1.Value = "FloorArea - Price Per Square Feet: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 3)];
            Cell2.Value = FloorArea_PricePSF_CorCoe;
            //4) Age - Fee
            double Age_Fee_CorCoe = Library.GetCorCoeValue(SourceSheet, "O", "P");
            Cell1 = DestSheet.Range["A" + (LastRow + 4)];
            Cell1.Value = "Age - Maint. Fee: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 4)];
            Cell2.Value = Age_Fee_CorCoe;
            //5) Age - Price Per Square Feet
            double Age_PricePSF_CorCoe = Library.GetCorCoeValue(SourceSheet, "O", "M");
            Cell1 = DestSheet.Range["A" + (LastRow + 5)];
            Cell1.Value = "Age - Price Per Square Feet: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 5)];
            Cell2.Value = Age_PricePSF_CorCoe;

            DestSheet.Select();
            Excel.Range rng = DestSheet.Range["A" + (LastRow + 1), "L" + (LastRow + 5)];
            rng.Select();
            rng.EntireRow.RowHeight *= 1.3;
            rng.Font.Bold = true;

        }

        public void AddDisclaimer(Excel.Worksheet Sheet)
        {
            Sheet.Select();
            System.DateTime moment = new System.DateTime(
                                1999, 1, 13, 3, 57, 32, 11);
            int LastRow = Library.GetLastRow(Sheet);
            int year = moment.Year;
            Excel.Range cell = Sheet.Range["A" + (LastRow + 1)];
            cell.Value = "Disclaimer: FOR REFERRENCE ONLY, NOT TO BE ANY INVESTMENT RECOMMENDATION. ";
            cell.Value += ("©" + year + " PIDREALTY.CA All Rights Reserved");
            cell.Font.Size = 8;
            cell.Font.Color = System.Drawing.Color.Red.ToArgb();
            cell.Font.Bold = true;
            cell.Font.Italic = true;
        }

    }
}

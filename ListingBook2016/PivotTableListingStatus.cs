using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListingBook2016
{

    public class PivotTableCMA : PivotTableListingStatus
    {
        private static readonly bool bShowUnitNoTrue = true;
        public PivotTableCMA(string pvSheetName, int TopPadding, string TableName, ListingStatus Status, ReportType cmaType)
                : base(pvSheetName, TopPadding, TableName, Status)
        {
            ReportType = cmaType;
            bShowUnitNo = bShowUnitNoTrue && (ReportType.ToString().IndexOf("Attached") > 0);
        }

        public void AddComparableCreteria(){
            //TO DO ADD AGE
            //FLOOR AREA, LOT AREA
            //ADD BEDROOMS, BATHROOMS, 
            //ADD STRUCTURE - CONCRETE OF WOOD FRAME?
        }
    }
    public class PivotTableListingStatus
    {
        public Excel.Worksheet PivotSheet;
        public Excel.Worksheet ListingSheet;
        public Excel.Workbook ListingBook;
        public int ListingDataRows;
        public decimal MaxSoldPricePerSF_Land;
        public decimal MaxSoldPricePerSF_Improve;
        public decimal MaxSoldPricePerSF_Total;
        public decimal AverageSoldPricePerSF_Land;
        public decimal AverageSoldPricePerSF_Improve;
        public decimal AverageSoldPricePerSF_Total;
        private string PivotSheetName;
        private string PivotTableName;
        private int PivotTableTopPaddingRows;
        private string PivotTableLocation;
        private char Status;
        protected bool bShowUnitNo;
        public bool FormatColumnsWidthDone = false;
        protected ReportType ReportType;
        public PivotTableListingStatus(string pvSheetName, int TopPadding, string TableName, ListingStatus Status)
        {
            this.PivotSheetName = pvSheetName;
            this.PivotTableName = TableName;
            this.PivotTableTopPaddingRows = TopPadding;
            this.Status = (char)Status; //Library.GetStatus(Status);

            this.ListingSheet = Globals.ThisAddIn.Application.Worksheets["Listings Table"];
            this.ListingBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            this.ListingSheet.AutoFilterMode = false;

            //TEST FILTERS IF NO RECORDS THEN PASS PIVOT TABLE FUNCTION
            int iCol = ListingSheet.Range[ListingDataColNames.Status + "1"].Column;
            string[] StatusArray = Library.StatusArray(Status);
            ListingSheet.Range["A1"].AutoFilter(iCol, StatusArray, Excel.XlAutoFilterOperator.xlFilterValues);
            int LastRow = Library.GetLastRow(ListingSheet);
            if (LastRow > 1)
            {
                ListingDataRows = LastRow - 1;
                //if (Library.SheetExist(PivotSheetName))
                //{
                //    Globals.ThisAddIn.Application.Worksheets[PivotSheetName].Delete();
                //}
                //Excel.Worksheet NewSheet = Globals.ThisAddIn.Application.Worksheets.Add();
                //NewSheet.Name = PivotSheetName;
                PivotSheet = Globals.ThisAddIn.Application.Worksheets[PivotSheetName];
                PivotSheet.Activate();
                int PivotTableFirstRow = Library.GetLastRow(PivotSheet) + PivotTableTopPaddingRows;
                this.PivotTableLocation = "A" + PivotTableFirstRow;
            }else
            {
                ListingDataRows = 0;
            }
            
        }

        public void Create()
        {
            Excel.Worksheet PivotSheet = this.PivotSheet;
            string Location = this.PivotTableLocation;
            string TableName = this.PivotTableName;
            char Status = this.Status;

            ListingSheet.Select();
            string LastRow = "";
            string LastCol = "";
            string LastCell = "";
            long lRow = 0;
            long lCol = 0;

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
            PivotSheet.Select();

            Excel.PivotField pvf = pvt.PivotFields("Status");
            pvf.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            switch ((ListingStatus)Status)
            {
                case ListingStatus.Active:
                case ListingStatus.Sold:
                    pvf.CurrentPage = Status.ToString();
                    break;
                case ListingStatus.OffMarket:
                    try { pvf.PivotItems(((char)ListingStatus.Active).ToString()).Visible = false; } catch (Exception e) { };
                    try { pvf.PivotItems(((char)ListingStatus.Sold).ToString()).Visible = false; } catch (Exception e) { };
                    try { pvf.PivotItems(((char)ListingStatus.Terminate).ToString()).Visible = true; } catch(Exception e) { };
                    try { pvf.PivotItems(((char)ListingStatus.Cancel).ToString()).Visible = true; } catch (Exception e) { };
                    try { pvf.PivotItems(((char)ListingStatus.Expire).ToString()).Visible = true; } catch (Exception e) { };
                    pvf.EnableMultiplePageItems = true;
                    break;
            }
            

            //Group 1 S/A
            pvt.PivotFields("S/A").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pvt.PivotFields("S/A").Name = "Neighborhood";
            //Group 2 Complex
            pvt.PivotFields("Complex/Subdivision Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pvt.PivotFields("Complex/Subdivision Name").Name = this.ReportType.ToString().IndexOf("Detached") < 0 ? "Complex" : "SubDivision";
            //Group 3 Address
            pvt.PivotFields("Address2").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pvt.PivotFields("Address2").Name = "Civic Address";
            //Group 4 UnitNo
            if (this.bShowUnitNo || this.ReportType.ToString().IndexOf("Detached")<0)
            {
                pvt.PivotFields("Unit#").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                pvt.PivotFields("Unit#").Name = "Unit No";
            }

            pvt.AddDataField(pvt.PivotFields("MLS"), "Count", Excel.XlConsolidationFunction.xlCount);
            pvt.AddDataField(pvt.PivotFields("Price0"), "Price", Excel.XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields("CDOM"), "Days On Mkt", Excel.XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields("TotFlArea"), "Floor Area", Excel.XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields("PrcSqft"), "$PSF", Excel.XlConsolidationFunction.xlAverage);
            //TEST Add Calculated Fields
            //Excel.PivotField ptField;
            //Excel.CalculatedFields cfField = pvt.CalculatedFields();
            //ptField = cfField.Add("New PSF", "='PrcSqft' * 'Age'", true);
            //pvt.AddDataField(ptField, " New PSF", Excel.XlConsolidationFunction.xlAverage);
            //
            pvt.AddDataField(pvt.PivotFields("Age"), "Building Age", Excel.XlConsolidationFunction.xlAverage);
            if (this.ReportType.ToString().IndexOf("Detached") < 0)
            {
                pvt.AddDataField(pvt.PivotFields("StratMtFee"), "Monthly Fee", Excel.XlConsolidationFunction.xlAverage);
            }
            else
            {
                pvt.AddDataField(pvt.PivotFields("Lot Sz (Sq.Ft.)"), "Land Size", Excel.XlConsolidationFunction.xlAverage);
                pvt.AddDataField(pvt.PivotFields("LandValue"), "Land Assess.", Excel.XlConsolidationFunction.xlAverage);
            }

            pvt.AddDataField(pvt.PivotFields("BCAValue"), "BC Assess.", Excel.XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields("Change%"), "Chg% to BCA", Excel.XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields("Lot$ PerSF"), "Lot$PSF", Excel.XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields("Improve$ PerSF"), "Improve$PSF", Excel.XlConsolidationFunction.xlAverage);

            pvt.PivotFields("Price").NumberFormat = "$#,##0";
            pvt.PivotFields("Days On Mkt").NumberFormat = "0";
            pvt.PivotFields("Floor Area").NumberFormat = "0";
            pvt.PivotFields("$PSF").NumberFormat = "$#,##0";
            pvt.PivotFields("Building Age").NumberFormat = "0";
            if (this.ReportType.ToString().IndexOf("Detached") < 0)
            {
                pvt.PivotFields("Monthly Fee").NumberFormat = "$#,##0";
            }
            else
            {
                pvt.PivotFields("Land Size").NumberFormat = "0";
                pvt.PivotFields("Land Assess.").NumberFormat = "$#,##0";
            }
            pvt.PivotFields("BC Assess.").NumberFormat = "$#,##0";
            pvt.PivotFields("Chg% to BCA").NumberFormat = "0%";
            pvt.PivotFields("Lot$PSF").NumberFormat = "$#,##0";
            pvt.PivotFields("Improve$PSF").NumberFormat = "$#,##0";

            pvt.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);

        }

        public void Format(Excel.Worksheet PivotSheet, string TableName, ListingStatus Status, string City)
        {
            Excel.PivotTable pvt = PivotSheet.PivotTables(TableName);
            int FirstRow = 0;
            int LastRow = 0;
            int LastCol = 0;
            int TitleRow = 0;

            FirstRow = pvt.TableRange1.Row + 1;
            LastRow = FirstRow + pvt.TableRange1.Rows.Count - 2;
            LastCol = pvt.ColumnRange.Columns.Count + pvt.ColumnRange.Column;
            TitleRow = pvt.TableRange2.Row - 1;
            //Todo: Format Title

            //Hide Values Row
            Excel.Range rng0 = PivotSheet.Range["A" + (FirstRow - 1)];
            rng0.EntireRow.Hidden = true;

            rng0 = PivotSheet.Range["A" + (FirstRow-2)];
            rng0.RowHeight = 32;
            rng0.Font.Size = 24;

            Excel.Range c1 = PivotSheet.Cells[FirstRow, 1];
            Excel.Range c2 = PivotSheet.Cells[FirstRow, LastCol];
            Excel.Range rng = PivotSheet.Range[c1, c2];
            rng.Select();
            rng.RowHeight = 60; // 38;
            rng.WrapText = true;
            rng.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rng.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rng.Font.Size = 16;
            rng.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue); ;

            rng = PivotSheet.Range["A" + LastRow, "O" + LastRow];
            rng.Select();
            rng.RowHeight = 36;
            rng.Font.Size = 16;

            rng = PivotSheet.Range["A" + (FirstRow + 1), "O" + (LastRow - 2)];
            rng.RowHeight = 32;
            rng.Font.Size = 16;

            // Create the table style
            //ListingBook.TableStyles.Add("Attached Report Style");
            Excel.TableStyle ptStyle = ListingBook.TableStyles["PivotStyleLight16"];
            //ptStyle.ShowAsAvailablePivotTableStyle = true;

            // Table style Header Row
            //Excel.TableStyleElement HeaderRow = ptStyle.TableStyleElements[Excel.XlTableStyleElementType.xlHeaderRow];
            //HeaderRow.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            //HeaderRow.Interior.TintAndShade = -0.249946592608417;
            //HeaderRow.Font.Bold = true;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Color = System.Drawing.Color.Black.ToArgb();
            //HeaderRow.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlInsideVertical].Color = System.Drawing.Color.Black.ToArgb();
            //HeaderRow.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.Color.Black.ToArgb();
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = System.Drawing.Color.Black.ToArgb();
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = System.Drawing.Color.Black.ToArgb();
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.Black.ToArgb();
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            //HeaderRow.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            //// Table style Row Stripe 1
            //Excel.TableStyleElement RowStripe1 = ptStyle.TableStyleElements[Excel.XlTableStyleElementType.xlRowStripe1];
            //RowStripe1.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.Red.ToArgb();
            //RowStripe1.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            //RowStripe1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            //// Table style Row Stripe 2
            //Excel.TableStyleElement RowStripe2 = ptStyle.TableStyleElements[Excel.XlTableStyleElementType.xlRowStripe2];
            //RowStripe2.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.Blue.ToArgb();
            //RowStripe2.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            //RowStripe2.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Set Table Style
            pvt.TableStyle2 = ptStyle;


            FormatColumnWidth();
            HideComplexSubTotal(PivotSheet, TableName);
            AddSubGroupBottomBorder(TableName);
            FormatMaxCells();
            FormatMinCells();
            
            AddSectionTitle(PivotSheet, TableName, City + " " + Status + " Records:");
        }

        private void AddSubGroupBottomBorder(string PivotTableName)
        {
            Excel.PivotTable pt = this.PivotSheet.PivotTables(PivotTableName);
            Excel.Range c = null;
            Excel.Range c1 = null;
            Excel.Range c2 = null;
            Excel.Range row = null;
            long LastCol = 0;
            int i = 0;
            int FirstRow = pt.TableRange1.Row;
            int LastRow = FirstRow + pt.TableRange1.Rows.Count - 1;
            LastCol = pt.TableRange1.Columns.Count;

            for (i = 1; i < LastRow - FirstRow; i++)
            {
                c = pt.TableRange1.Cells[i, 3];
                c.Select();
                if (!c.EntireRow.Hidden && c.Value != null && c.Value.ToString().IndexOf("Total") > 0)
                {
                    c1 = pt.TableRange1.Cells[i, 1];
                    c2 = pt.TableRange1.Cells[i, LastCol];
                    row = this.PivotSheet.Range[c1, c2];
                    row.Select();
                    row.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RoyalBlue);
                    row.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                    row.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    row.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RoyalBlue);
                    row.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                    row.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlDash;
                    row.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                }
            }
        }

        private void FormatColumnWidth()
        {
            if (FormatColumnsWidthDone)
            {
                return;
            }
            if (PivotSheet.PivotTables(1).RowRange.Columns.Count == 3)
            {
                PivotSheet.Columns["A"].ColumnWidth = 20; //neighborhood
                PivotSheet.Columns["B"].ColumnWidth = 17; //PlanNumb
                PivotSheet.Columns["C"].ColumnWidth = 18.5; //Address
            }
            else
            {
                PivotSheet.Columns["A"].ColumnWidth = 20; //neighborhood
                PivotSheet.Columns["B"].ColumnWidth = 17; //complex
                PivotSheet.Columns["C"].ColumnWidth = 18.5; //address
                PivotSheet.Columns["D"].ColumnWidth = 9; //Unit No
            }

            int FirstCol = PivotSheet.PivotTables(1).ColumnRange.Column;
            int TotalCols = PivotSheet.PivotTables(1).ColumnRange.Columns.Count;
            PivotSheet.Columns[FirstCol].ColumnWidth = 9; //count
            PivotSheet.Columns[++FirstCol].ColumnWidth = 17; //Price
            PivotSheet.Columns[++FirstCol].ColumnWidth = 7.5; //Days on Market
            PivotSheet.Columns[++FirstCol].ColumnWidth = 9; //Floor Area
            PivotSheet.Columns[++FirstCol].ColumnWidth = 9; //Price per SF
            PivotSheet.Columns[++FirstCol].ColumnWidth = 12; //Building Age
            if (this.ReportType.ToString().IndexOf("Detached") < 0)
            {
                PivotSheet.Columns[++FirstCol].ColumnWidth = 11; // Monthly Fee
                PivotSheet.Columns[++FirstCol].ColumnWidth = 13; // BC Assessment
                PivotSheet.Columns[++FirstCol].ColumnWidth = 7.5; //chg% to BCA
                PivotSheet.Columns[++FirstCol].ColumnWidth = 11; //Lot $PSF
                PivotSheet.Columns[++FirstCol].ColumnWidth = 11; //Improve $PSF
            }
            else
            {
                PivotSheet.Columns[++FirstCol].ColumnWidth = 9.5; //land size
                PivotSheet.Columns[++FirstCol].ColumnWidth = 17; //Land Assess
                PivotSheet.Columns[++FirstCol].ColumnWidth = 17; //BC Assess
                PivotSheet.Columns[++FirstCol].ColumnWidth = 9; //Chg% to BCA
                PivotSheet.Columns[++FirstCol].ColumnWidth = 11; //Lot $PSF
                PivotSheet.Columns[++FirstCol].ColumnWidth = 11; //Improve $PSF
            }
            FormatColumnsWidthDone = true;
        }

        public void AddSectionTitle(Excel.Worksheet WS, string PTName, string Title)
        {

            Excel.PivotTable PT = WS.PivotTables(PTName);
            int FirstRow = PT.TableRange2.Row;
            //ADD SECTION TITLE
            Excel.Range cell = WS.Range["A" + (FirstRow + 1)];
            cell.Value = Title;
            cell.Font.Size = 24;
            cell.Font.Color = System.Drawing.Color.Red.ToArgb();
            cell.Font.Bold = true;
            cell.Font.Italic = true;
            cell.EntireRow.RowHeight = 32;
            //HIDE PAGE GROUP FILTER
            cell = WS.Range["A" + FirstRow];
            cell.EntireRow.Hidden = true;
        }

        public void AddMedianSummary(Excel.Worksheet TableSheet, string TableName, ListingStatus Status)
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

            //TOTAL LISITNGS COUNT
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetCount(ListingSheet, ListingDataColNames.MLS, Status, "", "");
            //PRICE 
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.Price, Status, "", "");
            //DAYS ON MARKET DOM OR CDOM
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.DOM, Status, "", "");
            //
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.FlArTotFin, Status, "", "");
            //
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.PrcSqft, Status, "", "");
            //
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.Age, Status, "", "");
            //
            if (this.ReportType.ToString().IndexOf("Detached") < 0)
            {
                TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.StratMtFee, Status, "", "");
            }else
            {
                TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.Lot_Sz_Sq_Ft, Status, "", "");
                TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.LandValue , Status, "", "");
            }
                
            //
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.BCAValue, Status, "", "");
            //
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.Change_Percent, Status, "", "");
            //
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.LotPricePerSquareFeet, Status, "", "");
            TableSheet.Cells[medianRow, ++rw].Value = Library.GetMedianValue(ListingSheet, ListingDataColNames.ImproveValuePerSquareFeet, Status, "", "");

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
            double BCA_Price_CorCoe = Library.GetCorCoeValue(SourceSheet, ListingDataColNames.Price, ListingDataColNames.BCAValue);
            Cell1 = DestSheet.Range["A" + (LastRow + 1)];
            Cell1.Value = "BCA - Price: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 1)];
            Cell2.Value = BCA_Price_CorCoe;
            //2) BCA - Change
            double BCA_Change_CorCoe = Library.GetCorCoeValue(SourceSheet, ListingDataColNames.BCAValue, ListingDataColNames.Change_Percent);
            Cell1 = DestSheet.Range["A" + (LastRow + 2)];
            Cell1.Value = "BCA - Change: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 2)];
            Cell2.Value = BCA_Change_CorCoe;
            //3) FloorArea - Price Per Square Feet
            double FloorArea_PricePSF_CorCoe = Library.GetCorCoeValue(SourceSheet, ListingDataColNames.FlArTotFin, ListingDataColNames.PrcSqft);
            Cell1 = DestSheet.Range["A" + (LastRow + 3)];
            Cell1.Value = "FloorArea - Price Per Square Feet: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 3)];
            Cell2.Value = FloorArea_PricePSF_CorCoe;
            //4) Age - Fee
            double Age_Fee_CorCoe = Library.GetCorCoeValue(SourceSheet, ListingDataColNames.Age, ListingDataColNames.StratMtFee);
            Cell1 = DestSheet.Range["A" + (LastRow + 4)];
            Cell1.Value = "Age - Maint. Fee: CorCoe[-1, +1]";
            Cell2 = DestSheet.Range["E" + (LastRow + 4)];
            Cell2.Value = Age_Fee_CorCoe;
            //5) Age - Price Per Square Feet
            double Age_PricePSF_CorCoe = Library.GetCorCoeValue(SourceSheet, ListingDataColNames.Age, ListingDataColNames.PrcSqft);
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

        private void HideComplexSubTotal(Excel.Worksheet Sheet, string TableName)
        {
            Excel.Range Cell = null;
            Excel.Range CountCell = null;
            int subTotalCount = 0;

            foreach (Excel.Range row in Sheet.PivotTables(TableName).RowRange.Rows)
            {
                if (Sheet.Range["A" + row.Row].Value?.IndexOf("Total") > 0 && Sheet.Range["A" + row.Row].Value?.IndexOf("Grand Total") < 0)
                {
                    Cell = Sheet.Range["A" + row.Row];
                    Cell.Select();
                    Cell.Value = "SubTotal";
                    Cell.RowHeight = 19.5;
                    Cell.EntireRow.Font.Size = 11;
                    Cell.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                    CountCell = Sheet.Range["D" + row.Row];
                    
                    //subTotalCount++;
                    if(CountCell.Value == 1)
                    {
                        row.Select();
                        row.EntireRow.Hidden = true;
                        subTotalCount = 0;
                    }
                }
                subTotalCount++;
                //if (subTotalCount == 1)
                //{
                //    if (Sheet.Range["A" + row.Row].Value?.IndexOf("Total") > 0 && Sheet.Range["A" + row.Row].Value?.IndexOf("Grand Total") < 0)
                //    {
                //        row.Select();
                //        row.EntireRow.Hidden = true;
                //        subTotalCount = 0;
                //    }
                //}
                if (Sheet.Range["B" + row.Row].Value?.IndexOf("Total") > 0)
                {
                    row.Select();
                    row.EntireRow.Hidden = true;
                }
                if (Sheet.Range["C" + row.Row].Value?.IndexOf("Total") > 0)
                {
                    row.Select();
                    row.EntireRow.Hidden = true;
                }
            }
        }

        public void FormatMaxCells()
        {

            Excel.Range c = null;
            Excel.PivotTable PT = this.PivotSheet.PivotTables(this.PivotTableName);
            Excel.Worksheet WS = this.PivotSheet;
            long i = 0;
            long FirstRow = 0;
            long LastRow = 0;
            long FirstCol = 0;
            long LastCol = 0;
            string MaxCell = "";

            double Max = 0;

            //FIND THE LAST NON-BLANK CELL IN COLUMNA
            FirstRow = PT.TableRange1.Row + 2;
            LastRow = FirstRow + PT.TableRange1.Rows.Count - 4;
            FirstCol = PT.ColumnRange.Column + 1;
            LastCol = PT.ColumnRange.Column + PT.ColumnRange.Columns.Count - 1;

            for (long col = FirstCol; col <= LastCol; col++)
            {
                i = FirstRow;
                Max = Library.GetMax(this.PivotSheet, this.PivotTableName, col);
                MaxCell = "";
                for (i = FirstRow; i <= LastRow - 2; i++)
                {
                    c = WS.Cells[i, col];
                    if (c.Value2 != null && i <= LastRow - 2 && !((bool)c.Rows.Hidden) && (double)c.Value == Max)
                    {
                        if (WS.Cells[i, 1].Value == null || WS.Cells[i, 1].Value != "SubTotal")
                        {
                            MaxCell = WS.Range[c, c].Address;
                            WS.Range[MaxCell].Interior.ColorIndex = 0;
                            WS.Range[MaxCell].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            WS.Range[MaxCell].Font.Bold = true;
                            if(col == LastCol && this.Status == (char)ListingStatus.Sold)
                            {
                                this.MaxSoldPricePerSF_Improve = WS.Range[MaxCell].Value;
                                this.AverageSoldPricePerSF_Improve = WS.Cells[LastRow+1, col].Value;
                            }
                            if (col == LastCol - 1 && this.Status == (char)ListingStatus.Sold)
                            {
                                this.MaxSoldPricePerSF_Land = WS.Range[MaxCell].Value;
                                this.AverageSoldPricePerSF_Land = WS.Cells[LastRow+1, col].Value;
                            }
                            if (col == LastCol - 6 && this.Status == (char)ListingStatus.Sold && this.ReportType == ReportType.CMAAttached )
                            {
                                this.MaxSoldPricePerSF_Total = WS.Range[MaxCell].Value;
                                this.AverageSoldPricePerSF_Total = WS.Cells[LastRow+1 , col].Value;
                            }
                            if (col == LastCol - 7 && this.Status == (char)ListingStatus.Sold && this.ReportType == ReportType.CMADetached)
                            {
                                this.MaxSoldPricePerSF_Total = WS.Range[MaxCell].Value;
                                this.AverageSoldPricePerSF_Total = WS.Cells[LastRow+1, col].Value;
                            }
                        }
                    }
                }
            }
        }

        public void FormatMinCells()
        {
            Excel.Range c = null;
            Excel.PivotTable PT = this.PivotSheet.PivotTables(this.PivotTableName);
            Excel.Worksheet WS = this.PivotSheet;
            long i = 0;
            long FirstRow = 0;
            long LastRow = 0;
            long FirstCol = 0;
            long LastCol = 0;
            string MinCell = "";

            double Min = 0;

            //FIND THE LAST NON-BLANK CELL IN COLUMNA
            FirstRow = PT.TableRange1.Row + 2;
            LastRow = FirstRow + PT.TableRange1.Rows.Count - 4;
            FirstCol = PT.ColumnRange.Column + 1;
            LastCol = PT.ColumnRange.Column + PT.ColumnRange.Columns.Count - 1;
            for (long col = FirstCol; col <= LastCol; col++)
            {
                i = FirstRow;
                Min = Library.GetMin(this.PivotSheet, this.PivotTableName, col);
                MinCell = "";

                for (i = FirstRow; i <= LastRow - 2; i++)
                {
                    c = WS.Cells[i, col];
                    if (c.Value2 != null && i <= LastRow - 2 && !((bool)c.Rows.Hidden) && (double)c.Value == Min)
                    {
                        if (WS.Cells[i, 1].Value == null || WS.Cells[i, 1].Value != "SubTotal")
                        {
                            MinCell = WS.Range[c, c].Address;
                            WS.Range[MinCell].Interior.ColorIndex = 0;
                            WS.Range[MinCell].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                            WS.Range[MinCell].Font.Bold = true;
                        }
                    }
                }
            }
        }
    }
}

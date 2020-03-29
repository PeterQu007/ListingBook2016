using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListingBook2016
{

    public enum ReportType
    {
        MonthlyDetachedAllCities = 0,
        MonthlyDetachedAllCommunities = 1,
        MonthlyAttachedAllCities = 2,
        MonthlyAttachedAllCommunities = 3,
        MonthlyAttachedAllComplexes = 4,
        QuarterlyDetachedAllCities = 10,
        QuarterlyDetachedAllCommunities = 11,
        QuarterlyAttachedAllCities = 12,
        QuarterlyAttachedAllCommunities = 13,
        QuarterlyAttachedAllComplexes = 14,
        AnnuallyDetachedAllCities = 20,
        AnnuallyDetachedAllCommunities = 21,
        AnnuallyAttachedAllCities = 22,
        AnnuallyAttachedAllCommunities = 23,
        AnuuallyAttachedAllComplexes = 24,

        CMADetached = 30,
        CMAAttached = 31,

        SpecialityPriceChageTop10 = 40,
        SpecialtyActivePriceTop10 = 41,
        SpecialtySoldPriceTop10 = 42
    }

    public static class ListingDataSheet{
        public static string ParagonExport = "Spreadsheet";
        public static string MLSHelperExport = "Listings Table";
    }


    public enum ListingStatus
    {
        Sold = 'S',
        Active = 'A',
        Expire = 'X',
        Terminate = 'T',
        Cancel = 'C',
        OffMarket = 'Z'
    }
    public enum ReportDataSheet
    {
        ParagonExport = 0,
        MLSHelperExport = 1
    }


    public static class ListingDataColNames
    {
        public static string No = "A";
        public static string MLS = "B";
        public static string Status = "C";
        public static string Address = "D";
        public static string S_A = "E";
        public static string Price = "F";
        public static string PrcSqft = "G";
        public static string List_Date = "H";
        public static string DOM = "I";
        public static string CDOM = "J";
        public static string Complex_Subdivision = "K";
        public static string Tot_BR = "L";
        public static string Tot_Baths = "M";
        public static string FlArTotFin = "N";
        public static string Age = "O";
        public static string StratMtFee = "P";
        public static string TypeDwel = "Q";
        public static string Lot_Sz_Sq_Ft = "R";
        public static string PID = "S";
        public static string LandValue = "T";
        public static string ImproveValue = "U";
        public static string BCAValue = "V";
        public static string Change_Percent = "W";
        public static string Room27Dim1 = "X";
        public static string Address2 = "Y";
        public static string UnitNo = "Z";
        public static string City = "AA";
        public static string Area = "AB";
        public static string Postal_Code = "AC";
        public static string List_Price = "AD";
        public static string Prev_Price = "AE";
        public static string Price_Date = "AF";
        public static string Sold_Date = "AG";
        public static string Sold_Price = "AH";
        public static string SP_Sqft = "AI";
        public static string Processed_Date = "AJ";
        public static string Entry_Date = "AK";
        public static string Expiry_Date = "AL";
        public static string CDOMLS = "AM";
        public static string Search_Date = "AN";
        public static string SP_LP_Ratio = "AO";
        public static string SP_OLP_Ratio = "AP";
        public static string Yr_Blt = "AQ";
        public static string TotFlArea = "AR";
        public static string Kitchens = "AS";
        public static string Lot_Sz_Acres = "AT";
        public static string Frontage_Feet = "AU";
        public static string Depth = "AV";
        public static string Prop_Type = "AW";
        public static string Room27Type = "AX";
        public static string Parking_Places_Covered = "AY";
        public static string Legal_Description = "AZ";
        public static string Title_to_Land = "BA";
        public static string Units_in_Development = "BB";
        public static string Stories_in_Building = "BC";
        public static string Rentals_Allowed = "BD";
        public static string TotalPrkng = "BE";
        public static string Locker = "BF";
        public static string List_Firm_1_Code_Office_Name = "BG";
        public static string List_Sales_Rep_1_Agent_Name = "BH";
        public static string List_Firm_2_Code_Office_Name = "BI";
        public static string List_Sales_Rep_2_Agent_Name = "BJ";
        public static string Selling_Office_1_Office_Name = "BK";
        public static string Sell_Sales_Rep_1_Agent_Name = "BL";
        public static string Selling_Office_2_Office_Name_ = "BM";
        public static string Sell_Sales_Rep_2_Agent_Name = "BN";
        public static string Owner_Name = "BO";
        public static string Buyer = "BP";
        //ADD New Columns for lot price/square feet, improve price/square feet
        public static string LotValuePercentageOfBCA = "BQ";
        public static string LotValueMarketAssessment = "BR";
        public static string ImproveValueMarketAssessment = "BS";
        public static string LotPricePerSquareFeet = "BT";
        public static string ImproveValuePerSquareFeet = "BU";
    }
    public class DataProcessing
    {

        public Excel.Worksheet ListingSheet;

        public DataProcessing(Excel.Worksheet ListingSheet)
        {
            this.ListingSheet = ListingSheet;
        }

        public bool ValidateData(ReportType rptType)
        {
            bool bStop = false;
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            //Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.DisplayStatusBar = true;

            Excel.Range rng = null;
            rng = this.ListingSheet.Cells[2, ListingDataColNames.Prop_Type];
            rng.Select();

            if (rptType.ToString().IndexOf("Attached") > 0)
            {
                if (rng.Value2 != "Residential Attached")
                {
                    MessageBox.Show("Wrong Listings for Attached Homes");
                    return true;
                }
            }

            if (!ListingSheet.AutoFilterMode)
            {
                ListingSheet.Range["A1"].Select();
                ListingSheet.Range["A1"].AutoFilter(1);
            }

            //VALIDATE COMPLEX.NAME
            bStop |= ValidateColumnBlankCell(ListingDataColNames.Complex_Subdivision);
            //VALIDATE MAINT.FEE
            if (rptType.ToString().IndexOf("Attached") > 0)
            {
                bStop |= ValidateColumnZeroValue(ListingDataColNames.StratMtFee);
            }
            //VALIDATE AGE
            bStop |= ValidateColumnAge(ListingDataColNames.Age);
            //VALIDATE BCA.VALUE
            bStop |= ValidateColumnBlankCell(ListingDataColNames.BCAValue);
            bStop |= ValidateColumnZeroValue(ListingDataColNames.BCAValue);
            //VALIDATE CHANGE%
            bStop |= ValidateColumnBlankCell(ListingDataColNames.Change_Percent);
            //VALIDATE ADDRESS
            bStop |= ValidateColumnBlankCell(ListingDataColNames.Address2);

            Globals.ThisAddIn.Application.StatusBar = bStop? "Check DataSheet RED Highlights" : "Validate Data of Attached Done!";
            //Globals.ThisAddIn.Application.ScreenUpdating = true;
            //Globals.ThisAddIn.Application.DisplayAlerts = true;
            return bStop;
        }

        public bool AddLotAndImproveUnitPrice()
        {
            Excel.Range last = ListingSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int LastRow = last.Row;
            bool bAdded = false;
            string ColIndex = ListingDataColNames.LotValuePercentageOfBCA;

            Excel.Range TopCell = ListingSheet.Cells[2, ColIndex];
            TopCell.Select();
            Excel.Range BottomCell = ListingSheet.Cells[LastRow, ColIndex];
            Excel.Range Cell = null;
            Excel.Range rng = null;
            
            Cell = ListingSheet.Cells[1, ListingDataColNames.LotValuePercentageOfBCA];
            Cell.Value2 = "Lot%";
            ListingSheet.Cells[1, ListingDataColNames.LotValueMarketAssessment].Value = "Lot$ Market";
            ListingSheet.Cells[1, ListingDataColNames.ImproveValueMarketAssessment].Value = "Improve$ Market";
            ListingSheet.Cells[1, ListingDataColNames.LotPricePerSquareFeet].Value = "Lot$ PerSF";
            ListingSheet.Cells[1, ListingDataColNames.ImproveValuePerSquareFeet].Value = "Improve$ PerSF";

            rng = ListingSheet.Range[TopCell, BottomCell];
            rng.Select();

            double lotMarket = 0;

            foreach(Excel.Range c in rng)
            {
                try
                {
                    lotMarket = ListingSheet.Cells[c.Row, ListingDataColNames.LandValue].Value2 / ListingSheet.Cells[c.Row.ToString(), ListingDataColNames.BCAValue].Value2;
                    c.Value2 = lotMarket;
                }
                catch (Exception e)
                {
                    c.Value2 = e.Message;
                }
            }

            ColIndex = ListingDataColNames.LotValueMarketAssessment;
            TopCell = ListingSheet.Cells[2, ColIndex];
            BottomCell = ListingSheet.Cells[LastRow, ColIndex];
            rng = ListingSheet.Range[TopCell, BottomCell];
            foreach (Excel.Range c in rng)
            {
                try
                {
                    lotMarket = ListingSheet.Cells[c.Row, ListingDataColNames.Price].Value2
                        * (ListingSheet.Cells[c.Row, ListingDataColNames.LandValue].Value2 / ListingSheet.Cells[c.Row.ToString(), ListingDataColNames.BCAValue].Value2);
                    c.Value2 = lotMarket;

                }
                catch(Exception e)
                {
                    c.Value2 = e.Message;
                }
            }

            ColIndex = ListingDataColNames.ImproveValueMarketAssessment;
            TopCell = ListingSheet.Cells[2, ColIndex];
            BottomCell = ListingSheet.Cells[LastRow, ColIndex];
            rng = ListingSheet.Range[TopCell, BottomCell];
            foreach (Excel.Range c in rng)
            {
                try
                {
                    lotMarket = ListingSheet.Cells[c.Row, ListingDataColNames.Price].Value2 - ListingSheet.Cells[c.Row, ListingDataColNames.LotValueMarketAssessment].Value2;
                    c.Value2 = lotMarket;
                }
                catch(Exception e)
                {
                    c.Value2 = e.Message;
                }
            }
            //Lot Value Per Square Feet
            //Strata Property Use: Lot Market Value / Total Floor Area
            //Single House Use: Lot Market Value / Lot Size
            ColIndex = ListingDataColNames.LotPricePerSquareFeet;
            TopCell = ListingSheet.Cells[2, ColIndex];
            BottomCell = ListingSheet.Cells[LastRow, ColIndex];
            rng = ListingSheet.Range[TopCell, BottomCell];
            foreach (Excel.Range c in rng)
            {
                try
                {
                    if (ListingSheet.Cells[c.Row, ListingDataColNames.TypeDwel].Value2 == "HOUSE" || ListingSheet.Cells[c.Row, ListingDataColNames.Lot_Sz_Sq_Ft].Value2 != 0)
                    {
                        lotMarket = ListingSheet.Cells[c.Row, ListingDataColNames.LotValueMarketAssessment].Value2 / ListingSheet.Cells[c.Row, ListingDataColNames.Lot_Sz_Sq_Ft].Value2;
                    }
                    else
                    {
                        lotMarket = ListingSheet.Cells[c.Row, ListingDataColNames.LotValueMarketAssessment].Value2 / ListingSheet.Cells[c.Row, ListingDataColNames.FlArTotFin].Value2;
                    }
                    c.Value2 = lotMarket;
                }
                catch (Exception e)
                {
                    c.Value2 = e.Message;
                    Console.WriteLine(e.Message);
                }
            }

            ColIndex = ListingDataColNames.ImproveValuePerSquareFeet;
            TopCell = ListingSheet.Cells[2, ColIndex];
            BottomCell = ListingSheet.Cells[LastRow, ColIndex];
            rng = ListingSheet.Range[TopCell, BottomCell];
            foreach (Excel.Range c in rng)
            {
                try
                {
                    lotMarket = ListingSheet.Cells[c.Row, ListingDataColNames.ImproveValueMarketAssessment].Value2 / ListingSheet.Cells[c.Row, ListingDataColNames.FlArTotFin].Value2;
                    c.Value2 = lotMarket;
                }
                catch (Exception e)
                {
                    c.Value2 = e.Message;
                }
            }

            return bAdded;
        }
        public bool ValidateColumnBlankCell(string ColIndex)
        {
            Excel.Range last = ListingSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int LastRow = last.Row;
            bool bStop = false;

            Excel.Range TopCell = ListingSheet.Cells[2, ColIndex];
            Excel.Range BottomCell = ListingSheet.Cells[LastRow, ColIndex];
            Excel.Range Cells = null;
            Excel.Range rng = null;

            rng = ListingSheet.Range[TopCell, BottomCell];
            rng.Select();
            try
            {
                Cells = rng.SpecialCells(Excel.XlCellType.xlCellTypeBlanks).Cells;
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
                return false;
            }

            if (Cells.Count > 0)
            {
                foreach(Excel.Range c in Cells)
                {
                    if(ListingSheet.Cells[c.Row,1].Value > 0)
                    {
                        Cells.Interior.Color = Excel.XlRgbColor.rgbRed;
                        bStop = true;
                        break;
                    }
                }
            }
            return bStop;
        }

        public bool ValidateColumnZeroValue(string ColIndex)
        {
            int lRow = 0;
            bool bStop = false;
            Excel.Range rng = null;
            ListingSheet.AutoFilterMode = false;

            lRow = ListingSheet.Cells[ListingSheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            rng = ListingSheet.Range[ListingSheet.Cells[2, ColIndex], ListingSheet.Cells[lRow, ColIndex]];
            rng.Select();

            foreach (Excel.Range c in rng.Cells)
            {
                if (c.Value == 0 && ListingSheet.Cells[c.Row, 1].Value > 0)
                {
                    c.Interior.Color = Excel.XlRgbColor.rgbRed;  //System.Drawing.Color.Red;
                    bStop = true;
                };

            }
            return bStop;
        }

        public bool ValidateColumnAge(string ColIndex)
        {
            int lRow = 0;
            bool bStop = false;
            Excel.Range rng = null;

            lRow = ListingSheet.Cells[ListingSheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            rng = ListingSheet.Range[ListingSheet.Cells[2, ColIndex], ListingSheet.Cells[lRow, ColIndex]];
            rng.Select();

            foreach (Excel.Range c in rng.Cells)
            {
                if (c.Value > 200 && ListingSheet.Cells[c.Row, 1].Value > 0)
                {
                    c.Interior.Color = Excel.XlRgbColor.rgbRed;
                    bStop = true;
                }
            }
            return bStop;
        }
    }
    public static class Library

    {
        
        public static bool SheetExist(string SheetName)
        {
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                if (sheet.Name == SheetName)
                {
                    return true;
                }
            }
            return false;
        }

        public static string[] StatusArray(ListingStatus status)
        {
            List<string> StatusList = new List<string>();

            switch (status)
            {
                case ListingStatus.Active:
                case ListingStatus.Sold:
                    StatusList.Add(((char)status).ToString());
                    break;
                case ListingStatus.OffMarket:
                    StatusList.Add(((char)ListingStatus.Expire).ToString());
                    StatusList.Add(((char)ListingStatus.Terminate).ToString());
                    StatusList.Add(((char)ListingStatus.Cancel).ToString());
                    break;
            }

            return StatusList.ToArray();
        }

        public static int GetLastRow(Excel.Worksheet Sheet)
        {
            int row = 0;
            row = Sheet.Cells[Sheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            return row;
        }

        public static int GetLastCol(Excel.Worksheet Sheet)
        {
            Excel.Range last = Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = Sheet.Range["A1", last];

            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;

            return lastUsedColumn;
        }

        public static double GetMedianValue(Excel.Worksheet Sheet, string ColLabel, ListingStatus Status, string City, string DataName)
        {
            double MedianValue = 0;
            Excel.Range rng = null;
            int firstRow = 0;
            int lastRow = 0;
            int iCol = 0;
            int iFilterCol = 0;

            iCol = Sheet.Cells[1, ColLabel].Column; //Convert ColLable to ColNumber
            lastRow = GetLastRow(Sheet);
            Sheet.Select();

            string[] StatusArray = Library.StatusArray(Status);
            if (City == "")
            {
                //Status - column C
                iFilterCol = Sheet.Cells[1, "C"].Column;
                Sheet.Range["A1"].AutoFilter(iFilterCol, StatusArray, Excel.XlAutoFilterOperator.xlFilterValues);
                DataName = DataName + "_" + Status;
            }
            else
            {
                //City - column X
                iFilterCol = Sheet.Cells[1, "X"].Column;
                Sheet.Range["A1"].AutoFilter(iFilterCol, City, Excel.XlAutoFilterOperator.xlAnd);
                DataName = DataName + "_" + City;
            }

            lastRow = GetLastRow(Sheet);
            if (lastRow > 1)
            {
                firstRow = Sheet.Range["A2:A" + lastRow].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Row;
                rng = Sheet.Range[ColLabel + firstRow + ":" + ColLabel + lastRow];
                try
                {
                    MedianValue = Globals.ThisAddIn.Application.WorksheetFunction.Aggregate(12, 1, rng);
                }
                catch(Exception e)
                {
                    MedianValue = 0;
                    Console.WriteLine(e.Message);
                }
            }
            return Math.Round(MedianValue, 3);

        }

        public static int GetCount(Excel.Worksheet Sheet, string ColLabel, ListingStatus Status, string City, string DataName)
        {
            int CountValue = 0;
            Excel.Range rng = null;
            int firstRow = 0;
            int lastRow = 0;
            int iCol = 0;
            int iFilterCol = 0;

            iCol = Sheet.Cells[1, ColLabel].Column; //Convert ColLable to ColNumber
            lastRow = GetLastRow(Sheet);
            Sheet.Select();
           
            string[] StatusArray = Library.StatusArray(Status);
            if (City == "")
            {
                //Status - column C
                iFilterCol = Sheet.Cells[1, "C"].Column;
                Sheet.Range["A1"].AutoFilter(iFilterCol, StatusArray, Excel.XlAutoFilterOperator.xlFilterValues);
                DataName = DataName + "_" + Status;
            }
            else
            {
                //City - column X
                iFilterCol = Sheet.Cells[1, "X"].Column;
                Sheet.Range["A1"].AutoFilter(iFilterCol, City, Excel.XlAutoFilterOperator.xlAnd);
                DataName = DataName + "_" + City;
            }

            lastRow = GetLastRow(Sheet);
            if (lastRow > 1)
            {
                firstRow = Sheet.Range["A2:A" + lastRow].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Row;
                rng = Sheet.Range[ColLabel + firstRow + ":" + ColLabel + lastRow];
                CountValue = (int)Globals.ThisAddIn.Application.WorksheetFunction.Aggregate(3, 1, rng);
            }
            return CountValue;

        }

        public static double GetCorCoeValue(Excel.Worksheet Sheet, string ColLabel1, string ColLabel2)
        {
            double CorCoeValue = 0;
            Excel.Range rng1 = null;
            Excel.Range rng2 = null;
            int firstRow = 0;
            int lastRow = 0;
            int iCol1 = 0;
            int iCol2 = 0;

            Sheet.Select();

            iCol1 = Sheet.Cells[1, ColLabel1].Column; //Convert ColLable1 to ColNumber
            iCol2 = Sheet.Cells[1, ColLabel2].Column; //Convert ColLable2 to ColNumber
            lastRow = GetLastRow(Sheet);

            //Remove the Filters

            //if (City == "")
            //{
            //    //Status - column C
            //    iFilterCol = Sheet.Cells[1, "C"].Column;
            //    Sheet.Range["A1"].AutoFilter(iFilterCol, Status, Excel.XlAutoFilterOperator.xlAnd);
            //    DataName = DataName + "_" + Status;
            //}
            //else
            //{
            //    //City - column X
            //    iFilterCol = Sheet.Cells[1, "X"].Column;
            //    Sheet.Range["A1"].AutoFilter(iFilterCol, City, Excel.XlAutoFilterOperator.xlAnd);
            //    DataName = DataName + "_" + City;
            //}
            try
            {
                if (lastRow > 1)
                {
                    firstRow = Sheet.Range["A2:A" + lastRow].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Row;
                    rng1 = Sheet.Range[ColLabel1 + firstRow + ":" + ColLabel1 + lastRow];
                    rng2 = Sheet.Range[ColLabel2 + firstRow + ":" + ColLabel2 + lastRow];
                    CorCoeValue = Globals.ThisAddIn.Application.WorksheetFunction.Correl(rng1, rng2);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return 0;
            }

            return Math.Round(CorCoeValue, 3);
        }

        public static double GetMax(Excel.Worksheet WS, string TableName, long col)
        {
            Excel.Range c = null;
            Excel.PivotTable PT = WS.PivotTables(TableName);
            long i = 0;
            long FirstRow = 0;
            long LastRow = 0;

            //FIND THE LAST NON-BLANK CELL IN COLUMNA
            FirstRow = PT.TableRange1.Row + 2;
            LastRow = FirstRow + PT.TableRange1.Rows.Count - 4;

            double Max = 0;
            i = FirstRow;
            for (i = FirstRow; i <= LastRow - 2; i++)
            {
                c = WS.Cells[i, col];
                c.Select();
                if (c.Value2 != null && i <= LastRow - 2 && !((bool)c.Rows.Hidden) && (double)c.Value > Max)
                {
                    if (WS.Cells[i, 1].Value == null || WS.Cells[i, 1].Value != "SubTotal")
                    {
                        Max = (double)c.Value;
                    }
                }
            }

            return Max;
        }

        public static double GetMin(Excel.Worksheet WS, string TableName, long col)
        {
            Excel.Range c = null;
            Excel.PivotTable PT = WS.PivotTables(TableName);
            long i = 0;
            long FirstRow = 0;
            long LastRow = 0;

            //FIND THE LAST NON-BLANK CELL IN COLUMNA
            FirstRow = PT.TableRange1.Row + 2;
            LastRow = FirstRow + PT.TableRange1.Rows.Count - 4;

            double Min = 0;
            i = FirstRow;
            Min = (double)WS.Cells[i, col].Value;
            for (i = FirstRow; i <= LastRow - 2; i++)
            {
                c = WS.Cells[i, col];
                c.Select();
                if (c.Value2 != null && i <= LastRow - 2 && !((bool)c.Rows.Hidden) && (double)c.Value < Min)
                {
                    if (WS.Cells[i, 1].Value == null || WS.Cells[i, 1].Value != "SubTotal")
                    {
                        Min = (double)c.Value;
                    }

                }
            }
            return Min;
        }

        public static string[] GetCities(string CityColLable)
        {
            Excel.Worksheet WS = Globals.ThisAddIn.Application.Worksheets["Listings Table"];
            WS.AutoFilterMode = false;
            int LastRow = GetLastRow(WS);
            object[,] CityValues = WS.Range[CityColLable + "2:" + CityColLable + LastRow].Value2;
            List<string> CityList = new List<string>();
            foreach (string c in CityValues)
            {
                CityList.Add(c);
            }
            string[] UniqueCities = CityList.ToArray().Distinct().ToArray();
            Array.Sort(UniqueCities, StringComparer.InvariantCulture);
            return UniqueCities;
        }
    }
}

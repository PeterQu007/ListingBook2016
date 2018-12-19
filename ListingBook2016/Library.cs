﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListingBook2016
{

    namespace ReportType
    {
        public enum Detached
        {
            AllCities = 0,
            AllCommunities = 1
        }
        public enum Attached
        {
            AllCities = 2,
            AllCommunities = 3,
            AllNeighborhoods =4
        }
    }
    public enum ListingStatus
    {
        Sold = 0,
        Active =1,
        Expire = 2,
        Terminate =3,
        Cancel =4,
        OffMarket =5
    }
 
    public enum ReportDataSheet
    {
        ParagonExport = 0,
        MLSHelperExport = 1
    }

    public class DataProcessing
    {

        public Excel.Worksheet ListingSheet;

        public DataProcessing(Excel.Worksheet ListingSheet)
        {
            this.ListingSheet = ListingSheet;
        }

        public bool ValidateData_Attached()
        {
            bool bStop = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.DisplayStatusBar = true;

            Excel.Range rng = null;
            rng = this.ListingSheet.Cells[1, "V"];
            rng.Select();

            if (!rng.Value2.StartsWith("Unit"))
            {
                MessageBox.Show("Wrong Listings for Attached Homes");
                return true;
            }

            if (!ListingSheet.AutoFilterMode)
            {
                ListingSheet.Range["A1"].Select();
                ListingSheet.Range["A1"].AutoFilter(1);
            }

            //VALIDATE COMPLEX.NAME
            bStop |= ValidateColumnBlankCell("F");
            //VALIDATE MAINT.FEE
            bStop |= ValidateColumnZeroValue("P");
            //VALIDATE AGE
            bStop |= ValidateColumnAge("O");
            //VALIDATE BCA.VALUE
            bStop |= ValidateColumnBlankCell("R");
            bStop |= ValidateColumnZeroValue("R");
            //VALIDATE CHANGE%
            bStop |= ValidateColumnBlankCell("S");
            //VALIDATE ADDRESS
            bStop |= ValidateColumnBlankCell("U");

            Globals.ThisAddIn.Application.StatusBar = "Validate Data of Attached Done!";
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            return bStop;
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
                Cells.Interior.Color = Excel.XlRgbColor.rgbRed;
                bStop = true;
            }
            return bStop;
        }

        public bool ValidateColumnZeroValue(string ColIndex)
        {
            int lRow = 0;
            bool bStop = false;
            Excel.Range rng = null;

            lRow = ListingSheet.Cells[ListingSheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
            rng = ListingSheet.Range[ListingSheet.Cells[2, ColIndex], ListingSheet.Cells[lRow, ColIndex]];
            rng.Select();

            foreach (Excel.Range c in rng.Cells)
            {
                if (c.Value == 0)
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
                if (c.Value > 200)
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
        static char Sold = 'S';
        static char Active = 'A';
        static char Expire = 'X';
        static char Terminate = 'T';
        static char Cancel = 'C';
        static char OffMarket = 'Z';
        public static char GetStatus(ListingStatus status)
        {
            switch (status)
            {
                case ListingStatus.Active:
                    return Active;
                case ListingStatus.Sold:
                    return Sold;
                case ListingStatus.Expire:
                    return Expire;
                case ListingStatus.Terminate:
                    return Terminate;
                case ListingStatus.Cancel:
                    return Cancel;
                case ListingStatus.OffMarket:
                    return OffMarket;
                default:
                    return Sold;
            }
        }
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

        public static double GetMedianValue(Excel.Worksheet Sheet, string ColLabel, char Status, string City, string DataName)
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

            if (City == "")
            {
                //Status - column C
                iFilterCol = Sheet.Cells[1, "C"].Column;
                Sheet.Range["A1"].AutoFilter(iFilterCol, Status.ToString(), Excel.XlAutoFilterOperator.xlAnd);
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
                MedianValue = Globals.ThisAddIn.Application.WorksheetFunction.Aggregate(12, 1, rng);
            }
            return Math.Round(MedianValue, 3);

        }

        public static int GetCount(Excel.Worksheet Sheet, string ColLabel, char Status, string City, string DataName)
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

            if (City == "")
            {
                //Status - column C
                iFilterCol = Sheet.Cells[1, "C"].Column;
                Sheet.Range["A1"].AutoFilter(iFilterCol, Status.ToString(), Excel.XlAutoFilterOperator.xlAnd);
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

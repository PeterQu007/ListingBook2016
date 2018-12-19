using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListingBook2016
{
    public class ReportMonthlyDetached
    {
        public Excel.Worksheet ListingSheet;
        public Excel.Worksheet PivotSheet;
        
        public ReportMonthlyDetached(Excel.Worksheet ws)
        {
            this.ListingSheet = ws;

        }

        public void AllCities()
        {
            int PivotTableTopPaddingRows = 5;
            string PivotSheetName = "AllCitiesReport";
            string PivotTableName = "";

            /////////////////////
            //PIVOT TABLE
            PivotTableName = "PT_" + PivotSheetName;
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            PivotTableCities cpt = new PivotTableCities(PivotSheetName, PivotTableTopPaddingRows, PivotTableName, ReportType.Detached.AllCities);
            this.PivotSheet = cpt.PivotSheet;
            cpt.FormatPivotTable(cpt.PivotSheet, PivotTableName);

            cpt.PivotSheet.Select();
        }

        public void AllCommunities()
        {
            int PivotTableTopPaddingRows = 5;
            string PivotSheetName = "AllCommunitiesDetachedReport";
            string PivotTableName = "";

            /////////////////////
            //PIVOT TABLE
            PivotTableName = "PT_" + PivotSheetName;
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            PivotTableCities cpt = new PivotTableCities(PivotSheetName, PivotTableTopPaddingRows, PivotTableName, ReportType.Detached.AllCommunities);
            this.PivotSheet = cpt.PivotSheet;
            cpt.FormatPivotTable(cpt.PivotSheet, PivotTableName);

            cpt.PivotSheet.Select();
        }
       
    }
}

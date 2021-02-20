using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ListingBook2016
{
    public partial class ThisAddIn
    {
        private SQLEdit _tpSqlEdit;
        public Microsoft.Office.Tools.CustomTaskPane TpSqlEditCustomTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddTpSqlEdit();
            this.Application.WorkbookOpen  += new Excel.AppEvents_WorkbookOpenEventHandler(WorkbookOpenHandler);
            ((Excel.AppEvents_Event)Application).NewWorkbook += ThisAddIn_NewWorkbook;
        }

        private void WorkbookOpenHandler(Excel.Workbook wb)
        {
            //System.Windows.Forms.MessageBox.Show("OPEN workbook event" + wb.Name);
            //Microsoft.Office.Interop.Excel.Application myExcel;
            //myExcel = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            //System.Windows.Forms.MessageBox.Show(myExcel.ActiveWorkbook.FullName); // gives full path
            //System.Windows.Forms.MessageBox.Show(myExcel.ActiveWorkbook.Name);

            // Keeping track
            bool found = false;
            // Loop through all worksheets in the workbook
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name.Contains("(") && sheet.Name.Contains(")") && wb.Sheets.Count == 1)
                {
                    sheet.Name = "Spreadsheet";
                }
                if (sheet.Cells[1, 2].Value == "Pics")
                {
                    sheet.Columns["B"].Delete(); //Delete the picture address column
                }
                // Check the name of the current sheet
                switch (sheet.Name)
                {
                    case "Spreadsheet":
                    case "Listings Table":
                        Globals.Ribbons.Ribbon1.ReportDataSheet = sheet.Name;
                        found = true;
                        break;
                    default:
                        Globals.Ribbons.Ribbon1.ReportDataSheet = wb.ActiveSheet.Name;
                        found = false;
                        break;
                }
            }
        }

        private void ThisAddIn_NewWorkbook(Excel.Workbook wb)
        {
            //System.Windows.Forms.MessageBox.Show("NEW Workbook event");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)

        {
            Globals.ThisAddIn.Application.WorkbookOpen -= this.WorkbookOpenHandler;
            ((Excel.AppEvents_Event)Application).NewWorkbook -= ThisAddIn_NewWorkbook;
        }

        //Create the Custom TaskPane and Dock Bottom
        //You must Add your Control here
        private void AddTpSqlEdit()
        {
            // DISABLE SQL TASK PANE ON AUGUST 15 2020 DUE TO RARELY USE OF THE FEATURE
            //_tpSqlEdit = new SQLEdit();
            //TpSqlEditCustomTaskPane = CustomTaskPanes.Add(_tpSqlEdit, "SQL Editor");
            //TpSqlEditCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
            ////Show TaskPane
            //TpSqlEditCustomTaskPane.Visible = true;
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

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
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        //Create the Custom TaskPane and Dock Bottom
        //You must Add your Control here
        private void AddTpSqlEdit()
        {
            _tpSqlEdit = new SQLEdit();
            TpSqlEditCustomTaskPane = CustomTaskPanes.Add(_tpSqlEdit, "SQL Editor");
            TpSqlEditCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
            //Show TaskPane
            TpSqlEditCustomTaskPane.Visible = true;
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

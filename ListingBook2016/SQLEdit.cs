using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using ET = Microsoft.Office.Tools.Excel;

namespace ListingBook2016
{
    public partial class SQLEdit : UserControl
    {
        public SQLEdit()
        {
            InitializeComponent();
        }

        private void buttonGetData_Click(object sender, EventArgs e)
        {
            PopulateFromSql();
        }
        private void PopulateFromSql()
        {
            try
            {
                // DataTable Construction with Adapter and Connection 
                var conn = new SqlConnection(textBoxCS.Text);
                var strSql = richTextBoxSQLEdit.Text;
                conn.Open();
                var da = new SqlDataAdapter(strSql, conn);
                var dt = new System.Data.DataTable();
                da.Fill(dt);

                // Define the active Worksheet
                var sht = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

                var rowCount = 0;
                progressBarGetData.Minimum = 1;
                progressBarGetData.Maximum = dt.Rows.Count;

                // Loop thrue the Datatable and add it to Excel
                foreach (DataRow dr in dt.Rows)
                {
                    rowCount += 1;
                    for (var i = 1; i < dt.Columns.Count + 1; i++)
                    {
                        // Add the header the first time through 
                        if (rowCount == 2)
                        {
                            // Add the Columns using the foreach i++ to get the cell references
                            if (sht != null) sht.Cells[1, i] = dt.Columns[i - 1].ColumnName;
                        }
                        // Increment value in the Progress Bar
                        progressBarGetData.Value = rowCount;
                        // Add the Columns using the foreach i++ to get the cell references
                        if (sht != null) sht.Cells[rowCount, i] = dr[i - 1].ToString();
                        // Refresh the Progress Bar
                        progressBarGetData.Refresh();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }

        }
    }
}

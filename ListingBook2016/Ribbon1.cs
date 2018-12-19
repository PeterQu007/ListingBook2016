using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ListingBook2016
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnDetachedMonthlyCities_Click(object sender, RibbonControlEventArgs e)
        {
            ReportMonthlyDetached report = new ReportMonthlyDetached(Globals.ThisAddIn.Application.Worksheets["Sheet1"]);
            report.ListingSheet.Activate();
            report.ListingSheet.AutoFilterMode = false;

            report.AllCities();
        }

        private void btnDetachedMonthlyCommunities_Click(object sender, RibbonControlEventArgs e)
        {
            ReportMonthlyDetached report = new ReportMonthlyDetached(Globals.ThisAddIn.Application.Worksheets["Sheet1"]);
            report.ListingSheet.Activate();
            report.ListingSheet.AutoFilterMode = false;

            report.AllCommunities();
        }

        private void btmCondoSold_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets["Listings Table"]);
            bool bCMA = true;
            cma.Condo(ListingStatus.Sold, bCMA);
        }

        private void btnCondoActive_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets["Listings Table"]);
            bool bCMA = true;
            cma.Condo(ListingStatus.Active, bCMA);
        }
    }
}

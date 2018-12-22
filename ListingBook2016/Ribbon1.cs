using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace ListingBook2016
{
    public partial class Ribbon1
    {
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
        string CMADataSheet = string.Empty;
        string PublicReportDataSheet = string.Empty;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            CMADataSheet = ListingDataSheet.MLSHelperExport;
            PublicReportDataSheet = ListingDataSheet.ParagonExport;
        }

        #region MARKETING.REPORTS
        //PUBLIC.MARKETING.REPORTS
        private void btnDetachedMonthlyCities_Click(object sender, RibbonControlEventArgs e)
        {
            ReportMonthlyDetached report = new ReportMonthlyDetached(Globals.ThisAddIn.Application.Worksheets[PublicReportDataSheet]);
            report.ListingSheet.Activate();
            report.ListingSheet.AutoFilterMode = false;

            report.AllCities();
        }

        private void btnDetachedMonthlyCommunities_Click(object sender, RibbonControlEventArgs e)
        {
            ReportMonthlyDetached report = new ReportMonthlyDetached(Globals.ThisAddIn.Application.Worksheets[PublicReportDataSheet]);
            report.ListingSheet.Activate();
            report.ListingSheet.AutoFilterMode = false;

            report.AllCommunities();
        }
        #endregion

        #region RESIDENTIAL.ATTCHED.PROPERTY.CMA.REPORTS
        //RESIDENTIAL.ATTACHED.PROPERTY
        private void btmCondoSold_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMAAttached);
            cma.Residential(ListingStatus.Sold);
        }

        private void btnCondoActive_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMAAttached);
            cma.Residential(ListingStatus.Active);
        }

        private void btnCondoCMA_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMAAttached);
            try { cma.Residential(ListingStatus.Sold); } catch (Exception ex) { };
            try { cma.Residential(ListingStatus.Active); } catch (Exception ex) { };
            try { cma.Residential(ListingStatus.OffMarket); } catch (Exception ex) { };
        }
        #endregion

        #region RESIDENTIAL.DETACHED.PROPERTY.CMA.REPORTS
        //RESIDENTIAL.DETACHED
        //SOLD
        private void btnDetachedSoldCMA_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMADetached);
            cma.Residential(ListingStatus.Sold);
        }
        //ACTIVE
        private void btnDetachedActiveCMA_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMADetached);
            cma.Residential(ListingStatus.Active);
        }
        //SOLD, ACTIVE. OFFMARKET
        private void btnDetachedAllCMA_Click(object sender, RibbonControlEventArgs e)
        {
            ReportCMA cma = new ReportCMA(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMADetached);
            try { cma.Residential(ListingStatus.Sold); } catch (Exception ex) { };
            try { cma.Residential(ListingStatus.Active); } catch (Exception ex) { };
            try { cma.Residential(ListingStatus.OffMarket); } catch (Exception ex) { };
        }


        #endregion

        #region BUYER.REPORTS
        private void btnBuyerDetachedReport_Click(object sender, RibbonControlEventArgs e)
        {
            ReportBuyer buyerReport = new ReportBuyer(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMADetached);
            try { buyerReport.Residential(ListingStatus.Active); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.Sold); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.OffMarket); } catch (Exception ex) { throw ex; };
        }

        private void btnBuyerAttachedReport_Click(object sender, RibbonControlEventArgs e)
        {
            ReportBuyer buyerReport = new ReportBuyer(Globals.ThisAddIn.Application.Worksheets[CMADataSheet], ReportType.CMAAttached);
            try { buyerReport.Residential(ListingStatus.Active); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.Sold); } catch (Exception ex) { throw ex; };
            try { buyerReport.Residential(ListingStatus.OffMarket); } catch (Exception ex) { throw ex; };
        }
        #endregion
    }
}

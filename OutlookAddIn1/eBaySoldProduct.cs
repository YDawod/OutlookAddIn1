using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    class eBaySoldProduct
    {
        public string PaypalTransactionID { get; set; }
        public DateTime PaypalPaidDateTime { get; set; }
        public string PaypalPaidEmailPdf { get; set; }
        public string eBayItemNumber { get; set; }
        public DateTime eBaySoldDateTime { get; set; }
        public string eBayItemName { get; set; }
        public int eBayListingQuality { get; set; }
        public int eBaySoldQuality { get; set; }
        public int eBayRemainingQuality { get; set; }
        public decimal eBaySoldPrice { get; set; }
        public string eBayUrl { get; set; }
        public string eBaySoldEmailPdf { get; set;}
        public string BuyerName { get; set; }
        public string BuyerID { get; set; }
        public string BuyerAddress1 { get; set; }
        public string BuyerAddress2 { get; set; }
        public string BuyerCity { get; set; }
        public string BuyerState { get; set; }
        public string BuyerZip { get; set; }
        public string BuyerEmail { get; set; }
        public string BuyerNote { get; set; }
        public string CostcoUrlNumber { get; set; }
        public string CostcoUrl { get; set; }
        public decimal CostcoPrice { get; set; }
        public string CostcoOrderNumber { get; set; }
        public string CostcoItemName { get; set; }
        public string CostcoItemNumber { get; set; }
        public string CostcoTrackingNumber { get; set; }
        public string CostcoShipDate { get; set; }
        public string CostcoTaxExemptPdf { get; set; }
        public string CostcoOrderEmailPdf { get; set; }

    }
}


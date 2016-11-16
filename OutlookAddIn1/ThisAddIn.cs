using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Net;
using iTextSharp.text.pdf;
using iTextSharp.text;
using OpenQA.Selenium.Chrome;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        string connectionString;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

            SetConnectionString();

            //            string body = @"<html><head></head><body><div id='Header'><div><table border='0' cellpadding='0' cellspacing='0' width='100%'><tr><td width='100%' style='word-wrap:break-word'><table cellpadding='2' cellspacing='3' border='0' width='100%'><tr><td width='1%' nowrap='nowrap'><img src='http://q.ebaystatic.com/aw/pics/logos/ebay_95x39.gif' height='39' width='95' alt='eBay'></td><td align='left' valign='bottom'><span style='font-weight:bold; font-size:xx-small; font-family:verdana, sans-serif; color:#666'><b>eBay sent this message to Zhijun Ding (zjding2016).</b><br></span><span style='font-size:xx-small; font-family:verdana, sans-serif; color:#666'>Your registered name is included to show this message originated from eBay. <a href='http://pages.ebay.com/help/confidence/name-userid-emails.html'>Learn more</a>.</span></td></tr></table></td></tr></table></div></div><div id='Title'><div><table style='background-color:#ffe680' border='0' cellpadding='0' cellspacing='0' width='100%'><tr><td width='8' valign='top'><img src='http://q.ebaystatic.com/aw/pics/globalAssets/ltCurve.gif' height='8' width='8' alt=''></td><td valign='bottom' width='100%'><span style='font-weight:bold; font-size:14pt; font-family:arial, sans-serif; color:#000; margin:2px 0 2px 0'>Your item has been listed. Sell another item now!</span></td><td width='8' valign='top' align='right'><img src='http://p.ebaystatic.com/aw/pics/globalAssets/rtCurve.gif' height='8' width='8' alt=''></td></tr><tr><td style='background-color:#fc0' colspan='3' height='4'></td></tr></table></div></div><div id='SingleItemCTA'><div><table border='0' cellpadding='2' cellspacing='3' width='100%'><tr><td><font style='font-size:10pt; font-family:arial, sans-serif; color:#000'>Hi zjding2016,<table border='0' cellpadding='0' cellspacing='0' width='100%'><tr><td><img src='http://q.ebaystatic.com/aw/pics/s.gif' height='10' alt=' '></td></tr></table>Your item has been successfully listed on eBay. It may take some time for the item to appear on eBay search results. Here are the listing details:<table border='0' cellpadding='0' cellspacing='0' width='100%'><tr><td><img src='http://q.ebaystatic.com/aw/pics/s.gif' height='10' alt=' '></td></tr></table><div></div></font><div><table width='100%' cellpadding='0' cellspacing='3' border='0'><tr><td valign='top' align='center' width='100' nowrap='nowrap'><a href='http://rover.ebay.com/rover/0/e12000.m43.l1123/7?euid=db33b151a180449c92429caf42c24796&amp;loc=http%3A%2F%2Fcgi.ebay.com%2Fws%2FeBayISAPI.dll%3FViewItem%26item%3D152050500319%26ssPageName%3DADME%3AL%3ALCA%3AUS%3A1123'><img src='http://pics.ebaystatic.com/aw/pics/icon/iconPic_20x20.gif' alt='Cuisinart 14-Cup Programmable Coffeemaker' border='0'></a></td><td colspan='2' valign='top'><table width='100%' cellpadding='0' cellspacing='0' border='0'><tr><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' colspan='2'><a href='http://rover.ebay.com/rover/0/e12000.m43.l1123/7?euid=db33b151a180449c92429caf42c24796&amp;loc=http%3A%2F%2Fcgi.ebay.com%2Fws%2FeBayISAPI.dll%3FViewItem%26item%3D152050500319%26ssPageName%3DADME%3AL%3ALCA%3AUS%3A1123'>Cuisinart 14-Cup Programmable Coffeemaker</a></td></tr><tr><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' width='15%' nowrap='nowrap' valign='top'>Item Id:</td><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' valign='top'>152050500319</td></tr><tr><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' width='15%' nowrap='nowrap' valign='top'>Price:</td><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' valign='top'>$94.89</td></tr><tr><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' width='15%' nowrap='nowrap' valign='top'>End time:</td><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' valign='top'>May-11-16 10:49:05 PDT</td></tr><tr><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' width='15%' nowrap='nowrap' valign='top'>Listing fees:</td><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' valign='top'>0</td></tr><tr><td colspan='2'><font style='font-size:10pt; font-family:arial, sans-serif; color:#000'><a href='http://rover.ebay.com/rover/0/e12000.m43.l1125/7?euid=db33b151a180449c92429caf42c24796&amp;loc=http%3A%2F%2Fcgi5.ebay.com%2Fws2%2FeBayISAPI.dll%3FUserItemVerification%26%26item%3D152050500319%26ssPageName%3DADME%3AL%3ALCA%3AUS%3A1125'>Revise item</a>   |    <a href='http://rover.ebay.com/rover/0/e12000.m43.l1121/7?euid=db33b151a180449c92429caf42c24796&amp;loc=http%3A%2F%2Fmy.ebay.com%2Fws%2FeBayISAPI.dll%3FMyeBay%26%26CurrentPage%3DMyeBaySelling%26ssPageName%3DADME%3AL%3ALCA%3AUS%3A1121'>Go to My eBay</a></font></td></tr></table></td></tr></table></div><td valign='top' width='185'><div><span style='font-weight:bold; font-size:10pt; font-family:arial, sans-serif; color:#000'>Ready to List Your Next Item?</span><table border='0' cellpadding='0' cellspacing='0' width='100%'><tr><td><img src='http://q.ebaystatic.com/aw/pics/s.gif' height='4' alt=' '></td></tr></table><a href='http://rover.ebay.com/rover/0/e12000.m44.l1127/7?euid=db33b151a180449c92429caf42c24796&amp;loc=http%3A%2F%2Fcgi5.ebay.com%2Fws%2FeBayISAPI.dll%3FSellHub3%26ssPageName%3DADME%3AL%3ALCA%3AUS%3A1127' title='http://rover.ebay.com/rover/0/e12000.m44.l1127/7?euid=db33b151a180449c92429caf42c24796&amp;loc=http%3A%2F%2Fcgi5.ebay.com%2Fws%2FeBayISAPI.dll%3FSellHub3%26ssPageName%3DADME%3AL%3ALCA%3AUS%3A1127'><img src='http://p.ebaystatic.com/aw/pics/buttons/btnSellMore.gif' border='0' height='32' width='120'></img></a><br><span style='font-style:italic; font-size:8pt; font-family:arial, sans-serif; color:#000'>Click to list another item</span></div></td></td></tr></table><br></div></div><div id='OneClickUnsubscribe'><div><style>.cub-cwrp {display:block; border:1px solid #dedfde; font-family:arial, sans-serif; font-size:10pt; margin-bottom:20px}
            //h3.cub - chd {
            //            margin: 0px; padding: 5px; display: block; background:#e7e7e7; font-size:14px}
            //.cub - ccnt { padding: 0px 10px 10px 5px; display: block}
            //                ul.cub - ulst { margin: 0px 0px 0px 10px; padding: 0px 0px 0px 10px}
            //                ul.cub - ulst li, ul.cub - ulst li.cub - licn { list - style:square outside none; margin: 0px; padding: 10px 0px 0px 0px; line - height:16px}
            //.cub - ltxt {
            //                color:#333; display:block}
            //</ style >< div class='cub-cwrp'><h3 class='cub-chd'>Select your email preferences</h3><div class='cub-ccnt'><ul class='cub-ulst'><li><span class='cub-ltxt'><span>Want to reduce your inbox email volume? <a href = 'http://my.ebay.com/ws/eBayISAPI.dll?DigestEmail&amp;emailType=12000' > Receive this email as a daily digest</a>.</span><br><span>For other email digest options, go to<a href= 'http://my.ebay.com/ws/eBayISAPI.dll?MyEbayBeta&amp;CurrentPage=MyeBayNextNotificationPreferences' > Notification Preferences</a> in My eBay.</span><br></span></li><li><span class='cub-ltxt'><span>Don't want to receive this email? <a href='http://my.ebay.com/ws/eBayISAPI.dll?EmailUnsubscribe&amp;emailType=12000'>Unsubscribe from this email</a>.</span><br></span></li></ul></div></div></div></div><div id='Tips'></div><div id='RTMEducational'></div><div id='MST'><div><table style='border:1px solid #6b7b91' border='0' cellpadding='0' cellspacing='0' width='100%'><tr style='background-color:#c9d2dc' height='1'><td><img src='http://p.ebaystatic.com/aw/pics/securityCenter/imgShield_25x25.gif' height='25' width='25' alt='Marketplace Safety Tip' align='absmiddle'></td><td style='font-weight:bold; font-size:10pt; font-family:arial, sans-serif; color:#000' nowrap='nowrap' width='20%'>Marketplace Safety Tip</td><td><img src='http://p.ebaystatic.com/aw/pics/securityCenter/imgTabCorner_25x25.gif' height='25' width='25' alt='' align='absmiddle'></td><td background='http://q.ebaystatic.com/aw/pics/securityCenter/imgFlex_1x25.gif' height='1' width='80%'></td></tr><tr><td style='font-size:10pt; font-family:arial, sans-serif; color:#000' colspan='4'><ul style='margin-top: 5px; margin-bottom: 5px;'><li style='padding-bottom: 3px; line-height: 120%; padding-top: 3px; list-style-type: square;'>If you are contacted about buying a similar item outside of eBay, please do not respond. Outside-of-eBay transactions are against eBay policy, and they are not covered by eBay services such as feedback and eBay purchase protection programs.</li></ul></td></tr><tr><td style='background-color:#c9d2dc' colspan='4'><img src='http://q.ebaystatic.com/aw/pics/s.gif' height='1' width='1'></td></tr></table><br></div></div><div id='Footer'><div><hr style='HEIGHT: 1px'><table border='0' cellpadding='0' cellspacing='0' width='100%'><tr><td width='100%'><font style='font-size:8pt; font-family:arial, sans-serif; color:#000000'>Email reference id: [#db33b151a180449c92429caf42c24796#]</font></td></tr></table><br></div><hr style='HEIGHT: 1px'><table border='0' cellpadding='0' cellspacing='0' width='100%'><tr><td width='100%'><font style='font-size:xx-small; font-family:verdana; color:#666'><a href='http://pages.ebay.com/education/spooftutorial/index.html'>Learn More</a> to protect yourself from spoof (fake) emails.<br><br>eBay sent this email to you at zjding@outlook.com about your account registered on <a href='http://www.ebay.com'>www.ebay.com</a>.<br><br>eBay sends these emails based on the preferences you set for your account. To unsubscribe from this email, change your <a href='http://my.ebay.com/ws/eBayISAPI.dll?MyEbayBeta&amp;CurrentPage=MyeBayNextNotificationPreferences'>communication preferences</a>. Please note that it may take up to 10 days to process your request. Visit our <a href='http://pages.ebay.com/help/policies/privacy-policy.html'>Privacy Notice</a> and <a href='http://pages.ebay.com/help/policies/user-agreement.html'>User Agreement</a> if you have any questions.<br><br>Copyright © 2016 eBay Inc. All Rights Reserved. Designated trademarks and brands are the property of their respective owners. eBay and the eBay logo are trademarks of eBay Inc. eBay Inc. is located at 2145 Hamilton Avenue, San Jose, CA 95125.  </font></td></tr></table><img src='http://rover.ebay.com/roveropen/0/e12000/7?euid=db33b151a180449c92429caf42c24796&amp;' height='1' width='1'></div></body></html>";

            //            // ItemID
            //            int iItemId = body.IndexOf("Item Id:");
            //            int iStart = iItemId + 100;
            //            int iEnd = body.IndexOf("</td>", iStart);
            //            string id = body.Substring(iStart - 4, iEnd - iStart + 4);

            //            // ListPrice
            //            iItemId = body.IndexOf("Price:");
            //            iStart = iItemId + 95;
            //            iEnd = body.IndexOf("</td>", iStart);
            //            string price = body.Substring(iStart, iEnd - iStart);

            //            // EndTime
            //            iItemId = body.IndexOf("End time:");
            //            iStart = iItemId + 98;
            //            iEnd = body.IndexOf("</td>", iStart);
            //            string endTime = body.Substring(iStart-1, iEnd - iStart+1);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        public void SetConnectionString()
        {
            string azureConnectionString = "Server=tcp:zjding.database.windows.net,1433;Initial Catalog=Costco;Persist Security Info=False;User ID=zjding;Password=G4indigo;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";

            SqlConnection cn = new SqlConnection(azureConnectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();
            string sqlString = "SELECT ConnectionString FROM DatabaseToUse WHERE bUse = 1 and ApplicationName = 'Crawler'";
            cmd.CommandText = sqlString;
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    connectionString = (reader.GetString(0)).ToString();
                }
            }
            reader.Close();
            cn.Close();
        }

        void items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;

            try
            {
                if (Item != null)
                {
                    if (mail.TaskSubject.IndexOf("Your eBay listing is confirmed") == 0)
                    {
                        string subject = mail.Subject;

                        string body = mail.HTMLBody;

                        ProcessListingConfirmEmail(body, subject);
                    }
                    else if (mail.TaskSubject.IndexOf("Relist") == 0)
                    {
                        string subject = mail.Subject;

                        string productName = subject.Substring(7, subject.Length - 7);

                        SqlConnection cn = new SqlConnection(connectionString);
                        SqlCommand cmd = new SqlCommand();
                        cmd.Connection = cn;

                        Product product = new Product();

                        cn.Open();

                        string sqlString = "UPDATE eBay_CurrentListings SET DeleteDT = GETDATE() WHERE eBayListingName = '" + productName.Replace("'", "''") + "'";
                        cmd.CommandText = sqlString;
                        cmd.ExecuteNonQuery();

                        //sqlString = "UPDATE eBay_ToRemove SET DeleteTime = GETDATE() WHERE Name = '" + productName + "'";
                        //cmd.CommandText = sqlString;
                        //cmd.ExecuteNonQuery();

                        //sqlString = "DELETE eBayListingChange_Discontinue WHERE Name = '" + productName + "'";
                        //cmd.CommandText = sqlString;
                        //cmd.ExecuteNonQuery();

                        cn.Close();
                    }
                    else if (mail.TaskSubject.Contains("You received a payment from your buyer"))
                    {
                        //string body = mail.HTMLBody;

                        //File.WriteAllText(@"C:\temp\temp.html", body);

                        //body = body.Replace("\n", "");
                        //body = body.Replace("\t", "");
                        //body = body.Replace("\\", "");
                        //body = body.Replace("\"", "'");

                        //ProcessPaymentReceivedEmail(body);
                    }
                    else if (mail.TaskSubject.IndexOf("Your eBay item sold!") == 0)
                    {
                        //string body = mail.HTMLBody;

                        //File.WriteAllText(@"C:\temp\temp.html", body);

                        //body = body.Replace("\n", "");
                        //body = body.Replace("\r", "");
                        //body = body.Replace("\t", "");
                        //body = body.Replace("\\", "");
                        //body = body.Replace("\"", "'");

                        //ProcessItemSoldEmail(mail.TaskSubject, body);
                    }
                    else if (mail.TaskSubject.Contains("Your Costco.com Order Was Received"))
                    {
                        //string body = mail.HTMLBody;

                        //File.WriteAllText(@"C:\temp\temp.html", body);

                        //body = body.Replace("\n", "");
                        //body = body.Replace("\t", "");
                        //body = body.Replace("\\", "");
                        //body = body.Replace("\"", "'");

                        ////body = @"<html><head></head><body>" + body + @"</body></html>";

                        //ProcessCostcoOrderEmail(body);
                    }
                    else if (mail.TaskSubject.Contains("Your Costco.com order has been shipped"))
                    {
                        //string body = mail.HTMLBody;

                        //File.WriteAllText(@"C:\temp\temp.html", body);

                        //body = body.Replace("\n", "");
                        //body = body.Replace("\t", "");
                        //body = body.Replace("\\", "");
                        //body = body.Replace("\"", "'");

                        //ProcessCostcoShipEmail(body);
                    }

                    mail.UnRead = false;
                }
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
        }

        private void FillUpTaxPDF(eBaySoldProduct product)
        {
            string pdfTemplateFileName = @"c:\ebay\documents\" + "TaxExemption_Form.pdf";
            string newFileName = @"c:\temp\TaxExemption\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "-" + product.CostcoOrderNumber + ".pdf";
            PdfReader pdfReader = new PdfReader(pdfTemplateFileName);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(newFileName, FileMode.Create));
            AcroFields pdfFormFields = pdfStamper.AcroFields;

            pdfFormFields.SetField("txtLegalBusinessName", "eBay Business");
            pdfFormFields.SetField("txtDoingBusinessAs", "eBay Business");
            pdfFormFields.SetField("txtBusinessAddress", "1642 Crossgate Dr., Vestavia");
            pdfFormFields.SetField("txtCostcoMembership", "111775568587");
            pdfFormFields.SetField("txtBusinessPhone", "205-588-6960");
            pdfFormFields.SetField("txtSaleTaxRegistration", "11111111");
            pdfFormFields.SetField("txtStateRegistered", "AL");
            pdfFormFields.SetField("txtTotalRefundRequested", Convert.ToString(product.CostcoTax));
            pdfFormFields.SetField("txtPreciseNatureOfBusiness", "eCommerse");
            pdfFormFields.SetField("txtCategoriesOfItems", "Health and Beauty");
            pdfFormFields.SetField("txtDate", product.eBaySoldDateTime.ToShortDateString());

            pdfStamper.FormFlattening = true;
            pdfStamper.Close();
            pdfReader.Close();


            //List<string> fileNames = new List<string>();
            //fileNames.Add(@"c:\ebay\TaxExemption\TaxExemptionTotal.pdf");
            //fileNames.Add(newFileName);

            //string tempFileName = @"c:\ebay\TaxExemption\TaxExemptionTemp.pdf";

            //MergePDFs(fileNames, tempFileName);
        }

        public bool MergePDFs(IEnumerable<string> fileNames, string targetPdf)
        {
            bool merged = true;
            using (FileStream stream = new FileStream(targetPdf, FileMode.Create))
            {
                Document document = new Document();
                PdfCopy pdf = new PdfCopy(document, stream);
                PdfReader reader = null;
                try
                {
                    document.Open();
                    foreach (string file in fileNames)
                    {
                        reader = new PdfReader(file);
                        pdf.AddDocument(reader);
                        reader.Close();
                    }
                }
                catch (Exception)
                {
                    merged = false;
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
                finally
                {
                    if (document != null)
                    {
                        document.Close();
                    }
                }
            }
            return merged;
        }

        private void ProcessListingConfirmEmail(string body, string subject)
        {
            body = body.Replace("\n", "");
            body = body.Replace("\t", "");
            body = body.Replace("\\", "");
            body = body.Replace("\"", "'");

            //subject = subject.Replace("'", "''");

            string productName = subject.Substring(32, subject.Length - 32);

            string sqlString = "SELECT * FROM eBay_ToAdd WHERE eBayName = '" + productName.Replace("'", "''") + "' AND DeleteTime is NULL";

            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            Product product = new Product();

            cn.Open();
            cmd.CommandText = sqlString;
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                reader.Read();

                product.Name = Convert.ToString(reader["Name"]);
                product.eBayName = Convert.ToString(reader["eBayName"]);
                product.UrlNumber = Convert.ToString(reader["UrlNumber"]);
                product.ItemNumber = Convert.ToString(reader["ItemNumber"]);
                product.Category = Convert.ToString(reader["Category"]);
                product.Price = Convert.ToDecimal(reader["Price"]);
                product.Shipping = Convert.ToDecimal(reader["Shipping"]);
                product.Limit = Convert.ToString(reader["Limit"]);
                product.Details = Convert.ToString(reader["Details"]);
                product.Specification = Convert.ToString(reader["Specification"]);
                product.ImageLink = Convert.ToString(reader["ImageLink"]);
                product.NumberOfImage = Convert.ToInt16(reader["NumberOfImage"]);
                product.Url = Convert.ToString(reader["Url"]);
                product.ImageOptions = Convert.ToString(reader["ImageOptions"]);
                product.Options = Convert.ToString(reader["Options"]);
                product.Thumb = Convert.ToString(reader["Thumb"]);

                product.eBayCategoryID = Convert.ToString(reader["eBayCategoryID"]);
                product.eBayReferencePrice = Convert.ToDecimal(reader["eBayReferencePrice"]);
                product.eBayListingPrice = Convert.ToDecimal(reader["eBayListingPrice"]);
                product.DescriptionImageWidth = Convert.ToInt16(reader["DescriptionImageWidth"]);
                product.DescriptionImageHeight = Convert.ToInt16(reader["DescriptionImageHeight"]);
                product.eBayReferenceUrl = Convert.ToString(reader["eBayReferenceUrl"]);
            }

            reader.Close();

            //sqlString = @"DELETE FROM eBay_ToAdd WHERE name = '" + productName + "'";
            sqlString = @"UPDATE eBay_ToAdd SET DeleteTime = GETDATE() WHERE eBayName = '" + productName.Replace("'", "''") + "'";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            body = body.Replace("\r", "");
            body = body.Replace("\t", "");
            body = body.Replace("\n", "");

            string stItemId = SubstringInBetween(body, "Item Id:</td>", "</td>", false, true);
            stItemId = SubstringEndBack(stItemId, "</td>", ">", false, false);
            stItemId = stItemId.Trim();

            //string stListingUrl = SubstringEndBack(body, "Item Id:</td>", "<a href='", true, false);
            //stListingUrl = SubstringInBetween(stListingUrl, "<a href='", "target", false, false);
            //stListingUrl = stListingUrl.Trim();

            string stPrice = SubstringInBetween(body, "Price:</td>", "</td>", false, true);
            stPrice = SubstringEndBack(stPrice, "</td>", "$", false, false);
            stPrice = stPrice.Replace("$", "");
            stPrice = stPrice.Trim();

            string stEndTime = SubstringInBetween(body, "End time:</td>", "</td>", false, false);
            stEndTime = TrimTags(stEndTime);
            string stTimeZone = stEndTime.Substring(stEndTime.LastIndexOf(' ') + 1, stEndTime.Length - stEndTime.LastIndexOf(' ') - 1);
            DateTime dtEndTime = Convert.ToDateTime(stEndTime.Replace(stTimeZone, timeZones[stTimeZone]));
            //stEndTime = SubstringEndBack(stEndTime, "PDT", ">", false, true);
            //stEndTime = stEndTime.Trim();
            //string correctedTZ = stEndTime.Replace("PDT", "-0700");
            //DateTime dtEndTime = Convert.ToDateTime(correctedTZ);

            sqlString = @"INSERT INTO eBay_CurrentListings
                            (Name, eBayListingName, eBayCategoryID, eBayItemNumber, eBayListingPrice, eBayDescription, 
                             eBayEndTime, CostcoUrlNumber, CostcoItemNumber, CostcoUrl, CostcoPrice, ImageLink, ImageOptions, CostcoOptions, Thumb, eBayReferenceUrl, eBayReferencePrice) 
                          VALUES (@_name, @_eBayListingName, @_eBayCategoryID, @_eBayItemNumber, @_eBayListingPrice, @_eBayDescription,
                                @_eBayEndTime, @_CostcoUrlNumber, @_CostcoItemNumber, @_CostcoUrl, @_CostcoPrice, @_ImageLink, @_ImageOptions, @_Options, @_Thumb, @_eBayReferenceUrl, @_eBayReferencePrice)";

            cmd.CommandText = sqlString;
            cmd.Parameters.AddWithValue("@_name", product.Name);
            cmd.Parameters.AddWithValue("@_eBayListingName", product.eBayName);
            cmd.Parameters.AddWithValue("@_eBayCategoryID", product.eBayCategoryID);
            cmd.Parameters.AddWithValue("@_eBayItemNumber", stItemId);
            cmd.Parameters.AddWithValue("@_eBayListingPrice", product.eBayListingPrice);
            cmd.Parameters.AddWithValue("@_eBayDescription", product.Details);
            cmd.Parameters.AddWithValue("@_eBayEndTime", dtEndTime);
            //cmd.Parameters.AddWithValue("@_eBayUrl", stListingUrl);
            cmd.Parameters.AddWithValue("@_CostcoUrlNumber", product.UrlNumber);
            cmd.Parameters.AddWithValue("@_CostcoItemNumber", product.ItemNumber);
            cmd.Parameters.AddWithValue("@_CostcoUrl", product.Url);
            cmd.Parameters.AddWithValue("@_CostcoPrice", product.Price);
            cmd.Parameters.AddWithValue("@_ImageLink", product.ImageLink);
            cmd.Parameters.AddWithValue("@_ImageOptions", product.ImageOptions);
            cmd.Parameters.AddWithValue("@_Options", product.Options);
            cmd.Parameters.AddWithValue("@_Thumb", product.Thumb);
            cmd.Parameters.AddWithValue("@_eBayReferenceUrl", product.eBayReferenceUrl);
            cmd.Parameters.AddWithValue("@_eBayReferencePrice", product.eBayReferencePrice);

            cmd.ExecuteNonQuery();

            cn.Close();
        }

        private void ProcessCostcoOrderEmail(string body)
        {
            try
            {
                body = body.Replace("\r", "");
                body = body.Replace("\t", "");
                body = body.Replace("\n", "");
                string stOrderNumber = SubstringInBetween(body, "Order Number:</td>", "</td>", false, true);
                stOrderNumber = SubstringEndBack(stOrderNumber, "</td>", ">", false, false);
                stOrderNumber = stOrderNumber.Trim();

                string stDatePlaced = SubstringInBetween(body, "Date Placed:</td>", "</td>", false, true);
                stDatePlaced = SubstringEndBack(stDatePlaced, "</td>", ">", false, false);
                stDatePlaced = stDatePlaced.Trim();

                string stWorking = SubstringInBetween(body, "Item Total", "Shipping &amp; Terms", false, false);
                stWorking = TrimTags(stWorking);

                string stQuatity = stWorking.Substring(0, stWorking.IndexOf("<"));
                stQuatity = stQuatity.Trim();

                stWorking = stWorking.Substring(stQuatity.Length);
                stWorking = TrimTags(stWorking);
                string stProductName = stWorking.Substring(0, stWorking.IndexOf("<"));
                stProductName = stProductName.Trim();

                string stItemNum = stProductName.Substring(stProductName.IndexOf("Item#"));

                stItemNum = stItemNum.Replace("Item#", "");
                stItemNum = stItemNum.Trim();

                string stShipping = SubstringInBetween(body, "Shipping Address", "Note:", false, false);


                stWorking = TrimTags(stShipping);

                string stBuyerName = stWorking.Substring(0, stWorking.IndexOf("<"));

                stWorking = stWorking.Replace(stBuyerName, "");

                stWorking = TrimTags(stWorking);

                string stAddress1 = stWorking.Substring(0, stWorking.IndexOf("<"));

                stWorking = stWorking.Replace(stAddress1, "");

                stWorking = TrimTags(stWorking);

                string stAddress2 = stWorking.Substring(0, stWorking.IndexOf("<"));

                string stTax = SubstringInBetween(body, "Tax:", "</tr>", false, false);

                stTax = TrimTags(stTax);

                stTax = stTax.Substring(0, stTax.IndexOf("<"));

                stTax = stTax.Replace("$", "");

                string stTotal = SubstringInBetween(body, "Order Total:", "</tr>", false, false);
                stTotal = TrimTags(stTotal);
                stTotal = stTotal.Substring(0, stTotal.IndexOf("<"));
                stTotal = stTotal.Replace("$", "");

                // Generate PDF for email
                //File.WriteAllText(@"C:\temp\temp.html", body);

                FirefoxProfile profile = new FirefoxProfile();
                profile.SetPreference("print.always_print_silent", true);

                IWebDriver driver = new FirefoxDriver(profile);

                driver.Navigate().GoToUrl(@"file:///C:/temp/temp.html");

                IJavaScriptExecutor js = driver as IJavaScriptExecutor;

                js.ExecuteScript("window.print();");

                driver.Dispose();

                System.Threading.Thread.Sleep(3000);

                // Process files
                string[] files = Directory.GetFiles(@"C:\temp\tempPDF\");

                string sourceFileFullName = files[0];

                string sourceFileName = sourceFileFullName.Replace(@"C:\temp\tempPDF\", "");

                string destinationFileName = Convert.ToDateTime(stDatePlaced).ToString("yyyyMMddHHmmss") + "_" + stOrderNumber + ".pdf";

                File.Delete(@"C:\temp\CostcoOrderEmails\" + destinationFileName);

                File.Move(sourceFileFullName, @"C:\temp\CostcoOrderEmails\" + destinationFileName);

                sourceFileFullName = destinationFileName;
                destinationFileName = Convert.ToDateTime(stDatePlaced).ToString("yyyyMMdd") + "-" + stTotal + "-" + "Costco" + stOrderNumber + ".pdf";
                File.Delete(@"C:\Users\Jason Ding\Dropbox\Bookkeeping\Cost\" + destinationFileName);
                File.Copy(@"C:\temp\CostcoOrderEmails\" + sourceFileFullName, @"C:\Users\Jason Ding\Dropbox\Bookkeeping\Cost\" + destinationFileName);

                // db stuff
                string sqlString;
                bool bExist = false;

                SqlConnection cn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cn.Open();

                if (stItemNum != "")
                {
                    sqlString = @"SELECT * FROM eBay_SoldTransactions WHERE CostcoItemNumber = @_costcoItemNumber 
                                AND BuyerName = @_buyerName AND CostcoOrderNumber IS NULL";

                    cmd.CommandText = sqlString;
                    cmd.Parameters.AddWithValue("@_costcoItemNumber", stItemNum);
                    cmd.Parameters.AddWithValue("@_buyerName", stBuyerName);

                    SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        bExist = true;
                    }
                    reader.Close();

                    if (bExist)
                    {
                        sqlString = @"UPDATE eBay_SoldTransactions SET CostcoOrderNumber = @_costcoOrderNumber, 
                                CostcoOrderDate = @_costcoOrderDate, 
                                CostcoOrderEmailPdf = @_costcoOrderEmailPdf, CostcoTax = @_costcoTax 
                                WHERE CostcoItemNumber = @_costcoItemNumber 
                                AND BuyerName = @_buyerName AND  CostcoOrderNumber IS NULL";

                        cmd.CommandText = sqlString;
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@_costcoOrderNumber", stOrderNumber);
                        cmd.Parameters.AddWithValue("@_costcoOrderDate", stDatePlaced);
                        cmd.Parameters.AddWithValue("@_costcoOrderEmailPdf", destinationFileName);
                        cmd.Parameters.AddWithValue("@_costcoTax", stTax);
                        cmd.Parameters.AddWithValue("@_costcoItemNumber", stItemNum);
                        cmd.Parameters.AddWithValue("@_buyerName", stBuyerName);

                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    sqlString = @"SELECT * FROM eBay_SoldTransactions WHERE CostcoItemName = @_costcoItemName
                                AND BuyerName = @_buyerName AND CostcoOrderNumber IS NULL";

                    cmd.CommandText = sqlString;
                    cmd.Parameters.AddWithValue("@_costcoItemName", stProductName);
                    cmd.Parameters.AddWithValue("@_buyerName", stBuyerName);

                    SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        bExist = true;
                    }
                    reader.Close();

                    if (bExist)
                    {
                        sqlString = @"UPDATE eBay_SoldTransactions SET CostcoOrderNumber = @_costcoOrderNumber,
                                CostcoOrderDate = @_costcoOrderDate, 
                                CostcoOrderEmailPdf = @_costcoOrderEmailPdf, CostcoTax = @_costcoTax 
                                WHERE CostcoItemName = @_costcoItemName 
                                AND BuyerName = @_buyerName AND  CostcoOrderNumber IS NULL";

                        cmd.CommandText = sqlString;
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@_costcoOrderNumber", stOrderNumber);
                        cmd.Parameters.AddWithValue("@_costcoOrderDate", stDatePlaced);
                        cmd.Parameters.AddWithValue("@_costcoOrderEmailPdf", destinationFileName);
                        cmd.Parameters.AddWithValue("@_costcoTax", stTax);
                        cmd.Parameters.AddWithValue("@_costcoItemName", stProductName);
                        cmd.Parameters.AddWithValue("@_buyerName", stBuyerName);

                        cmd.ExecuteNonQuery();
                    }
                }

                cn.Close();

                eBaySoldProduct product = new eBaySoldProduct();
                product.eBaySoldDateTime = Convert.ToDateTime(stDatePlaced);
                product.CostcoTax = Convert.ToDecimal(stTax);
                product.CostcoOrderNumber = stOrderNumber;

                FillUpTaxPDF(product);
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
        }

        private void ProcessItemSoldEmail(string subject, string body)
        {
            try
            {
                int iLastSpace = subject.LastIndexOf(' ');
                string stItemNum = subject.Substring(iLastSpace, subject.Length - iLastSpace);
                stItemNum = stItemNum.Trim();

                //string stItemNum = SubstringInBetween(subject, "(", ")", false, false);

                subject = subject.Replace(stItemNum, "");
                subject = subject.Replace("Your eBay item sold!", "");

                string stItemName = subject.Trim();

                stItemName = WebUtility.HtmlDecode(stItemName);

                body = WebUtility.HtmlDecode(body);

                body = SubstringInBetween(body, @"<body", @"</body>", true, true);

                //string stUrl = SubstringEndBack(body, ">" + stItemName, "<a href =", false, false);
                //stUrl = SubstringInBetween(stUrl, "'", "'", false, false);

                string stPaid = SubstringInBetween(body, "Paid:", @"<br>", false, false);
                stPaid = stPaid.Replace("$", "");
                stPaid = stPaid.Trim();

                string stColor = string.Empty;
                if (body.Contains(@"Color:"))
                {
                    stColor = SubstringInBetween(body, @"Color:", @"</td>", false, false);
                    stColor = stColor.Trim();
                }

                string stSize = string.Empty;
                if (body.Contains(@"Size:"))
                {
                    stSize = SubstringInBetween(body, @"Size:", @"</td>", false, false);
                    stSize = stSize.Trim();
                }

                string stVariation = string.Empty;

                if (string.IsNullOrEmpty(stColor))
                {
                    if (string.IsNullOrEmpty(stSize))
                    {

                    }
                    else
                    {
                        stVariation = stSize;
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(stSize))
                    {
                        stVariation = stColor;
                    }
                    else
                    {
                        stVariation = stColor + ";" + stSize;
                    }
                }

                string stDateSold = SubstringInBetween(body, @"Date Sold:", @"</td>", false, false);
                stDateSold = stDateSold.Trim();

                string stQuantitySold = SubstringInBetween(body, @"Quantity Sold:", @"</td>", false, false);
                stQuantitySold = stQuantitySold.Trim();

                string stBuyer = SubstringEndBack(body, "Contact Buyer", "Buyer:", false, false);
                stBuyer = stBuyer.Replace(@"&nbsp", "");
                stBuyer = TrimTags(stBuyer);
                stBuyer = stBuyer.Substring(0, stBuyer.IndexOf("<"));

                FirefoxProfile profile = new FirefoxProfile();
                profile.SetPreference("print.always_print_silent", true);

                IWebDriver driver = new FirefoxDriver(profile);

                driver.Navigate().GoToUrl(@"file:///C:/temp/temp.html");

                IJavaScriptExecutor js = driver as IJavaScriptExecutor;

                js.ExecuteScript("window.print();");

                driver.Dispose();

                System.Threading.Thread.Sleep(3000);

                // Process files
                string[] files = Directory.GetFiles(@"C:\temp\tempPDF\");

                string sourceFileFullName = files[0];

                string sourceFileName = sourceFileFullName.Replace(@"C:\temp\tempPDF\", "");

                string destinationFileName = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + stItemNum + ".pdf";

                File.Delete(@"C:\temp\eBaySoldEmails\" + destinationFileName);

                File.Move(sourceFileFullName, @"C:\eBayApp\Files\Emails\eBaySoldEmails\" + destinationFileName);


                // db stuff
                SqlConnection cn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;

                string sqlString = @"SELECT * FROM eBay_CurrentListings WHERE eBayItemNumber = " + stItemNum;

                eBayListingProduct eBayProduct = new eBayListingProduct();

                cn.Open();
                cmd.CommandText = sqlString;
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();

                    eBayProduct.Name = Convert.ToString(reader["Name"]);
                    eBayProduct.eBayListingName = Convert.ToString(reader["eBayListingName"]);
                    eBayProduct.eBayCategoryID = Convert.ToString(reader["eBayCategoryID"]);
                    eBayProduct.eBayItemNumber = Convert.ToString(reader["eBayItemNumber"]);
                    eBayProduct.eBayListingPrice = Convert.ToDecimal(reader["eBayListingPrice"]);
                    eBayProduct.eBayDescription = Convert.ToString(reader["eBayDescription"]);
                    // eBayProduct.eBayListingDT = Convert.ToDateTime(reader["eBayListingDT"]);
                    eBayProduct.eBayUrl = Convert.ToString(reader["eBayUrl"]);
                    eBayProduct.CostcoUrlNumber = Convert.ToString(reader["CostcoUrlNumber"]);
                    eBayProduct.CostcoItemNumber = Convert.ToString(reader["CostcoItemNumber"]);
                    eBayProduct.CostcoUrl = Convert.ToString(reader["CostcoUrl"]);
                    eBayProduct.CostcoPrice = Convert.ToDecimal(reader["CostcoPrice"]);
                    eBayProduct.ImageLink = Convert.ToString(reader["ImageLink"]);
                }
                reader.Close();

                // check exist

                bool bExist = false;

                sqlString = "SELECT * FROM eBay_SoldTransactions WHERE eBayItemNumber = " + stItemNum;
                cmd.CommandText = sqlString;
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                    bExist = true;
                reader.Close();

                if (!bExist)
                {
                    sqlString = @"INSERT INTO eBay_SoldTransactions 
                              (eBayItemNumber, eBaySoldDateTime, eBayItemName, eBaySoldVariation, eBaySoldPrice, eBaySoldQuality, eBaySoldEmailPdf,
                               BuyerID, CostcoUrlNumber, CostcoItemNumber, CostcoUrl, CostcoPrice)
                              VALUES (@_eBayItemNumber, @_eBaySoldDateTime, @_eBayItemName, @_eBaySoldVariation, @_eBayPrice, @_eBaySoldQuality,  @_eBaySoldEmailPdf,
                               @_BuyerID, @_CostcoUrlNumber, @_CostcoItemNumber, @_CostcoUrl, @_CostcoPrice)";

                    cmd.CommandText = sqlString;

                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@_eBayItemNumber", stItemNum);
                    cmd.Parameters.AddWithValue("@_eBaySoldDateTime", Convert.ToDateTime(stDateSold));
                    cmd.Parameters.AddWithValue("@_eBayItemName", stItemName);
                    cmd.Parameters.AddWithValue("@_eBaySoldVariation", stVariation);
                    cmd.Parameters.AddWithValue("@_eBayPrice", Convert.ToDecimal(eBayProduct.eBayListingPrice));
                    cmd.Parameters.AddWithValue("@_eBaySoldQuality", Convert.ToInt16(stQuantitySold));
                    cmd.Parameters.AddWithValue("@_eBaySoldEmailPdf", destinationFileName);
                    cmd.Parameters.AddWithValue("@_BuyerID", stBuyer);
                    cmd.Parameters.AddWithValue("@_CostcoUrlNumber", eBayProduct.CostcoUrlNumber);
                    cmd.Parameters.AddWithValue("@_CostcoItemNumber", eBayProduct.CostcoItemNumber);
                    cmd.Parameters.AddWithValue("@_CostcoUrl", eBayProduct.CostcoUrl);
                    cmd.Parameters.AddWithValue("@_CostcoPrice", eBayProduct.CostcoPrice);

                    cmd.ExecuteNonQuery();
                }
                else
                {

                }

                cn.Close();
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
        }

        private bool hasElement(IWebElement webElement, By by)
        {
            try
            {
                webElement.FindElement(by);
                return true;
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        private bool hasElement(IWebDriver webDriver, By by)
        {
            try
            {
                webDriver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        private bool hasToOrderItem()
        {
            int result = 0;

            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();

            string sqlString = @"SELECT count(*) FROM eBay_SoldTransactions WHERE CostcoOrderStatus = 'To order";
            cmd.CommandText = sqlString;
            object value = cmd.ExecuteScalar();

            if (value != null)
                result = Convert.ToInt16(value);

            cn.Close();

            return result > 0;
        }

        private void DoOrderCostcoProduct(eBaySoldProduct soldProduct)
        {
            IWebDriver driver = new ChromeDriver();

            try
            {
                driver.Navigate().GoToUrl("https://www.costco.com/LogonForm");
                driver.FindElement(By.Id("logonId")).SendKeys("zjding@gmail.com");
                driver.FindElement(By.Id("logonPassword")).SendKeys("721123");
                driver.FindElements(By.ClassName("submit"))[0].Click();

                driver.Navigate().GoToUrl("http://www.costco.com/");
                driver.FindElement(By.Id("cart-d")).Click();

                while (driver.FindElements(By.LinkText("Remove from cart")).Count > 0)
                {
                    driver.FindElements(By.LinkText("Remove from cart"))[0].Click();
                    System.Threading.Thread.Sleep(3000);
                }

                driver.Navigate().GoToUrl(soldProduct.CostcoUrl);

                IWebElement eProductDetails = driver.FindElement(By.Id("product-details"));
                if (hasElement(eProductDetails, By.Id("variants")))
                {
                    var eVariants = eProductDetails.FindElement(By.Id("variants"));

                    var productOptions = eVariants.FindElements(By.ClassName("swatchDropdown"));

                    List<string> selectList = new List<string>();

                    foreach (var productOption in productOptions)
                    {
                        selectList.Add(productOption.FindElement(By.TagName("select")).GetAttribute("id").ToString());
                    }

                    if (selectList.Count == 1)
                    {
                        IWebElement selectElement0 = eProductDetails.FindElement(By.Id(selectList[0]));
                        var options0 = selectElement0.FindElements(By.TagName("option"));
                        foreach (IWebElement option0 in options0)
                        {
                            if (options0.IndexOf(option0) > 0)
                            {
                                if (option0.Text.Contains("$"))
                                {
                                    int index = option0.Text.LastIndexOf("- $");
                                    if (option0.Text.Substring(0, index - 1).Trim() == soldProduct.eBaySoldVariation)
                                    {
                                        option0.Click();
                                        break;
                                    }
                                }
                                else
                                {
                                    if (option0.Text.Trim() == soldProduct.eBaySoldVariation)
                                    {
                                        option0.Click();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                driver.FindElement(By.Id("minQtyText")).Clear();
                driver.FindElement(By.Id("minQtyText")).SendKeys("1");
                driver.FindElement(By.Name("add-to-cart")).Click();

                //if (isAlertPresents(ref driver))
                //    driver.SwitchTo().Alert().Accept();

                driver.Navigate().GoToUrl("https://www.costco.com/CheckoutCartView");

                //driver.FindElement(By.Id("cart-d")).Click();

                if (isAlertPresents(ref driver))
                    driver.SwitchTo().Alert().Accept();

                string buyerFirstName = soldProduct.BuyerName.Split(' ')[0];
                string buyerMiddleInitial = soldProduct.BuyerName.Split(' ').Count() == 3 ? soldProduct.BuyerName.Split(' ')[1] : "";
                string buyerLastName = soldProduct.BuyerName.Split(' ')[soldProduct.BuyerName.Split(' ').Count() - 1];


                driver.FindElement(By.Id("shopCartCheckoutSubmitButton")).Click();

                driver.FindElement(By.Id("addressFormInlineFirstName")).SendKeys(buyerFirstName);
                driver.FindElement(By.Id("addressFormInlineMiddleInitial")).SendKeys(buyerMiddleInitial);
                driver.FindElement(By.Id("addressFormInlineLastName")).SendKeys(buyerLastName);
                driver.FindElement(By.Id("addressFormInlineAddressLine1")).SendKeys(soldProduct.BuyerAddress1);
                driver.FindElement(By.Id("addressFormInlineCity")).SendKeys(soldProduct.BuyerCity);

                string state = GetState(soldProduct.BuyerState);

                driver.FindElement(By.XPath("//select[@id='" + "addressFormInlineState" + "']/option[contains(.,'" + state + "')]")).Click();
                driver.FindElement(By.Id("addressFormInlineZip")).SendKeys(soldProduct.BuyerZip);
                driver.FindElement(By.Id("addressFormInlinePhoneNumber")).SendKeys("2056175063");
                driver.FindElement(By.Id("addressFormInlineAddressNickName")).SendKeys(DateTime.Now.ToString());

                if (driver.FindElement(By.Id("saveAddressCheckboxInline")).Selected)
                {
                    driver.FindElement(By.Id("saveAddressCheckboxInline")).Click();
                }

                driver.FindElement(By.Id("addressFormInlineButton")).Click();

                System.Threading.Thread.Sleep(3000);

                if (driver.FindElements(By.XPath("//span[contains(text(), 'Continue')]")).Count > 0)
                {
                    driver.FindElement(By.XPath("//span[contains(text(), 'Continue')]")).Click();
                }

                System.Threading.Thread.Sleep(3000);

                if (hasElement(driver, By.Id("deliverySubmitButton")))
                    driver.FindElement(By.Id("deliverySubmitButton")).Click();

                driver.FindElement(By.Id("cc_cvc")).SendKeys("819");

                driver.FindElement(By.Id("paymentSubButtonBot")).Click();

                //if (isAlertPresents(ref driver))
                //    driver.SwitchTo().Alert().Accept();

                //driver.FindElement(By.Id("orderButton")).Click();
            }
            catch (Exception e)
            {

            }
            finally
            {
                driver.Dispose();
            }
        }

        private void OrderCostcoProduct(eBaySoldProduct soldProduct)
        {
            IWebDriver driver = new ChromeDriver();
            //IWebDriver driver = new FirefoxDriver();

            try
            {
                driver.Navigate().GoToUrl("https://www.costco.com/LogonForm");
                driver.FindElement(By.Id("logonId")).SendKeys("zjding@gmail.com");
                driver.FindElement(By.Id("logonPassword")).SendKeys("721123");
                driver.FindElements(By.ClassName("submit"))[0].Click();

                driver.Navigate().GoToUrl("http://www.costco.com/");
                driver.FindElement(By.Id("cart-d")).Click();

                while (driver.FindElements(By.LinkText("Remove from cart")).Count > 0)
                {
                    driver.FindElements(By.LinkText("Remove from cart"))[0].Click();
                    System.Threading.Thread.Sleep(3000);
                }

                driver.Navigate().GoToUrl(soldProduct.CostcoUrl);

                IWebElement eProductDetails = driver.FindElement(By.Id("product-details"));

                string select0 = string.Empty;
                string select1 = string.Empty;

                if (soldProduct.eBaySoldVariation.Length > 0)
                {
                    if (soldProduct.eBaySoldVariation.Contains(";"))
                    {
                        select0 = soldProduct.eBaySoldVariation.Split(';')[0];
                        select1 = soldProduct.eBaySoldVariation.Split(';')[1];
                    }
                    else
                    {
                        select0 = soldProduct.eBaySoldVariation;
                    }
                }

                if (hasElement(eProductDetails, By.Id("variants")))
                {
                    var eVariants = eProductDetails.FindElement(By.Id("variants"));

                    var productOptions = eVariants.FindElements(By.ClassName("swatchDropdown"));

                    List<string> selectList = new List<string>();

                    foreach (var productOption in productOptions)
                    {
                        selectList.Add(productOption.FindElement(By.TagName("select")).GetAttribute("id").ToString());
                    }

                    if (selectList.Count == 1)
                    {
                        IWebElement selectElement0 = eProductDetails.FindElement(By.Id(selectList[0]));
                        var options0 = selectElement0.FindElements(By.TagName("option"));
                        foreach (IWebElement option0 in options0)
                        {
                            if (options0.IndexOf(option0) > 0)
                            {
                                if (option0.Text.Contains("$"))
                                {
                                    int index = option0.Text.LastIndexOf("- $");
                                    if (option0.Text.Substring(0, index - 1).Trim() == soldProduct.eBaySoldVariation)
                                    {
                                        option0.Click();
                                        break;
                                    }
                                }
                                else
                                {
                                    if (option0.Text.Trim() == soldProduct.eBaySoldVariation)
                                    {
                                        option0.Click();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else if (selectList.Count == 2)
                    {
                        IWebElement selectElement0 = eProductDetails.FindElement(By.Id(selectList[0]));
                        var options0 = selectElement0.FindElements(By.TagName("option"));
                        foreach (IWebElement option0 in options0)
                        {
                            if (options0.IndexOf(option0) > 0)
                            {
                                if (option0.Text.Trim() == select0)
                                {
                                    option0.Click();
                                    break;
                                }
                            }
                        }

                        IWebElement selectElement1 = eProductDetails.FindElement(By.Id(selectList[1]));
                        var options1 = selectElement1.FindElements(By.TagName("option"));
                        foreach (IWebElement option1 in options1)
                        {
                            if (options1.IndexOf(option1) > 0)
                            {
                                if (option1.Text.Contains("$"))
                                {
                                    int index = option1.Text.LastIndexOf("- $");
                                    if (option1.Text.Substring(0, index - 1).Trim() == select1)
                                    {
                                        option1.Click();
                                        break;
                                    }
                                }
                                else
                                {
                                    if (option1.Text.Trim() == select1)
                                    {
                                        option1.Click();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                driver.FindElement(By.Id("minQtyText")).Clear();
                driver.FindElement(By.Id("minQtyText")).SendKeys("1");
                driver.FindElement(By.Name("add-to-cart")).Click();

                if (isAlertPresents(ref driver))
                    driver.SwitchTo().Alert().Accept();

                System.Threading.Thread.Sleep(3000);

                driver.Navigate().GoToUrl("https://www.costco.com/CheckoutCartView");

                //driver.FindElement(By.Id("cart-d")).Click();

                if (isAlertPresents(ref driver))
                    driver.SwitchTo().Alert().Accept();

                string buyerFirstName = soldProduct.BuyerName.Split(' ')[0];
                string buyerMiddleInitial = soldProduct.BuyerName.Split(' ').Count() == 3 ? soldProduct.BuyerName.Split(' ')[1] : "";
                string buyerLastName = soldProduct.BuyerName.Split(' ')[soldProduct.BuyerName.Split(' ').Count() - 1];


                driver.FindElement(By.Id("shopCartCheckoutSubmitButton")).Click();

                driver.FindElement(By.Id("addressFormInlineFirstName")).SendKeys(buyerFirstName);
                driver.FindElement(By.Id("addressFormInlineMiddleInitial")).SendKeys(buyerMiddleInitial);
                driver.FindElement(By.Id("addressFormInlineLastName")).SendKeys(buyerLastName);
                driver.FindElement(By.Id("addressFormInlineAddressLine1")).SendKeys(soldProduct.BuyerAddress1);
                driver.FindElement(By.Id("addressFormInlineCity")).SendKeys(soldProduct.BuyerCity);

                string state = GetState(soldProduct.BuyerState);

                driver.FindElement(By.XPath("//select[@id='" + "addressFormInlineState" + "']/option[contains(.,'" + state + "')]")).Click();
                driver.FindElement(By.Id("addressFormInlineZip")).SendKeys(soldProduct.BuyerZip);
                driver.FindElement(By.Id("addressFormInlinePhoneNumber")).SendKeys("2056175063");
                driver.FindElement(By.Id("addressFormInlineAddressNickName")).SendKeys(DateTime.Now.ToString());

                if (driver.FindElement(By.Id("saveAddressCheckboxInline")).Selected)
                {
                    driver.FindElement(By.Id("saveAddressCheckboxInline")).Click();
                }

                driver.FindElement(By.Id("addressFormInlineButton")).Click();

                System.Threading.Thread.Sleep(3000);

                if (driver.FindElements(By.XPath("//span[contains(text(), 'Continue')]")).Count > 0)
                {
                    driver.FindElement(By.XPath("//span[contains(text(), 'Continue')]")).Click();
                }

                System.Threading.Thread.Sleep(3000);

                if (hasElement(driver, By.Id("deliverySubmitButton")))
                    driver.FindElement(By.Id("deliverySubmitButton")).Click();

                driver.FindElement(By.Id("cc_cvc")).SendKeys("819");

                driver.FindElement(By.Id("paymentSubButtonBot")).Click();

                System.Threading.Thread.Sleep(3000);

                if (isAlertPresents(ref driver))
                    driver.SwitchTo().Alert().Accept();

                //driver.FindElement(By.Id("orderButton")).Click();
            }
            catch (Exception e)
            {

            }
            finally
            {
                driver.Dispose();
            }
        }

        private void ProcessCostcoShipEmail(string body)
        {
            try
            {
                body = body.Replace("\r", "");
                body = body.Replace("\t", "");
                body = body.Replace("\n", "");
                string stOrderNumber = SubstringInBetween(body, "Order Number:</td>", "</td>", false, true);
                stOrderNumber = SubstringEndBack(stOrderNumber, "</td>", ">", false, false);
                stOrderNumber = stOrderNumber.Trim();

                string stWorking = SubstringInBetween(body, "Tracking #", "Shipping &amp; Terms", false, false);
                stWorking = TrimTags(stWorking);

                string stProductName = stWorking.Substring(0, stWorking.IndexOf("<"));
                stProductName = stProductName.Trim();

                stWorking = stWorking.Substring(stProductName.Length);
                stWorking = TrimTags(stWorking);

                string stQuatity = stWorking.Substring(0, stWorking.IndexOf("<"));
                stQuatity = stQuatity.Trim();

                stWorking = stWorking.Substring(stQuatity.Length);
                stWorking = TrimTags(stWorking);

                string stShipDate = stWorking.Substring(0, stWorking.IndexOf("<"));
                stShipDate = stShipDate.Trim();

                stWorking = stWorking.Substring(stShipDate.Length);
                stWorking = TrimTags(stWorking);

                string stArriveDate = stWorking.Substring(0, stWorking.IndexOf("<"));
                stArriveDate = stArriveDate.Trim();

                stWorking = stWorking.Substring(stArriveDate.Length);
                stWorking = TrimTags(stWorking);

                string stTrackNum = stWorking.Substring(0, stWorking.IndexOf("<"));
                stTrackNum = stTrackNum.Trim();


                // Generate PDF for email
                File.WriteAllText(@"C:\temp\temp.html", body);

                FirefoxProfile profile = new FirefoxProfile();
                profile.SetPreference("print.always_print_silent", true);

                IWebDriver driver = new FirefoxDriver(profile);

                driver.Navigate().GoToUrl(@"file:///C:/temp/temp.html");

                IJavaScriptExecutor js = driver as IJavaScriptExecutor;

                js.ExecuteScript("window.print();");

                driver.Dispose();

                System.Threading.Thread.Sleep(3000);

                // Process files
                string[] files = Directory.GetFiles(@"C:\temp\tempPDF\");

                string sourceFileFullName = files[0];

                string sourceFileName = sourceFileFullName.Replace(@"C:\temp\tempPDF\", "");

                string destinationFileName = Convert.ToDateTime(stShipDate).ToString("yyyyMMddHHmmss") + "_" + stOrderNumber + ".pdf";

                File.Delete(@"C:\temp\CostcoShipEmails\" + destinationFileName);

                File.Move(sourceFileFullName, @"C:\temp\CostcoShipEmails\" + destinationFileName);

                // db stuff
                string sqlString;

                SqlConnection cn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cn.Open();

                sqlString = @"UPDATE eBay_SoldTransactions SET CostcoTrackingNumber = @_costcoTrackingNumber,
                        CostcoShipEmailPdf = @_costcoShipEmailPdf, CostcoShipDate = @_costcoShipDate 
                        WHERE CostcoOrderNumber = @_costcoOrderNumber";

                cmd.CommandText = sqlString;
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@_costcoTrackingNumber", stTrackNum);
                cmd.Parameters.AddWithValue("@_costcoShipEmailPdf", destinationFileName);
                cmd.Parameters.AddWithValue("@_costcoOrderNumber", stOrderNumber);
                cmd.Parameters.AddWithValue("@_costcoShipDate", stShipDate);

                cmd.ExecuteNonQuery();

                cn.Close();
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
        }

        public bool isAlertPresents(ref IWebDriver driver)
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }// try
            catch (Exception e)
            {
                return false;
            }// catch
        }

        private void ProcessPaymentReceivedEmail(string html)
        {
            {
                string body = html;

                // TransactionID
                string stTime = SubstringEndBack(body, "Transaction ID:", "<td ", true, false);
                stTime = TrimTags(stTime);
                stTime = stTime.Replace("<br>", "");


                string stTimeZone = stTime.Substring(stTime.LastIndexOf(' ') + 1, stTime.Length - stTime.LastIndexOf(' ') - 1);

                DateTime dtTime = Convert.ToDateTime(stTime.Replace(stTimeZone, timeZones[stTimeZone]));

                string stTransactionID = SubstringInBetween(body, "Transaction ID:", "</a>", true, true);

                stTransactionID = SubstringEndBack(stTransactionID, "</a>", ">", false, false);

                // Buyer Name
                string stBuyer = SubstringInBetween(body, "Buyer", @"</a>", false, true);

                string stFullName = SubstringInBetween(stBuyer, "<br>", "<br>", false, false);

                stBuyer = stBuyer.Replace("<br>" + stFullName + "<br>", "");

                string stUserID = SubstringInBetween(stBuyer, @"</span>", "<br>", false, false);

                stBuyer = stBuyer.Replace(stUserID, "");
                stBuyer = TrimTags(stBuyer);

                string stUserEmail = stBuyer.Substring(0, stBuyer.IndexOf('<'));


                // Shipping Address
                string stShippingAddress = SubstringInBetween(body, "Shipping address", "</td>", true, false);

                string stShippingName = SubstringInBetween(stShippingAddress, "<br>", "<br>", false, false);

                stShippingAddress = stShippingAddress.Replace("<br>" + stShippingName, "");

                string stShippingAddress1 = SubstringInBetween(stShippingAddress, "<br>", "<br>", false, false);

                stShippingAddress = stShippingAddress.Replace("<br>" + stShippingAddress1, "");

                string stShippingAddress2 = SubstringInBetween(stShippingAddress, "<br>", "<br>", false, false);

                string stShippingCity = stShippingAddress2.Substring(0, stShippingAddress2.IndexOf(","));

                string stShippingState = SubstringInBetween(stShippingAddress2, "&nbsp;", "&nbsp;", false, false);

                stShippingAddress2 = stShippingAddress2.Replace(stShippingCity, "");
                stShippingAddress2 = stShippingAddress2.Replace(stShippingState, "");
                stShippingAddress2 = stShippingAddress2.Replace("&nbsp;", "");
                stShippingAddress2 = stShippingAddress2.Replace(",", "");

                string stShippingZip = stShippingAddress2;

                // Buyer note
                string stBuyerNote = SubstringInBetween(body, "Note to seller", "</td>", false, true);
                stBuyerNote = SubstringInBetween(stBuyerNote, "<br>", "</td>", false, false);
                stBuyerNote = stBuyerNote.Replace("The buyer hasn't sent a note.", "");

                // Item 
                string stItemNum = SubstringInBetween(body, "Item#", "</td>", false, false);
                stItemNum = stItemNum.Trim();

                string stItemName = SubstringEndBack(body, "Item# " + stItemNum, "<a target='new' href='http://cgi.ebay.com/ws/eBayISAPI.dll?ViewItem&amp;item=" + stItemNum, true, false);
                stItemName = TrimTags(stItemName);
                stItemName = stItemName.Substring(0, stItemName.IndexOf('<'));

                // Amount
                string stAmount = SubstringInBetween(body, "Item# " + stItemNum, @"</table>", false, false);
                stAmount = TrimTags(stAmount);
                //stAmount = stAmount.Substring(0, stAmount.IndexOf('<'));

                string stUnitePrice = stAmount.Substring(0, stAmount.IndexOf("<"));
                stUnitePrice = stUnitePrice.Replace("$", "");
                stUnitePrice = stUnitePrice.Replace("USD", "");
                stUnitePrice = stUnitePrice.Trim();

                stAmount = stAmount.Substring(stUnitePrice.Length + 5);

                stAmount = TrimTags(stAmount);

                string stQuatity = stAmount.Substring(0, stAmount.IndexOf("<"));

                stAmount = stAmount.Substring(stQuatity.Length);

                stAmount = TrimTags(stAmount);

                string stTotal = stAmount.Substring(0, stAmount.IndexOf("<"));
                stTotal = stTotal.Replace("$", "");
                stTotal = stTotal.Replace("USD", "");
                stTotal = stTotal.Trim();

                // Tax 
                string stTax = SubstringInBetween(body, "Tax", "Total", false, false);
                stTax = TrimTags(stTax);
                stTax = stTax.Substring(0, stTax.IndexOf("<"));
                stTax = stTax.Replace("$", "");
                stTax = stTax.Replace("USD", "");
                stTax = stTax.Trim();

                //// Generate PDF for email
                //File.WriteAllText(@"C:\temp\temp.html", body);

                FirefoxProfile profile = new FirefoxProfile();
                profile.SetPreference("print.always_print_silent", true);

                IWebDriver driver = new FirefoxDriver(profile);

                driver.Navigate().GoToUrl(@"file:///C:/temp/temp.html");

                IJavaScriptExecutor js = driver as IJavaScriptExecutor;

                js.ExecuteScript("window.print();");

                driver.Dispose();

                System.Threading.Thread.Sleep(3000);

                // Process files
                string[] files = Directory.GetFiles(@"C:\temp\tempPDF\");

                string sourceFileFullName = files[0];

                string sourceFileName = sourceFileFullName.Replace(@"C:\temp\tempPDF\", "");

                string destinationFileName = dtTime.ToString("yyyyMMddHHmmss") + "_" + stTransactionID + ".pdf";

                File.Delete(@"C:\eBayApp\Files\Emails\PaypalPaidEmails\" + destinationFileName);

                File.Move(sourceFileFullName, @"C:\eBayApp\Files\Emails\PaypalPaidEmails\" + destinationFileName);

                sourceFileFullName = destinationFileName;
                destinationFileName = dtTime.ToString("yyyyMMdd") + "-" + stTotal + "-" + "Paypal" + stTransactionID + ".pdf";
                File.Delete(@"C:\Users\Jason Ding\Dropbox\Bookkeeping\Income\" + destinationFileName);
                File.Copy(@"C:\eBayApp\Files\Emails\PaypalPaidEmails\" + sourceFileFullName, @"C:\Users\Jason Ding\Dropbox\Bookkeeping\Income\" + destinationFileName);

                // db stuff
                SqlConnection cn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cn.Open();

                string sqlString = @"UPDATE eBay_SoldTransactions SET PaypalTransactionID = @_paypalTransactionID, 
                                PaypalPaidDateTime = @_paypalPaidDateTime, PaypalPaidEmailPdf = @_paypalPaidEmailPdf,
                                BuyerEmail = @_buyerEmail,
                                BuyerName = @_buyerName,
                                BuyerAddress1 = @_buyerAddress1, 
                                BuyerCity = @_buyerCity, 
                                BuyerState = @_buyerState, BuyerZip = @_buyerZip, BuyerNote = @_buyerNote,
                                eBaySoldQuality = @_eBaySoldQuality, eBaySaleTax = @_eBaySaleTax
                                WHERE eBayItemNumber = @_eBayItemNumber AND BuyerID = @_buyerID";

                cmd.CommandText = sqlString;
                cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
                cmd.Parameters.AddWithValue("@_paypalPaidDateTime", dtTime);
                cmd.Parameters.AddWithValue("@_paypalPaidEmailPdf", destinationFileName);
                cmd.Parameters.AddWithValue("@_buyerEmail", stUserEmail);
                cmd.Parameters.AddWithValue("@_buyerName", stFullName);
                cmd.Parameters.AddWithValue("@_buyerAddress1", stShippingAddress1);
                //cmd.Parameters.AddWithValue("@_buyAddress2", stShippingAddress2);
                cmd.Parameters.AddWithValue("@_buyerCity", stShippingCity);
                cmd.Parameters.AddWithValue("@_buyerState", stShippingState);
                cmd.Parameters.AddWithValue("@_buyerZip", stShippingZip);
                cmd.Parameters.AddWithValue("@_buyerNote", stBuyerNote);
                cmd.Parameters.AddWithValue("@_eBaySoldQuality", stQuatity);
                cmd.Parameters.AddWithValue("@_eBayItemNumber", stItemNum);
                cmd.Parameters.AddWithValue("@_eBaySaleTax", stTax);
                cmd.Parameters.AddWithValue("@_buyerID", stUserID);

                cmd.ExecuteNonQuery();

                sqlString = @"SELECT * FROM eBay_SoldTransactions WHERE eBayItemNumber = @_eBayItemNumber AND BuyerID = @_buyerID";

                cmd.CommandText = sqlString;
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@_eBayItemNumber", stItemNum);
                cmd.Parameters.AddWithValue("@_buyerID", stUserID);

                eBaySoldProduct soldProduct = new eBaySoldProduct();

                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();

                    soldProduct.PaypalTransactionID = Convert.ToString(reader["PaypalTransactionID"]);
                    soldProduct.PaypalPaidDateTime = Convert.ToDateTime(reader["PaypalPaidDateTime"]);
                    soldProduct.PaypalPaidEmailPdf = Convert.ToString(reader["PaypalPaidEmailPdf"]);
                    soldProduct.eBayItemNumber = Convert.ToString(reader["eBayItemNumber"]);
                    soldProduct.eBaySoldDateTime = Convert.ToDateTime(reader["eBaySoldDateTime"]);
                    soldProduct.eBayItemName = Convert.ToString(reader["eBayItemName"]);
                    //soldProduct.eBayListingQuality = Convert.ToInt16(reader["eBayListingQuality"]);
                    soldProduct.eBaySoldQuality = Convert.ToInt16(reader["eBaySoldQuality"]);
                    soldProduct.eBaySoldEmailPdf = Convert.ToString(reader["eBaySoldEmailPdf"]);
                    soldProduct.BuyerName = Convert.ToString(reader["BuyerName"]);
                    soldProduct.BuyerID = Convert.ToString(reader["BuyerID"]);
                    soldProduct.BuyerAddress1 = Convert.ToString(reader["BuyerAddress1"]);
                    soldProduct.BuyerAddress2 = Convert.ToString(reader["BuyerAddress2"]);
                    soldProduct.BuyerCity = Convert.ToString(reader["BuyerCity"]);
                    soldProduct.BuyerState = Convert.ToString(reader["BuyerState"]);
                    soldProduct.BuyerZip = Convert.ToString(reader["BuyerZip"]);
                    soldProduct.BuyerEmail = Convert.ToString(reader["BuyerEmail"]);
                    soldProduct.BuyerNote = Convert.ToString(reader["BuyerNote"]);
                    soldProduct.CostcoUrlNumber = Convert.ToString(reader["CostcoUrlNumber"]);
                    soldProduct.CostcoUrl = Convert.ToString(reader["CostcoUrl"]);
                    soldProduct.CostcoPrice = Convert.ToDecimal(reader["CostcoPrice"]);

                    //soldProduct.CostcoUrl = "http://www.costco.com/Vasanti-Gel-Matte-Lipstick-with-Lipline-Extreme-Lipliner.product.100243171.html";
                }
                reader.Close();

                cn.Close();
            }
        }

        private string SubstringInBetween(string input, string start, string end, bool bIncludeStart, bool bIncludeEnd)
        {
            int iStart = input.IndexOf(start);

            if (bIncludeStart)
                input = input.Substring(iStart);
            else
                input = input.Substring(iStart + start.Length);

            int iEnd = input.IndexOf(end);

            if (bIncludeEnd)
                input = input.Substring(0, iEnd + end.Length);
            else
                input = input.Substring(0, iEnd);

            return input;
        }

        private string SubstringEndBack(string input, string end, string start, bool bIncludeStart, bool bIncludeEnd)
        {
            int iEnd = input.IndexOf(end);

            if (bIncludeEnd)
                input = input.Substring(0, iEnd + end.Length);
            else
                input = input.Substring(0, iEnd);

            int iStart = input.LastIndexOf(start);

            if (bIncludeStart)
                input = input.Substring(iStart, input.Length - iStart);
            else
                input = input.Substring(iStart + start.Length, input.Length - iStart - start.Length);

            return input;
        }

        private string TrimTags(string input)
        {
            int iStart = input.IndexOf("<");
            string stTag;
            input = input.Substring(iStart);

            while (input.IndexOf("<") == 0)
            {
                stTag = SubstringInBetween(input, "<", ">", true, true);
                input = input.Substring(stTag.Length);
                input = input.TrimStart();
            }

            return input;
        }

        public string GetState(string state)
        {
            switch (state.ToUpper())
            {
                case "AL":
                    return "Alabama";

                case "AK":
                    return "Alaska";

                case "AS":
                    return "American Samoa";

                case "AZ":
                    return "Arizona";

                case "AR":
                    return "Arkansas";

                case "CA":
                    return "California";

                case "CO":
                    return "Colorado";

                case "CT":
                    return "Connecticut";

                case "DE":
                    return "Delaware";

                case "DC":
                    return "District of Columbia";

                case "FL":
                    return "Florida";

                case "GA":
                    return "Georgia";

                case "GU":
                    return "Guam";

                case "HI":
                    return "Hawaii";

                case "ID":
                    return "Idaho";

                case "IL":
                    return "Illinois";

                case "IN":
                    return "Indiana";

                case "IA":
                    return "Iowa";

                case "KS":
                    return "Kansas";

                case "KY":
                    return "Kentucky";

                case "LA":
                    return "Louisiana";

                case "ME":
                    return "Maine";

                case "MH":
                    return "Narshall Islands";

                case "MD":
                    return "Maryland";

                case "MA":
                    return "Massachusetts";

                case "MI":
                    return "Michigan";

                case "MN":
                    return "Minnesota";

                case "MS":
                    return "Mississippi";

                case "MO":
                    return "Missouri";

                case "MT":
                    return "Montana";

                case "NE":
                    return "Nebraska";

                case "NV":
                    return "Nevada";

                case "NH":
                    return "New Hampshire";

                case "NJ":
                    return "New Jersey";

                case "NM":
                    return "New Mexico";

                case "NY":
                    return "New York";

                case "NC":
                    return "North Carolina";

                case "ND":
                    return "North Dakota";

                case "OH":
                    return "Ohio";

                case "OK":
                    return "Oklahoma";

                case "OR":
                    return "Oregon";

                case "PW":
                    return "Palau";

                case "PA":
                    return "Pennsylvania";

                case "PR":
                    return "Puerto Rico";

                case "RI":
                    return "Rhode Island";

                case "SC":
                    return "South Carolina";

                case "SD":
                    return "South Dakota";

                case "TN":
                    return "Tennessee";

                case "TX":
                    return "Texas";

                case "UT":
                    return "Utah";

                case "VT":
                    return "Vermont";

                case "VI":
                    return "Virgin Islands";

                case "VA":
                    return "Virginia";

                case "WA":
                    return "Washington";

                case "WV":
                    return "West Virginia";

                case "WI":
                    return "Wisconsin";

                case "WY":
                    return "Wyoming";
            }

            throw new Exception("Not Available");
        }

        public Dictionary<string, string> timeZones = new Dictionary<string, string>() {
        {"ACDT", "+1030"},
        {"ACST", "+0930"},
        {"ADT", "-0300"},
        {"AEDT", "+1100"},
        {"AEST", "+1000"},
        {"AHDT", "-0900"},
        {"AHST", "-1000"},
        {"AST", "-0400"},
        {"AT", "-0200"},
        {"AWDT", "+0900"},
        {"AWST", "+0800"},
        {"BAT", "+0300"},
        {"BDST", "+0200"},
        {"BET", "-1100"},
        {"BST", "-0300"},
        {"BT", "+0300"},
        {"BZT2", "-0300"},
        {"CADT", "+1030"},
        {"CAST", "+0930"},
        {"CAT", "-1000"},
        {"CCT", "+0800"},
        {"CDT", "-0500"},
        {"CED", "+0200"},
        {"CET", "+0100"},
        {"CEST", "+0200"},
        {"CST", "-0600"},
        {"EAST", "+1000"},
        {"EDT", "-0400"},
        {"EED", "+0300"},
        {"EET", "+0200"},
        {"EEST", "+0300"},
        {"EST", "-0500"},
        {"FST", "+0200"},
        {"FWT", "+0100"},
        {"GMT", "GMT"},
        {"GST", "+1000"},
        {"HDT", "-0900"},
        {"HST", "-1000"},
        {"IDLE", "+1200"},
        {"IDLW", "-1200"},
        {"IST", "+0530"},
        {"IT", "+0330"},
        {"JST", "+0900"},
        {"JT", "+0700"},
        {"MDT", "-0600"},
        {"MED", "+0200"},
        {"MET", "+0100"},
        {"MEST", "+0200"},
        {"MEWT", "+0100"},
        {"MST", "-0700"},
        {"MT", "+0800"},
        {"NDT", "-0230"},
        {"NFT", "-0330"},
        {"NT", "-1100"},
        {"NST", "+0630"},
        {"NZ", "+1100"},
        {"NZST", "+1200"},
        {"NZDT", "+1300"},
        {"NZT", "+1200"},
        {"PDT", "-0700"},
        {"PST", "-0800"},
        {"ROK", "+0900"},
        {"SAD", "+1000"},
        {"SAST", "+0900"},
        {"SAT", "+0900"},
        {"SDT", "+1000"},
        {"SST", "+0200"},
        {"SWT", "+0100"},
        {"USZ3", "+0400"},
        {"USZ4", "+0500"},
        {"USZ5", "+0600"},
        {"USZ6", "+0700"},
        {"UT", "-0000"},
        {"UTC", "-0000"},
        {"UZ10", "+1100"},
        {"WAT", "-0100"},
        {"WET", "-0000"},
        {"WST", "+0800"},
        {"YDT", "-0800"},
        {"YST", "-0900"},
        {"ZP4", "+0400"},
        {"ZP5", "+0500"},
        {"ZP6", "+0600"}
        };

        public enum State
        {
            AL,
            AK,
            AS,
            AZ,
            AR,
            CA,
            CO,
            CT,
            DE,
            DC,
            FM,
            FL,
            GA,
            GU,
            HI,
            ID,
            IL,
            IN,
            IA,
            KS,
            KY,
            LA,
            ME,
            MH,
            MD,
            MA,
            MI,
            MN,
            MS,
            MO,
            MT,
            NE,
            NV,
            NH,
            NJ,
            NM,
            NY,
            NC,
            ND,
            MP,
            OH,
            OK,
            OR,
            PW,
            PA,
            PR,
            RI,
            SC,
            SD,
            TN,
            TX,
            UT,
            VT,
            VI,
            VA,
            WA,
            WV,
            WI,
            WY
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
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

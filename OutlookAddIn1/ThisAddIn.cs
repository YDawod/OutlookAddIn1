﻿using System;
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

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        string connectionString = "Data Source=DESKTOP-ABEPKAT;Initial Catalog=Costco;Integrated Security=False;User ID=sa;Password=G4indigo;Connect Timeout=15;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

            string a = DateTime.Now.AddDays(10).ToString();

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

        void items_ItemAdd(object Item)
        {
            string filter = "USED CARS";
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                if (mail.TaskSubject.IndexOf("Your eBay listing is confirmed") == 0)
                {
                    string subject = mail.Subject;

                    string productName = subject.Substring(32, subject.Length - 32);

                    string sqlString = "select top 1 * from Archieve where name = '" + productName + "' order by ImportedDT desc";

                    string categoryID = "";

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
                        product.UrlNumber = Convert.ToString(reader["UrlNumber"]);
                        product.ItemNumber = Convert.ToString(reader["ItemNumber"]);
                        product.Category = Convert.ToString(reader["Category"]);
                        product.Price = Convert.ToDecimal(reader["Price"]);
                        product.Shipping = Convert.ToDecimal(reader["Shipping"]);
                        product.Discount = Convert.ToString(reader["Discount"]);
                        product.Details = Convert.ToString(reader["Details"]);
                        product.Specification = Convert.ToString(reader["Specification"]);
                        product.ImageLink = Convert.ToString(reader["ImageLink"]);
                        product.Url = Convert.ToString(reader["Url"]);
                    }
                    reader.Close();

                    string body = mail.HTMLBody;

                    // ItemID
                    int iItemId = body.IndexOf("Item Id:");
                    int iStart = iItemId + 100;
                    int iEnd = body.IndexOf("</td>", iStart);
                    string id = body.Substring(iStart - 4, iEnd - iStart + 4);

                    // ListPrice
                    iItemId = body.IndexOf("Price:");
                    iStart = iItemId + 95;
                    iEnd = body.IndexOf("</td>", iStart);
                    string price = body.Substring(iStart, iEnd - iStart);

                    // EndTime
                    iItemId = body.IndexOf("End time:");
                    iStart = iItemId + 98;
                    iEnd = body.IndexOf("</td>", iStart);
                    string endTime = body.Substring(iStart - 1, iEnd - iStart + 1);

                    //MessageBox.Show(id + "|" + price + "|" + endTime);

                    sqlString = "INSERT INTO eBay_CurrentListings (Name, eBayItemNumber, eBayListingPrice, " +
                                "eBayListingDT, CostcoUrlNumber, CostcoUrl, eBayDescription, ImageLink) " +
                                "VALUES ('" + product.Name + "', '" + id + "', '" + price + "', '" + DateTime.Now.AddDays(10).ToString() + "', '" +
                                product.UrlNumber + "', '" + product.Url + "', '" +
                                product.Specification + "', '" + product.ImageLink + "')";

                    cmd.CommandText = sqlString;
                    cmd.ExecuteNonQuery();

                    cn.Close();
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

                    string sqlString = "DELETE FROM eBay_CurrentListings WHERE Name = '" + productName + "'";
                    cmd.CommandText = sqlString;
                    cmd.ExecuteNonQuery();

                    cn.Close();
                }
                else if (mail.TaskSubject.Contains("Instant payment received"))
                {
                    string body = mail.HTMLBody;
                    body = body.Replace("\n", "");
                    body = body.Replace("\t", "");
                    body = body.Replace("\\", "");

                    ProcessPaymentReceivedEmail(body);
                }
                else if (mail.TaskSubject.IndexOf("Your eBay item sold!") == 0)
                {
                    string body = mail.HTMLBody;
                    body = body.Replace("\n", "");
                    body = body.Replace("\t", "");
                    body = body.Replace("\\", "");
                    body = body.Replace("\"", "'");

                    ProcessItemSoldEmail(mail.TaskSubject, body);
                }

                else if (mail.TaskSubject.Contains("Your Costco.com Order Was Received"))
                {
                    string body = mail.HTMLBody;
                    body = body.Replace("\n", "");
                    body = body.Replace("\t", "");
                    body = body.Replace("\\", "");
                    body = body.Replace("\"", "'");

                    body = @"<html><head></head><body>" + body + @"</body></html>";

                    ProcessCostcoOrderEmail(body);
                }
            } 
        }

        private void ProcessCostcoOrderEmail(string body)
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

            string stBuyerName = TrimTags(stShipping);

            // Generate PDF for email
            string destinationFileName = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + stOrderNumber;

            File.WriteAllText(@"C:\temp\" + @"\" + destinationFileName + ".html", body);

            FirefoxProfile profile = new FirefoxProfile();
            profile.SetPreference("print.always_print_silent", true);

            IWebDriver driver = new FirefoxDriver(profile);

            driver.Navigate().GoToUrl(@"file:///" + @"C:\temp\" + @"\" + destinationFileName + ".html");

            IJavaScriptExecutor js = driver as IJavaScriptExecutor;

            js.ExecuteScript("window.print();");

            driver.Dispose();

            System.Threading.Thread.Sleep(10000);

            // Process files
            string sourceFileName = @"C:\temp\tempPDF\file__C__temp_" + destinationFileName + @"\" + "file_C_temp_" + destinationFileName + ".pdf";

            File.Move(sourceFileName, @"C:\temp\CostcoOrderEmails\" + destinationFileName + ".pdf");

            File.Delete(@"C:\temp\" + destinationFileName + ".html");
            Directory.Delete(@"C:\temp\tempPDF\file__C__temp_" + destinationFileName);

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
                                AND BuyerName = @_buyerName AND  CostcoOrderNumber IS NULL";

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
                                CostcoOrderEmailPdf = @_costcoOrderEmailPdf
                                WHERE WHERE CostcoItemNumber = @_costcoItemNumber 
                                AND BuyerName = @_buyerName AND  CostcoOrderNumber IS NULL";

                    cmd.CommandText = sqlString;
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@_costcoOrderNumber", stOrderNumber);
                    cmd.Parameters.AddWithValue("@_costcoOrderEmailPdf", destinationFileName);
                    cmd.Parameters.AddWithValue("@_costcoItemNumber", stItemNum);
                    cmd.Parameters.AddWithValue("@_buyerName", stBuyerName);

                    cmd.ExecuteNonQuery();
                }
            }
            else
            {
                sqlString = @"SELECT * FROM eBay_SoldTransactions WHERE CostcoItemName = @_costcoItemName
                                AND CostcoOrderNumber IS NULL";

                cmd.CommandText = sqlString;
                cmd.Parameters.AddWithValue("@_costcoItemName", stProductName);

                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    bExist = true;
                }
                reader.Close();

                if (bExist)
                {
                    sqlString = @"UPDATE eBay_SoldTransactions SET CostcoOrderNumber = @_costcoOrderNumber,
                                CostcoOrderEmailPdf = @_costcoOrderEmailPdf
                                WHERE WHERE CostcoItemName = @_costcoItemName 
                                AND BuyerName = @_buyerName AND  CostcoOrderNumber IS NULL";

                    cmd.CommandText = sqlString;
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@_costcoOrderNumber", stOrderNumber);
                    cmd.Parameters.AddWithValue("@_costcoOrderEmailPdf", destinationFileName);
                    cmd.Parameters.AddWithValue("@_costcoItemName", stProductName);
                    cmd.Parameters.AddWithValue("@_buyerName", stBuyerName);

                    cmd.ExecuteNonQuery();
                }
            }

            cn.Close();


        }

        private void ProcessItemSoldEmail(string subject, string body)
        {
            string stItemNum = SubstringInBetween(subject, "(", ")", false, false);

            subject = subject.Replace("(" + stItemNum + ")", "");
            subject = subject.Replace("Your eBay item sold!", "");

            string stItemName = subject.Trim();

            string stUrl = SubstringEndBack(body, ">" + stItemName, "<a href=", false, false);

            stUrl = SubstringInBetween(stUrl, "'", "'", false, false);

            string stEndTime = SubstringInBetween(body, "End time:", "PDT", false, true);

            stEndTime = SubstringEndBack(stEndTime, "PDT", ">", false, true);

            string correctedTZ = stEndTime.Replace("PDT", "-0700");
            DateTime dt = Convert.ToDateTime(correctedTZ);

            string stPrice = SubstringInBetween(body, "Sale price:", "Quantity:", false, false);

            stPrice = SubstringInBetween(stPrice, "$", "<", false, false);

            string stQuantity = SubstringInBetween(body, "Quantity:", "Quantity sold:", false, false);

            stQuantity = TrimTags(stQuantity);

            stQuantity = stQuantity.Substring(0, stQuantity.IndexOf("<"));

            string stQuantitySold = SubstringInBetween(body, "Quantity sold:", "Quantity remaining:", false, false);

            stQuantitySold = TrimTags(stQuantitySold);

            stQuantitySold = stQuantitySold.Substring(0, stQuantitySold.IndexOf("<"));

            string stQuantityRemaining = SubstringInBetween(body, "Quantity remaining:", "Buyer:", false, false);

            stQuantityRemaining = TrimTags(stQuantityRemaining);

            stQuantityRemaining = stQuantityRemaining.Substring(0, stQuantityRemaining.IndexOf("<"));

            string stBuyerName = SubstringInBetween(body, "Buyer:", "<div>", false, false);

            stBuyerName = TrimTags(stBuyerName);

            stBuyerName = stBuyerName.Substring(0, stBuyerName.IndexOf("<"));

            string stBuyerId = SubstringInBetween(body, stBuyerName, "(<a href='mailto", false, true);

            stBuyerId = TrimTags(stBuyerId);

            stBuyerId = stBuyerId.Substring(0, stBuyerId.IndexOf("(<a href='mailto"));

            stBuyerId = stBuyerId.Trim();

            string stBuyerEmail = SubstringInBetween(body, "(<a href='mailto:", "'", false, false);

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

            string destinationFileName = dt.ToString("yyyyMMddHHmmss") + "_" + stItemNum + ".pdf";

            File.Move(sourceFileFullName, @"C:\temp\eBaySoldEmails\" + destinationFileName);

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
                eBayProduct.eBayListingDT = Convert.ToDateTime(reader["eBayListingDT"]);
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
                              (eBayItemNumber, eBaySoldDateTime, eBayItemName, eBayUrl, eBayPrice, eBayListingQuality, eBaySoldQuality, eBayRemainingQuality, eBaySoldEmailPdf,
                               BuyerName, BuyerID, BuyerEmail, CostcoUrlNumber, CostcoUrl, CostcoPrice)
                              VALUES (@_eBayItemNumber, @_eBaySoldDateTime, @_eBayItemName, @_eBayUrl, @_eBayPrice, @_eBayListingQuality, @_eBaySoldQuality, @_eBayRemainingQuality, @_eBaySoldEmailPdf,
                               @_BuyerName, @_BuyerID, @_BuyerEmail, @_CostcoUrlNumber, @_CostcoUrl, @_CostcoPrice)";

                cmd.CommandText = sqlString;
                cmd.Parameters.AddWithValue("@_eBayItemNumber", stItemNum);
                cmd.Parameters.AddWithValue("@_eBaySoldDateTime", dt);
                cmd.Parameters.AddWithValue("@_eBayItemName", stItemName);
                cmd.Parameters.AddWithValue("@_eBayUrl", stUrl);
                cmd.Parameters.AddWithValue("@_eBayPrice", Convert.ToDecimal(stPrice));
                cmd.Parameters.AddWithValue("@_eBayListingQuality", Convert.ToInt16(stQuantity));
                cmd.Parameters.AddWithValue("@_eBaySoldQuality", Convert.ToInt16(stQuantitySold));
                cmd.Parameters.AddWithValue("@_eBayRemainingQuality", Convert.ToInt16(stQuantityRemaining));
                cmd.Parameters.AddWithValue("@_eBaySoldEmailPdf", destinationFileName);
                cmd.Parameters.AddWithValue("@_BuyerName", stBuyerName);
                cmd.Parameters.AddWithValue("@_BuyerID", stBuyerId);
                cmd.Parameters.AddWithValue("@_BuyerEmail", stBuyerEmail);
                cmd.Parameters.AddWithValue("@_CostcoUrlNumber", eBayProduct.CostcoUrlNumber);
                cmd.Parameters.AddWithValue("@_CostcoUrl", eBayProduct.CostcoUrl);
                cmd.Parameters.AddWithValue("@_CostcoPrice", eBayProduct.CostcoPrice);

                cmd.ExecuteNonQuery();
            }
            else
            {

            }

            cn.Close();
        }

        private void ProcessPaymentReceivedEmail(string html)
        {
            string body = SubstringInBetween(html, "<body", @"</body>", true, true);

            // TransactionID
            string stTime = SubstringEndBack(body, "PDT", ">", false, true);

            DateTime dtTime = Convert.ToDateTime(stTime.Replace("PDT", "-0700"));

            string stTransactionID = SubstringInBetween(body, "Transaction ID:", "</a>", true, true);

            stTransactionID = SubstringEndBack(stTransactionID, "</a>", ">", false, false);

            // Buyer Name
            string stBuyer = SubstringInBetween(body, "Buyer", @"</a>", true, true);

            string stFullName = SubstringInBetween(stBuyer, "<br>", "<br>", false, false);

            stBuyer = stBuyer.Replace("<br>" + stFullName + "<br>", "");

            string stUserID = SubstringInBetween(stBuyer, @"</b>", "<br>", false, false);

            string stUserEmail = SubstringEndBack(stBuyer, @"</a>", ">", false, false);

            // Shipping Address
            string stShippingAddress = SubstringInBetween(body, "Shipping address", "<o:p>", true, false);

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
            string stBuyerNote = SubstringInBetween(body, "Note to seller", "<o:p>", false, true);
            stBuyerNote = SubstringInBetween(stBuyerNote, "<br>", "<o:p>", false, false);
            stBuyerNote = stBuyerNote.Replace("The buyer hasn't sent a note.", "");

            // Item 
            string stItemNum = SubstringInBetween(body, "Item#", "<o:p>", false, false);
            stItemNum = stItemNum.Trim();

            string stItemName = SubstringInBetween(body, "<a href='http://cgi.ebay.com/ws/eBayISAPI.dll?ViewItem&amp;item=" + stItemNum + "' target='_blank'>", @"</a>", false, false);

            // Amount
            string stAmount = SubstringInBetween(body, "Item# " + stItemNum, @"</table>", false, false);

            stAmount = TrimTags(stAmount);

            string stUnitePrice = stAmount.Substring(0, stAmount.IndexOf("<"));
            stUnitePrice = stUnitePrice.Replace("$", "");
            stUnitePrice = stUnitePrice.Replace("USD", "");
            stUnitePrice = stUnitePrice.Trim();

            stAmount = stAmount.Substring(stUnitePrice.Length);

            stAmount = TrimTags(stAmount);

            string stQuatity = stAmount.Substring(0, stAmount.IndexOf("<"));

            stAmount = stAmount.Substring(stQuatity.Length);

            stAmount = TrimTags(stAmount);

            string stTotal = stAmount.Substring(0, stAmount.IndexOf("<"));
            stTotal = stTotal.Replace("$", "");
            stTotal = stTotal.Replace("USD", "");
            stTotal = stTotal.Trim();

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

            string destinationFileName = dtTime.ToString("yyyyMMddHHmmss") + "_" + stTransactionID + ".pdf";

            File.Move(sourceFileFullName, @"C:\temp\PaypalPaidEmails\" + destinationFileName);

            // db stuff
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cn.Open();

            string sqlString = @"UPDATE eBay_SoldTransactions SET PaypalTransactionID = @_paypalTransactionID, 
                                PaypalPaidDateTime = @_paypalPaidDateTime, PaypalPaidEmailPdf = @_paypalPaidEmailPdf,
                                BuyerAddress1 = @_buyerAddress1, 
                                BuyerAddress2 = @_buyAddress2, BuyerState = @_buyerState, BuyerNote = @_buyerNote 
                                WHERE eBayItemNumber = @_eBayItemNumber AND BuyerID = @_buyerID";

            cmd.CommandText = sqlString;
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalPaidDateTime", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);




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

            int iStart = input.LastIndexOf(">");

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
            }

            return input;
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

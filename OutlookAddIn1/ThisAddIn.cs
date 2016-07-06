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

                        string sqlString = "DELETE FROM eBay_CurrentListings WHERE eBayListingName = '" + productName + "'";
                        //string sqlString = "UPDATE eBay_CurrentListings SET DeleteDT = GETDATE() WHERE eBayListingName = '" + productName + "'";
                        cmd.CommandText = sqlString;
                        cmd.ExecuteNonQuery();

                        sqlString = "DELETE FROM eBay_ToRemove WHERE eBayListingName = '" + productName + "'";
                        //sqlString = "UPDATE eBay_ToRemove SET DeleteDT = GETDATE() WHERE eBayListingName = '" + productName + "'";
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
                        body = body.Replace("\"", "'");

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

                        //body = @"<html><head></head><body>" + body + @"</body></html>";

                        ProcessCostcoOrderEmail(body);
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

        private void ProcessListingConfirmEmail(string body, string subject)
        {
            body = body.Replace("\n", "");
            body = body.Replace("\t", "");
            body = body.Replace("\\", "");
            body = body.Replace("\"", "'");

            string productName = subject.Substring(32, subject.Length - 32);

            string sqlString = "SELECT * FROM eBay_ToAdd WHERE Name = '" + productName + "'";

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
                product.Limit = Convert.ToString(reader["Limit"]);
                product.Details = Convert.ToString(reader["Details"]);
                product.Specification = Convert.ToString(reader["Specification"]);
                product.ImageLink = Convert.ToString(reader["ImageLink"]);
                product.NumberOfImage = Convert.ToInt16(reader["NumberOfImage"]);
                product.Url = Convert.ToString(reader["Url"]);
                product.eBayCategoryID = Convert.ToString(reader["eBayCategoryID"]);
                product.eBayReferencePrice = Convert.ToDecimal(reader["eBayReferencePrice"]);
                product.eBayListingPrice = Convert.ToDecimal(reader["eBayListingPrice"]);
                product.DescriptionImageWidth = Convert.ToInt16(reader["DescriptionImageWidth"]);
                product.DescriptionImageHeight = Convert.ToInt16(reader["DescriptionImageHeight"]);
            }

            reader.Close();

            sqlString = @"DELETE FROM eBay_ToAdd WHERE name = '" + productName + "'";
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
            stEndTime = SubstringEndBack(stEndTime, "PDT", ">", false, true);
            stEndTime = stEndTime.Trim();
            string correctedTZ = stEndTime.Replace("PDT", "-0700");
            DateTime dtEndTime = Convert.ToDateTime(correctedTZ);

            sqlString = @"INSERT INTO eBay_CurrentListings
                            (Name, eBayListingName, eBayCategoryID, eBayItemNumber, eBayListingPrice, eBayDescription, 
                             eBayEndTime, CostcoUrlNumber, CostcoItemNumber, CostcoUrl, CostcoPrice, ImageLink) 
                          VALUES (@_name, @_eBayListingName, @_eBayCategoryID, @_eBayItemNumber, @_eBayListingPrice, @_eBayDescription,
                                @_eBayEndTime, @_CostcoUrlNumber, @_CostcoItemNumber, @_CostcoUrl, @_CostcoPrice, @_ImageLink)";

            cmd.CommandText = sqlString;
            cmd.Parameters.AddWithValue("@_name", product.Name);
            cmd.Parameters.AddWithValue("@_eBayListingName", productName);
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

                string destinationFileName = Convert.ToDateTime(stDatePlaced).ToString("yyyyMMddHHmmss") + "_" + stOrderNumber + ".pdf";

                File.Delete(@"C:\temp\CostcoOrderEmails\" + destinationFileName);

                File.Move(sourceFileFullName, @"C:\temp\CostcoOrderEmails\" + destinationFileName);

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
                string stItemNum = SubstringInBetween(subject, "(", ")", false, false);

                subject = subject.Replace("(" + stItemNum + ")", "");
                subject = subject.Replace("Your eBay item sold!", "");

                string stItemName = subject.Trim();

                stItemName = WebUtility.HtmlDecode(stItemName);

                body = WebUtility.HtmlDecode(body);

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
                //string destinationFileName = dt.ToString("yyyyMMddHHmmss") + "_" + stItemNum + ".html";
                //File.WriteAllText(@"C:\temp\eBaySoldEmails\" + destinationFileName, body);

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
                              (eBayItemNumber, eBaySoldDateTime, eBayItemName, eBayUrl, eBaySoldPrice, eBayListingQuality, eBaySoldQuality, eBayRemainingQuality, eBaySoldEmailPdf,
                               BuyerName, BuyerID, BuyerEmail, CostcoUrlNumber, CostcoUrl, CostcoPrice)
                              VALUES (@_eBayItemNumber, @_eBaySoldDateTime, @_eBayItemName, @_eBayUrl, @_eBayPrice, @_eBayListingQuality, @_eBaySoldQuality, @_eBayRemainingQuality, @_eBaySoldEmailPdf,
                               @_BuyerName, @_BuyerID, @_BuyerEmail, @_CostcoUrlNumber, @_CostcoUrl, @_CostcoPrice)";

                    cmd.CommandText = sqlString;

                    cmd.Parameters.Clear();
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

                    //eBaySoldProduct soldProduct = new eBaySoldProduct();
                    //soldProduct.eBayItemName = stItemNum;
                    //soldProduct.eBaySoldDateTime = dt;
                    //soldProduct.eBayItemName = stItemName;
                    //soldProduct.eBayUrl = stUrl;
                    //soldProduct.eBaySoldPrice = Convert.ToDecimal(stPrice);
                    //soldProduct.eBayListingQuality = Convert.ToInt16(stQuantity);
                    //soldProduct.eBaySoldQuality = Convert.ToInt16(stQuantitySold);
                    //soldProduct.eBayRemainingQuality = Convert.ToInt16(stQuantityRemaining);
                    //soldProduct.eBaySoldEmailPdf = destinationFileName;
                    //soldProduct.BuyerName = stBuyerName;
                    //soldProduct.BuyerID = stBuyerId;
                    //soldProduct.BuyerEmail = stBuyerEmail;
                    //soldProduct.CostcoUrlNumber = eBayProduct.CostcoUrlNumber;
                    //soldProduct.CostcoUrl = eBayProduct.CostcoUrl;
                    //soldProduct.CostcoPrice = eBayProduct.CostcoPrice;
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

        private void OrderCostcoProduct(eBaySoldProduct soldProduct)
        {
            try
            {
                IWebDriver driver = new FirefoxDriver();

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
                driver.FindElement(By.Id("minQtyText")).Clear();
                driver.FindElement(By.Id("minQtyText")).SendKeys("1");
                driver.FindElement(By.Id("addToCartBtn")).Click();

                if (isAlertPresents(ref driver))
                    driver.SwitchTo().Alert().Accept();

                driver.FindElement(By.Id("cart-d")).Click();

                if (isAlertPresents(ref driver))
                    driver.SwitchTo().Alert().Accept();

                string buyerFirstName = soldProduct.BuyerName.Split(' ')[0];
                string buyerLastName = soldProduct.BuyerName.Split(' ')[soldProduct.BuyerName.Split(' ').Count() - 1];


                driver.FindElement(By.Id("shopCartCheckoutSubmitButton")).Click();

                driver.FindElement(By.Id("addressFormInlineFirstName")).SendKeys(buyerFirstName);
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
            string body = html;

            // TransactionID
            string stTime = SubstringEndBack(body, "PDT", ">", false, true);

            DateTime dtTime = Convert.ToDateTime(stTime.Replace("PDT", "-0700"));

            string stTransactionID = SubstringInBetween(body, "Transaction ID:", "</a>", true, true);

            stTransactionID = SubstringEndBack(stTransactionID, "</a>", ">", false, false);

            // Buyer Name
            string stBuyer = SubstringInBetween(body, "Buyer", @"</a>", true, true);

            string stFullName = SubstringInBetween(stBuyer, "<br>", "<br>", false, false);

            stBuyer = stBuyer.Replace("<br>" + stFullName + "<br>", "");

            string stUserID = SubstringInBetween(stBuyer, @"</span>", "<br>", false, false);

            string stUserEmail = SubstringEndBack(stBuyer, @"</a>", ">", false, false);

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

            string stItemName = SubstringInBetween(body, "<a href='http://cgi.ebay.com/ws/eBayISAPI.dll?ViewItem&amp;item=" + stItemNum + "' target='_blank'>", @"</a>", false, false);

            // Amount
            string stAmount = SubstringInBetween(body, "Item# " + stItemNum, @"</table>", false, false);

            stAmount = TrimTags(stAmount);

            string stUnitePrice = stAmount.Substring(0, stAmount.IndexOf("<"));
            stUnitePrice = stUnitePrice.Replace("$", "");
            stUnitePrice = stUnitePrice.Replace("USD", "");
            stUnitePrice = stUnitePrice.Trim();

            stAmount = stAmount.Substring(stUnitePrice.Length+5);

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

            File.Delete(@"C:\temp\PaypalPaidEmails\" + destinationFileName);

            File.Move(sourceFileFullName, @"C:\temp\PaypalPaidEmails\" + destinationFileName);

            // db stuff
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cn.Open();

            string sqlString = @"UPDATE eBay_SoldTransactions SET PaypalTransactionID = @_paypalTransactionID, 
                                PaypalPaidDateTime = @_paypalPaidDateTime, PaypalPaidEmailPdf = @_paypalPaidEmailPdf,
                                BuyerAddress1 = @_buyerAddress1, 
                                BuyerAddress2 = @_buyAddress2, BuyerCity = @_buyerCity, 
                                BuyerState = @_buyerState, BuyerZip = @_buyerZip, BuyerNote = @_buyerNote,
                                eBaySoldQuality = @_eBaySoldQuality
                                WHERE eBayItemNumber = @_eBayItemNumber AND BuyerID = @_buyerID";

            cmd.CommandText = sqlString;
            cmd.Parameters.AddWithValue("@_paypalTransactionID", stTransactionID);
            cmd.Parameters.AddWithValue("@_paypalPaidDateTime", dtTime);
            cmd.Parameters.AddWithValue("@_paypalPaidEmailPdf", destinationFileName);
            cmd.Parameters.AddWithValue("@_buyerAddress1", stShippingAddress1);
            cmd.Parameters.AddWithValue("@_buyAddress2", stShippingAddress2);
            cmd.Parameters.AddWithValue("@_buyerCity", stShippingCity);
            cmd.Parameters.AddWithValue("@_buyerState", stShippingState);
            cmd.Parameters.AddWithValue("@_buyerZip", stShippingZip);
            cmd.Parameters.AddWithValue("@_buyerNote", stBuyerNote);
            cmd.Parameters.AddWithValue("@_eBaySoldQuality", stQuatity);
            cmd.Parameters.AddWithValue("@_eBayItemNumber", stItemNum);
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
                soldProduct.eBayListingQuality = Convert.ToInt16(reader["eBayListingQuality"]);
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

            OrderCostcoProduct(soldProduct);
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

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

                    


                    string body = mail.HTMLBody;
                    int iItemId = body.IndexOf("Item Id:");

                    int iStart = iItemId + 100;

                    int iEnd = body.IndexOf("</td>", iStart);

                    string id = body.Substring(iStart - 4, iEnd - iStart + 4);

                    MessageBox.Show(id);
                }

                //if (mail.MessageClass == "IPM.Note" &&
                //           mail.Subject.ToUpper().Contains(filter.ToUpper()))
                //{
                //    mail.Move(outlookNameSpace.GetDefaultFolder(
                //        Microsoft.Office.Interop.Outlook.
                //        OlDefaultFolders.olFolderJunk));
                //}
            }

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

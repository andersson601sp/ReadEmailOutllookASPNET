using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Text;

namespace Worker.ReadEmailOutllookASPNET
{
    public class OutlookEmails
    {
        public string EmailFrom { get; set; }
        public string EmailSubject{ get; set; }
        public string EmailBody { get; set; }

        public static List<OutlookEmails> ReadEmailItems()
        {
            Application outlookAplication = null;
            NameSpace outlookNameSpace = null;
            MAPIFolder inboxFolder = null;

            Items mailItems = null;
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails emailDetails;

            try
            {
                outlookAplication = new Application();
                outlookNameSpace = outlookAplication.GetNamespace("MAPI");
                inboxFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;

                foreach (MailItem item in mailItems)
                {
                    emailDetails = new OutlookEmails();
                    emailDetails.EmailFrom = item.SenderEmailAddress;
                    emailDetails.EmailSubject = item.Subject;
                    emailDetails.EmailBody = item.Body;
                    var x = item.Attachments;
                    listEmailDetails.Add(emailDetails);
                }
            }
            catch (System.Exception ex)
            {
                Console.Write(ex);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNameSpace);
                ReleaseComObject(outlookAplication);
            }
            return listEmailDetails;
        }

        public static void ReleaseComObject(object obj)
        {
            if(obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}



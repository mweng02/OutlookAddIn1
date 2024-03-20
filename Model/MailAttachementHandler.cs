using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    public class MailAttachementHandler
    {
        public MailItem mailItem; 

        public MailAttachementHandler(MailItem mailItem)
        {
            this.mailItem = mailItem;
            this.mailItem.AttachmentAdd += new ItemEvents_10_AttachmentAddEventHandler(MailItem_AttachmentAdd); 
        }

        public void MailItem_AttachmentAdd(Attachment attachment)
        {

            string subjectName = UtilsModel.ExtractRegex(attachment.FileName, @"(Q|C)\d{4}-\d{4,}-[A-Z]{2}");
            if (!string.IsNullOrEmpty(subjectName))
            {
                mailItem.Subject = subjectName;
                mailItem.Save();
            }
           
        }
    }
}

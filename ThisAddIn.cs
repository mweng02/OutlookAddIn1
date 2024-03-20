using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using Newtonsoft.Json;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            var _view = Globals.Ribbons.GetRibbon<Ribbon_>();
            Controller controller = new Controller(_view);
            if (controller.IsAuthorized)
            {
                _view.group3.Visible = true;
                _view.group2.Visible = true;
                _view.group1.Visible = true;
            }
            else
            {
                _view.group3.Visible = true;
            }


        }


        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {

            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if(mailItem.ReceivedByEntryID == null)
                {
                    new MailAttachementHandler(mailItem);
                }
            }
        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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

using Microsoft.Office.Tools.Ribbon;

namespace OutlookAddIn1
{
    public partial class Ribbon_
    {
      
        private void btn_source_Click(object sender, RibbonControlEventArgs e)
        {
          Controller controller = new Controller(this);
            controller.SaveADminListPath(); 
        }

        private void btn_ARAApproval_Click(object sender, RibbonControlEventArgs e)
        {
            Controller controller = new Controller(this);
            controller.SaveEmails("ARA Approval");
        }

        private void btn_LegalApproval_Click(object sender, RibbonControlEventArgs e)
        {
            Controller controller = new Controller(this);
            controller.SaveEmails("Legal Approval");
        }

        private void btn_LOE_Click(object sender, RibbonControlEventArgs e)
        {
            Controller controller = new Controller(this);
            controller.SaveEmails("LOE");
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
           Controller controller = new Controller(this);
            controller.SaveEmails("High Risk Approval");
        }

        private void btn_POD_Click(object sender, RibbonControlEventArgs e)
        {
            Controller controller = new Controller(this);
            controller.SaveEmails("POD");
        }

       
    }
}

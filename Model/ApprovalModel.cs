using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Windows.Forms;

namespace OutlookAddIn1.Model
{
    public class ApprovalModel
    {

        private readonly Ribbon_ view;
        private readonly ExcelModel model;
        private int HKFolderPathCol = 50;
        private int ECFolderPathCol = 48 ;
        private int ECContractPathCol = 49;

        public string cbo_email { get; set; }
        public bool isLOEStage { get; set; }
        public bool isJobStage { get; set; }
        Microsoft.Office.Interop.Excel.Workbook workbook;
        Microsoft.Office.Interop.Excel.Range searchRange;

        public ApprovalModel(Ribbon_ view, ExcelModel model)
        {
            this.view = view;
            this.model = model; 
        }
        public void HandleARAApproval(MailItem mailItem, string QuoteNumber,string folderPath)
        {

            string compliancePath = Path.Combine(folderPath, "4.Compliance");
      
            CreateDirectory(compliancePath);

            this.model.UpdateExcel(this.model.OpenWorkbook(), folderPath, QuoteNumber);


            mailItem.SaveAs(compliancePath + @"\" + this.cbo_email + "_" + QuoteNumber + ".msg", OlSaveAsType.olMSG);

            MessageBox.Show("Done");
        }

 
        public void BaseApproval(MailItem mailItem, string QuoteNumber, string Task)
        {

            int PathRow = GetSavePathRow(mailItem, QuoteNumber);

            string folderPath = workbook.Sheets["SH"].Cells[PathRow, HKFolderPathCol].Value;
            string EcPath_ = workbook.Sheets["SH"].Cells[PathRow, getQuoteNumber(this.cbo_email)].Value;
            //Get The Path 

            SaveEmails(folderPath, Task, QuoteNumber, mailItem);
            SaveEmails(EcPath_, Task, QuoteNumber, mailItem);


            if (!String.IsNullOrEmpty(workbook.Sheets["SH"].Cells[PathRow, 5].value) && isLOEStage == true)
            {
                string destination = this.model.ChangeProposalNumber(folderPath, QuoteNumber, searchRange.Offset[0,1]);
                workbook.Sheets["SH"].Cells[PathRow, HKFolderPathCol].Value = destination;
                Directory.Move(folderPath, destination);
            }

            MessageBox.Show("Done");

        }

        private int GetSavePathRow(MailItem mailItem, string QuoteNumber) {


            //Read from Excel 
            workbook = this.model.OpenWorkbook();
            int lastRow = workbook.Sheets["SH"].Cells[workbook.Sheets["SH"].Rows.Count, "D"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;


            if (isJobStage == false){
                searchRange = this.model.FindQuoteNumber(workbook.Sheets["SH"].Range["D3:D" + lastRow], QuoteNumber);
            }
            else
            {
                searchRange = this.model.FindQuoteNumber(workbook.Sheets["SH"].Range["D3:D" + lastRow].offset[0,1], QuoteNumber);
            }

            return searchRange.Row; 
        }


        private int getQuoteNumber(string Task)
        {

            switch (Task)
            {
                case "ARA Approval":case "Legal Approval":case "LOE":
                    return ECFolderPathCol;
 
                case "High Risk Approval":case "POD":
                    return ECContractPathCol;

                default:
                    return ECFolderPathCol;
            }
            
        }

        private void SaveEmails(string folderPath,string Task, string QuoteNumber, MailItem mailItem)
        {

            if (folderPath != null)
            {
                string HKPath = Path.Combine(folderPath, Task);
                CreateDirectory(HKPath);
                mailItem.SaveAs(HKPath + @"\" + this.cbo_email + "_" + QuoteNumber + ".msg", OlSaveAsType.olMSG);//Save to HK 
            }

        }

        private void CreateDirectory(string Path)
        {
            if (!Directory.Exists(Path) && !string.IsNullOrEmpty(Path))
            {
                try
                {
                    Directory.CreateDirectory(Path);
                }
                catch
                {
                    MessageBox.Show("No access to path");
                }
            }
        }
    }
}

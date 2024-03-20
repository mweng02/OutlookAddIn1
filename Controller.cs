using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using OutlookAddIn1.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;


namespace OutlookAddIn1
{
    public class Controller
    {

        private readonly Ribbon_ view;
        private readonly ExcelModel model;
        private readonly ApprovalModel approval;
        private readonly AuthorizedService authorizedService; 
      
        private bool isAuthorized;
        public bool IsAuthorized 
        {   get { return isAuthorized; }
            set { isAuthorized = value; } 
        }

        public Controller(Ribbon_ view)
        {

            this.view = view;
            this.model = new ExcelModel();
            this.authorizedService = new AuthorizedService();
            this.isAuthorized = this.authorizedService.CheckAuthorized(); 
            this.approval = new ApprovalModel(this.view, this.model);

        }
        public void SaveEmails(string Type)
        {
            #region Save Email 

            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            if (explorer != null)
            {
                Selection selection = explorer.Selection;

                if (selection != null && selection.Count > 0)
                {
                    object item = selection[1];
                    if (item is MailItem)
                    {
                        MailItem mailItem = (MailItem)item;
                        this.approval.cbo_email = Type;
           
                        string quoteNumber = getQuoteNumber(mailItem, this.approval.cbo_email);
                        if (string.IsNullOrEmpty(quoteNumber))
                        {
                            MessageBox.Show("Unable to find Proposal Number/Job Number.Please Check");
                            return;
                        }

                        string folderPath = string.Empty;

                        if (Type == "ARA Approval")
                        {

                            folderPath = UtilsModel.ExtractRegex(mailItem.Body, @"(?<=file:///)(.*)(?=>)|(?<=file:)(.*)(?=>)");
                            folderPath = folderPath.Trim().Replace(@"//ap.cbre.net/DFS", "N:").Replace(@"/", @"\").Replace("%20", " ");

                            this.approval.HandleARAApproval(mailItem, quoteNumber, folderPath);

                        }
                        else if (Type == "Legal Approval")
                        {
                            //this.approval.HandleLegalApproval(mailItem,quoteNumber);
                            this.approval.BaseApproval(mailItem, quoteNumber, "2.LOE");

                        }
                        else if (Type == "LOE")
                        {
                            this.approval.isLOEStage = true;
                            this.approval.BaseApproval(mailItem, quoteNumber, "2.LOE");
                        }
                        else if (Type == "High Risk Approval")
                        {
                            this.approval.isJobStage = true;
                            this.approval.BaseApproval(mailItem, quoteNumber, "4.Compliance");
                        }

                        else
                        {
                            this.approval.isJobStage = true;
                            this.approval.BaseApproval(mailItem, quoteNumber, "8.Invoice_POD");
                        }

                    }
                }

            }
            #endregion
        }

        private string getQuoteNumber(MailItem mailItem, string Task)
        {

            switch (Task)
            {
                case "ARA Approval":
                    return UtilsModel.ExtractRegex(mailItem.Body, @"Q\d{4}-\d{5}-[A-Z]{2}");

                case "Legal Approval":
                    return UtilsModel.ExtractRegex(mailItem.Subject, @"(Q|C)\d{4}-\d{4,}-[A-Z]{2}");

                case "LOE":
                    return UtilsModel.ExtractRegex(mailItem.Subject, @"(Q|C)\d{4}-\d{4,}-[A-Z]{2}");
                case "High Risk Approval":
                    return UtilsModel.ExtractRegex(mailItem.Subject, @"(Q|C)\d{4}-\d{4,}-[A-Z]{2}");
                case "POD":
                    return UtilsModel.ExtractRegex(mailItem.Subject, @"(Q|C)\d{4}-\d{4,}-[A-Z]{2}");
                default:
                    return UtilsModel.ExtractRegex(mailItem.Subject, @"(Q|C)\d{4}-\d{4,}-[A-Z]{2}");
            }


        }

        public void SaveADminListPath()
        {
            #region Save the Source: AdminList Path
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All Files|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFilePath = openFileDialog.FileName;
                Dictionary<string, string> jsonString = new Dictionary<string, string> { { "FilePath", $"{selectedFilePath}" } };
                string jsonString_ = JsonConvert.SerializeObject(jsonString);
                File.WriteAllText(this.model.jsonFilePath, jsonString_);

            }
            #endregion
        }
    }

    public class AuthorizedService{

        private Dictionary<string, string> _cacheJsonContent = null;
        private static Dictionary<string, string> _checkAuthorized = null;
        private string DirectoryPathSource = "CBRE_China_Tool_settings.json";
        private string ReplaceTerm = "List_Check.json";
        private string CheckTerm = "ChinaRetail.accdb"; 

        public bool CheckAuthorized()
        {
            var authorizedUsers = GetJsonContent()["AuthorizedUsers"];
            return authorizedUsers.Contains(Environment.UserName);
        }

        public dynamic GetJsonContent()
        {
            if (_checkAuthorized == null)
            {
                LoadJsonContent();
            }

            return _checkAuthorized;
        }

        private void LoadJsonContent()
        {
            if (_cacheJsonContent == null)
            {
                string jsonFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Excel\XlStart\Data", DirectoryPathSource);
                _cacheJsonContent = LoadJsonFromFile(jsonFilePath);
            }

            string transactionDataPath = _cacheJsonContent["transaction_data_filepath"].Replace(CheckTerm, ReplaceTerm);
            _checkAuthorized = LoadJsonFromFile(transactionDataPath);
        }

        private Dictionary<string, string> LoadJsonFromFile(string FilePath)
        {
            string jsonString = File.ReadAllText(FilePath);
            return JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);
        }
    }
}

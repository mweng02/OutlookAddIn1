
namespace OutlookAddIn1
{
    partial class Ribbon_ : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
           
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_ARAApproval = this.Factory.CreateRibbonButton();
            this.btn_LegalApproval = this.Factory.CreateRibbonButton();
            this.btn_LOE = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btn_POD = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_source = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "CBRE GC VAS";
            this.tab1.Name = "tab1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.label1);
            this.group3.Label = "Menu";
            this.group3.Name = "group3";
            this.group3.Visible = false;
            // 
            // label1
            // 
            this.label1.Label = "Working in Progress";
            this.label1.Name = "label1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_ARAApproval);
            this.group1.Items.Add(this.btn_LegalApproval);
            this.group1.Items.Add(this.btn_LOE);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.btn_POD);
            this.group1.Label = "Save Emails";
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // btn_ARAApproval
            // 
            this.btn_ARAApproval.Label = "ARA Approval";
            this.btn_ARAApproval.Name = "btn_ARAApproval";
            this.btn_ARAApproval.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ARAApproval_Click);
            // 
            // btn_LegalApproval
            // 
            this.btn_LegalApproval.Label = "Legal Approval";
            this.btn_LegalApproval.Name = "btn_LegalApproval";
            this.btn_LegalApproval.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LegalApproval_Click);
            // 
            // btn_LOE
            // 
            this.btn_LOE.Label = "LOE";
            this.btn_LOE.Name = "btn_LOE";
            this.btn_LOE.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LOE_Click);
            // 
            // button1
            // 
            this.button1.Label = "High Risk Approval";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btn_POD
            // 
            this.btn_POD.Label = "POD";
            this.btn_POD.Name = "btn_POD";
            this.btn_POD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_POD_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_source);
            this.group2.Label = "Admin";
            this.group2.Name = "group2";
            this.group2.Visible = false;
            // 
            // btn_source
            // 
            this.btn_source.Label = "Load Admin Source";
            this.btn_source.Name = "btn_source";
            this.btn_source.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_source_Click);
            // 
            // Ribbon_
            // 
            this.Name = "Ribbon_";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mai" +
    "l.Read";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_source;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ARAApproval;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LegalApproval;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LOE;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_POD;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_ Ribbon1
        {
            get { return this.GetRibbon<Ribbon_>(); }
        }
    }
}

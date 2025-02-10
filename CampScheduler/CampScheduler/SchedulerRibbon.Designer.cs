namespace CampScheduler
{
    partial class SchedulerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SchedulerRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.GenerateInputButton = this.Factory.CreateRibbonDropDown();
            this.GenerateEmptyInputButton = this.Factory.CreateRibbonButton();
            this.GenerateExampleInputButton = this.Factory.CreateRibbonButton();
            this.OpenInputButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.GenerateOutputButton = this.Factory.CreateRibbonButton();
            this.OpenInputFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Camp Scheduler";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.GenerateInputButton);
            this.group1.Items.Add(this.OpenInputButton);
            this.group1.Label = "Input Tools";
            this.group1.Name = "group1";
            // 
            // GenerateInputButton
            // 
            this.GenerateInputButton.Buttons.Add(this.GenerateEmptyInputButton);
            this.GenerateInputButton.Buttons.Add(this.GenerateExampleInputButton);
            this.GenerateInputButton.Label = "Generate Input File";
            this.GenerateInputButton.Name = "GenerateInputButton";
            this.GenerateInputButton.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateInputButton_SelectionChanged);
            // 
            // GenerateEmptyInputButton
            // 
            this.GenerateEmptyInputButton.Label = "Empty";
            this.GenerateEmptyInputButton.Name = "GenerateEmptyInputButton";
            // 
            // GenerateExampleInputButton
            // 
            this.GenerateExampleInputButton.Label = "Example";
            this.GenerateExampleInputButton.Name = "GenerateExampleInputButton";
            // 
            // OpenInputButton
            // 
            this.OpenInputButton.Label = "Open Input File";
            this.OpenInputButton.Name = "OpenInputButton";
            this.OpenInputButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenInputButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.GenerateOutputButton);
            this.group2.Label = "Output";
            this.group2.Name = "group2";
            // 
            // GenerateOutputButton
            // 
            this.GenerateOutputButton.Label = "Generate Output";
            this.GenerateOutputButton.Name = "GenerateOutputButton";
            this.GenerateOutputButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateOutputButton_Click);
            // 
            // OpenInputFileDialog
            // 
            this.OpenInputFileDialog.FileName = "Scheduler_Parameter_File";
            // 
            // SchedulerRibbon
            // 
            this.Name = "SchedulerRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SchedulerRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown GenerateInputButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton GenerateEmptyInputButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton GenerateExampleInputButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenInputButton;
        private System.Windows.Forms.OpenFileDialog OpenInputFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateOutputButton;
    }

    partial class ThisRibbonCollection
    {
        internal SchedulerRibbon SchedulerRibbon
        {
            get { return this.GetRibbon<SchedulerRibbon>(); }
        }
    }
}

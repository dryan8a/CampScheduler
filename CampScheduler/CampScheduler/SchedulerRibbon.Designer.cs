﻿namespace CampScheduler
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
            this.GenerateExampleInput2Button = this.Factory.CreateRibbonButton();
            this.GenerateEmptyWeekButton = this.Factory.CreateRibbonButton();
            this.GenerateExampleWeekButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.GenerateDayOutputButton = this.Factory.CreateRibbonButton();
            this.GenerateWeekOutputButton = this.Factory.CreateRibbonButton();
            this.OpenInputFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.FormatOutputDropDown = this.Factory.CreateRibbonDropDown();
            this.FormatOutputButton = this.Factory.CreateRibbonButton();
            this.UnFormatOutputButton = this.Factory.CreateRibbonButton();
            this.DoTallyButton = this.Factory.CreateRibbonCheckBox();
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
            this.group1.Label = "Input";
            this.group1.Name = "group1";
            // 
            // GenerateInputButton
            // 
            this.GenerateInputButton.Buttons.Add(this.GenerateEmptyInputButton);
            this.GenerateInputButton.Buttons.Add(this.GenerateExampleInputButton);
            this.GenerateInputButton.Buttons.Add(this.GenerateExampleInput2Button);
            this.GenerateInputButton.Buttons.Add(this.GenerateEmptyWeekButton);
            this.GenerateInputButton.Buttons.Add(this.GenerateExampleWeekButton);
            this.GenerateInputButton.Label = "Generate Input File";
            this.GenerateInputButton.Name = "GenerateInputButton";
            this.GenerateInputButton.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateInputButton_SelectionChanged);
            // 
            // GenerateEmptyInputButton
            // 
            this.GenerateEmptyInputButton.Label = "Empty Day";
            this.GenerateEmptyInputButton.Name = "GenerateEmptyInputButton";
            // 
            // GenerateExampleInputButton
            // 
            this.GenerateExampleInputButton.Label = "Example Day 1";
            this.GenerateExampleInputButton.Name = "GenerateExampleInputButton";
            // 
            // GenerateExampleInput2Button
            // 
            this.GenerateExampleInput2Button.Label = "Example Day 2";
            this.GenerateExampleInput2Button.Name = "GenerateExampleInput2Button";
            // 
            // GenerateEmptyWeekButton
            // 
            this.GenerateEmptyWeekButton.Label = "Empty Week";
            this.GenerateEmptyWeekButton.Name = "GenerateEmptyWeekButton";
            // 
            // GenerateExampleWeekButton
            // 
            this.GenerateExampleWeekButton.Label = "Example Week";
            this.GenerateExampleWeekButton.Name = "GenerateExampleWeekButton";
            // 
            // group2
            // 
            this.group2.Items.Add(this.GenerateDayOutputButton);
            this.group2.Items.Add(this.GenerateWeekOutputButton);
            this.group2.Items.Add(this.FormatOutputDropDown);
            this.group2.Items.Add(this.DoTallyButton);
            this.group2.Label = "Output";
            this.group2.Name = "group2";
            // 
            // GenerateDayOutputButton
            // 
            this.GenerateDayOutputButton.Label = "Generate Output (Day)";
            this.GenerateDayOutputButton.Name = "GenerateDayOutputButton";
            this.GenerateDayOutputButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateDayOutputButton_Click);
            // 
            // GenerateWeekOutputButton
            // 
            this.GenerateWeekOutputButton.Label = "Generate Output (Week)";
            this.GenerateWeekOutputButton.Name = "GenerateWeekOutputButton";
            this.GenerateWeekOutputButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateWeekOutputButton_Click);
            // 
            // OpenInputFileDialog
            // 
            this.OpenInputFileDialog.FileName = "Scheduler_Parameter_File";
            // 
            // FormatOutputDropDown
            // 
            this.FormatOutputDropDown.Buttons.Add(this.FormatOutputButton);
            this.FormatOutputDropDown.Buttons.Add(this.UnFormatOutputButton);
            this.FormatOutputDropDown.Label = "Format Options";
            this.FormatOutputDropDown.Name = "FormatOutputDropDown";
            // 
            // FormatOutputButton
            // 
            this.FormatOutputButton.Label = "Format";
            this.FormatOutputButton.Name = "FormatOutputButton";
            // 
            // UnFormatOutputButton
            // 
            this.UnFormatOutputButton.Label = "Undo Format";
            this.UnFormatOutputButton.Name = "UnFormatOutputButton";
            // 
            // DoTallyButton
            // 
            this.DoTallyButton.Checked = true;
            this.DoTallyButton.Label = "Schedule With Tally";
            this.DoTallyButton.Name = "DoTallyButton";
            this.DoTallyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DoTallyButton_Click);
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
        private System.Windows.Forms.OpenFileDialog OpenInputFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateDayOutputButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton GenerateEmptyWeekButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton GenerateExampleWeekButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateWeekOutputButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton GenerateExampleInput2Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown FormatOutputDropDown;
        private Microsoft.Office.Tools.Ribbon.RibbonButton FormatOutputButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton UnFormatOutputButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox DoTallyButton;
    }

    partial class ThisRibbonCollection
    {
        internal SchedulerRibbon SchedulerRibbon
        {
            get { return this.GetRibbon<SchedulerRibbon>(); }
        }
    }
}

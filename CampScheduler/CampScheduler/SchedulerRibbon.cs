using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace CampScheduler
{
    public partial class SchedulerRibbon
    {
        private void SchedulerRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            GenerateEmptyInputButton.Click += GenerateEmptyInputButton_Click;
            GenerateExampleInputButton.Click += GenerateExampleInputButton_Click;

        }

        private void GenerateInputButton_SelectionChanged(object sender, RibbonControlEventArgs e)
        {   
        }

        private void GenerateEmptyInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var emptyInputSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            emptyInputSheet.Range["A1"].Value2 = "This is a generated empty input for the scheduler!";
        }

        private void GenerateExampleInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var exampleInputSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            exampleInputSheet.Range["A1"].Value2 = "This is a generated example input for the scheduler!";
        }

        private void GenerateOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var inputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            int blockBottom = 2;
            while (inputSheet.Range["A" + ++blockBottom].Value2 != null) ; 
            var blockData = inputSheet.Range["A2", "H" + (blockBottom - 1)];

            int activityBottom = 2;
            while (inputSheet.Range["J" + ++activityBottom].Value2 != null) ;
            var activityData = inputSheet.Range["J2", "Q" + (activityBottom - 1)];

            int groupBottom = 2;
            while (inputSheet.Range["S" + ++groupBottom].Value2 != null) ;
            var groupData = inputSheet.Range["S2", "X" + (groupBottom - 1)]; 

            Schedule schedule;
            //error handling commented out for testing purposes
            try
            {
                schedule = Schedule.GenerateSchedule(blockData, activityData, groupData);
            }
            catch (Exception ex)
            {
                var errorSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                errorSheet.Range["A1"].Value2 = "An Error occured while generating schedule:";
                errorSheet.Range["A2"].Value2 = ex.Message;
                return;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inputSheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(blockData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(activityData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(groupData);


            var outputSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            schedule.OutputSchedule(outputSheet.Range["A1","Z100"]);
        }

        private void OpenInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            OpenInputFileDialog.ShowDialog();
        }
    }
}

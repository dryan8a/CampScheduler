using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
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
            GenerateEmptyWeekButton.Click += GenerateEmptyWeekButton_Click;
            GenerateExampleWeekButton.Click += GenerateExampleWeekButton_Click;
        }
        private void GenerateInputButton_SelectionChanged(object sender, RibbonControlEventArgs e)
        {   
        }

        private void GenerateEmptyInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated empty input for the scheduler!";
            
            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];
            
            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[2]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);
        }

        private void GenerateExampleInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated example input for the scheduler!";

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[1]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);
        }

        private void GenerateExampleWeekButton_Click(object sender, RibbonControlEventArgs e)
        {
            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated empty input for the scheduler!";

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[2]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);
        }

        private void GenerateEmptyWeekButton_Click(object sender, RibbonControlEventArgs e)
        {
            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated empty input for the scheduler!";

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[2]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);
        }


        private void GenerateDayOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var inputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            int blockBottom = 3;
            while (inputSheet.Range["A" + ++blockBottom].Value2 != null) ; 
            var blockData = inputSheet.Range["A3", "H" + (blockBottom - 1)];

            int activityBottom = 3;
            while (inputSheet.Range["J" + ++activityBottom].Value2 != null) ;
            var activityData = inputSheet.Range["J3", "Q" + (activityBottom - 1)];

            int groupBottom = 3;
            while (inputSheet.Range["S" + ++groupBottom].Value2 != null) ;
            var groupData = inputSheet.Range["S3", "W" + (groupBottom - 1)];

            int rulesBottom = 3;
            while (inputSheet.Range["Y" + ++rulesBottom].Value2 != null) ;
            var rulesData = inputSheet.Range["Y3", "AA" + (rulesBottom - 1)];

            DaySchedule schedule;
            //error handling commented out for testing purposes
            try
            {
                schedule = SchedulerParser.GenerateDaySchedule(blockData, activityData, groupData, rulesData);
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

        private void GenerateWeekOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var inputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            int blockBottom = 3;
            while (inputSheet.Range["A" + ++blockBottom].Value2 != null) ;
            var blockData = inputSheet.Range["A3", "I" + (blockBottom - 1)];

            int activityBottom = 3;
            while (inputSheet.Range["K" + ++activityBottom].Value2 != null) ;
            var activityData = inputSheet.Range["K3", "R" + (activityBottom - 1)];

            int groupBottom = 3;
            while (inputSheet.Range["T" + ++groupBottom].Value2 != null) ;
            var groupData = inputSheet.Range["T3", "X" + (groupBottom - 1)];

            int rulesBottom = 3;
            while (inputSheet.Range["Z" + ++rulesBottom].Value2 != null) ;
            var rulesData = inputSheet.Range["Z3", "AC" + (rulesBottom - 1)];

            //var errorSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            //errorSheet.Range["A1"].Value2 = "Week Generation Not Available. Launching soon.";

            WeekSchedule schedule;

            //error handling commented out for testing purposes
            //try
            //{
                schedule = SchedulerParser.GenerateWeekSchedule(blockData, activityData, groupData, rulesData);
            //}
            //catch (Exception ex)
            //{
            //    var errorSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            //    errorSheet.Range["A1"].Value2 = "An Error occured while generating schedule:";
            //    errorSheet.Range["A2"].Value2 = ex.Message;
            //    return;
            //}

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inputSheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(blockData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(activityData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(groupData);

            var outputRanges = new Excel.Range[schedule.NumOfDays];

            for(int i = 0; i < outputRanges.Length; i++)
            {
                var outputSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                outputRanges[i] = outputSheet.Range["A1", "Z100"];
            }
            
            schedule.OutputSchedule(outputRanges);
        }

        private void FormatOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            var outputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            int columnsWidth = -1;
            while (outputSheet.Range[(char)('A' + ++columnsWidth) + "4"].Value2 != null)
            {
                outputSheet.Range[(char)('A' + columnsWidth) + "4"].ColumnWidth = 22;
            }

            int rows = 3;
            outputSheet.Range["A1","A3"].RowHeight = 20;

            bool isColorRow = false;
            while (outputSheet.Range["A" + ++rows].Value2 != null)
            {
                outputSheet.Range["A" + rows].RowHeight = 46;

                if(isColorRow) outputSheet.Range["A" + rows, ((char)('A' + columnsWidth - 1)).ToString() + rows].Interior.Color = Excel.XlRgbColor.rgbLightGrey;
                
                isColorRow = !isColorRow;
            }

            var outputRange = outputSheet.Range["A1", ((char)('A' + columnsWidth)).ToString() + (rows - 1)];

            outputRange.Cells.Font.Name = "Arial";

            var firstColRange = outputRange.Range["A1", "A" + (rows - 1)];
            firstColRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            outputSheet.PageSetup.Zoom = false;

        }
    }
}

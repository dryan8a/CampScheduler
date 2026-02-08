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
using System.Windows.Forms.VisualStyles;
using Microsoft.Office.Core;

namespace CampScheduler
{
    public partial class SchedulerRibbon
    {
        public Dictionary<string, int> sheetNameToIndex;

        public Excel.Workbook UserIntendedWorkbook;

        private void SchedulerRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            GenerateEmptyInputButton.Click += GenerateEmptyInputButton_Click;
            GenerateExampleInputButton.Click += GenerateExampleInputButton_Click;
            GenerateExampleInput2Button.Click += GenerateExampleInput2Button_Click;
            GenerateEmptyWeekButton.Click += GenerateEmptyWeekButton_Click;
            GenerateExampleWeekButton.Click += GenerateExampleWeekButton_Click;
            GenerateEmptyBumpButton.Click += GenerateEmptyBumpButton_Click;
            GenerateExampleBumpButton.Click += GenerateExampleBumpButton_Click;
            

            FormatOutputButton.Click += FormatOutputButton_Click;
            FormatBumpButton.Click += FormatBumpButton_Click;
            UnFormatOutputButton.Click += UnFormatOutputButton_Click;

            //TallyNameBox.TextChanged += TallyNameBox_TextChanged;

            Globals.ThisAddIn.Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            Globals.ThisAddIn.Application.SheetBeforeDelete += Application_SheetBeforeDelete;

            sheetNameToIndex = new Dictionary<string, int>
            {
                { "", 0 }
            };

            UserIntendedWorkbook = null;

           // foreach(string SheetName in GetWorksheetsNames()) AddWorkbookSheet(SheetName);
        }

        private void GenerateInputButton_SelectionChanged(object sender, RibbonControlEventArgs e)
        {   
        }

        private void GenerateEmptyInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated empty input for the scheduler!";
            
            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];
            
            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[1]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);

            AddWorksheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Name);
        }

        private void GenerateExampleInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated example input for the scheduler!";

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[2]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);

            AddWorksheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Name);
        }

        private void GenerateExampleInput2Button_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated example input for the scheduler!";

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[3]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);

            AddWorksheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Name);
        }

        private void GenerateExampleWeekButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated empty input for the scheduler!";

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[5]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);

            AddWorksheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Name);
        }

        private void GenerateEmptyWeekButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //emptyInputSheet.Range["A1"].Value2 = "This is a generated empty input for the scheduler!";

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[4]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);

            AddWorksheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Name);
        }

        private void GenerateEmptyBumpButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[6]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);

            AddWorksheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Name);
        }

        private void GenerateExampleBumpButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            var currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            Globals.ThisAddIn.Application.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "CampSchedulerInputExamples.xlsx");

            var exampleWorkbookIndex = Globals.ThisAddIn.Application.Workbooks.Count; //THIS IS NOT THREAD SAFE, PLEASE PROGRAMATICALLY OPEN ANY OTHER WORKBOOKS (luckily it should hopefully just crash and not harm data)
            var exampleWorkbook = Globals.ThisAddIn.Application.Workbooks[exampleWorkbookIndex];

            exampleWorkbook.Windows[1].Visible = false;
            ((Excel.Worksheet)exampleWorkbook.Worksheets[7]).Copy(Type.Missing, currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);

            exampleWorkbook.Close(false);

            AddWorksheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Name);
        }


        private string[] GetWorksheetsNames()
        {
            var workSheetNames = new string[Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count];
            for(int i = 1; i <= workSheetNames.Length; i++)
            {
                workSheetNames[i-1] = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[i].Name;
            }
            return workSheetNames;
        }

        private void AddWorksheet(string sheetName)
        {
            RibbonDropDownItem newItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            newItem.Label = sheetName;
            sheetNameToIndex.Add(sheetName, TallyInputBox.Items.Count);
            TallyInputBox.Items.Add(newItem);
        }

        private void Application_WorkbookNewSheet(Excel.Workbook Wb, object Sh)
        {
            if (UserIntendedWorkbook != null && UserIntendedWorkbook.Name != Wb.Name) return;

            AddWorksheet(((Excel.Worksheet)Sh).Name);
        }

        private void Application_SheetBeforeDelete(object Sh)
        {
            for (int i = 1; i < TallyInputBox.Items.Count; i++)
            {
                if(TallyInputBox.Items[i].Label == ((Excel.Worksheet)Sh).Name) 
                {
                    sheetNameToIndex.Remove(TallyInputBox.Items[i].Label);
                    TallyInputBox.Items.RemoveAt(i);
                    break;
                }
            }
        }

        private void UpdateLastListSheetName(string sheetName)
        {
            sheetNameToIndex.Remove(TallyInputBox.Items.Last().Label);
            sheetNameToIndex.Add(sheetName, TallyInputBox.Items.Count() - 1);
            TallyInputBox.Items.Last().Label = sheetName;
        }

        private void UpdateLastListSheetNames(string[] sheetNames)
        {
            int sheetNameI = 0;
            for (int i = TallyInputBox.Items.Count - sheetNames.Length; i < TallyInputBox.Items.Count; i++)
            {
                sheetNameToIndex.Remove(TallyInputBox.Items[i].Label);
                sheetNameToIndex.Add(sheetNames[sheetNameI], i);
                TallyInputBox.Items[i].Label = sheetNames[sheetNameI];
                sheetNameI++;
            }
        }

        private void RefreshTallyInput()
        {
            var sheetNames = GetWorksheetsNames();

            for(int i = 1; i < TallyInputBox.Items.Count; i++)
            {
                if (!sheetNames.Contains(TallyInputBox.Items[i].Label))
                {
                    sheetNameToIndex.Remove(TallyInputBox.Items[i].Label);
                    TallyInputBox.Items.RemoveAt(i);
                }
            }

            foreach (var sheetName in sheetNames)
            {
                if(!sheetNameToIndex.ContainsKey(sheetName))
                {
                    AddWorksheet(sheetName);
                }
            }
        }

        private void RefreshTallyInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            RefreshTallyInput();
            TallyInputBox.Text = "";
        }

        private Excel.Range GetTallyInput()
        {
            if (TallyInputBox.Text == "") return ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range["A1"];

            var tallySheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Cast<Excel.Worksheet>().SingleOrDefault(w => w.Name == TallyInputBox.Text);

            if(tallySheet == null) return ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range["A1"]; ;

            //int tallyBottom = 1;
            //while (tallySheet.Range["B" + ++tallyBottom].Value2 != null) ;

            //int tallyRight = tallySheet.UsedRange.Columns.Count;
            //while (tallySheet.Range[(char)('B' + ++tallyRight) + "1"].Value2 != null);

            return tallySheet.UsedRange;//tallySheet.Range["A1", (char)('B' + tallyRight) + tallyBottom];
        }

        private void GenerateDayOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1,18,Type.Missing,Type.Missing,true); 

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            UserIntendedWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

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

            var tallyData = GetTallyInput();

            DaySchedule schedule;

            try
            {
                schedule = SchedulerParser.GenerateDaySchedule(blockData, activityData, groupData, rulesData, tallyData);
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

            var takenNames = GetWorksheetsNames();

            var outputSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            var outputName = schedule.OutputSchedule(outputSheet, takenNames);

            UpdateLastListSheetName(outputName);

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheet);

            if (DoTallyButton.Checked)
            {
                var tallySheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                var TallyName = schedule.OutputTally(tallySheet, takenNames);

                UpdateLastListSheetName(TallyName);

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tallySheet);
            }

            if (GroupSchedulesBox.Checked)
            {
                Excel.Workbook GroupsWorkbook = (Excel.Workbook)Globals.ThisAddIn.Application.Workbooks.Add();

                var groupSheets = new Excel.Worksheet[schedule.NumOfGroups];

                for (int i = 0; i < groupSheets.Length; i++)
                {
                    groupSheets[i] = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                }

                schedule.OutputGroups(groupSheets, takenNames);

                ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[groupSheets.Length + 1]).Delete();

                for (int i = 0; i < groupSheets.Length; i++)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(groupSheets[i]);
                }

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(GroupsWorkbook);
            }

            UserIntendedWorkbook = null;
        }

        private void GenerateWeekOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            UserIntendedWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

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

            var tallyData = GetTallyInput();

            WeekSchedule schedule;

            try
            {
                schedule = SchedulerParser.GenerateWeekSchedule(blockData, activityData, groupData, rulesData,tallyData);
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
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rulesData);

            var outputSheets = new Excel.Worksheet[schedule.NumOfDays];

            for (int i = 0; i < outputSheets.Length; i++)
            {
                outputSheets[i] = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            }

            var takenNames = GetWorksheetsNames();

            var outputSheetNames = schedule.OutputSchedule(outputSheets, takenNames);

            UpdateLastListSheetNames(outputSheetNames);

            foreach(var outputSheet in outputSheets)
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheet);
            }


            //Output Tally
            if (DoTallyButton.Checked)
            {
                var tallySheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                var tallyName = schedule.OutputTally(tallySheet, takenNames);

                UpdateLastListSheetName(tallyName);

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tallySheet);
            }

            //Output Groups
            if (GroupSchedulesBox.Checked)
            {
                Excel.Workbook GroupsWorkbook = (Excel.Workbook)Globals.ThisAddIn.Application.Workbooks.Add();

                var groupSheets = new Excel.Worksheet[schedule.NumOfGroups];

                for (int i = groupSheets.Length-1; i >= 0; i--)
                {
                    groupSheets[i] = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                }

                var groupSheetNames = schedule.OutputGroups(groupSheets, takenNames);

                ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[groupSheets.Length+1]).Delete();

                for (int i = 0; i < groupSheets.Length; i++)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(groupSheets[i]);
                }

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(GroupsWorkbook);
            }

            UserIntendedWorkbook = null;
        }

        private void GenerateBumpButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            UserIntendedWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            var inputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            int blockBottom = 3;
            while (inputSheet.Range["A" + ++blockBottom].Value2 != null) ;
            var blockData = inputSheet.Range["A3", "C" + (blockBottom - 1)];

            int activityBottom = 3;
            while (inputSheet.Range["E" + ++activityBottom].Value2 != null) ;
            var activityData = inputSheet.Range["E3", "K" + (activityBottom - 1)];

            int counselorBottom = 3;
            while (inputSheet.Range["M" + ++counselorBottom].Value2 != null) ;
            var counselorData = inputSheet.Range["M3", "R" + (counselorBottom - 1)];

            Bump bump;
            
            try
            {
                bump = SchedulerParser.GenerateBump(blockData,activityData,counselorData);
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
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(counselorData);

            var takenNames = GetWorksheetsNames();

            var outputSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            string bumpName = bump.OutputBump(outputSheet, takenNames);

            UpdateLastListSheetName(bumpName);

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheet);

            UserIntendedWorkbook = null;
        }

        private void FormatOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            try
            {
                var outputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

                int columnsWidth = -1;
                while (outputSheet.Range[(char)('A' + ++columnsWidth) + "4"].Value2 != null)
                {
                    outputSheet.Range[(char)('A' + columnsWidth) + "4"].ColumnWidth = 22;
                }

                int rows = 3;
                outputSheet.Range["A1", "A3"].RowHeight = 20;

                bool isColorRow = false;
                while (outputSheet.Range["A" + ++rows].Value2 != null)
                {
                    outputSheet.Range["A" + rows].RowHeight = 46;

                    if (isColorRow) outputSheet.Range["A" + rows, ((char)('A' + columnsWidth - 1)).ToString() + rows].Interior.Color = Excel.XlRgbColor.rgbLightGrey;

                    isColorRow = !isColorRow;
                }

                var outputRange = outputSheet.Range["A1", ((char)('A' + columnsWidth)).ToString() + (rows - 1)];

                outputRange.Cells.Font.Name = "Arial";

                var firstColRange = outputRange.Range["A1", "A" + (rows - 1)];
                firstColRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                outputSheet.PageSetup.Zoom = false;

            }
            catch (Exception ex) { }
        }

        private void FormatBumpButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            try
            {
                var outputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

                outputSheet.Range["A1"].ColumnWidth = 9;
                outputSheet.Range["A1"].RowHeight = 30;

                int columnsWidth = -1;
                while (outputSheet.Range[(char)('B' + ++columnsWidth) + "3"].Value2 != null)
                {
                    outputSheet.Range[(char)('B' + columnsWidth) + "3"].ColumnWidth = 21.13;
                }

                int rows = 2;
                while (outputSheet.Range["B" + rows].Value2 != null)
                {
                    outputSheet.Range["B" + rows].Font.Underline = true;
                    outputSheet.Range["B" + rows].RowHeight = 30;

                    while (outputSheet.Range["B" + ++rows].Value2 != null)
                    {
                        outputSheet.Range["B" + rows, "B" + (rows + 3)].RowHeight = 30;

                        var headerRange = outputSheet.Range["B" + rows, ((char)('B' + columnsWidth - 1)).ToString() + rows];
                        headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGrey;
                        headerRange.Font.Bold = true;

                        rows += 2;
                        outputSheet.Range["B" + rows, ((char)('B' + columnsWidth - 1)).ToString() + rows].Font.Italic = true;
                    }
                    rows++;
                }

                var outputRange = outputSheet.Range["B2", ((char)('B' + columnsWidth)).ToString() + (rows - 1)];

                outputRange.Cells.Font.Name = "Aptos Narrow";
   
                outputSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                outputSheet.PageSetup.TopMargin = 54;
                outputSheet.PageSetup.BottomMargin = 54;
                outputSheet.PageSetup.RightMargin = 18;
                outputSheet.PageSetup.LeftMargin = 18;
                outputSheet.PageSetup.HeaderMargin = 21.6;
                outputSheet.PageSetup.FooterMargin = 21.6;
                outputSheet.PageSetup.Zoom = false;
                outputSheet.PageSetup.FitToPagesTall = false;
                outputSheet.PageSetup.FitToPagesWide = 1;
                //outputSheet.Columns.AutoFit();

            }
            catch (Exception ex) { }
        }

        private void UnFormatOutputButton_Click(object sender, RibbonControlEventArgs e)
        {
            CommandBarControl oNewMenu = Globals.ThisAddIn.Application.CommandBars["Worksheet Menu Bar"].FindControl(1, 18, Type.Missing, Type.Missing, true);

            if (oNewMenu != null)
            {
                if (!oNewMenu.Enabled)
                {
                    return;
                }
            }

            try
            {
                var outputSheet = Globals.ThisAddIn.GetActiveWorkSheet();

                int columnsWidth = -1;
                do { columnsWidth++; } while (outputSheet.Range[(char)('A' + columnsWidth) + "4"].Value2 != null || outputSheet.Range[(char)('B' + columnsWidth) + "3"].Value2 != null);

                int rows = 3;

                do { rows++; } while (outputSheet.Range["A" + rows].Value2 != null || outputSheet.Range["B" + rows].Value2 != null || outputSheet.Range["B" + (rows + 1)].Value2 != null);

                var outputRange = outputSheet.Range["A1", ((char)('A' + columnsWidth)).ToString() + (rows - 1)];
                outputRange.ColumnWidth = 16;
                outputRange.RowHeight = 14.3;
                outputRange.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;

                outputRange.Columns.AutoFit();
                outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                outputRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                outputRange.Cells.Font.Name = "Aptos Narrow";
                outputRange.Font.Bold = false;
                outputRange.Font.Italic = false;
                outputRange.Font.Underline = false;


                var firstColRange = outputRange.Range["A1", "A" + (rows - 1)];
                firstColRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                outputSheet.PageSetup.Orientation = XlPageOrientation.xlPortrait;
                outputSheet.PageSetup.TopMargin = 54;
                outputSheet.PageSetup.BottomMargin = 54;
                outputSheet.PageSetup.RightMargin = 50.4;
                outputSheet.PageSetup.LeftMargin = 50.4;
                outputSheet.PageSetup.HeaderMargin = 21.6;
                outputSheet.PageSetup.FooterMargin = 21.6;
                outputSheet.PageSetup.FitToPagesTall = 1;
                outputSheet.PageSetup.FitToPagesWide = 1;
            }
            catch (Exception ex) { }
        }

        private void DoTallyButton_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void GroupSchedulesBox_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void FormatOutputDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

    }
}

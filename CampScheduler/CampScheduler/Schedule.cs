using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace CampScheduler
{
    public enum Grade
    {
        PK,
        K,
        First,
        Second,
        Third,
        Fourth,
        Fifth,
        Sixth,
        MS,
        NA
    }

    public struct Activity
    {
        public byte Id { get; }
        public string Name { get; }
        public bool WaterActivity { get; }
        public bool Overflow { get; }
        public byte NumofGroups { get; }
        public bool Open { get; }
        public Grade[] GradeOnly { get; }
        public Grade[] GradeStrike { get; }
        public bool IsSpecialist { get; }

        public Activity(byte id,string name, bool waterActivity, bool overflow, byte numOfGroups, bool open, Grade[] gradeOnly, Grade[] gradeStrike, bool specialist)
        {
            Id = id;
            Name = name;
            WaterActivity = waterActivity;
            Overflow = overflow;
            NumofGroups = numOfGroups;
            Open = open;
            GradeOnly = gradeOnly;
            GradeStrike = gradeStrike;
            IsSpecialist = specialist;
        }
    }

    public struct Group
    {
        public byte RowNum { get; }
        public string Name { get; }
        public Grade Grade { get; }
        public byte Unit { get; }
        public bool SpecialGroup { get; }
        public byte LunchNum { get; }
        public byte WaterGrouping { get; }

        public Group(byte rowNum, string name, Grade grade, byte unit, bool specialGroup, byte lunchNum, byte waterGrouping)
        {
            RowNum = rowNum;
            Name = name;
            Grade = grade;
            Unit = unit;
            SpecialGroup = specialGroup;
            LunchNum = lunchNum;
            WaterGrouping = waterGrouping;
        }
    }

    public struct SpecialActivityPrefs
    {
        public char OpeningCirclePref;
        public char MiddleCirclePref;
        public char PopsicleTimePref;
        public char ClosingCirclePref;
        public char OpenPref;
        public string SpecialEntPrefs;

        public SpecialActivityPrefs(char openingCirclePref, char middleCirclePref, char popsicleTimePref, char closingCirclePref, char openPref, string specialEntPrefs)
        {
            OpeningCirclePref = openingCirclePref;
            MiddleCirclePref = middleCirclePref;
            PopsicleTimePref = popsicleTimePref;
            ClosingCirclePref = closingCirclePref;
            OpenPref = openPref;
            SpecialEntPrefs = specialEntPrefs;
        }
    }

    public class Schedule
    {
        public string[,] ScheduleData; //change to internal
        internal List<Activity> Activities { get; }

        private Dictionary<byte, byte> LunchNumToTimeIndex;
        private SpecialActivityPrefs[] SpecActPrefs;

        internal Group[] Groups { get; }
        internal Dictionary<Grade, byte> GradeToUnit;
        internal Dictionary<byte, List<byte>> WaterGroupingToGroups;

        internal string[] Times { get; }

        private Random Gen;

        internal Schedule(int numOfBlocks, Group[] groups, string[] times, Dictionary<byte,byte> lunchNumToTimeIndex, Dictionary<Grade,byte> gradeToUnit, SpecialActivityPrefs[] specActPrefs, Dictionary<byte,List<byte>> waterGroupingToGroups)
        {
            ScheduleData = new string[numOfBlocks, groups.Length];

            Activities = new List<Activity>();
            //AddActivity("Lunch 1", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Lunch 2", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Lunch 3", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Opening Circle", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Middle Circle", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Closing Circle", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Popsicle Time", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Special Entertainment", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });
            //AddActivity("Open Activity", false, false, 1, false, new[] { Grade.NA }, new[] { Grade.NA });

            Groups = groups;

            LunchNumToTimeIndex = lunchNumToTimeIndex;
            Times = times;
            GradeToUnit = gradeToUnit;
            WaterGroupingToGroups = waterGroupingToGroups;

            SpecActPrefs = specActPrefs;

            Gen = new Random();
        }

        public void AddActivity(string name, bool waterActivity, bool overflow, byte numOfGroups, bool open, Grade[] gradeOnly, Grade[] gradeStrike, bool specialist)
        {
            Activities.Add(new Activity((byte)Activities.Count, name, waterActivity, overflow, numOfGroups, open, gradeOnly,gradeStrike,specialist));
        }

        public static Grade ParseGrade(string gradeInput)
        {
            switch(gradeInput)
            {
                case "PK":
                    return Grade.PK;
                case "K":
                    return Grade.K;
                case "1":
                    return Grade.First;
                case "2":
                    return Grade.Second;
                case "3":
                    return Grade.Third;
                case "4":
                    return Grade.Fourth;
                case "5":
                    return Grade.Fifth;
                case "6":
                    return Grade.Sixth;
                case "MS":
                    return Grade.MS;
                default:
                    return Grade.NA;
            }
        }

        public static Grade[] ParseGrades(string gradesInput)
        {
            var gradeStrings = gradesInput.Split(',');

            Grade[] grades = new Grade[gradeStrings.Length];

            for(int i = 0; i < grades.Length; i++)
            {
                grades[i] = ParseGrade(gradeStrings[i].Trim());
                if (grades[i] == Grade.NA) return new Grade[0];
            }

            return grades;
        }

        public static bool YNParse(string ynInput) => ynInput == "y";

        public static Schedule GenerateSchedule(Excel.Range blockData, Excel.Range activityData, Excel.Range groupData)
        {
            Group[] groups = new Group[groupData.Rows.Count];
            var GradeToUnit = new Dictionary<Grade,byte>();
            var WaterGroupingToGroups = new Dictionary<byte,byte>();

            for (byte i = 0; i < groups.Length; i++)
            {
                var name = groupData.Cells.Value2[i+1,1];
                var grade = ParseGrade(groupData.Cells.Value2[i+1, 3].ToString());
                var unit = groupData.Cells.Value2[i+1, 4];
                bool specialGroup = YNParse(groupData.Cells.Value2[i+1, 2]);
                var lunch = groupData.Cells.Value2[i+1, 5];
                var waterGrouping = groupData.Cells.Value2[i + 1, 6];
                groups[i] = new Group(i, name, grade, (byte)unit , specialGroup, (byte)lunch, (byte)waterGrouping);

                //deal with groupings

                if (GradeToUnit.ContainsKey(grade)) continue;
                GradeToUnit.Add(grade, (byte)unit);
            }


            Dictionary<byte, byte> lunchNumToTimeIndex = new Dictionary<byte, byte>();
            string[] times = new string[blockData.Rows.Count];

            var specActPrefs = new SpecialActivityPrefs[blockData.Rows.Count];

            for (byte i = 0; i < blockData.Rows.Count; i++)
            {
                var timeName = blockData.Cells.Value2[i + 1, 1];
                times[i] = timeName;

                var lunchNum = blockData.Cells.Value2[i + 1, 2];
                if (lunchNum != 0) lunchNumToTimeIndex.Add((byte)lunchNum, i);

                char openingCirclePref = blockData.Cells.Value2[i + 1, 3].ToString()[0];
                char middleCirclePref = blockData.Cells.Value2[i + 1, 4].ToString()[0];
                char popsicleTimePref = blockData.Cells.Value2[i + 1, 5].ToString()[0];
                char closingCirclePref = blockData.Cells.Value2[i + 1, 6].ToString()[0];
                char openPref = blockData.Cells.Value2[i + 1, 7].ToString()[0];
                string specialEntPrefs = blockData.Cells.Value2[i + 1, 8].ToString();

                specActPrefs[i] = new SpecialActivityPrefs(openingCirclePref, middleCirclePref, popsicleTimePref, closingCirclePref, openPref, specialEntPrefs);
            }


            var schedule = new Schedule(blockData.Rows.Count, groups, times,lunchNumToTimeIndex, GradeToUnit, specActPrefs);

            for (byte i = 0; i < activityData.Rows.Count; i++)
            {
                var name = activityData.Cells.Value2[i + 1, 1];
                bool water = YNParse(activityData.Cells.Value2[i + 1, 2]);
                bool overflow = YNParse(activityData.Cells.Value2[i + 1, 3]);
                var numOfGroups = activityData.Cells.Value2[i + 1, 4];
                var onlyGrades = ParseGrades(activityData.Cells.Value2[i + 1, 5]);
                var strikedGrades = ParseGrades(activityData.Cells.Value2[i + 1, 6]);
                bool open = YNParse(activityData.Cells.Value2[i + 1, 7]);
                bool isSpecialist = YNParse(activityData.Cells.Value2[i + 1, 8]);
                schedule.AddActivity(name, water, overflow, (byte)numOfGroups, open, onlyGrades, strikedGrades, isSpecialist);
            }

            schedule.GenerateSchedule();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(blockData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(activityData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(groupData);


            return schedule;
        }

        private void ScheduleSpecialActivity(byte blockNum, char unitPref, string ActivityName)
        {
            switch (unitPref)
            {
                case 'b':
                    foreach (var group in Groups)
                    {
                        ScheduleData[blockNum, group.RowNum] = ActivityName;
                    }
                    break;
                case '2':
                    foreach (var group in Groups)
                        if (GradeToUnit[group.Grade] == 2)
                            ScheduleData[blockNum, group.RowNum] = ActivityName;
                    break;
                case '1':
                    foreach (var group in Groups)
                        if (GradeToUnit[group.Grade] == 1)
                            ScheduleData[blockNum, group.RowNum] = ActivityName;
                    break;
                default:
                    break;
            }
        }
        private void ScheduleSpecialActivity(byte blockNum, string gradePrefs, string ActivityName)
        {
            var grades = ParseGrades(gradePrefs);

            foreach(var group in Groups)
            {
                if(grades.Contains(group.Grade))
                {
                    ScheduleData[blockNum,group.RowNum] = ActivityName;
                }
            }
        }

        private void GenerateSchedule()
        {
            for (byte block = 0; block < Times.Length; block++)
            {
                ScheduleSpecialActivity(block, SpecActPrefs[block].OpenPref, "Open Activity");

                ScheduleSpecialActivity(block, SpecActPrefs[block].OpeningCirclePref, "Opening Circle");
                ScheduleSpecialActivity(block, SpecActPrefs[block].MiddleCirclePref, "Middle Circle");
                ScheduleSpecialActivity(block, SpecActPrefs[block].PopsicleTimePref, "Popsicle Time");
                ScheduleSpecialActivity(block, SpecActPrefs[block].ClosingCirclePref, "Closing Circle");

                ScheduleSpecialActivity(block, SpecActPrefs[block].SpecialEntPrefs, "Special Entertainment");
            }

            foreach (Group group in Groups)
            {
                ScheduleData[LunchNumToTimeIndex[group.LunchNum], group.RowNum] = "Lunch " + group.LunchNum;
            }

            Stack<int> lunchStack = new Stack<int>();
            for(int i = 0; i < Activities.Count; i++)
            {
                if (Activities[i].IsSpecialist)
                {
                    lunchStack.Push(i);
                }
            }

            Stack<int> BookableActivitiesIndexes = new Stack<int>();
            for(byte blockIndex = 0; blockIndex < Times.Length; blockIndex++)
            {
                BookableActivitiesIndexes.Clear();

                for(int i = 0; i < Activities.Count; i++)
                {
                    if (Activities[i].Overflow || Activities[i].WaterActivity) continue;

                    //if (LunchNumToTimeIndex.ContainsValue(blockIndex)) continue; //do something special here
                    BookableActivitiesIndexes.Push(i);
                }
                BookableActivitiesIndexes = new Stack<int>(BookableActivitiesIndexes.OrderBy(_ => Gen.Next()));

                foreach (Group group in Groups)
                {
                    
                    //if(Activities[BookableActivitiesIndexes.Peek()].GradeStrike.Contains(group.Grade) || !Activities[BookableActivitiesIndexes.Peek()].GradeOnly.Contains(group.Grade))
                }
            }
        }

        public void OutputSchedule(Excel.Range outputRange)
        {
            for(int column = 0;column < Times.Length;column++)
            {
                outputRange.Cells[1, column + 2].Value2 = Times[column];
                outputRange.Cells[2, column + 2].Value2 = column + 1;
            }

            for (int row = 0; row < Groups.Length; row++)
            {
                outputRange.Cells[row + 3, 1].Value2 = Groups[row].Name;
            }

            for (int row = 0; row < Times.Length; row++)
            {
                for (int column = 0; column < Groups.Length; column++)
                {
                    outputRange.Cells[column+3, row+2].Value2 = ScheduleData[row, column];
                }
            }

            outputRange.Columns.AutoFit();
            outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
        
    }
}

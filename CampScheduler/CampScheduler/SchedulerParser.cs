using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CampScheduler
{
    public static class SchedulerParser
    {
        public static Grade ParseGrade(string gradeInput)
        {
            switch (gradeInput)
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

        public static ChangingRoomCode ParseChangingRoom(string roomInput)
        {
            switch(roomInput)
            {
                case "f":
                    return ChangingRoomCode.F;
                case "m":
                    return ChangingRoomCode.M;
                case "b":
                    return ChangingRoomCode.B;
                default:
                    return ChangingRoomCode.n;
            }
        }

        public static Grade[] ParseGrades(string gradesInput)
        {
            var gradeStrings = gradesInput.Split(',');

            if (gradeStrings[0][0] == 'a')
            {
                return new[] { Grade.PK, Grade.K, Grade.First, Grade.Second, Grade.Third, Grade.Fourth, Grade.Fifth, Grade.Sixth, Grade.MS };
            }

            Grade[] grades = new Grade[gradeStrings.Length];

            for (int i = 0; i < grades.Length; i++)
            {
                grades[i] = ParseGrade(gradeStrings[i].Trim());
                if (grades[i] == Grade.NA) return new Grade[0];
            }

            return grades;
        }

        public static bool YNParse(string ynInput) => ynInput == "y";

        public static DaySchedule GenerateDaySchedule(Excel.Range blockData, Excel.Range activityData, Excel.Range groupData, Excel.Range rulesData)
        {
            Group[] groups = new Group[groupData.Rows.Count];
            var GradeToUnit = new Dictionary<Grade, byte>();
            var WaterGroupingToGroups = new Dictionary<byte, byte>();

            try
            {
                for (byte i = 0; i < groups.Length; i++)
                {
                    var name = groupData.Cells.Value2[i + 1, 1];
                    var grade = ParseGrade(groupData.Cells.Value2[i + 1, 3].ToString());
                    var unit = groupData.Cells.Value2[i + 1, 4];
                    bool specialGroup = YNParse(groupData.Cells.Value2[i + 1, 2]);
                    var lunch = groupData.Cells.Value2[i + 1, 5];

                    groups[i] = new Group(i, name, grade, (byte)unit, specialGroup, (byte)lunch);

                    if (GradeToUnit.ContainsKey(grade)) continue;
                    GradeToUnit.Add(grade, (byte)unit);
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse groups table; check for empty or invalid inputs");
            }


            Dictionary<byte, byte> lunchNumToTimeIndex = new Dictionary<byte, byte>();
            string[] times = new string[blockData.Rows.Count];

            var specActPrefs = new SpecialActivityPrefs[blockData.Rows.Count];
            try
            {
                for (byte i = 0; i < blockData.Rows.Count; i++)
                {
                    var timeName = blockData.Cells[i + 1, 1].Text;
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
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse blocks table; check for empty or invalid inputs");
            }

            DayInfo dayInfo = new DayInfo("Day",times,specActPrefs,lunchNumToTimeIndex);

            var schedule = new DaySchedule(groups, dayInfo, GradeToUnit);

            try
            {
                for (byte i = 0; i < activityData.Rows.Count; i++)
                {
                    var name = activityData.Cells.Value2[i + 1, 1];
                    bool water = YNParse(activityData.Cells.Value2[i + 1, 2]);
                    bool overflow = YNParse(activityData.Cells.Value2[i + 1, 3]);
                    string numOfGroups = Convert.ToString(activityData.Cells.Value2[i + 1, 4]);
                    var onlyGrades = ParseGrades(activityData.Cells.Value2[i + 1, 5]);
                    var strikedGrades = ParseGrades(activityData.Cells.Value2[i + 1, 6]);
                    bool open = YNParse(activityData.Cells.Value2[i + 1, 7]);
                    bool isSpecialist = YNParse(activityData.Cells.Value2[i + 1, 8]);
                    schedule.AddActivity(name, water, overflow, numOfGroups.Trim(' ').Split(',').Select(n => byte.Parse(n.Trim(' '))).ToArray(), open, onlyGrades, strikedGrades, isSpecialist);
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse activities table; check for empty or invalid inputs");
            }

            try
            {
                for (byte i = 0; i < rulesData.Rows.Count; i++)
                {
                    var groupIDs = schedule.ParseGroupOrGrade(rulesData.Cells.Value2[i + 1, 1]);
                    var actIds = schedule.ParseActivities(rulesData.Cells.Value2[i + 1, 2]);
                    var timeIndeces = schedule.ParseTimes(rulesData.Cells.Value2[i + 1, 3]);
                    schedule.AddRule(groupIDs, actIds, timeIndeces);
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse rules table; check for empty or invalid inputs");
            }


            schedule.GenerateSchedule();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(blockData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(activityData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(groupData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rulesData);

            return schedule;
        }

        public static WeekSchedule GenerateWeekSchedule(Excel.Range blockData, Excel.Range activityData, Excel.Range groupData, Excel.Range rulesData)
        {
            Group[] groups = new Group[groupData.Rows.Count];
            var GradeToUnit = new Dictionary<Grade, byte>();
            var WaterGroupingToGroups = new Dictionary<byte, byte>();

            try
            {
                for (byte i = 0; i < groups.Length; i++)
                {
                    var name = groupData.Cells.Value2[i + 1, 1];
                    var grade = ParseGrade(groupData.Cells.Value2[i + 1, 3].ToString());
                    var unit = groupData.Cells.Value2[i + 1, 4];
                    bool specialGroup = YNParse(groupData.Cells.Value2[i + 1, 2]);
                    var lunch = groupData.Cells.Value2[i + 1, 5];

                    groups[i] = new Group(i, name, grade, (byte)unit, specialGroup, (byte)lunch);

                    if (GradeToUnit.ContainsKey(grade)) continue;
                    GradeToUnit.Add(grade, (byte)unit);
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse groups table; check for empty or invalid inputs");
            }


            var dayInputs = new Dictionary<string, (List<string> times, Dictionary<byte, byte> lunchNumToTimeIndex, List<SpecialActivityPrefs> specActPrefs)>();   
            try
            {
                for (byte i = 0; i < blockData.Rows.Count; i++)
                {
                    string dayName = blockData.Cells[i + 1, 1].Text;
                    if(!dayInputs.ContainsKey(dayName))
                    {
                        dayInputs.Add(dayName, (new List<string>(), new Dictionary<byte,byte>(),new List<SpecialActivityPrefs>()));
                    }

                    string timeName = blockData.Cells[i + 1, 2].Text;
                    dayInputs[dayName].times.Add(timeName);

                    var lunchNum = blockData.Cells.Value2[i + 1, 3];
                    if (lunchNum != 0) dayInputs[dayName].lunchNumToTimeIndex.Add((byte)lunchNum, (byte)(dayInputs[dayName].times.Count-1));

                    char openingCirclePref = blockData.Cells.Value2[i + 1, 4].ToString()[0];
                    char middleCirclePref = blockData.Cells.Value2[i + 1, 5].ToString()[0];
                    char popsicleTimePref = blockData.Cells.Value2[i + 1, 6].ToString()[0];
                    char closingCirclePref = blockData.Cells.Value2[i + 1, 7].ToString()[0];
                    char openPref = blockData.Cells.Value2[i + 1, 8].ToString()[0];
                    string specialEntPrefs = blockData.Cells.Value2[i + 1, 9].ToString();

                    dayInputs[dayName].specActPrefs.Add(new SpecialActivityPrefs(openingCirclePref, middleCirclePref, popsicleTimePref, closingCirclePref, openPref, specialEntPrefs));
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse blocks table; check for empty or invalid inputs");
            }

            Dictionary<string, DayInfo> WeekInfo = new Dictionary<string, DayInfo>();

            foreach(var dayInput in dayInputs)
            {
                var dayInfo = new DayInfo(dayInput.Key, dayInput.Value.times.ToArray(),dayInput.Value.specActPrefs.ToArray(),dayInput.Value.lunchNumToTimeIndex);

                WeekInfo.Add(dayInput.Key, dayInfo);
            }

            var schedule = new WeekSchedule(groups, WeekInfo, GradeToUnit);

            try
            {
                for (byte i = 0; i < activityData.Rows.Count; i++)
                {
                    var name = activityData.Cells.Value2[i + 1, 1];
                    bool water = YNParse(activityData.Cells.Value2[i + 1, 2]);
                    bool overflow = YNParse(activityData.Cells.Value2[i + 1, 3]);
                    string numOfGroups = Convert.ToString(activityData.Cells.Value2[i + 1, 4]);
                    var onlyGrades = ParseGrades(activityData.Cells.Value2[i + 1, 5]);
                    var strikedGrades = ParseGrades(activityData.Cells.Value2[i + 1, 6]);
                    string[] open = schedule.ParseDays(activityData.Cells.Value2[i + 1, 7]);
                    bool isSpecialist = YNParse(activityData.Cells.Value2[i + 1, 8]);
                    schedule.AddActivity(name, water, overflow, numOfGroups.Trim(' ').Split(',').Select(n => byte.Parse(n.Trim(' '))).ToArray(), open, onlyGrades, strikedGrades, isSpecialist);
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse activities table; check for empty or invalid inputs");
            }

            try
            {
                for (byte i = 0; i < rulesData.Rows.Count; i++)
                {
                    var groupIDs = schedule.ParseGroupOrGrade(rulesData.Cells.Value2[i + 1, 1]);
                    var actIds = schedule.ParseActivities(rulesData.Cells.Value2[i + 1, 2]);
                    var days = schedule.ParseDays(rulesData.Cells.Value2[i + 1, 4]);

                    var timeListInput = rulesData.Cells.Value2[i + 1, 3];

                    byte[] timeIndeces;
                    foreach (var day in days)
                    {
                        timeIndeces = schedule.ParseTimes(day, timeListInput);
                        schedule.AddRule(day, groupIDs, actIds, timeIndeces);
                    }
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse rules table; check for empty or invalid inputs");
            }


            schedule.GenerateSchedule();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(blockData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(activityData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(groupData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(rulesData);

            return schedule;
        }

        public static Bump GenerateBump(Excel.Range blockData, Excel.Range activityData, Excel.Range counselorData)
        {
            List<BumpActivity> activities = new List<BumpActivity>();
            List<Counselor> counselors = new List<Counselor>();

            List<string> times = new List<string>();
            Dictionary<byte, byte> lunchNumToTimeIndex = new Dictionary<byte, byte>();
            try
            {
                for (byte i = 0; i < blockData.Rows.Count; i++)
                {
                    bool isOpen = YNParse(blockData.Cells.Value2[i + 1, 2]);
                    if(!isOpen) continue;
                    
                    var timeName = blockData.Cells[i + 1, 1].Text;
                    times.Add(timeName);

                    var lunchNum = blockData.Cells.Value2[i + 1, 3];
                    if (lunchNum != 0) lunchNumToTimeIndex.Add((byte)lunchNum, i);
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse blocks table; check for empty or invalid inputs");
            }

            DayInfo dayInfo = new DayInfo("Day", times.ToArray(), new SpecialActivityPrefs[0], lunchNumToTimeIndex);

            try
            {
                byte id = 0;
                for (byte i = 0; i < activityData.Rows.Count; i++)
                {
                    bool open = YNParse(activityData.Cells.Value2[i + 1, 2]);
                    if (!open) continue;

                    var name = activityData.Cells.Value2[i + 1, 1];
                    var numOfPaid = activityData.Cells.Value2[i + 1, 3];
                    var numOfUnpaid = activityData.Cells.Value2[i + 1, 4];
                    var required = ParseChangingRoom(activityData.Cells.Value2[i + 1, 5]);
                    bool accessible = YNParse(activityData.Cells.Value2[i + 1, 6]);
                    bool overflow = YNParse(activityData.Cells.Value2[i + 1, 7]);

                    activities.Add(new BumpActivity(id, name, (byte)numOfPaid, (byte)numOfUnpaid, required, accessible, overflow));
                    id++;
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse activities table; check for empty or invalid inputs");
            }

            try
            {
                
                for (byte i = 0; i < counselorData.Rows.Count; i++)
                {
                    bool working = YNParse(counselorData.Cells.Value2[i + 1, 6]);
                    if (!working) continue;

                    var name = counselorData.Cells.Value2[i + 1, 1];
                    bool paid = YNParse(counselorData.Cells.Value2[i + 1, 2]);
                    var changingRoom = ParseChangingRoom(counselorData.Cells.Value2[i + 1, 3]);
                    var lunch = counselorData.Cells.Value2[i + 1, 4];
                    bool handicap = YNParse(counselorData.Cells.Value2[i + 1, 5]);

                    counselors.Add(new Counselor(i, name, paid, changingRoom, (byte)lunch, handicap));
                }
            }
            catch (Exception)
            {
                throw new Exception("Failed to parse counselors table; check for empty or invalid inputs");
            }

            var bump = new Bump(dayInfo, activities, counselors);

            bump.GenerateBump();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(blockData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(activityData);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(counselorData);

            return bump;

        }

    }
}

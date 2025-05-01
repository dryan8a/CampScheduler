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
using Microsoft.Office.Core;
using System.Linq.Expressions;
using System.Security;
using System.Xml;
using System.Drawing;
using System.Windows.Forms.VisualStyles;

namespace CampScheduler
{
    public class WeekSchedule : Schedule
    {
        public Dictionary<string,string[,]> ScheduleData; //change to internal
        internal List<WeekActivity> Activities { get; }
        internal List<WeekActivity> WaterActivities { get; }

        internal Dictionary<string,DayInfo> WeekInfo { get; }

        public int NumOfDays => WeekInfo.Count;

        internal Dictionary<string,List<Rule>> Rules {get;}

        internal WeekSchedule(Group[] groups, Dictionary<string,DayInfo> weekInfo, Dictionary<Grade, byte> gradeToUnit) : base(groups, gradeToUnit)
        {
            ScheduleData = new Dictionary<string, string[,]>();
            foreach(var dayInfo in weekInfo.Values)
            {
                ScheduleData.Add(dayInfo.DayName,new string[dayInfo.Times.Length,groups.Length]);
            }

            Activities = new List<WeekActivity>();
            WaterActivities = new List<WeekActivity>();

            WeekInfo = weekInfo;

            Rules = new Dictionary<string, List<Rule>>();
        }

        public void AddActivity(string name, bool waterActivity, bool overflow, byte[] numOfGroups, string[] open, Grade[] gradeOnly, Grade[] gradeStrike, bool specialist)
        {
            WeekActivity activity = new WeekActivity((byte)Activities.Count, name, waterActivity, overflow, numOfGroups, open, gradeOnly, gradeStrike, specialist);
            Activities.Add(activity);
            if (waterActivity)
            {
                WaterActivities.Add(activity);
                if (activity.NumofGroups.Length > WActMaxNumofGroups) WActMaxNumofGroups = activity.NumofGroups.Length;
            }

            if (!waterActivity && specialist) NumOfSpecialists++;

        }

        public void AddRule(string day, byte[] groupIDs, byte[] actIDs, byte[] timeIDs)
        {
            if (!Rules.ContainsKey(day)) Rules.Add(day, new List<Rule>());

            Rules[day].Add(new Rule(groupIDs, actIDs, timeIDs));
        }

        public string[] ParseDays(string dayListInput)
        {
            if(string.Equals("a",dayListInput.Trim(),StringComparison.CurrentCultureIgnoreCase))
            {
                return WeekInfo.Keys.ToArray();
            }

            var dayStrings = dayListInput.Split(',');
            return dayStrings.Select(dayString => dayString.Trim()).Where(dayString => WeekInfo.Keys.Contains(dayString)).ToArray();
        }

        public byte[] ParseActivities(string activityNamesInput)
        {
            var activityStrings = activityNamesInput.Split(',');
            byte[] activityIds = new byte[activityStrings.Length];
            for (int i = 0; i < activityStrings.Length; i++)
            {
                activityIds[i] = Activities.First(act => act.Name.Equals(activityStrings[i].Trim())).Id;
            }
            return activityIds;
        }
        public byte[] ParseTimes(string day, string timesInput)
        {
            var timesStrings = timesInput.Split(',');
            byte[] activityIds = new byte[timesStrings.Length];
            for (int i = 0; i < timesStrings.Length; i++)
            {
                activityIds[i] = (byte)Array.IndexOf(WeekInfo[day].Times, timesStrings[i].Trim());
                if (activityIds[i] == 255) throw new Exception();
            }
            return activityIds;
        }

        private void ScheduleSpecialActivity(string day, byte blockNum, char unitPref, string ActivityName)
        {
            switch (unitPref)
            {
                case 'b':
                    foreach (var group in Groups)
                    {
                        ScheduleData[day][blockNum, group.RowNum] = ActivityName;
                    }
                    break;
                case '2':
                    foreach (var group in Groups)
                        if (GradeToUnit[group.Grade] == 2)
                            ScheduleData[day][blockNum, group.RowNum] = ActivityName;
                    break;
                case '1':
                    foreach (var group in Groups)
                        if (GradeToUnit[group.Grade] == 1)
                            ScheduleData[day][blockNum, group.RowNum] = ActivityName;
                    break;
                default:
                    break;
            }
        }
        private void ScheduleSpecialActivity(string day, byte blockNum, string gradePrefs, string ActivityName)
        {
            var grades = SchedulerParser.ParseGrades(gradePrefs);

            foreach (var group in Groups)
            {
                if (grades.Contains(group.Grade))
                {
                    ScheduleData[day][blockNum, group.RowNum] = ActivityName;
                }
            }
        }

        private bool IsBookedInBlock(byte ActID, string day, int TimeIndex)
        {
            for (int i = 0; i < Groups.Length; i++)
            {
                if (ScheduleData[day][TimeIndex, i] == Activities[ActID].Name) return true;
            }

            return false;
        }

        //private ScheduleActivityReturnCode CanScheduleWater(byte wActID, int TimeIndex, byte startGroupID, int numOfGroupsIndex)
        //{
        //    int endGroupID = startGroupID + Activities[wActID].GetNumOfGroups(numOfGroupsIndex);
        //    for (byte groupID = startGroupID; groupID < endGroupID; groupID++)
        //    {
        //        if (groupID >= Groups.Count()) break;
        //        if (Activities[wActID].GradeOnly.Length > 0 && !Activities[wActID].GradeOnly.Contains(Groups[groupID].Grade))
        //        {
        //            return ScheduleActivityReturnCode.NotGradeOnly;
        //        }
        //        if (Activities[wActID].GradeStrike.Length > 0 && Activities[wActID].GradeStrike.Contains(Groups[groupID].Grade))
        //        {
        //            return ScheduleActivityReturnCode.GradeStriked;
        //        }
        //        if (!string.IsNullOrEmpty(ScheduleData[TimeIndex, groupID]))
        //        {
        //            return ScheduleActivityReturnCode.Overlapped;
        //        }
        //        if (Activities[wActID].Open && DayInfo.SpecialActivityPrefs[TimeIndex].OpenPref != 'n')
        //        {
        //            return ScheduleActivityReturnCode.BookedOpen;
        //        }
        //        if (IsBookedInBlock(wActID, TimeIndex) || GroupIDsWithRuleWActs.Contains(groupID)) return ScheduleActivityReturnCode.Duplicate;
        //    }
        //    return ScheduleActivityReturnCode.Success;
        //}

        //private void ScheduleWaterActivities()
        //{
        //    var waterActivityTimesAvailable = new List<(byte, int)>();
        //    byte lunchNum;
        //    foreach (DayActivity wAct in WaterActivities)
        //    {
        //        lunchNum = (byte)(Gen.Next(DayInfo.LunchNumToTimeIndex.Count) + 1);
        //        if (IsBookedInBlock(wAct.Id, DayInfo.LunchNumToTimeIndex[lunchNum]))
        //        {
        //            lunchNum = (byte)((lunchNum + 1) % DayInfo.LunchNumToTimeIndex.Count);
        //        }

        //        for (int i = 0; i < DayInfo.Times.Count(); i++)
        //        {
        //            if (i == DayInfo.LunchNumToTimeIndex[lunchNum] && wAct.IsSpecialist) continue;
        //            waterActivityTimesAvailable.Add((wAct.Id, i));
        //        }
        //    }
        //    waterActivityTimesAvailable = new List<(byte, int)>(waterActivityTimesAvailable.OrderBy(_ => Gen.Next()));

        //    int WActNumOfGroupCombos = WActMaxNumofGroups * waterActivityTimesAvailable.Count;

        //    var UnitOpen = new[] { false, false };
        //    foreach (var specActPref in DayInfo.SpecialActivityPrefs)
        //    {
        //        switch (specActPref.OpenPref)
        //        {
        //            case 'b':
        //                UnitOpen[0] = true;
        //                UnitOpen[1] = true;
        //                break;
        //            case '1':
        //                UnitOpen[0] = true;
        //                break;
        //            case '2':
        //                UnitOpen[1] = true;
        //                break;
        //        }
        //    }

        //    int failCount = 0;
        //    var ScheduledWaters = new List<(int TimeIndex, byte groupId, string ActName)>();
        //    while (true)
        //    {
        //        ScheduledWaters.Clear();
        //        bool failed = false;

        //        int minGroupsIndex = Math.DivRem(failCount, waterActivityTimesAvailable.Count(), out int numOfMaxGroupsIndex);

        //        for (byte groupID = 0, availableIndex = 0; ; availableIndex++)
        //        {
        //            while (groupID < Groups.Length && UnitOpen[Groups[groupID].Unit - 1]) groupID++;

        //            if (groupID >= Groups.Length) break;

        //            if (GroupIDsWithRuleWActs.Contains(groupID))
        //            {
        //                groupID++;
        //                availableIndex--;
        //                continue;
        //            }

        //            if (availableIndex == waterActivityTimesAvailable.Count)
        //            {
        //                failCount++;
        //                failed = true;
        //                break;
        //            }

        //            //int numOfGroupsIndex = (int)Math.Ceiling((float)(failCount < availableIndex ? 0 : failCount - availableIndex) / waterActivityTimesAvailable.Count());
        //            int numOfGroupsIndex = minGroupsIndex;
        //            if (Gen.Next(waterActivityTimesAvailable.Count() - availableIndex) < numOfMaxGroupsIndex)
        //            {
        //                numOfGroupsIndex++;
        //                numOfMaxGroupsIndex--;
        //            }

        //            (byte wActID, int TimeIndex) = (0, 0);
        //            //(byte firstActID, int FirstTime) = waterActivityTimesAvailable[availableIndex];
        //            for (int i = 0; ; i++)
        //            {
        //                (wActID, TimeIndex) = waterActivityTimesAvailable[availableIndex];

        //                var ScheduleCode = CanScheduleWater(wActID, TimeIndex, groupID, numOfGroupsIndex);

        //                //Activity overlapped when it didn't need to
        //                if (ScheduleCode == ScheduleActivityReturnCode.Overlapped && numOfGroupsIndex > minGroupsIndex && numOfMaxGroupsIndex + 1 < waterActivityTimesAvailable.Count() - availableIndex)
        //                {
        //                    numOfMaxGroupsIndex++;
        //                    numOfGroupsIndex = minGroupsIndex;
        //                    ScheduleCode = CanScheduleWater(wActID, TimeIndex, groupID, numOfGroupsIndex);
        //                }

        //                //Activity failed to place, rotate to next possible activity 
        //                if (ScheduleCode != ScheduleActivityReturnCode.Success)
        //                {

        //                    var temp = waterActivityTimesAvailable[availableIndex];
        //                    waterActivityTimesAvailable.RemoveAt(availableIndex);
        //                    waterActivityTimesAvailable.Add(temp);
        //                    if (i >= waterActivityTimesAvailable.Count - availableIndex - 1)
        //                    {
        //                        failed = true;
        //                        break;
        //                    }
        //                    continue;
        //                }
        //                break;
        //            }
        //            if (failed)
        //            {
        //                failCount++;
        //                break;
        //            }

        //            for (int i = 0; i < Activities[wActID].GetNumOfGroups(numOfGroupsIndex); i++)
        //            {
        //                ScheduledWaters.Add((TimeIndex, groupID, Activities[wActID].Name));
        //                groupID++;
        //                if (groupID >= Groups.Length) break;
        //            }
        //        }
        //        if (failCount > WActNumOfGroupCombos) throw new Exception("Couldn't schedule water activities for all groups; try freeing up schedule");
        //        if (!failed)
        //        {
        //            break;
        //        }
        //    }
        //    foreach (var scheduledWater in ScheduledWaters)
        //    {
        //        ScheduleData[scheduledWater.TimeIndex, scheduledWater.groupId] = scheduledWater.ActName;
        //    }
        //}

        //private ScheduleActivityReturnCode CanScheduleRegular(byte ActID, int TimeIndex, byte startGroupID)
        //{
        //    int endGroupID = startGroupID + Activities[ActID].NumofGroups[0]; //"prioritizes first num of groups"
        //    for (int groupID = startGroupID; groupID < endGroupID; groupID++)
        //    {
        //        var Act = Activities[ActID];
        //        if (groupID >= Groups.Count() || !string.IsNullOrEmpty(ScheduleData[TimeIndex, groupID]))
        //        {
        //            return ScheduleActivityReturnCode.Overlapped;
        //        }
        //        if (Groups[groupID].SpecialGroup)
        //        {
        //            return ScheduleActivityReturnCode.SpecialGroup;
        //        }
        //        if (Act.GradeStrike.Length > 0 && Act.GradeStrike.Contains(Groups[groupID].Grade))
        //        {
        //            return ScheduleActivityReturnCode.GradeStriked;
        //        }
        //        if (Act.GradeOnly.Length > 0 && !Act.GradeOnly.Contains(Groups[groupID].Grade))
        //        {
        //            return ScheduleActivityReturnCode.NotGradeOnly;
        //        }
        //        if (Act.Open && DayInfo.SpecialActivityPrefs[TimeIndex].OpenPref != 'n')
        //        {
        //            return ScheduleActivityReturnCode.BookedOpen;
        //        }
        //        for (int i = 0; i < DayInfo.Times.Length; i++)
        //        {
        //            if (ScheduleData[i, groupID] == Act.Name)
        //            {
        //                return ScheduleActivityReturnCode.Duplicate;
        //            }
        //        }
        //        if (IsBookedInBlock(ActID, TimeIndex)) return ScheduleActivityReturnCode.Duplicate;
        //    }

        //    return ScheduleActivityReturnCode.Success;
        //}

        //private void ScheduleRegularActivities(int[] LunchNumsCount)
        //{
        //    var BookableActInds = new List<byte>();
        //    var GroupInds = Enumerable.ToList(Enumerable.Range(0, Groups.Length));
        //    var BookableActivityToLunchNum = new Dictionary<byte, byte>();
        //    var OverflowActInds = new List<byte>();

        //    double numOfRegGroups = Groups.Sum(g => g.SpecialGroup ? 0 : 1);

        //    for (int i = 0; i < LunchNumsCount.Length; i++)
        //    {
        //        LunchNumsCount[i] = (int)Math.Round(LunchNumsCount[i] / numOfRegGroups * NumOfSpecialists);
        //    }
        //    LunchNumsCount[LunchNumsCount.Length - 1]++;

        //    foreach (byte ActId in new List<int>(Enumerable.ToList(Enumerable.Range(0, Activities.Count)).OrderBy(_ => Gen.Next())))
        //    {
        //        if (Activities[ActId].WaterActivity) continue;
        //        if (Activities[ActId].Overflow)
        //        {
        //            OverflowActInds.Add(ActId);
        //            continue;
        //        }

        //        BookableActInds.Add(ActId);

        //        if (Activities[ActId].IsSpecialist)
        //        {
        //            //randomly choose lunch for specialist based off of lunch counts
        //            byte currentLunchNum = 1;

        //            while (LunchNumsCount[currentLunchNum - 1] == 0 || IsBookedInBlock(ActId, DayInfo.LunchNumToTimeIndex[currentLunchNum]))
        //            {
        //                currentLunchNum = (byte)(currentLunchNum % DayInfo.LunchNumToTimeIndex.Count + 1);
        //                if (currentLunchNum == 1) throw new Exception($"Couldn't give specialist for {Activities[ActId].Name} a lunch; check rules table to see if they were overbooked");
        //            }

        //            LunchNumsCount[currentLunchNum - 1]--;

        //            BookableActivityToLunchNum.Add(ActId, currentLunchNum);
        //        }
        //    }

        //    int currentBookableActIndInd;
        //    for (byte blockIndex = 0; blockIndex < DayInfo.Times.Length; blockIndex++)
        //    {
        //        currentBookableActIndInd = 0;
        //        BookableActInds = new List<byte>(BookableActInds.OrderBy(_ => Gen.Next()));
        //        GroupInds = new List<int>(GroupInds.OrderBy(_ => Gen.Next()));

        //        for (int GroupIndInd = 0; GroupIndInd < Groups.Length; GroupIndInd++)
        //        {
        //            var GroupInd = GroupInds[GroupIndInd];
        //            var group = Groups[GroupInd];

        //            if (group.SpecialGroup) continue;
        //            if (!string.IsNullOrEmpty(ScheduleData[blockIndex, group.RowNum])) continue;

        //            bool needsOverflow = false;
        //            if (currentBookableActIndInd >= BookableActInds.Count)
        //            {
        //                needsOverflow = true;
        //            }
        //            var currentAct = Activities[BookableActInds[currentBookableActIndInd >= BookableActInds.Count ? 0 : currentBookableActIndInd]];
        //            var originalBookableName = currentAct.Name;
        //            ScheduleActivityReturnCode ScheduleCode = ScheduleActivityReturnCode.NotReturned;
        //            while (!needsOverflow)
        //            {
        //                ScheduleCode = CanScheduleRegular(BookableActInds[currentBookableActIndInd], blockIndex, group.RowNum);
        //                if ((currentAct.IsSpecialist && DayInfo.LunchNumToTimeIndex[BookableActivityToLunchNum[BookableActInds[currentBookableActIndInd]]] == blockIndex)
        //                    || ScheduleCode != ScheduleActivityReturnCode.Success)
        //                {
        //                    var temp = BookableActInds[currentBookableActIndInd];
        //                    BookableActInds.RemoveAt(currentBookableActIndInd);
        //                    BookableActInds.Add(temp);

        //                    currentAct = Activities[BookableActInds[currentBookableActIndInd]];

        //                    if (originalBookableName == currentAct.Name)
        //                    {
        //                        needsOverflow = true;
        //                        break;
        //                    }
        //                    continue;
        //                }
        //                break;
        //            }

        //            if (!needsOverflow && ScheduleCode == ScheduleActivityReturnCode.SpecialGroup)
        //            {
        //                continue;
        //            }
        //            else if (needsOverflow)
        //            {
        //                if (OverflowActInds.Count == 0) throw new Exception("Couldn't schedule all activities; please add an overflow activity");
        //                ScheduleData[blockIndex, group.RowNum] = Activities[OverflowActInds[Gen.Next(OverflowActInds.Count)]].Name;
        //            }
        //            else
        //            {
        //                for (int i = 0; i < currentAct.NumofGroups[0]; i++)
        //                {
        //                    ScheduleData[blockIndex, group.RowNum + i] = currentAct.Name;
        //                }
        //                currentBookableActIndInd += currentAct.NumofGroups[0];
        //            }
        //        }
        //    }
        //}

        public override void GenerateSchedule()
        {

            foreach (var dayInfo in WeekInfo.Values)
            {
                //Special Activity Scheduling
                for (byte block = 0; block < dayInfo.Times.Length; block++)
                {
                    ScheduleSpecialActivity(dayInfo.DayName, block, dayInfo.SpecialActivityPrefs[block].OpenPref, "Open Activity");

                    ScheduleSpecialActivity(dayInfo.DayName, block, dayInfo.SpecialActivityPrefs[block].OpeningCirclePref, "Opening Circle");
                    ScheduleSpecialActivity(dayInfo.DayName, block, dayInfo.SpecialActivityPrefs[block].MiddleCirclePref, "Middle Circle");
                    ScheduleSpecialActivity(dayInfo.DayName, block, dayInfo.SpecialActivityPrefs[block].PopsicleTimePref, "Popsicle Time");
                    ScheduleSpecialActivity(dayInfo.DayName, block, dayInfo.SpecialActivityPrefs[block].ClosingCirclePref, "Closing Circle");

                    ScheduleSpecialActivity(dayInfo.DayName, block, dayInfo.SpecialActivityPrefs[block].SpecialEntPrefs, "Special Entertainment");
                }


                //Group Lunch Scheduling
                //int[] LunchNumsCount = new int[dayInfo.LunchNumToTimeIndex.Count];
                foreach (Group group in Groups)
                {
                    if (!dayInfo.LunchNumToTimeIndex.TryGetValue(group.LunchNum, out byte timeIndex))
                    {
                        throw new Exception("Invalid Lunch Number entered in groups table; change groups table or add time to blocks table");
                    }
                    ScheduleData[dayInfo.DayName][timeIndex, group.RowNum] = "Lunch " + group.LunchNum;

                    //LunchNumsCount[group.LunchNum - 1]++; //just by the by, this means that lunch num has to start at 1
                }

                //Rules Scheduling
                foreach (Rule rule in Rules[dayInfo.DayName])
                {
                    var actID = rule.ActIDs[Gen.Next(rule.ActIDs.Length)];
                    var timeIndex = rule.TimeIDs[Gen.Next(rule.TimeIDs.Length)];

                    foreach (var groupID in rule.GroupIDs)
                    {
                        ScheduleData[dayInfo.DayName][timeIndex, groupID] = Activities[actID].Name;

                        if (Activities[actID].WaterActivity)
                        {
                            GroupIDsWithRuleWActs.Add(groupID);
                        }
                    }
                }
            }

            //Water Scheduling
            //ScheduleWaterActivities();

            //Regular Activity Scheduling
            //ScheduleRegularActivities(LunchNumsCount);
        }

        public void OutputSchedule(Excel.Range[] outputRanges)
        {
            int i = 0;
            foreach (var dayInfo in WeekInfo.Values)
            {
                outputRanges[i].Range["A1", (char)('A' + dayInfo.Times.Length) + "1"].Merge();
                outputRanges[i].Cells[1, 1].Value2 = dayInfo.DayName;

                for (int column = 0; column < dayInfo.Times.Length; column++)
                {
                    outputRanges[i].Cells[2, column + 2].Value2 = dayInfo.Times[column];
                    outputRanges[i].Cells[3, column + 2].Value2 = column + 1;
                }

                for (int row = 0; row < Groups.Length; row++)
                {
                    outputRanges[i].Cells[row + 4, 1].Value2 = Groups[row].Name;
                }

                for (int row = 0; row < Groups.Length; row++)
                {
                    for (int column = 0; column < dayInfo.Times.Length; column++)
                    {
                        outputRanges[i].Cells[row + 4, column + 2].Value2 = ScheduleData[dayInfo.DayName][column, row];
                    }
                }

                outputRanges[i].Columns.AutoFit();
                outputRanges[i].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                outputRanges[i].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                i++;
            }
        }
    }
}

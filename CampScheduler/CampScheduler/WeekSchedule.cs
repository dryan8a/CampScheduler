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
using System.Web;

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

        internal Dictionary<int, List<byte>> OffBlockRules { get; }

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

            OffBlockRules = new Dictionary<int, List<byte>>(); //hash ("day",ActId)
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
            if (groupIDs.Length == 0)
            {
                foreach (var actID in actIDs)
                {
                    var key = (day, actID).GetHashCode();

                    if (!OffBlockRules.ContainsKey(key)) OffBlockRules.Add(key, new List<byte>());

                    foreach (var timeID in timeIDs)
                    {
                        OffBlockRules[key].Add(timeID);
                    }
                }
            }
            else
            {
                if (!Rules.ContainsKey(day)) Rules.Add(day, new List<Rule>());

                Rules[day].Add(new Rule(groupIDs, actIDs, timeIDs));
            }
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

        private ScheduleActivityReturnCode CanScheduleWater(byte wActID, string day, int TimeIndex, byte startGroupID, int numOfGroupsIndex)
        {
            int endGroupID = startGroupID + Activities[wActID].GetNumOfGroups(numOfGroupsIndex);
            for (byte groupID = startGroupID; groupID < endGroupID; groupID++)
            {
                if (groupID >= Groups.Count()) break;
                if (Activities[wActID].GradeOnly.Length > 0 && !Activities[wActID].GradeOnly.Contains(Groups[groupID].Grade))
                {
                    return ScheduleActivityReturnCode.NotGradeOnly;
                }
                if (Activities[wActID].GradeStrike.Length > 0 && Activities[wActID].GradeStrike.Contains(Groups[groupID].Grade))
                {
                    return ScheduleActivityReturnCode.GradeStriked;
                }
                if (!string.IsNullOrEmpty(ScheduleData[day][TimeIndex, groupID])) //&& !string.Equals(ScheduleData[TimeIndex,groupID], "Open Activity"))
                {
                    return ScheduleActivityReturnCode.Overlapped;
                }
                if (Activities[wActID].Open.Contains(day) && WeekInfo[day].SpecialActivityPrefs[TimeIndex].OpenPref != 'n')
                {
                    return ScheduleActivityReturnCode.BookedOpen;
                }
                if (IsBookedInBlock(wActID, day, TimeIndex) || GroupIDsWithRuleWActs.Contains(groupID)) return ScheduleActivityReturnCode.Duplicate;
            }
            return ScheduleActivityReturnCode.Success;
        }

        private void ScheduleWaterActivities(string day)
        {
            var waterActivityTimesAvailable = new List<(byte, int)>();
            byte lunchNum = 0;
            foreach (WeekActivity wAct in WaterActivities)
            {
                bool canHaveLunchRule = OffBlockRules.TryGetValue((day,wAct.Id).GetHashCode(), out List<byte> offBlocks);
                if (canHaveLunchRule)
                {
                    lunchNum = WeekInfo[day].LunchNumToTimeIndex.FirstOrDefault(x => offBlocks.Contains(x.Value)).Key;

                    if (lunchNum == 0) canHaveLunchRule = false;
                }

                if (!canHaveLunchRule)
                {
                    lunchNum = (byte)(Gen.Next(WeekInfo[day].LunchNumToTimeIndex.Count) + 1);
                    if (IsBookedInBlock(wAct.Id, day, WeekInfo[day].LunchNumToTimeIndex[lunchNum]))
                    {
                        lunchNum = (byte)((lunchNum % WeekInfo[day].LunchNumToTimeIndex.Count) + 1);
                    }
                }

                for (byte i = 0; i < WeekInfo[day].Times.Count(); i++)
                {
                    if (i == WeekInfo[day].LunchNumToTimeIndex[lunchNum] && wAct.IsSpecialist
                        || wAct.Open.Contains(day)
                        || OffBlockRules.ContainsKey((day,wAct.Id).GetHashCode()) && OffBlockRules[(day, wAct.Id).GetHashCode()].Contains(i)) continue;
                    waterActivityTimesAvailable.Add((wAct.Id, i));
                }
            }
            waterActivityTimesAvailable = new List<(byte, int)>(waterActivityTimesAvailable.OrderBy(_ => Gen.Next()));

            int WActNumOfGroupCombos = WActMaxNumofGroups * waterActivityTimesAvailable.Count;

            var UnitOpen = new[] { false, false };
            foreach (var specActPref in WeekInfo[day].SpecialActivityPrefs)
            {
                switch (specActPref.OpenPref)
                {
                    case 'b':
                        UnitOpen[0] = true;
                        UnitOpen[1] = true;
                        break;
                    case '1':
                        UnitOpen[0] = true;
                        break;
                    case '2':
                        UnitOpen[1] = true;
                        break;
                }
            }

            int failCount = 0;
            var ScheduledWaters = new List<(int TimeIndex, byte groupId, byte ActId)>();
            while (true)
            {
                ScheduledWaters.Clear();
                bool failed = false;

                int minGroupsIndex = Math.DivRem(failCount, waterActivityTimesAvailable.Count(), out int numOfMaxGroupsIndex);

                for (byte groupID = 0, availableIndex = 0; ; availableIndex++)
                {
                    while (groupID < Groups.Length && UnitOpen[Groups[groupID].Unit - 1]) groupID++;

                    if (groupID >= Groups.Length) break;

                    if (GroupIDsWithRuleWActs.Contains(groupID))
                    {
                        groupID++;
                        availableIndex--;
                        continue;
                    }

                    if (availableIndex == waterActivityTimesAvailable.Count)
                    {
                        failCount++;
                        failed = true;
                        break;
                    }

                    //int numOfGroupsIndex = (int)Math.Ceiling((float)(failCount < availableIndex ? 0 : failCount - availableIndex) / waterActivityTimesAvailable.Count());
                    int numOfGroupsIndex = minGroupsIndex;
                    if (Gen.Next(waterActivityTimesAvailable.Count() - availableIndex) < numOfMaxGroupsIndex)
                    {
                        numOfGroupsIndex++;
                        numOfMaxGroupsIndex--;
                    }

                    (byte wActID, int TimeIndex) = (0, 0);
                    //(byte firstActID, int FirstTime) = waterActivityTimesAvailable[availableIndex];
                    for (int i = 0; ; i++)
                    {
                        (wActID, TimeIndex) = waterActivityTimesAvailable[availableIndex];

                        var ScheduleCode = CanScheduleWater(wActID, day, TimeIndex, groupID, numOfGroupsIndex);

                        //Activity overlapped when it didn't need to
                        if (ScheduleCode == ScheduleActivityReturnCode.Overlapped && numOfGroupsIndex > minGroupsIndex && numOfMaxGroupsIndex + 1 < waterActivityTimesAvailable.Count() - availableIndex)
                        {
                            numOfMaxGroupsIndex++;
                            numOfGroupsIndex = minGroupsIndex;
                            ScheduleCode = CanScheduleWater(wActID, day, TimeIndex, groupID, numOfGroupsIndex);
                        }

                        //Activity failed to place, rotate to next possible activity 
                        if (ScheduleCode != ScheduleActivityReturnCode.Success)
                        {
                            var temp = waterActivityTimesAvailable[availableIndex];
                            waterActivityTimesAvailable.RemoveAt(availableIndex);
                            waterActivityTimesAvailable.Add(temp);
                            if (i >= waterActivityTimesAvailable.Count - availableIndex - 1)
                            {
                                failed = true;
                                break;
                            }
                            continue;
                        }
                        break;
                    }
                    if (failed)
                    {
                        failCount++;
                        break;
                    }

                    for (int i = 0; i < Activities[wActID].GetNumOfGroups(numOfGroupsIndex); i++)
                    {
                        ScheduledWaters.Add((TimeIndex, groupID, wActID));
                        groupID++;
                        if (groupID >= Groups.Length) break;
                    }
                }
                if (failCount > WActNumOfGroupCombos) throw new Exception($"Couldn't schedule water activities for all groups; try freeing up {day} schedule");
                if (!failed)
                {
                    break;
                }
            }
            foreach (var scheduledWater in ScheduledWaters)
            {
                ScheduleData[day][scheduledWater.TimeIndex, scheduledWater.groupId] = Activities[scheduledWater.ActId].Name;
                GroupByActivityCount[scheduledWater.groupId, scheduledWater.ActId]++;
            }
        }

        private void ScheduleWaterActivities()
        {
            foreach(var day in WeekInfo.Keys)
            {
                ScheduleWaterActivities(day);
            }
        }

        private ScheduleActivityReturnCode CanScheduleRegular(byte ActID, string day, int TimeIndex, byte startGroupID)
        {
            int endGroupID = startGroupID + Activities[ActID].NumofGroups[0]; //"prioritizes first num of groups"
            for (int groupID = startGroupID; groupID < endGroupID; groupID++)
            {
                var Act = Activities[ActID];
                if (groupID >= Groups.Count() || !string.IsNullOrEmpty(ScheduleData[day][TimeIndex, groupID]))
                {
                    return ScheduleActivityReturnCode.Overlapped;
                }
                if (Groups[groupID].SpecialGroup)
                {
                    return ScheduleActivityReturnCode.SpecialGroup;
                }
                if (Act.GradeStrike.Length > 0 && Act.GradeStrike.Contains(Groups[groupID].Grade))
                {
                    return ScheduleActivityReturnCode.GradeStriked;
                }
                if (Act.GradeOnly.Length > 0 && !Act.GradeOnly.Contains(Groups[groupID].Grade))
                {
                    return ScheduleActivityReturnCode.NotGradeOnly;
                }
                if (OffBlockRules.TryGetValue((day, ActID).GetHashCode(), out List<byte> offBlocks) && offBlocks.Contains((byte)TimeIndex))
                {
                    return ScheduleActivityReturnCode.OffBlock;
                }
                if (Act.Open.Contains(day) && WeekInfo[day].SpecialActivityPrefs[TimeIndex].OpenPref != 'n')
                {
                    return ScheduleActivityReturnCode.BookedOpen;
                }
                for (int i = 0; i < WeekInfo[day].Times.Length; i++)
                {
                    if (ScheduleData[day][i, groupID] == Act.Name)
                    {
                        return ScheduleActivityReturnCode.Duplicate;
                    }
                }
                if (IsBookedInBlock(ActID, day, TimeIndex)) return ScheduleActivityReturnCode.Duplicate;
            }

            return ScheduleActivityReturnCode.Success;
        }

        private void ScheduleRegularActivities(Dictionary<string,int[]> LunchNumsCount)
        {
            var BookableActInds = new List<byte>();
            var GroupInds = Enumerable.ToList(Enumerable.Range(0, Groups.Length));
            var BookableActivityToLunchNum = new Dictionary<byte, byte>();
            var OverflowActInds = new List<byte>();

            double numOfRegGroups = Groups.Sum(g => g.SpecialGroup ? 0 : 1);

            //Find regular non-overflow activities
            foreach (var Act in Activities)
            {
                if (Activities[Act.Id].WaterActivity) continue;
                if (Activities[Act.Id].Overflow)
                {
                    OverflowActInds.Add(Act.Id);
                    continue;
                }

                BookableActInds.Add(Act.Id);
            }

            var ActGroupScheduledCounts = new byte[Activities.Count, Groups.Length];

            //Schedule by day
            foreach (var day in WeekInfo.Values)
            {
                for (int i = 0; i < LunchNumsCount[day.DayName].Length; i++)
                {
                    LunchNumsCount[day.DayName][i] = (int)Math.Round(LunchNumsCount[day.DayName][i] / numOfRegGroups * NumOfSpecialists);
                }
                LunchNumsCount[day.DayName][LunchNumsCount[day.DayName].Length - 1]++; //prevent rounding error

                BookableActivityToLunchNum.Clear();

                //account for off block rules in selecting lunch
                foreach (var Act in Activities)
                {
                    if (!OffBlockRules.ContainsKey((day.DayName, Act.Id).GetHashCode()) || !Act.IsSpecialist) continue;

                    var lunchNum = WeekInfo[day.DayName].LunchNumToTimeIndex.FirstOrDefault(x => OffBlockRules[(day.DayName, Act.Id).GetHashCode()].Contains(x.Value)).Key;

                    if (lunchNum == 0) continue;

                    LunchNumsCount[day.DayName][lunchNum - 1]--;

                    BookableActivityToLunchNum.Add(Act.Id, lunchNum);
                }

                //select lunch for other activities
                foreach (byte ActId in new List<byte>(BookableActInds.OrderBy(_ => Gen.Next())))
                {
                    if (Activities[ActId].IsSpecialist && !BookableActivityToLunchNum.ContainsKey(ActId))
                    {
                        //randomly choose lunch for specialist based off of lunch counts
                        byte currentLunchNum = 1;

                        while (LunchNumsCount[day.DayName][currentLunchNum - 1] == 0 || IsBookedInBlock(ActId, day.DayName, WeekInfo[day.DayName].LunchNumToTimeIndex[currentLunchNum]))
                        {
                            currentLunchNum = (byte)(currentLunchNum % WeekInfo[day.DayName].LunchNumToTimeIndex.Count + 1);
                            if (currentLunchNum == 1) throw new Exception($"Couldn't give specialist for {Activities[ActId].Name} a lunch on {day.DayName}; check rules table to see if they were overbooked");
                        }

                        LunchNumsCount[day.DayName][currentLunchNum - 1]--;

                        BookableActivityToLunchNum.Add(ActId, currentLunchNum);
                    }
                }

                //schedule activitiy by block by group
                int currentBookableActIndInd;
                for (byte blockIndex = 0; blockIndex < WeekInfo[day.DayName].Times.Length; blockIndex++)
                {
                    //randomize activities and groups for each block
                    currentBookableActIndInd = 0;
                    BookableActInds = new List<byte>(BookableActInds.OrderBy(_ => Gen.Next()));
                    GroupInds = new List<int>(GroupInds.OrderBy(_ => Gen.Next()));

                    for (int GroupIndInd = 0; GroupIndInd < Groups.Length; GroupIndInd++)
                    {
                        var GroupInd = GroupInds[GroupIndInd];
                        var group = Groups[GroupInd];

                        if (group.SpecialGroup) continue;
                        if (!string.IsNullOrEmpty(ScheduleData[day.DayName][blockIndex, group.RowNum])) continue;

                        bool needsOverflow = false;
                        if (currentBookableActIndInd >= BookableActInds.Count)
                        {
                            needsOverflow = true;
                        }

                        var ActToSchedule = Activities[BookableActInds[currentBookableActIndInd >= BookableActInds.Count ? 0 : currentBookableActIndInd]];
                        ScheduleActivityReturnCode ScheduleCode = ScheduleActivityReturnCode.NotReturned;

                        if (!needsOverflow)
                        {
                            (int actIndInd, byte ScheduleCount) bestSchedule = (-1,255);
                            for (int i = currentBookableActIndInd; i < BookableActInds.Count; i++)
                            {
                                ScheduleCode = CanScheduleRegular(BookableActInds[i], day.DayName, blockIndex, group.RowNum);

                                if ((Activities[BookableActInds[i]].IsSpecialist && WeekInfo[day.DayName].LunchNumToTimeIndex[BookableActivityToLunchNum[BookableActInds[i]]] == blockIndex)
                                || ScheduleCode != ScheduleActivityReturnCode.Success)
                                {
                                    continue;
                                }

                                if(bestSchedule.actIndInd == -1 || bestSchedule.ScheduleCount > ActGroupScheduledCounts[BookableActInds[i],GroupInd])
                                {
                                    bestSchedule = (i, ActGroupScheduledCounts[BookableActInds[i], GroupInd]);
                                }
                            }

                            if (bestSchedule.actIndInd == -1) needsOverflow = true;
                            else if (bestSchedule.actIndInd != currentBookableActIndInd)
                            {
                                ActToSchedule = Activities[BookableActInds[bestSchedule.actIndInd]];

                                var temp = BookableActInds[currentBookableActIndInd];
                                BookableActInds[currentBookableActIndInd] = BookableActInds[bestSchedule.actIndInd];
                                BookableActInds[bestSchedule.actIndInd] = temp;
                            }
                        }

                        if (needsOverflow)
                        {
                            if (OverflowActInds.Count == 0) throw new Exception("Couldn't schedule all activities; please add an overflow activity");

                            var overflowId = OverflowActInds[Gen.Next(OverflowActInds.Count)];
                            ScheduleData[day.DayName][blockIndex, group.RowNum] = Activities[overflowId].Name;
                            GroupByActivityCount[group.RowNum, overflowId]++;
                        }
                        else
                        {
                            for (int i = 0; i < ActToSchedule.NumofGroups[0]; i++)
                            {
                                ScheduleData[day.DayName][blockIndex, group.RowNum + i] = ActToSchedule.Name;
                                ActGroupScheduledCounts[ActToSchedule.Id, group.RowNum + i]++;
                                GroupByActivityCount[group.RowNum + i, ActToSchedule.Id]++;
                            }
                            currentBookableActIndInd++;
                        }
                    }
                }
            }
        }

        public override void GenerateSchedule()
        {
            //Initialize Activity Counter
            GroupByActivityCount = new byte[Groups.Length, Activities.Count];

            var LunchNumsCount = new Dictionary<string, int[]>();

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
                LunchNumsCount.Add(dayInfo.DayName, new int[dayInfo.LunchNumToTimeIndex.Count]);
                foreach (Group group in Groups)
                {
                    if (!dayInfo.LunchNumToTimeIndex.TryGetValue(group.LunchNum, out byte timeIndex))
                    {
                        throw new Exception("Invalid Lunch Number entered in groups table; change groups table or add time to blocks table");
                    }
                    ScheduleData[dayInfo.DayName][timeIndex, group.RowNum] = "Lunch " + group.LunchNum;

                    LunchNumsCount[dayInfo.DayName][group.LunchNum - 1]++; //just by the by, this means that lunch num has to start at 1
                }

                //Rules Scheduling
                foreach (Rule rule in Rules[dayInfo.DayName])
                {
                    var actID = rule.ActIDs[Gen.Next(rule.ActIDs.Length)];
                    var timeIndex = rule.TimeIDs[Gen.Next(rule.TimeIDs.Length)];

                    foreach (var groupID in rule.GroupIDs)
                    {
                        ScheduleData[dayInfo.DayName][timeIndex, groupID] = Activities[actID].Name;
                        GroupByActivityCount[groupID, actID]++;

                        if (Activities[actID].WaterActivity)
                        {
                            GroupIDsWithRuleWActs.Add(groupID);
                        }
                    }
                }
            }

            //Water Scheduling
            ScheduleWaterActivities();

            //Regular Activity Scheduling
            ScheduleRegularActivities(LunchNumsCount);
        }

        public void OutputSchedule(Excel.Worksheet[] outputSheets, string[] takenSheetNames)
        {
            int i = 0;
            foreach (var dayInfo in WeekInfo.Values)
            {
                var outputRange = outputSheets[i].Range["A1", "Z100"];

                outputRange.Range["A1", (char)('A' + dayInfo.Times.Length) + "1"].Merge();
                outputRange.Cells[1, 1].Value2 = dayInfo.DayName;

                string baseName = $"{dayInfo.DayName} Output";
                string currentName = baseName;
                for (int j = 0; ; j++)
                {
                    currentName = j == 0 ? baseName : baseName + $" ({j})";
                    if (!takenSheetNames.Contains(currentName)) break;
                }
                outputSheets[i].Name = currentName;

                for (int column = 0; column < dayInfo.Times.Length; column++)
                {
                    outputRange.Cells[2, column + 2].Value2 = dayInfo.Times[column];
                    outputRange.Cells[3, column + 2].Value2 = column + 1;
                }

                for (int row = 0; row < Groups.Length; row++)
                {
                    outputRange.Cells[row + 4, 1].Value2 = Groups[row].Name;
                }

                for (int row = 0; row < Groups.Length; row++)
                {
                    for (int column = 0; column < dayInfo.Times.Length; column++)
                    {
                        outputRange.Cells[row + 4, column + 2].Value2 = ScheduleData[dayInfo.DayName][column, row];
                    }
                }

                outputRange.Columns.AutoFit();
                outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                outputRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputRange);

                i++;
            }
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(outputSheets);
        }

        public void OutputTally(Excel.Worksheet tallySheet, string[] takenSheetNames)
        {
            string bottomRightIndex;
            if (Activities.Count >= 25)
            {
                var quo = Math.DivRem(Activities.Count + 1, 26, out int rem);
                bottomRightIndex = (char)('A' + quo - 1) + "" + (char)('A' + rem) + (Activities.Count + 1).ToString();
            }
            else bottomRightIndex = (char)('A' + Activities.Count + 1) + (Groups.Length + 1).ToString();

            var outputRange = tallySheet.Range["A1", bottomRightIndex];

            string baseName = $"Activities Tally";
            string currentName = baseName;
            for (int i = 0; ; i++)
            {
                currentName = i == 0 ? baseName : baseName + $" ({i})";
                if (!takenSheetNames.Contains(currentName)) break;
            }
            tallySheet.Name = currentName;

            for (int column = 0; column < Activities.Count; column++)
            {
                outputRange.Cells[1, column + 2].Value2 = Activities[column].Name;
            }

            for (int row = 0; row < Groups.Length; row++)
            {
                outputRange.Cells[row + 2, 1].Value2 = Groups[row].Name;
            }

            for (int row = 0; row < Groups.Length; row++)
            {
                for (int column = 0; column < Activities.Count; column++)
                {
                    outputRange.Cells[row + 2, column + 2].Value2 = GroupByActivityCount[row, column].ToString();
                }
            }

            outputRange.Columns.AutoFit();
            outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            outputRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            ColorScale colorScale = outputRange.FormatConditions.AddColorScale(2);
            colorScale.ColorScaleCriteria[1].Type = XlConditionValueTypes.xlConditionValueLowestValue;
            colorScale.ColorScaleCriteria[1].FormatColor.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            colorScale.ColorScaleCriteria[2].Type = XlConditionValueTypes.xlConditionValueHighestValue;
            colorScale.ColorScaleCriteria[2].FormatColor.Color = Color.FromArgb(0, 158, 62);

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputRange);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tallySheet);
        }
    }
}

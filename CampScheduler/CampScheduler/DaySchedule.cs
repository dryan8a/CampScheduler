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

namespace CampScheduler
{
    public class DaySchedule : Schedule
    {
        public string[,] ScheduleData; //change to internal
        internal List<DayActivity> Activities { get; }
        internal List<DayActivity> WaterActivities { get; }

        internal DayInfo DayInfo { get; }

        internal List<Rule> Rules { get; }

        internal Dictionary<byte, List<byte>> OffBlockRules { get; }

        internal DaySchedule(Group[] groups, DayInfo dayInfo, Dictionary<Grade,byte> gradeToUnit) : base(groups, gradeToUnit)
        {
            ScheduleData = new string[dayInfo.Times.Length, groups.Length];

            Activities = new List<DayActivity>();
            WaterActivities = new List<DayActivity>();

            DayInfo = dayInfo;

            Rules = new List<Rule>();
            OffBlockRules = new Dictionary<byte, List<byte>>();
        }

        public void AddActivity(string name, bool waterActivity, bool overflow, byte[] numOfGroups, bool open, Grade[] gradeOnly, Grade[] gradeStrike, bool specialist)
        {
            DayActivity activity = new DayActivity((byte)Activities.Count, name, waterActivity, overflow, numOfGroups, open, gradeOnly, gradeStrike, specialist);
            Activities.Add(activity); 
            if (waterActivity)
            {
                WaterActivities.Add(activity);
                if (activity.NumofGroups.Length > WActMaxNumofGroups) WActMaxNumofGroups = activity.NumofGroups.Length;               
            }

            if (!waterActivity && specialist) NumOfSpecialists++;
            
        }
        
        public void AddRule(byte[] groupIDs, byte[] actIDs, byte[] timeIDs)
        {
            if(groupIDs.Length == 0)
            {
                foreach (var actID in actIDs)
                {
                    if (!OffBlockRules.ContainsKey(actID)) OffBlockRules.Add(actID, new List<byte>());

                    foreach (var timeID in timeIDs)
                    {
                        OffBlockRules[actID].Add(timeID);
                    }
                }
            }
            else Rules.Add(new Rule(groupIDs, actIDs, timeIDs));
        }
        
        public byte[] ParseActivities(string activityNamesInput)
        {
            var activityStrings = activityNamesInput.Split(',');
            byte[] activityIds = new byte[activityStrings.Length];
            for(int i = 0; i <  activityStrings.Length; i++)
            {
                activityIds[i] = Activities.First(act => act.Name.Equals(activityStrings[i].Trim())).Id;
            }
            return activityIds;
        }
        public byte[] ParseTimes(string timesInput)
        {
            var timesStrings = timesInput.Split(',');
            byte[] activityIds = new byte[timesStrings.Length];
            for (int i = 0; i < timesStrings.Length; i++)
            {
                activityIds[i] = (byte)Array.IndexOf(DayInfo.Times,timesStrings[i].Trim());
                if (activityIds[i] == 255) throw new Exception();
            }
            return activityIds;
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
            var grades = SchedulerParser.ParseGrades(gradePrefs);

            foreach(var group in Groups)
            {
                if(grades.Contains(group.Grade))
                {
                    ScheduleData[blockNum,group.RowNum] = ActivityName;
                }
            }
        }

        private bool IsBookedInBlock(byte ActID, int TimeIndex)
        {
            for (int i = 0; i < Groups.Length; i++)
            {
                if (ScheduleData[TimeIndex, i] == Activities[ActID].Name) return true;
            }

            return false;
        }

        private ScheduleActivityReturnCode CanScheduleWater(byte wActID, int TimeIndex, byte startGroupID, int numOfGroupsIndex)
        {
            int endGroupID = startGroupID + Activities[wActID].GetNumOfGroups(numOfGroupsIndex);
            for (byte groupID = startGroupID; groupID < endGroupID; groupID++)
            {
                if(groupID >= Groups.Count()) break;
                if(Activities[wActID].GradeOnly.Length > 0 && !Activities[wActID].GradeOnly.Contains(Groups[groupID].Grade))
                {
                    return ScheduleActivityReturnCode.NotGradeOnly;
                }
                if (Activities[wActID].GradeStrike.Length > 0 && Activities[wActID].GradeStrike.Contains(Groups[groupID].Grade))
                {
                    return ScheduleActivityReturnCode.GradeStriked;
                }  
                if (!string.IsNullOrEmpty(ScheduleData[TimeIndex,groupID])) //&& !string.Equals(ScheduleData[TimeIndex,groupID], "Open Activity"))
                {
                    return ScheduleActivityReturnCode.Overlapped;
                }
                if (Activities[wActID].Open && DayInfo.SpecialActivityPrefs[TimeIndex].OpenPref != 'n')
                {
                    return ScheduleActivityReturnCode.BookedOpen;
                }
                if (IsBookedInBlock(wActID, TimeIndex) || GroupIDsWithRuleWActs.Contains(groupID)) return ScheduleActivityReturnCode.Duplicate;
            }
            return ScheduleActivityReturnCode.Success;
        }

        private void ScheduleWaterActivities()
        {
            var waterActivityTimesAvailable = new List<(byte, int)>();
            byte lunchNum = 0;
            foreach (DayActivity wAct in WaterActivities)
            {
                bool canHaveLunchRule = OffBlockRules.TryGetValue(wAct.Id, out List<byte> offBlocks);
                if (canHaveLunchRule)
                {
                    lunchNum = DayInfo.LunchNumToTimeIndex.FirstOrDefault(x => offBlocks.Contains(x.Value)).Key;

                    if (lunchNum == 0) canHaveLunchRule = false;
                }

                if(!canHaveLunchRule)
                {
                    lunchNum = (byte)(Gen.Next(DayInfo.LunchNumToTimeIndex.Count) + 1);
                    if (IsBookedInBlock(wAct.Id, DayInfo.LunchNumToTimeIndex[lunchNum]))
                    {
                        lunchNum = (byte)((lunchNum + 1) % DayInfo.LunchNumToTimeIndex.Count);
                    }
                }

                for (byte i = 0; i < DayInfo.Times.Count(); i++)
                {
                    if (i == DayInfo.LunchNumToTimeIndex[lunchNum] && wAct.IsSpecialist 
                        || wAct.Open
                        || OffBlockRules.ContainsKey(wAct.Id) && OffBlockRules[wAct.Id].Contains(i)) continue;
                    waterActivityTimesAvailable.Add((wAct.Id, i));
                }
            }
            waterActivityTimesAvailable = new List<(byte, int)>(waterActivityTimesAvailable.OrderBy(_ => Gen.Next()));

            int WActNumOfGroupCombos = WActMaxNumofGroups * waterActivityTimesAvailable.Count;

            var UnitOpen = new [] { false, false };
            foreach(var specActPref in DayInfo.SpecialActivityPrefs)
            {
                switch(specActPref.OpenPref)
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

                for (byte groupID = 0, availableIndex = 0;; availableIndex++)
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
                    if(Gen.Next(waterActivityTimesAvailable.Count() - availableIndex) < numOfMaxGroupsIndex)
                    {
                        numOfGroupsIndex++;
                        numOfMaxGroupsIndex--;
                    }

                    (byte wActID, int TimeIndex) = (0, 0);
                    //(byte firstActID, int FirstTime) = waterActivityTimesAvailable[availableIndex];
                    for (int i = 0; ; i++)
                    {
                        (wActID, TimeIndex) = waterActivityTimesAvailable[availableIndex];

                        var ScheduleCode = CanScheduleWater(wActID, TimeIndex, groupID, numOfGroupsIndex);

                        //Activity overlapped when it didn't need to
                        if(ScheduleCode == ScheduleActivityReturnCode.Overlapped && numOfGroupsIndex > minGroupsIndex && numOfMaxGroupsIndex + 1 < waterActivityTimesAvailable.Count() - availableIndex)
                        {
                            numOfMaxGroupsIndex++;
                            numOfGroupsIndex = minGroupsIndex;
                            ScheduleCode = CanScheduleWater(wActID,TimeIndex, groupID, numOfGroupsIndex);
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
                if (failCount > WActNumOfGroupCombos) throw new Exception("Couldn't schedule water activities for all groups; try freeing up schedule");
                if (!failed)
                {
                    break;
                }
            }
            foreach (var scheduledWater in ScheduledWaters)
            {
                ScheduleData[scheduledWater.TimeIndex, scheduledWater.groupId] = Activities[scheduledWater.ActId].Name;
                GroupByActivityCount[scheduledWater.groupId, scheduledWater.ActId]++;
            }
        }

        private ScheduleActivityReturnCode CanScheduleRegular(byte ActID, int TimeIndex, byte startGroupID)
        {
            int endGroupID = startGroupID + Activities[ActID].NumofGroups[0]; //"prioritizes first num of groups"
            for (int groupID = startGroupID; groupID < endGroupID; groupID++)
            {
                var Act = Activities[ActID];
                if (groupID >= Groups.Count() || !string.IsNullOrEmpty(ScheduleData[TimeIndex, groupID]))
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
                if(OffBlockRules.TryGetValue(ActID, out List<byte> offBlocks) && offBlocks.Contains((byte)TimeIndex))
                {
                    return ScheduleActivityReturnCode.OffBlock;
                }
                if (Act.Open && DayInfo.SpecialActivityPrefs[TimeIndex].OpenPref != 'n')
                {
                    return ScheduleActivityReturnCode.BookedOpen;
                }
                for (int i = 0; i < DayInfo.Times.Length; i++)
                {
                    if (ScheduleData[i, groupID] == Act.Name)
                    {
                        return ScheduleActivityReturnCode.Duplicate;
                    }
                }
                if (IsBookedInBlock(ActID, TimeIndex)) return ScheduleActivityReturnCode.Duplicate;
            }

            return ScheduleActivityReturnCode.Success;
        }

        private void ScheduleRegularActivities(int[] LunchNumsCount)
        {
            var BookableActInds = new List<byte>();
            var GroupInds = Enumerable.ToList(Enumerable.Range(0,Groups.Length));
            var BookableActivityToLunchNum = new Dictionary<byte, byte>();
            var OverflowActInds = new List<byte>();

            double numOfRegGroups = Groups.Sum(g => g.SpecialGroup ? 0 : 1);

            for(int i = 0;i < LunchNumsCount.Length;i++)
            {
                LunchNumsCount[i] = (int)Math.Round(LunchNumsCount[i] / numOfRegGroups * NumOfSpecialists);
            }
            LunchNumsCount[LunchNumsCount.Length - 1]++; //prevent rounding error

            //account for off block rules in selecting lunch
            foreach (var actId in OffBlockRules.Keys)
            {
                if (Activities[actId].IsSpecialist)
                {
                    var lunchNum = DayInfo.LunchNumToTimeIndex.FirstOrDefault(x => OffBlockRules[actId].Contains(x.Value)).Key;

                    if (lunchNum == 0) continue;

                    LunchNumsCount[lunchNum - 1]--;

                    BookableActivityToLunchNum.Add(actId, lunchNum);
                }
            }

            foreach (byte ActId in new List<int>(Enumerable.ToList(Enumerable.Range(0, Activities.Count)).OrderBy(_ => Gen.Next())))
            {
                if (Activities[ActId].WaterActivity) continue;
                if (Activities[ActId].Overflow)
                {
                    OverflowActInds.Add(ActId);
                    continue;
                }

                BookableActInds.Add(ActId);

                if (Activities[ActId].IsSpecialist && !BookableActivityToLunchNum.ContainsKey(ActId))
                {
                    //randomly choose lunch for specialist based off of lunch counts
                    byte currentLunchNum = 1;

                    while(LunchNumsCount[currentLunchNum - 1] == 0 || IsBookedInBlock(ActId, DayInfo.LunchNumToTimeIndex[currentLunchNum]))
                    {
                        currentLunchNum = (byte)(currentLunchNum % DayInfo.LunchNumToTimeIndex.Count + 1);
                        if (currentLunchNum == 1) throw new Exception($"Couldn't give specialist for {Activities[ActId].Name} a lunch; check rules table to see if they were overbooked");
                    }

                    LunchNumsCount[currentLunchNum - 1]--;

                    BookableActivityToLunchNum.Add(ActId, currentLunchNum);
                }
            }

            int currentBookableActIndInd;
            for (byte blockIndex = 0; blockIndex < DayInfo.Times.Length; blockIndex++)
            {
                currentBookableActIndInd = 0;
                BookableActInds = new List<byte>(BookableActInds.OrderBy(_ => Gen.Next()));
                GroupInds = new List<int>(GroupInds.OrderBy(_ => Gen.Next()));

                for (int GroupIndInd = 0; GroupIndInd < Groups.Length; GroupIndInd++)
                {
                    var GroupInd = GroupInds[GroupIndInd];
                    var group = Groups[GroupInd];

                    if (group.SpecialGroup) continue;
                    if (!string.IsNullOrEmpty(ScheduleData[blockIndex, group.RowNum])) continue;

                    bool needsOverflow = false;
                    if(currentBookableActIndInd >= BookableActInds.Count)
                    {
                        needsOverflow = true;
                    }
                    var currentAct = Activities[BookableActInds[currentBookableActIndInd >= BookableActInds.Count ? 0 : currentBookableActIndInd]];
                    var originalBookableName = currentAct.Name;
                    ScheduleActivityReturnCode ScheduleCode = ScheduleActivityReturnCode.NotReturned;
                    while (!needsOverflow)
                    {
                        ScheduleCode = CanScheduleRegular(BookableActInds[currentBookableActIndInd], blockIndex, group.RowNum);
                        if ((currentAct.IsSpecialist && DayInfo.LunchNumToTimeIndex[BookableActivityToLunchNum[BookableActInds[currentBookableActIndInd]]] == blockIndex)
                            || ScheduleCode != ScheduleActivityReturnCode.Success)
                        {
                            var temp = BookableActInds[currentBookableActIndInd];
                            BookableActInds.RemoveAt(currentBookableActIndInd);
                            BookableActInds.Add(temp);

                            currentAct = Activities[BookableActInds[currentBookableActIndInd]];

                            if (originalBookableName == currentAct.Name)
                            {
                                needsOverflow = true;
                                break;
                            }
                            continue;
                        }
                        break;
                    }

                    if(!needsOverflow && ScheduleCode == ScheduleActivityReturnCode.SpecialGroup)
                    {
                        continue;
                    }
                    else if (needsOverflow)
                    {
                        if (OverflowActInds.Count == 0) throw new Exception("Couldn't schedule all activities; please add an overflow activity");

                        var overflowIndex = OverflowActInds[Gen.Next(OverflowActInds.Count)];
                        ScheduleData[blockIndex, group.RowNum] = Activities[overflowIndex].Name;
                        GroupByActivityCount[group.RowNum, overflowIndex]++;
                    }
                    else
                    {
                        for (int i = 0; i < currentAct.NumofGroups[0]; i++)
                        {
                            ScheduleData[blockIndex, group.RowNum + i] = currentAct.Name;
                            GroupByActivityCount[group.RowNum + i, currentAct.Id]++;
                        }
                        currentBookableActIndInd++;
                    }
                }
            }
        }

        public override void GenerateSchedule()
        {
            //Initialize Activity Counter
            GroupByActivityCount = new byte[Groups.Length, Activities.Count];

            //Special Activity Scheduling
            for (byte block = 0; block < DayInfo.Times.Length; block++)
            {
                ScheduleSpecialActivity(block, DayInfo.SpecialActivityPrefs[block].OpenPref, "Open Activity");

                ScheduleSpecialActivity(block, DayInfo.SpecialActivityPrefs[block].OpeningCirclePref, "Opening Circle");
                ScheduleSpecialActivity(block, DayInfo.SpecialActivityPrefs[block].MiddleCirclePref, "Middle Circle");
                ScheduleSpecialActivity(block, DayInfo.SpecialActivityPrefs[block].PopsicleTimePref, "Popsicle Time");
                ScheduleSpecialActivity(block, DayInfo.SpecialActivityPrefs[block].ClosingCirclePref, "Closing Circle");

                ScheduleSpecialActivity(block, DayInfo.SpecialActivityPrefs[block].SpecialEntPrefs, "Special Entertainment");
            }

            //Group Lunch Scheduling
            int[] LunchNumsCount = new int[DayInfo.LunchNumToTimeIndex.Count];
            foreach (Group group in Groups)
            {
                if (!DayInfo.LunchNumToTimeIndex.TryGetValue(group.LunchNum, out byte timeIndex))
                {
                    throw new Exception("Invalid Lunch Number entered in groups table; change groups table or add time to blocks table");
                }
                ScheduleData[timeIndex, group.RowNum] = "Lunch " + group.LunchNum;

                LunchNumsCount[group.LunchNum - 1]++; //just by the by, this means that lunch num has to start at 1
            }

            //Rules Scheduling
            foreach(Rule rule in Rules)
            {
                var actID = rule.ActIDs[Gen.Next(rule.ActIDs.Length)];
                var timeIndex = rule.TimeIDs[Gen.Next(rule.TimeIDs.Length)];

                foreach (var groupID in rule.GroupIDs)
                {
                    ScheduleData[timeIndex, groupID] = Activities[actID].Name;
                    GroupByActivityCount[groupID, actID]++;

                    if (Activities[actID].WaterActivity)
                    {
                        GroupIDsWithRuleWActs.Add(groupID);
                    }
                }
            }

            //Water Scheduling
            ScheduleWaterActivities();

            //Regular Activity Scheduling
            ScheduleRegularActivities(LunchNumsCount);
        }

        public void OutputSchedule(Excel.Worksheet outputSheet, string[] takenSheetNames)
        {
            var outputRange = outputSheet.Range["A1", "Z100"];

            outputRange.Range["A1", (char)('A' + DayInfo.Times.Length) + "1"].Merge();
            outputRange.Cells[1, 1].Value2 = DayInfo.DayName;

            string baseName = $"{DayInfo.DayName} Output";
            string currentName = baseName;
            for(int i = 0; ; i++)
            {
                currentName = i == 0 ? baseName : baseName + $" ({i})";
                if (!takenSheetNames.Contains(currentName)) break;
            }
            outputSheet.Name = currentName;

            for(int column = 0;column < DayInfo.Times.Length;column++)
            {
                outputRange.Cells[2, column + 2].Value2 = DayInfo.Times[column];
                outputRange.Cells[3, column + 2].Value2 = column + 1;
            }

            for (int row = 0; row < Groups.Length; row++)
            {
                outputRange.Cells[row + 4, 1].Value2 = Groups[row].Name;
            }

            for (int row = 0; row < Groups.Length; row++)
            {
                for (int column = 0; column < DayInfo.Times.Length; column++)
                {
                    outputRange.Cells[row+4, column+2].Value2 = ScheduleData[column, row];
                }
            }

            outputRange.Columns.AutoFit();
            outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            outputRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputRange);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheet);
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
            colorScale.ColorScaleCriteria[2].FormatColor.Color = Color.FromArgb(138, 255, 149);


            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputRange);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tallySheet);
        }

        public void OutputGroup(Excel.Worksheet outputSheet, string[] takenSheetNames, Group group)
        {
            string bottomRightIndex;
            if (Activities.Count >= 25)
            {
                var quo = Math.DivRem(Activities.Count + 1, 26, out int rem);
                bottomRightIndex = (char)('A' + quo - 1) + "" + (char)('A' + rem) + (Activities.Count + 1).ToString();
            }
            else bottomRightIndex = (char)('A' + Activities.Count + 1) + (Groups.Length + 1).ToString();

            var outputRange = outputSheet.Range["A1", bottomRightIndex];

            outputRange.Range["A1", (char)('A' + DayInfo.Times.Length) + "1"].Merge();
            outputRange.Cells[1, 1].Value2 = group.Name;

            string baseName = $"{group.Name} Output";
            string currentName = baseName;
            for (int i = 0; ; i++)
            {
                currentName = i == 0 ? baseName : baseName + $" ({i})";
                if (!takenSheetNames.Contains(currentName)) break;
            }
            outputSheet.Name = currentName;

            for (int column = 0; column < DayInfo.Times.Length; column++)
            {
                outputRange.Cells[2, column + 2].Value2 = DayInfo.Times[column];
                outputRange.Cells[3, column + 2].Value2 = column + 1;
            }

            outputRange.Cells[4, 1].Value2 = DayInfo.DayName;

            
            for (int column = 0; column < DayInfo.Times.Length; column++)
            {
                outputRange.Cells[4, column + 2].Value2 = ScheduleData[column, group.RowNum];
            }

            outputRange.Columns.AutoFit();
            outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            outputRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputRange);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheet);

        }
        
        public void OutputGroups(Excel.Worksheet[] outputSheets, string[] takenSheetNames)
        {
            for(int i = 0; i < Math.Min(outputSheets.Length, Groups.Length);i++)
            {
                OutputGroup(outputSheets[i], takenSheetNames, Groups[i]);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheets[i]);
            }
        }

    }
}

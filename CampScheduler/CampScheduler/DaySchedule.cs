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

namespace CampScheduler
{
    public class DaySchedule
    {
        public string[,] ScheduleData; //change to internal
        internal List<DayActivity> Activities { get; }
        internal List<DayActivity> WaterActivities { get; }
        private int WActMaxNumofGroups;

        private int NumOfSpecialists;

        private readonly Dictionary<byte, byte> LunchNumToTimeIndex;
        private readonly SpecialActivityPrefs[] SpecActPrefs;

        internal Group[] Groups { get; }
        internal Dictionary<Grade, byte> GradeToUnit { get; }


        internal List<Rule> Rules { get; }
        internal List<byte> GroupIDsWithRuleWActs { get; }

        internal string[] Times { get; }

        private Random Gen { get; }

        internal DaySchedule(int numOfBlocks, Group[] groups, string[] times, Dictionary<byte,byte> lunchNumToTimeIndex, Dictionary<Grade,byte> gradeToUnit, SpecialActivityPrefs[] specActPrefs)
        {
            ScheduleData = new string[numOfBlocks, groups.Length];

            Activities = new List<DayActivity>();
            WaterActivities = new List<DayActivity>();
            WActMaxNumofGroups = 0;  //try to fix this nonsense to make it a little faster

            NumOfSpecialists = 0;

            Groups = groups;

            LunchNumToTimeIndex = lunchNumToTimeIndex;
            Times = times;
            GradeToUnit = gradeToUnit;

            SpecActPrefs = specActPrefs;

            Gen = new Random();

            Rules = new List<Rule>();
            GroupIDsWithRuleWActs = new List<byte>();
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
            Rules.Add(new Rule(groupIDs, actIDs, timeIDs));
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
                activityIds[i] = (byte)Array.IndexOf(Times,timesStrings[i].Trim());
                if (activityIds[i] == 255) throw new Exception();
            }
            return activityIds;
        }

        public byte[] ParseGroupOrGrade(string groupOrGradeInput)
        {
            var groupOrGradeStrings = groupOrGradeInput.Split(',');
            List<byte> groupIds = new List<byte>();
            foreach(string groupOrGradeString in groupOrGradeStrings)
            {
                string groupOrGradeStringTrim = groupOrGradeString.Trim();

                Grade grade = SchedulerParser.ParseGrade(groupOrGradeStringTrim);
                if(grade == Grade.NA)
                {
                    byte groupID = (byte)Array.FindIndex(Groups, x => x.Name == groupOrGradeStringTrim);
                    if (groupID == 255) throw new Exception();
                    groupIds.Add(groupID);
                    continue;
                }

                for(byte i = 0; i < Groups.Length;i++) 
                {
                    if (Groups[i].Grade == grade && !groupIds.Contains(i)) groupIds.Add(i);
                }
            }
            return groupIds.ToArray();
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
                if (!string.IsNullOrEmpty(ScheduleData[TimeIndex,groupID]))
                {
                    return ScheduleActivityReturnCode.Overlapped;
                }
                if (Activities[wActID].Open && SpecActPrefs[TimeIndex].OpenPref != 'n')
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
            byte lunchNum;
            foreach (DayActivity wAct in WaterActivities)
            {
                lunchNum = (byte)(Gen.Next(LunchNumToTimeIndex.Count) + 1);
                if (IsBookedInBlock(wAct.Id, LunchNumToTimeIndex[lunchNum]))
                {
                    lunchNum = (byte)((lunchNum + 1) % LunchNumToTimeIndex.Count);
                }

                for (int i = 0; i < Times.Count(); i++)
                {
                    if (i == LunchNumToTimeIndex[lunchNum] && wAct.IsSpecialist) continue;
                    waterActivityTimesAvailable.Add((wAct.Id, i));
                }
            }
            waterActivityTimesAvailable = new List<(byte, int)>(waterActivityTimesAvailable.OrderBy(_ => Gen.Next()));

            int WActNumOfGroupCombos = WActMaxNumofGroups * waterActivityTimesAvailable.Count;

            var UnitOpen = new [] { false, false };
            foreach(var specActPref in SpecActPrefs)
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
            var ScheduledWaters = new List<(int TimeIndex, byte groupId, string ActName)>();
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
                        ScheduledWaters.Add((TimeIndex, groupID, Activities[wActID].Name));
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
                ScheduleData[scheduledWater.TimeIndex, scheduledWater.groupId] = scheduledWater.ActName;
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
                if (Act.Open && SpecActPrefs[TimeIndex].OpenPref != 'n')
                {
                    return ScheduleActivityReturnCode.BookedOpen;
                }
                for (int i = 0; i < Times.Length; i++)
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
            LunchNumsCount[LunchNumsCount.Length - 1]++;

            foreach (byte ActId in new List<int>(Enumerable.ToList(Enumerable.Range(0, Activities.Count)).OrderBy(_ => Gen.Next())))
            {
                if (Activities[ActId].WaterActivity) continue;
                if (Activities[ActId].Overflow)
                {
                    OverflowActInds.Add(ActId);
                    continue;
                }

                BookableActInds.Add(ActId);

                if (Activities[ActId].IsSpecialist)
                {
                    //randomly choose lunch for specialist based off of lunch counts
                    byte currentLunchNum = 1;
                    
                    while(LunchNumsCount[currentLunchNum - 1] == 0 || IsBookedInBlock(ActId, LunchNumToTimeIndex[currentLunchNum]))
                    {
                        currentLunchNum = (byte)(currentLunchNum % LunchNumToTimeIndex.Count + 1);
                        if (currentLunchNum == 1) throw new Exception($"Couldn't give specialist for {Activities[ActId].Name} a lunch; check rules table to see if they were overbooked");
                    }

                    LunchNumsCount[currentLunchNum - 1]--;

                    BookableActivityToLunchNum.Add(ActId, currentLunchNum);
                }
            }

            int currentBookableActIndInd;
            for (byte blockIndex = 0; blockIndex < Times.Length; blockIndex++)
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
                        if ((currentAct.IsSpecialist && LunchNumToTimeIndex[BookableActivityToLunchNum[BookableActInds[currentBookableActIndInd]]] == blockIndex)
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
                        ScheduleData[blockIndex, group.RowNum] = Activities[OverflowActInds[Gen.Next(OverflowActInds.Count)]].Name;
                    }
                    else
                    {
                        for (int i = 0; i < currentAct.NumofGroups[0]; i++)
                        {
                            ScheduleData[blockIndex, group.RowNum + i] = currentAct.Name;
                        }
                        currentBookableActIndInd += currentAct.NumofGroups[0];
                    }
                }
            }
        }

        public void GenerateSchedule()
        {
            //Special Activity Scheduling
            for (byte block = 0; block < Times.Length; block++)
            {
                ScheduleSpecialActivity(block, SpecActPrefs[block].OpenPref, "Open Activity");

                ScheduleSpecialActivity(block, SpecActPrefs[block].OpeningCirclePref, "Opening Circle");
                ScheduleSpecialActivity(block, SpecActPrefs[block].MiddleCirclePref, "Middle Circle");
                ScheduleSpecialActivity(block, SpecActPrefs[block].PopsicleTimePref, "Popsicle Time");
                ScheduleSpecialActivity(block, SpecActPrefs[block].ClosingCirclePref, "Closing Circle");

                ScheduleSpecialActivity(block, SpecActPrefs[block].SpecialEntPrefs, "Special Entertainment");
            }

            //Group Lunch Scheduling
            int[] LunchNumsCount = new int[LunchNumToTimeIndex.Count];
            foreach (Group group in Groups)
            {
                if (!LunchNumToTimeIndex.TryGetValue(group.LunchNum, out byte timeIndex))
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

        public void OutputSchedule(Excel.Range outputRange)
        {
            outputRange.Range["A1", (char)('A' + Times.Length) + "1"].Merge();
            outputRange.Cells[1,1].Value2 = "Day";

            for(int column = 0;column < Times.Length;column++)
            {
                outputRange.Cells[2, column + 2].Value2 = Times[column];
                outputRange.Cells[3, column + 2].Value2 = column + 1;
            }

            for (int row = 0; row < Groups.Length; row++)
            {
                outputRange.Cells[row + 4, 1].Value2 = Groups[row].Name;
            }

            for (int row = 0; row < Groups.Length; row++)
            {
                for (int column = 0; column < Times.Length; column++)
                {
                    outputRange.Cells[row+4, column+2].Value2 = ScheduleData[column, row];
                }
            }

            outputRange.Columns.AutoFit();
            outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            outputRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }
        
    }
}

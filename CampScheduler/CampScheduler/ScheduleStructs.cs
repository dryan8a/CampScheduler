using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

    public readonly struct DayActivity
    {
        public byte Id { get; }
        public string Name { get; }
        public bool WaterActivity { get; }
        public bool Overflow { get; }
        public byte[] NumofGroups { get; }
        public bool Open { get; }
        public Grade[] GradeOnly { get; }
        public Grade[] GradeStrike { get; }
        public bool IsSpecialist { get; }

        public DayActivity(byte id, string name, bool waterActivity, bool overflow, byte[] numOfGroups, bool open, Grade[] gradeOnly, Grade[] gradeStrike, bool specialist)
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

        public byte GetNumOfGroups(int index)
        {
            return NumofGroups[index >= NumofGroups.Length ? NumofGroups.Length - 1 : index];
        }
    }

    public readonly struct WeekActivity
    {
        public byte Id { get; }
        public string Name { get; }
        public bool WaterActivity { get; }
        public bool Overflow { get; }
        public byte[] NumofGroups { get; }
        public string[] Open { get; }
        public Grade[] GradeOnly { get; }
        public Grade[] GradeStrike { get; }
        public bool IsSpecialist { get; }

        public WeekActivity(byte id, string name, bool waterActivity, bool overflow, byte[] numOfGroups, string[] open, Grade[] gradeOnly, Grade[] gradeStrike, bool specialist)
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

        public byte GetNumOfGroups(int index)
        {
            return NumofGroups[index >= NumofGroups.Length ? NumofGroups.Length - 1 : index];
        }
    }

    public readonly struct Group
    {
        public byte RowNum { get; }
        public string Name { get; }
        public Grade Grade { get; }
        public byte Unit { get; }
        public bool SpecialGroup { get; }
        public byte LunchNum { get; }

        public Group(byte rowNum, string name, Grade grade, byte unit, bool specialGroup, byte lunchNum)
        {
            RowNum = rowNum;
            Name = name;
            Grade = grade;
            Unit = unit;
            SpecialGroup = specialGroup;
            LunchNum = lunchNum;
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

    public struct Rule
    {
        public byte[] GroupIDs;
        public byte[] ActIDs;
        public byte[] TimeIDs;

        public Rule(byte[] groupIDs, byte[] actIDs, byte[] timeIDs)
        {
            GroupIDs = groupIDs;
            ActIDs = actIDs;
            TimeIDs = timeIDs;
        }
    }

    public enum ScheduleActivityReturnCode
    {
        NotReturned,
        Success,
        GradeStriked,
        NotGradeOnly,
        Overlapped,
        Duplicate,
        BookedOpen,
        SpecialGroup
    }
}

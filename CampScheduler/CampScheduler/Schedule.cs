using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CampScheduler
{
    public abstract class Schedule
    {
        internal Random Gen { get; }

        internal Group[] Groups { get; }
        public int NumOfGroups => Groups.Length;

        internal Dictionary<Grade, byte> GradeToUnit { get; }

        internal Dictionary<string,byte> GroupNameToID { get; }

        internal Dictionary<string, byte> ActivityNameToID { get; }

        internal List<byte> GroupIDsWithRuleWActs { get; }

        internal int WActMaxNumofGroups;

        internal int NumOfSpecialists;

        internal byte[,] GroupByActivityCount { get; set; }

        public Schedule(Group[] groups, Dictionary<Grade,byte> gradeToUnit)
        {
            Gen = new Random();

            Groups = groups;

            GradeToUnit = gradeToUnit;

            GroupNameToID = new Dictionary<string, byte>();

            ActivityNameToID = new Dictionary<string, byte>();

            GroupIDsWithRuleWActs = new List<byte>();

            WActMaxNumofGroups = 0;  //try to fix this nonsense to make it a little faster

            NumOfSpecialists = 0;
        }

        public byte[] ParseGroupOrGrade(string groupOrGradeInput)
        {
            var groupOrGradeStrings = groupOrGradeInput.Split(',');
            List<byte> groupIds = new List<byte>();
            foreach (string groupOrGradeString in groupOrGradeStrings)
            {
                string groupOrGradeStringTrim = groupOrGradeString.Trim();

                Grade grade = SchedulerParser.ParseGrade(groupOrGradeStringTrim);
                if (grade == Grade.NA)
                {
                    byte groupID = (byte)Array.FindIndex(Groups, x => x.Name == groupOrGradeStringTrim);
                    if (groupID != 255) groupIds.Add(groupID);
                    continue;
                }

                for (byte i = 0; i < Groups.Length; i++)
                {
                    if (Groups[i].Grade == grade && !groupIds.Contains(i)) groupIds.Add(i);
                }
            }
            return groupIds.ToArray();
        }

        public abstract void GenerateSchedule();

        /// <summary>
        /// Add previous Tally Data to Schedule for generation. Note: please only use after finishing adding activities
        /// </summary>
        /// <param name="GroupName"></param>
        /// <param name="ActivityName"></param>
        /// <param name="Data"></param>
        /// <returns>Returns false when an invalid group or activity name is featured in tally</returns>
        public abstract bool InitTallyData(string GroupName, string ActivityName, byte Data);
    }
}

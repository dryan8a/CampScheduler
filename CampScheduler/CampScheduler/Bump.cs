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
    public class Bump
    {
        internal Random Gen = new Random();

        public List<byte>[,] BumpData;

        internal DayInfo DayInfo;

        internal List<BumpActivity> Activities;

        internal List<Counselor> Counselors;
        
        internal Bump(DayInfo dayInfo, List<BumpActivity> activities, List<Counselor> counselors)
        {
            DayInfo = dayInfo;
            Activities = activities;
            Counselors = counselors;

            BumpData = new List<byte>[dayInfo.Times.Length, activities.Count];
        }

        private bool IsBookedInBlock(byte CounselorId, int TimeIndex)
        {
            for (int i = 0; i < Activities.Count; i++)
            {
                if (BumpData[TimeIndex,i] != null && BumpData[TimeIndex, i].Contains(CounselorId)) return true;
            }

            return false;
        }

        private ScheduleBumpReturnCode CanScheduleCounselor(byte CounselorId, byte BlockIndex, int ActInd, short paidCount, short unpaidCount)
        {
            var Counselor = Counselors[CounselorId];
            var Activity = Activities[ActInd];

            if ((Counselor.Paid && paidCount <= 0) || (!Counselor.Paid && unpaidCount <= 0)) return ScheduleBumpReturnCode.IncorrectPay;
            if (Counselor.Handicapped && !Activity.IsAccessible) return ScheduleBumpReturnCode.Inaccessible;
            for (int i = 0; i < DayInfo.Times.Length; i++)
            {
                if (BumpData[i, ActInd] != null && BumpData[i, ActInd].Contains(CounselorId))
                {
                    return ScheduleBumpReturnCode.Duplicate;
                }
            }
            if (IsBookedInBlock(CounselorId, BlockIndex)) return ScheduleBumpReturnCode.Duplicate;

            return ScheduleBumpReturnCode.Success;
        }

        private ScheduleBumpReturnCode CanScheduleCounselor(byte CounselorId, byte BlockIndex, int ActInd, short paidCount, short unpaidCount, bool needM, bool needF)
        {
            var Counselor = Counselors[CounselorId];
            var Activity = Activities[ActInd];

            if ((Counselor.Paid && paidCount <= 0) || (!Counselor.Paid && unpaidCount <= 0)) return ScheduleBumpReturnCode.IncorrectPay;
            if ((Counselor.ChangingRoom == ChangingRoomCode.M && !needM) || (Counselor.ChangingRoom == ChangingRoomCode.F && !needF)) return ScheduleBumpReturnCode.WrongRoom;
            if (Counselor.Handicapped && !Activity.IsAccessible) return ScheduleBumpReturnCode.Inaccessible;
            for (int i = 0; i < DayInfo.Times.Length; i++)
            {
                if (BumpData[i,ActInd] != null && BumpData[i, ActInd].Contains(CounselorId))
                {
                    return ScheduleBumpReturnCode.Duplicate;
                }
            }
            if (IsBookedInBlock(CounselorId, BlockIndex)) return ScheduleBumpReturnCode.Duplicate;

            return ScheduleBumpReturnCode.Success;
        }


        public void GenerateBump()
        {
            List<byte> BookableCounInds;
            var OverflowCounInds = new List<byte>();

            List<BumpActivity> OverflowJobs = new List<BumpActivity>();

            int currentBookableCounIndInd;
            for (byte blockIndex = 0; blockIndex < DayInfo.Times.Length; blockIndex++)
            {
                currentBookableCounIndInd = 0;
                BookableCounInds = new List<int>(Enumerable.ToList(Enumerable.Range(0, Counselors.Count)).OrderBy(_ => Gen.Next())).ConvertAll(x => (byte)x);

                BumpActivity Act;
                short paidCount;
                short unpaidCount;

                //Priority (changing room) Scheduling
                bool needM, needF = false;
                for(int ActInd = 0; ActInd < Activities.Count; ActInd++)
                {
                    Act = Activities[ActInd];

                    paidCount = Act.NumPaid;
                    unpaidCount = Act.NumUnpaid;

                    BumpData[blockIndex, ActInd] = new List<byte>();

                    switch (Act.Required)
                    {
                        case ChangingRoomCode.F:
                            needF = true;
                            needM = false;
                            break;
                        case ChangingRoomCode.M:
                            needF = false;
                            needM = true;
                            break;
                        case ChangingRoomCode.B:
                            needF = true;
                            needM = true;
                            break;
                        default:
                            continue;
                    }

                    while (needF || needM)
                    {
                        if (currentBookableCounIndInd >= BookableCounInds.Count)
                        {
                            throw new Exception("Not enough Counselors to fully fill bump; try adding more, or reducing required staff");
                        }

                        var currentCoun = Counselors[BookableCounInds[currentBookableCounIndInd >= BookableCounInds.Count ? 0 : currentBookableCounIndInd]];
                        var originalBookableName = currentCoun.Name;

                        ScheduleBumpReturnCode CanSchedule;
                        while (true)
                        {
                            CanSchedule = CanScheduleCounselor(BookableCounInds[currentBookableCounIndInd], blockIndex, ActInd, paidCount, unpaidCount, needM, needF);
                            if (CanSchedule != ScheduleBumpReturnCode.Success) //|| DayInfo.LunchNumToTimeIndex[currentCoun.LunchNum] == blockIndex)
                            {
                                var temp = BookableCounInds[currentBookableCounIndInd];
                                BookableCounInds.RemoveAt(currentBookableCounIndInd);
                                BookableCounInds.Add(temp);

                                currentCoun = Counselors[BookableCounInds[currentBookableCounIndInd]];

                                if (originalBookableName == currentCoun.Name)
                                {
                                    throw new Exception("Criteria too strict to schedule bump; likely not enough paid/unpaid staff");
                                }
                                continue;
                            }
                            break;
                        }

                        BumpData[blockIndex, ActInd].Add(currentCoun.Id);

                        if (currentCoun.ChangingRoom == ChangingRoomCode.M) needM = false;
                        else needF = false;

                        if (currentCoun.Paid) paidCount--;
                        else unpaidCount--;

                        currentBookableCounIndInd++;
                    }


                }

                //Regular Scheduling
                for (int ActInd = 0; ActInd < Activities.Count;ActInd++)
                {
                    Act = Activities[ActInd];

                    if (Act.Overflow)
                    {
                        OverflowJobs.Add(Act);
                        continue;
                    }

                    paidCount = Act.NumPaid;
                    unpaidCount = Act.NumUnpaid;

                    foreach(byte CounId in BumpData[blockIndex,ActInd])
                    {
                        if (Counselors[CounId].Paid) paidCount--;
                        else unpaidCount--;
                    }

                    while (paidCount > 0 || unpaidCount > 0)
                    {
                        if (currentBookableCounIndInd >= BookableCounInds.Count)
                        {
                            throw new Exception("Not enough Counselors to fully fill bump; try adding more, or reducing required staff");
                        }

                        var currentCoun = Counselors[BookableCounInds[currentBookableCounIndInd >= BookableCounInds.Count ? 0 : currentBookableCounIndInd]];
                        var originalBookableName = currentCoun.Name;

                        ScheduleBumpReturnCode CanSchedule;
                        while (true)
                        {
                            CanSchedule = CanScheduleCounselor(BookableCounInds[currentBookableCounIndInd], blockIndex, ActInd, paidCount, unpaidCount);
                            if (CanSchedule != ScheduleBumpReturnCode.Success) //|| DayInfo.LunchNumToTimeIndex[currentCoun.LunchNum] == blockIndex)
                            {
                                var temp = BookableCounInds[currentBookableCounIndInd];
                                BookableCounInds.RemoveAt(currentBookableCounIndInd);
                                BookableCounInds.Add(temp);

                                currentCoun = Counselors[BookableCounInds[currentBookableCounIndInd]];

                                if (originalBookableName == currentCoun.Name)
                                {
                                    throw new Exception("Criteria too strict to schedule bump; likely not enough paid/unpaid staff");
                                }
                                continue;
                            }
                            break;
                        }

                        BumpData[blockIndex, ActInd].Add(currentCoun.Id);

                        if (currentCoun.Paid) paidCount--;
                        else unpaidCount--;

                        currentBookableCounIndInd++;
                    }
                }

                //Overflow job booking
                int counsPerOverflow = (BookableCounInds.Count - currentBookableCounIndInd) / OverflowJobs.Count;

                for(int overflowInd = 0; overflowInd < OverflowJobs.Count; overflowInd++)
                {
                    if(overflowInd == OverflowJobs.Count - 1)
                    {
                        while (currentBookableCounIndInd < BookableCounInds.Count)
                        {
                            BumpData[blockIndex, OverflowJobs[overflowInd].Id].Add(BookableCounInds[currentBookableCounIndInd]);
                            currentBookableCounIndInd++;
                        }
                        break;
                    }
                    
                    for(int i = 0; i < counsPerOverflow; i++)
                    {
                        BumpData[blockIndex, OverflowJobs[overflowInd].Id].Add(BookableCounInds[currentBookableCounIndInd]);
                        currentBookableCounIndInd++;
                    }
                }
            }
        }

        public void OutputBump(Excel.Worksheet outputSheet, string[] takenSheetNames)
        {
            var outputRange = outputSheet.Range["A1", "Z100"];

            outputRange.Range["A1", (char)('A' + DayInfo.Times.Length) + "1"].Merge();
            outputRange.Cells[1, 1].Value2 = DayInfo.DayName;

            string baseName = $"{DayInfo.DayName} Bump Output";
            string currentName = baseName;
            for (int i = 0; ; i++)
            {
                currentName = i == 0 ? baseName : baseName + $" ({i})";
                if (!takenSheetNames.Contains(currentName)) break;
            }
            outputSheet.Name = currentName;


            int currentRow = 1;
            int currentCol = 2;
            StringBuilder paid, unpaid;
            Counselor currentCoun;
            for (int timeInd = 0; timeInd < DayInfo.Times.Length; timeInd++)
            {
                currentRow++;
                outputRange.Cells[currentRow++,2].Value2 = DayInfo.Times[timeInd];
                for (int ActInd = 0; ActInd < Activities.Count; ActInd++)
                {
                    if (currentCol > 7)
                    {
                        currentCol = 2;
                        currentRow += 3;
                    }
                    
                    paid = new StringBuilder();
                    unpaid = new StringBuilder();
                    
                    for(int i = 0; i < BumpData[timeInd,ActInd].Count;i++)
                    {
                        currentCoun = Counselors[BumpData[timeInd, ActInd][i]];
                        if (currentCoun.Paid)
                        {
                            if(paid.Length != 0) paid.Append(", ");
                            paid.Append(currentCoun.Name);
                        }
                        else
                        {
                            if (unpaid.Length != 0) unpaid.Append(", ");
                            unpaid.Append(currentCoun.Name);
                        }
                    }

                    outputRange.Cells[currentRow, currentCol].Value = Activities[ActInd].Name;
                    outputRange.Cells[currentRow + 1, currentCol].Value = paid.ToString();
                    outputRange.Cells[currentRow + 2, currentCol].Value = unpaid.ToString();

                    currentCol++;
                }

                currentCol = 2;
                currentRow += 3;
            }

            outputRange.Columns.AutoFit();
            outputRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            outputRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputRange);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheet);
        }
    }
}

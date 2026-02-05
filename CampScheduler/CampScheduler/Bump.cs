using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CampScheduler
{
    public class Bump
    {
        internal Random Gen = new Random();

        public List<string>[,] BumpData;

        internal DayInfo DayInfo;

        internal List<BumpActivity> Activities;

        internal List<Counselor> Counselors;
        
        internal Bump(DayInfo dayInfo, List<BumpActivity> activities, List<Counselor> counselors)
        {
            DayInfo = dayInfo;
            Activities = activities;
            Counselors = counselors;

            BumpData = new List<string>[dayInfo.Times.Length, activities.Count];
        }

        private bool IsBookedInBlock(byte CounselorId, int TimeIndex)
        {
            for (int i = 0; i < Activities.Count; i++)
            {
                if (BumpData[TimeIndex,i] != null && BumpData[TimeIndex, i].Contains(Counselors[CounselorId].Name)) return true;
            }

            return false;
        }

        private ScheduleBumpReturnCode CanScheduleCounselor(byte CounselorId, byte BlockIndex, int ActInd, short paidCount, short unpaidCount)
        {
            var Counselor = Counselors[CounselorId];
            var Activity = Activities[ActInd];

            if ((Counselor.Paid && paidCount <= 0) || (!Counselor.Paid && unpaidCount <= 0)) return ScheduleBumpReturnCode.IncorrectPay;
            if (Counselor.Handicapped && !Activity.IsAccessible) return ScheduleBumpReturnCode.Inaccessible;
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

                for (int ActInd = 0; ActInd < Activities.Count;ActInd++)
                {
                    var Act = Activities[ActInd];

                    BumpData[blockIndex, ActInd] = new List<string>();

                    if (Act.Overflow)
                    {
                        OverflowJobs.Add(Act);
                        continue;
                    }

                    short paidCount = Act.NumPaid;
                    short unpaidCount = Act.NumUnpaid;

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

                        BumpData[blockIndex, ActInd].Add(currentCoun.Name);

                        if (currentCoun.Paid) paidCount--;
                        else unpaidCount--;

                        currentBookableCounIndInd++;
                    }
                }

                int counsPerOverflow = (BookableCounInds.Count - currentBookableCounIndInd) / OverflowJobs.Count;

                for(int overflowInd = 0; overflowInd < OverflowJobs.Count; overflowInd++)
                {
                    if(overflowInd == OverflowJobs.Count - 1)
                    {
                        while (currentBookableCounIndInd < BookableCounInds.Count)
                        {
                            BumpData[blockIndex, OverflowJobs[overflowInd].Id].Add(Counselors[BookableCounInds[currentBookableCounIndInd]].Name);
                            currentBookableCounIndInd++;
                        }
                        break;
                    }
                    
                    for(int i = 0; i < counsPerOverflow; i++)
                    {
                        BumpData[blockIndex, OverflowJobs[overflowInd].Id].Add(Counselors[BookableCounInds[currentBookableCounIndInd]].Name);
                        currentBookableCounIndInd++;
                    }
                }
            }
        }

        
    }
}

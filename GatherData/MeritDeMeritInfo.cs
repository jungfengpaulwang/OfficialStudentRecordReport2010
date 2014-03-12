using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using SHSchool.Data;

namespace OfficialStudentRecordReport2010
{
    public class MeritDeMeritInfo : IDisposable
    {
        //  學生獎勵資料
        private Dictionary<string, List<MeritRecord>> _PersonalMeritRecordInfo;
        //  學生懲戒資料
        private Dictionary<string, List<DemeritRecord>> _PersonalDemeritRecordInfo;
        //  功過換算對照表
        MeritDemeritReduceRecord _MeritDemeritReduceRecord;
        //  大過換算為小過
        int? _DemeritAToDemeritB;
        //  小過換算為警告
        int? _DemeritBToDemeritC;
        //  大功換算為小功
        int? _MeritAToMeritB;
        //  小功換算為嘉獎
        int? _MeritBToMeritC;

        public MeritDeMeritInfo(List<SHStudentRecord> pStudents)
        {
            _PersonalMeritRecordInfo = new Dictionary<string, List<MeritRecord>>();
            _PersonalDemeritRecordInfo = new Dictionary<string, List<DemeritRecord>>();
            _MeritDemeritReduceRecord = MeritDemeritReduce.Select();

            _DemeritAToDemeritB = _MeritDemeritReduceRecord.DemeritAToDemeritB.HasValue ? _MeritDemeritReduceRecord.DemeritAToDemeritB.Value : 0;
            _DemeritBToDemeritC = _MeritDemeritReduceRecord.DemeritBToDemeritC.HasValue ? _MeritDemeritReduceRecord.DemeritBToDemeritC.Value : 0;
            _MeritAToMeritB = _MeritDemeritReduceRecord.MeritAToMeritB.HasValue ? _MeritDemeritReduceRecord.MeritAToMeritB.Value : 0;
            _MeritBToMeritC = _MeritDemeritReduceRecord.MeritBToMeritC.HasValue ? _MeritDemeritReduceRecord.MeritBToMeritC.Value : 0;

            ProduceData(pStudents);
        }

        public void Dispose()
        {
            _PersonalMeritRecordInfo.GetEnumerator().Dispose();
            _PersonalDemeritRecordInfo.GetEnumerator().Dispose();
            _MeritDemeritReduceRecord = null;

            GC.SuppressFinalize(this);
        }

        private void ProduceData(List<SHStudentRecord> pStudents)
        {                  
            List<MeritRecord> meritRecords = Merit.SelectByStudentIDs(pStudents.Select(x => x.ID));
            foreach (MeritRecord mr in meritRecords)
            {
                if (!_PersonalMeritRecordInfo.ContainsKey(mr.RefStudentID))
                    _PersonalMeritRecordInfo.Add(mr.RefStudentID, new List<MeritRecord>());

                _PersonalMeritRecordInfo[mr.RefStudentID].Add(mr);
            }

            List<DemeritRecord> demeritRecords = Demerit.SelectByStudentIDs(pStudents.Select(x => x.ID));
            foreach (DemeritRecord dr in demeritRecords)
            {
                if (!_PersonalDemeritRecordInfo.ContainsKey(dr.RefStudentID))
                    _PersonalDemeritRecordInfo.Add(dr.RefStudentID, new List<DemeritRecord>());

                _PersonalDemeritRecordInfo[dr.RefStudentID].Add(dr);
            }
        }

        public bool NotExceedThreeMajorDemerits(string pStudentID)
        {
            //  大功
            int? pMeritA = 0;
            //  小功
            int? pMeritB = 0;
            //  嘉獎
            int? pMeritC = 0;
            //  大過
            int? pDemeritA = 0;
            //  小過
            int? pDemeritB = 0;
            //  警告
            int? pDemeritC = 0;

            //  功的累加
            int? pAccumulatedMerit = 0;
            //  過的累加
            int? pAccumulatedDemerit = 0;

            if (_PersonalMeritRecordInfo.ContainsKey(pStudentID))
            {
                foreach (MeritRecord mr in _PersonalMeritRecordInfo[pStudentID])
                {
                    pMeritA = mr.MeritA.HasValue ? mr.MeritA.Value : 0;
                    pMeritB = mr.MeritB.HasValue ? mr.MeritB.Value : 0;
                    pMeritC = mr.MeritC.HasValue ? mr.MeritC.Value : 0;

                    pAccumulatedMerit += pMeritA * _MeritAToMeritB * _MeritBToMeritC + pMeritB * _MeritBToMeritC + pMeritC;
                }
            }

            if (_PersonalDemeritRecordInfo.ContainsKey(pStudentID))
            {
                foreach (DemeritRecord dr in _PersonalDemeritRecordInfo[pStudentID])
                {
                    if (dr.Cleared == "是")
                        continue;

                    pDemeritA = dr.DemeritA.HasValue ? dr.DemeritA.Value : 0;
                    pDemeritB = dr.DemeritB.HasValue ? dr.DemeritB.Value : 0;
                    pDemeritC = dr.DemeritC.HasValue ? dr.DemeritC.Value : 0;

                    pAccumulatedDemerit += pDemeritA * _DemeritAToDemeritB * _DemeritBToDemeritC + pDemeritB * _DemeritBToDemeritC + pDemeritC;
                }
            }
            //  剛好三大過也算喔
            if ((pAccumulatedDemerit - pAccumulatedMerit) >= (3 * (_DemeritAToDemeritB * _DemeritBToDemeritC)))
                return false;
            else
                return true;
        }
    }
}
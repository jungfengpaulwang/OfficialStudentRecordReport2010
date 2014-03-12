using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHSchool.Data;
using System.Xml.Linq;
using K12.Data;

namespace OfficialStudentRecordReport2010
{
    #region 課程學年成績資訊之 ValueObject
    /// <summary>
    /// 課程成績資訊之 ValueObject
    /// </summary>
    public class SHSubjectYearScoreInfo
    {
        public SHSubjectYearScoreInfo() { }

        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { set; get; }
        /// <summary>
        /// 學年度
        /// </summary>
        public int SchoolYear { set; get; }
        /// <summary>
        /// 科目名稱
        /// </summary>  
        public string SubjectName { set; get; }
        /// <summary>
        /// 原始成績
        /// </summary>
        public decimal? Score { set; get; }
        /// <summary>
        /// 學年實得學分=上、下學期獲得學分之加總
        /// </summary>
        public decimal? Credit { set; get; }
        /// <summary>
        /// 學年應得學分
        /// </summary>
        public decimal? AccumulatedCredit { set; get; }
        /// <summary>
        /// 學年平均課時
        /// </summary>
        public string Hour { set; get; }
        /// <summary>
        /// 成績年級
        /// </summary>
        public int? GradeYear { set; get; } 
    }
    #endregion

    #region 取得學生學年科目成績內容之物件
    public class SHYearSubjectScoreInfo : IDisposable
    {            
        //  學生科目學年成績資料
        private Dictionary<string, SHSubjectYearScoreInfo> _PersonalSubjectYearScoreInfo;
        //  學生學年科目成績資料
        private Dictionary<string, List<SHSubjectYearScoreInfo>> _PersonalYearSubjectScoreInfo;
        //  所有學生學年歷程資料
        private Dictionary<string, Dictionary<int, int>> _PersonalYearSubjectScoreHistoryInfo;
        //  單一學生重讀學年
        private Dictionary<string, List<int>> _PersonalReReadSchoolYear;

        public void Dispose()
        {
            _PersonalSubjectYearScoreInfo.GetEnumerator().Dispose();
            _PersonalYearSubjectScoreInfo.GetEnumerator().Dispose();
            _PersonalYearSubjectScoreHistoryInfo.GetEnumerator().Dispose();
            _PersonalReReadSchoolYear.GetEnumerator().Dispose();

            GC.SuppressFinalize(this);
        }

        public enum Entry
        {
            學業, 體育, 實習科目, 國防通識, 健康與護理, 專業科目
        }

        public SHYearSubjectScoreInfo(List<SHStudentRecord> pStudents)
        {
            _PersonalSubjectYearScoreInfo = new Dictionary<string, SHSubjectYearScoreInfo>();
            _PersonalYearSubjectScoreInfo = new Dictionary<string, List<SHSubjectYearScoreInfo>>();
            _PersonalYearSubjectScoreHistoryInfo = new Dictionary<string, Dictionary<int, int>>();
            _PersonalReReadSchoolYear = new Dictionary<string, List<int>>();

            ProduceData(pStudents);
        }

        public Dictionary<int, int> GetPersonalYearSubjectScoreHistoryInfo(string pStudentID)
        {
            if (_PersonalYearSubjectScoreHistoryInfo.ContainsKey(pStudentID))
                return _PersonalYearSubjectScoreHistoryInfo[pStudentID];
            else
                return new Dictionary<int,int>();
        }

        public List<int> GetPersonalReReadSchoolYear(string pStudentID)
        {
            if (_PersonalReReadSchoolYear.ContainsKey(pStudentID))
                return _PersonalReReadSchoolYear[pStudentID];
            else
                return new List<int>();
        }

        #region 取得個人學年科目成績資訊
        /// <summary>
        /// 由個人學期科目成績資訊取得個人學年科目成績資訊
        /// </summary>
        /// <param name="pSHSubject1stSemesterScoreInfo">個人上學期科目成績資訊</param>
        /// <param name="pSHSubject2ndSemesterScoreInfo">個人下學期科目成績資訊</param>
        /// <returns>個人學年科目成績資訊</returns>
        public List<SHSubjectYearScoreInfo> GetPersonalYearSubjectScoreInfo(Dictionary<int, List<SHSubjectSemesterScoreInfo>> dicSHSubjectSemesterScoreInfos)
        {
            List<SHSubjectSemesterScoreInfo> pSHSubject1stSemesterScoreInfo = new List<SHSubjectSemesterScoreInfo>();
            List<SHSubjectSemesterScoreInfo> pSHSubject2ndSemesterScoreInfo = new List<SHSubjectSemesterScoreInfo>();

            if (dicSHSubjectSemesterScoreInfos.ContainsKey(1))
                pSHSubject1stSemesterScoreInfo = dicSHSubjectSemesterScoreInfos[1];
            if (dicSHSubjectSemesterScoreInfos.ContainsKey(2))
                pSHSubject2ndSemesterScoreInfo = dicSHSubjectSemesterScoreInfos[2];

             if (pSHSubject1stSemesterScoreInfo.Count == 0 && pSHSubject2ndSemesterScoreInfo.Count == 0)
                return null;

            if (pSHSubject1stSemesterScoreInfo.Count == 0 && pSHSubject2ndSemesterScoreInfo.Count > 0)
                return this.GetPersonalYearSubjectScoreInfo(pSHSubject2ndSemesterScoreInfo);

            if (pSHSubject2ndSemesterScoreInfo.Count == 0 && pSHSubject1stSemesterScoreInfo.Count > 0)
                return this.GetPersonalYearSubjectScoreInfo(pSHSubject1stSemesterScoreInfo);

            List<SHSubjectYearScoreInfo> _SHSubjectYearScoreInfo = new List<SHSubjectYearScoreInfo>();

            int i = -1;
            do
            {
                i++;
                SHSubjectSemesterScoreInfo _SHSubjectSemesterScoreInfo;

                if ((pSHSubject1stSemesterScoreInfo.Count <= i) && (pSHSubject2ndSemesterScoreInfo.Count <= i))
                    break;

                //if (pSHSubject1stSemesterScoreInfo.Count > i)
                //    _SHSubjectSemesterScoreInfo = pSHSubject1stSemesterScoreInfo[i];
                //else 
                //    _SHSubjectSemesterScoreInfo = pSHSubject2ndSemesterScoreInfo[i];

                if (pSHSubject2ndSemesterScoreInfo.Count > i)
                    _SHSubjectSemesterScoreInfo = pSHSubject2ndSemesterScoreInfo[i];
                else
                    _SHSubjectSemesterScoreInfo = pSHSubject1stSemesterScoreInfo[i];

                string keySingleSubject = _SHSubjectSemesterScoreInfo.StudentID + "_" + _SHSubjectSemesterScoreInfo.SchoolYear.ToString() + "_" + _SHSubjectSemesterScoreInfo.SubjectName;

                SHSubjectYearScoreInfo pSHSubjectYearScoreInfo = this.GetPersonalSubjectYearScoreInfo(_SHSubjectSemesterScoreInfo.StudentID, _SHSubjectSemesterScoreInfo.SchoolYear, _SHSubjectSemesterScoreInfo.SubjectName, _SHSubjectSemesterScoreInfo.GradeYear);

                if (pSHSubjectYearScoreInfo == null)
                {
                    if (pSHSubject1stSemesterScoreInfo.Count > i)
                        pSHSubjectYearScoreInfo = this.GetPersonalSubjectYearScoreInfo(pSHSubject1stSemesterScoreInfo[i].StudentID, pSHSubject1stSemesterScoreInfo[i].SchoolYear, pSHSubject1stSemesterScoreInfo[i].SubjectName, pSHSubject1stSemesterScoreInfo[i].GradeYear);
                    if (pSHSubjectYearScoreInfo == null)               
                    {
                        if (pSHSubject1stSemesterScoreInfo.Count > i && pSHSubject2ndSemesterScoreInfo.Count > i)
                            pSHSubjectYearScoreInfo = this.GetPersonalSubjectYearScoreInfo(pSHSubject1stSemesterScoreInfo[i].StudentID, pSHSubject2ndSemesterScoreInfo[i].SchoolYear, pSHSubject1stSemesterScoreInfo[i].SubjectName, pSHSubject1stSemesterScoreInfo[i].GradeYear);
                        if (pSHSubjectYearScoreInfo == null)
                        {
                            pSHSubjectYearScoreInfo = new SHSubjectYearScoreInfo();

                            pSHSubjectYearScoreInfo.StudentID = _SHSubjectSemesterScoreInfo.StudentID;
                            pSHSubjectYearScoreInfo.SchoolYear = _SHSubjectSemesterScoreInfo.SchoolYear;
                            pSHSubjectYearScoreInfo.SubjectName = _SHSubjectSemesterScoreInfo.SubjectName;
                            //pSHSubjectYearScoreInfo.Score = _SHSubjectSemesterScoreInfo.Score;
                            //pSHSubjectYearScoreInfo.Credit = _SHSubjectSemesterScoreInfo.StudentID;
                            pSHSubjectYearScoreInfo.GradeYear = _SHSubjectSemesterScoreInfo.GradeYear;
                        }
                    }
                }
                //else
                //{
                    pSHSubjectYearScoreInfo.AccumulatedCredit = 0;

                    decimal? credit1stSemester = 0;
                    decimal? credit2ndSemester = 0;

                    if (pSHSubject1stSemesterScoreInfo.Count > i)
                    {
                        pSHSubjectYearScoreInfo.AccumulatedCredit += (pSHSubject1stSemesterScoreInfo[i].Credit.HasValue ? pSHSubject1stSemesterScoreInfo[i].Credit.Value : 0);
                        if (pSHSubject1stSemesterScoreInfo[i].Pass.HasValue)
                            if (pSHSubject1stSemesterScoreInfo[i].Pass.Value)
                                credit1stSemester = (pSHSubject1stSemesterScoreInfo[i].Credit.HasValue ? pSHSubject1stSemesterScoreInfo[i].Credit.Value : 0);
                        pSHSubjectYearScoreInfo.Hour = (pSHSubject1stSemesterScoreInfo[i].Credit.HasValue ? pSHSubject1stSemesterScoreInfo[i].Credit.Value + "" : string.Empty);
                    }

                    if (pSHSubject2ndSemesterScoreInfo.Count > i)
                    {
                        pSHSubjectYearScoreInfo.AccumulatedCredit += (pSHSubject2ndSemesterScoreInfo[i].Credit.HasValue ? pSHSubject2ndSemesterScoreInfo[i].Credit.Value : 0);
                        if (pSHSubject2ndSemesterScoreInfo[i].Pass.HasValue)
                            if (pSHSubject2ndSemesterScoreInfo[i].Pass.Value)
                                credit2ndSemester = (pSHSubject2ndSemesterScoreInfo[i].Credit.HasValue ? pSHSubject2ndSemesterScoreInfo[i].Credit.Value : 0);

                        decimal hour1st = 0;
                        decimal hour2nd = 0;

                        if (decimal.TryParse(pSHSubjectYearScoreInfo.Hour, out hour1st) && decimal.TryParse(pSHSubject2ndSemesterScoreInfo[i].Credit + "", out hour2nd))
                            pSHSubjectYearScoreInfo.Hour = (hour1st + hour2nd)/2 + "";
                        else if (!decimal.TryParse(pSHSubjectYearScoreInfo.Hour, out hour1st) && decimal.TryParse(pSHSubject2ndSemesterScoreInfo[i].Credit + "", out hour2nd))
                            pSHSubjectYearScoreInfo.Hour = (pSHSubject2ndSemesterScoreInfo[i].Credit.HasValue ? pSHSubject2ndSemesterScoreInfo[i].Credit.Value + "" : string.Empty);
                    }

                    pSHSubjectYearScoreInfo.Credit = credit1stSemester + credit2ndSemester;
                //}

                _SHSubjectYearScoreInfo.Add(pSHSubjectYearScoreInfo);
            }
            while (true);

            return _SHSubjectYearScoreInfo;
        }

        private List<SHSubjectYearScoreInfo> GetPersonalYearSubjectScoreInfo(List<SHSubjectSemesterScoreInfo> pSHSubjectSemesterScoreInfo)
        {
            List<SHSubjectYearScoreInfo> _SHSubjectYearScoreInfo = new List<SHSubjectYearScoreInfo>();
            int i = -1;
            do
            {
                i++;
                SHSubjectSemesterScoreInfo _SHSubjectSemesterScoreInfo;

                if ((pSHSubjectSemesterScoreInfo.Count <= i))
                    break;

                _SHSubjectSemesterScoreInfo = pSHSubjectSemesterScoreInfo[i];

                string keySingleSubject = _SHSubjectSemesterScoreInfo.StudentID + "_" + _SHSubjectSemesterScoreInfo.SchoolYear.ToString() + "_" + _SHSubjectSemesterScoreInfo.SubjectName;

                SHSubjectYearScoreInfo pSHSubjectYearScoreInfo = this.GetPersonalSubjectYearScoreInfo(_SHSubjectSemesterScoreInfo.StudentID, _SHSubjectSemesterScoreInfo.SchoolYear, _SHSubjectSemesterScoreInfo.SubjectName, _SHSubjectSemesterScoreInfo.GradeYear);

                if (pSHSubjectYearScoreInfo == null)
                {
                    pSHSubjectYearScoreInfo = new SHSubjectYearScoreInfo();

                    pSHSubjectYearScoreInfo.StudentID = _SHSubjectSemesterScoreInfo.StudentID;
                    pSHSubjectYearScoreInfo.SchoolYear = _SHSubjectSemesterScoreInfo.SchoolYear;
                    pSHSubjectYearScoreInfo.SubjectName = _SHSubjectSemesterScoreInfo.SubjectName;
                    //pSHSubjectYearScoreInfo.Score = _SHSubjectSemesterScoreInfo
                    //pSHSubjectYearScoreInfo.Credit = _SHSubjectSemesterScoreInfo.StudentID;
                    pSHSubjectYearScoreInfo.GradeYear = _SHSubjectSemesterScoreInfo.GradeYear;
                    pSHSubjectYearScoreInfo.Hour = (_SHSubjectSemesterScoreInfo.Credit.HasValue ? _SHSubjectSemesterScoreInfo.Credit.Value + "" : string.Empty);
                }
                else
                {
                    pSHSubjectYearScoreInfo.AccumulatedCredit = 0;

                    decimal? creditSemester = 0;

                    pSHSubjectYearScoreInfo.AccumulatedCredit += (_SHSubjectSemesterScoreInfo.Credit.HasValue ? _SHSubjectSemesterScoreInfo.Credit.Value : 0);

                    if (_SHSubjectSemesterScoreInfo.Pass.HasValue)
                        if (_SHSubjectSemesterScoreInfo.Pass.Value)
                            creditSemester = (_SHSubjectSemesterScoreInfo.Credit.HasValue ? _SHSubjectSemesterScoreInfo.Credit.Value : 0);

                    pSHSubjectYearScoreInfo.Credit = creditSemester;
                    pSHSubjectYearScoreInfo.Hour = (_SHSubjectSemesterScoreInfo.Credit.HasValue ? _SHSubjectSemesterScoreInfo.Credit.Value + "" : string.Empty);
                }
                _SHSubjectYearScoreInfo.Add(pSHSubjectYearScoreInfo);
            }
            while (true);
            return _SHSubjectYearScoreInfo;
        }


        #endregion 

        public List<SHSubjectYearScoreInfo> GetPersonalYearSubjectScoreInfo(string pStudentID, int pSchoolYear)
        {
            string keyYearSubject = pStudentID + "_" + pSchoolYear.ToString();

            if (_PersonalYearSubjectScoreInfo.ContainsKey(keyYearSubject))
                return _PersonalYearSubjectScoreInfo[keyYearSubject];
            else
                return new List<SHSubjectYearScoreInfo>();
        }

        public SHSubjectYearScoreInfo GetPersonalSubjectYearScoreInfo(string pStudentID, int pSchoolYear, string pSubjectName, int? pGradeYear)
        {
            string keySingleSubject = pStudentID + "_" + pSchoolYear.ToString() + "_" + pSubjectName + "_" + (pGradeYear.HasValue ? pGradeYear.Value : 0).ToString();

            if (_PersonalSubjectYearScoreInfo.ContainsKey(keySingleSubject))
                return _PersonalSubjectYearScoreInfo[keySingleSubject];
            else
                return null;
        }

        #region 取得個人單科學年成績

        private decimal? GetPersonalSubjectYearScore(string pStudentID, int pSchoolYear, string pSubjectName, int? pGradeYear)
        {
            string keySingleSubject = pStudentID + "_" + pSchoolYear.ToString() + "_" + pSubjectName + "_" + pGradeYear;

            if (_PersonalSubjectYearScoreInfo.ContainsKey(keySingleSubject))
                return _PersonalSubjectYearScoreInfo[keySingleSubject].Score;
            else
                return null;
        }

        #endregion

        //  初始化所需資料
        private void ProduceData(List<SHStudentRecord> pStudents)
        {
            IEnumerable<string> pStudentIDs = pStudents.Select(x => x.ID);
            Dictionary<string, SHScoreCalcRuleRecord> _SHScoreCalcRule = SHScoreCalcRule.SelectAll().ToDictionary(x => x.ID);

            //  學生修習科目學年成績資料
            List<SHSchoolYearScoreRecord> _SHYearScore = SHSchoolYearScore.SelectByStudentIDs(pStudentIDs);
            if (_SHYearScore.Count > 0)
                _SHYearScore = _SHYearScore.OrderByDescending(x => x.SchoolYear).Select(x => x).ToList();

            //  收集學生修習科目學年成績資料
            foreach (SHSchoolYearScoreRecord sr in _SHYearScore)
            {
                //  學生個人學年歷程記錄
                if (!_PersonalYearSubjectScoreHistoryInfo.ContainsKey(sr.RefStudentID))
                    _PersonalYearSubjectScoreHistoryInfo.Add(sr.RefStudentID, new Dictionary<int, int>());

                if (!_PersonalYearSubjectScoreHistoryInfo[sr.RefStudentID].ContainsKey(sr.GradeYear))
                    _PersonalYearSubjectScoreHistoryInfo[sr.RefStudentID].Add(sr.GradeYear, sr.SchoolYear);
                else
                {
                    if (!_PersonalReReadSchoolYear.ContainsKey(sr.RefStudentID))
                        _PersonalReReadSchoolYear.Add(sr.RefStudentID, new List<int>());

                    _PersonalReReadSchoolYear[sr.RefStudentID].Add(sr.SchoolYear);
                }

                if (_PersonalReReadSchoolYear.ContainsKey(sr.RefStudentID) && _PersonalReReadSchoolYear[sr.RefStudentID].Contains(sr.SchoolYear))
                    continue;

                foreach (SHSchoolYearScoreSubject ss in sr.Subjects)
                {
                    string keySingleSubject = sr.RefStudentID + "_" + sr.SchoolYear.ToString() + "_" + ss.Subject + "_" + sr.GradeYear;

                    string keyYearSubject = sr.RefStudentID + "_" + sr.SchoolYear.ToString();
                    
                    if (!_PersonalSubjectYearScoreInfo.ContainsKey(keySingleSubject))
                        _PersonalSubjectYearScoreInfo.Add(keySingleSubject, new SHSubjectYearScoreInfo());

                    _PersonalSubjectYearScoreInfo[keySingleSubject].SchoolYear = sr.SchoolYear;
                    _PersonalSubjectYearScoreInfo[keySingleSubject].StudentID = sr.RefStudentID;
                    _PersonalSubjectYearScoreInfo[keySingleSubject].SubjectName = ss.Subject;
                    _PersonalSubjectYearScoreInfo[keySingleSubject].Score = ss.Score;
                    _PersonalSubjectYearScoreInfo[keySingleSubject].GradeYear = sr.GradeYear;

                    if (!_PersonalYearSubjectScoreInfo.ContainsKey(keyYearSubject))
                        _PersonalYearSubjectScoreInfo.Add(keyYearSubject, new List<SHSubjectYearScoreInfo>());

                    _PersonalYearSubjectScoreInfo[keyYearSubject].Add(_PersonalSubjectYearScoreInfo[keySingleSubject]);

                    //  學生個人學年歷程記錄
                    //if (!_PersonalYearSubjectScoreHistoryInfo.ContainsKey(sr.RefStudentID))
                    //    _PersonalYearSubjectScoreHistoryInfo.Add(sr.RefStudentID, new Dictionary<int, int>());

                    //if (!_PersonalYearSubjectScoreHistoryInfo[sr.RefStudentID].ContainsKey(sr.SchoolYear))
                    //    _PersonalYearSubjectScoreHistoryInfo[sr.RefStudentID].Add(sr.SchoolYear, sr.GradeYear);
                }
            }
            //Dictionary<string, Dictionary<int, int>> dd = new Dictionary<string, Dictionary<int, int>>();
            //foreach (string studentID in _PersonalYearSubjectScoreHistoryInfo.Keys)
            //{
            //    var _SchoolYearHistoryItems = from pair in _PersonalYearSubjectScoreHistoryInfo[studentID] orderby pair.Key descending select pair;
            //    Dictionary<int, int> d = new Dictionary<int, int>();
            //    int yIndex = 0;
            //    foreach (KeyValuePair<int, int> sh in _SchoolYearHistoryItems)
            //    {
            //        if (yIndex != sh.Value)
            //        {
            //            d.Add(sh.Key, sh.Value);
            //            yIndex = sh.Value;
            //        }
            //    }
            //    dd.Add(studentID, d);
            //}
            //_PersonalYearSubjectScoreHistoryInfo = dd;
            _SHYearScore = null;
        }
    }
    #endregion
}






//XDocument doc = XDocument.Parse("<root>" + sr.TextScore.InnerXml + "</root>");
//        foreach (XElement e in doc.Document.Descendants("Morality"))
//        {
//            moralScore += e.Attribute("Face").Value + "：" + e.Value + "；";
//        }
//XDocument doc = XDocument.Parse("<root>" + sr.TextScore.InnerXml + "</root>");

////  成績計算規則
//_ScoreCalcRule = SHScoreCalcRule.SelectAll();   

////  課程規劃表
//List<SHProgramPlanRecord> _SHProgramPlan = SHProgramPlan.SelectAll();
//foreach (SHProgramPlanRecord pr in _SHProgramPlan)
//{
//    foreach (ProgramSubject ps in pr.Subjects)
//    {
//        string key = pr.Name + "_" + ps.SubjectName + "_" + ((ps.Level.HasValue) ? ps.Level.ToString() : "");

//        if (!_SHProgramPlanRecord.ContainsKey(key))
//            _SHProgramPlanRecord.Add(key, ps);
//    }
//}

//  課程規劃表
//private Dictionary<string, ProgramSubject> _SHProgramPlanRecord;
//  學生課程規劃表屬性
//private Dictionary<string, string> _StudentProgramPlan;
//  成績計算規則
//private List<SHScoreCalcRuleRecord> _ScoreCalcRule;
//  學生學期科目成績屬性採計方式
//private Dictionary<string, string> _StudentSemesterSubjectScoreCalcRuleUsingType;  

        ////  課程規劃之「校/部訂」屬性
        //private string GetPersonalSubjectRequiredByUsingProgramPlan(string pProgramPlan, string pSubjectName, int? pLevel)
        //{
        //    if (pProgramPlan == string.Empty)
        //        return string.Empty;

        //    string key = pProgramPlan + "_" + pSubjectName + "_" + ((pLevel.HasValue) ? pLevel.ToString() : "");

        //    if (_SHProgramPlanRecord.ContainsKey(key))
        //        return _SHProgramPlanRecord[key].RequiredBy;
        //    else
        //        return null;
        //}

        ////  課程規劃之「必/選修」屬性
        //private bool? GetPersonalSubjectRequiredUsingProgramPlan(string pProgramPlan, string pSubjectName, int? pLevel)
        //{
        //    if (pProgramPlan == string.Empty)
        //        return null;

        //    string key = pProgramPlan + "_" + pSubjectName + "_" + ((pLevel.HasValue) ? pLevel.ToString() : "");

        //    if (_SHProgramPlanRecord.ContainsKey(key))
        //        return _SHProgramPlanRecord[key].Required;
        //    else
        //        return null;
        //}

        ////  課程規劃之「學分」屬性
        //private decimal? GetPersonalSubjectCreditUsingProgramPlan(string pProgramPlan, string pSubjectName, int? pLevel)
        //{
        //    if (pProgramPlan == string.Empty)
        //        return null;
            
        //    string key = pProgramPlan + "_" + pSubjectName + "_" + ((pLevel.HasValue) ? pLevel.ToString() : "");

        //    if (_SHProgramPlanRecord.ContainsKey(key))
        //        return _SHProgramPlanRecord[key].Credit;
        //    else
        //        return null;
        //}

        ////  成績計算規則是否「以實際學期科目成績內容為準」
        //private bool IsScoreCalcBySubject(SHStudentRecord pStudent)
        //{
        //    if (!_StudentSemesterSubjectScoreCalcRuleUsingType.ContainsKey(pStudent.ID))
        //        SetScoreCalcRuleUsingType(pStudent);

        //    if (_StudentSemesterSubjectScoreCalcRuleUsingType[pStudent.ID] == "以實際學期科目成績內容為準" || _StudentSemesterSubjectScoreCalcRuleUsingType[pStudent.ID] == "")
        //        return true;
        //    else
        //        return false;
        //}

        ////  設定「學期科目成績屬性採計方式」
        //private void SetScoreCalcRuleUsingType(SHStudentRecord pStudent)
        //{
        //    if (pStudent.ScoreCalcRule != null)
        //        _StudentSemesterSubjectScoreCalcRuleUsingType.Add(pStudent.ID, XElement.Parse(pStudent.ScoreCalcRule.Content.ToString()).Element("學期科目成績屬性採計方式").Value);
        //    else
        //        _StudentSemesterSubjectScoreCalcRuleUsingType.Add(pStudent.ID, string.Empty);
        //} 

//  學期對照表：SHSemesterHistoryRecord-->考慮 外部 處理
//List<SHSemesterHistoryRecord> _SHSemesterHistory = SHSemesterHistory.SelectByStudentIDs(pStudentIDs);
//Dictionary<string, SHSemesterHistoryRecord> sHSemesterHistory = new Dictionary<string, SHSemesterHistoryRecord>();
//foreach (SHSemesterHistoryRecord sr in _SHSemesterHistory)
//    if (!sHSemesterHistory.ContainsKey(sr.RefStudentID))
//        sHSemesterHistory.Add(sr.RefStudentID, sr);

///畢業計算規則內的所有資料
//<畢業學分數><學科累計總學分數>160</學科累計總學分數><必修學分數>120</必修學分數><選修學分數>40</選修學分數></畢業學分數>
//"<學期科目成績屬性採計方式>以實際學期科目成績內容為準</學期科目成績屬性採計方式><各項成績計算位數><科目成績計算位數 位數=\"0\" 四捨五入=\"True\" 無條件捨去=\"False\" 無條件進位=\"False\" /><學期分項成績計算位數 位數=\"1\" 四捨五入=\"True\" 無條件捨去=\"False\" 無條件進位=\"False\" /><學年科目成績計算位數 位數=\"0\" 四捨五入=\"True\" 無條件捨去=\"False\" 無條件進位=\"False\" /><學年分項成績計算位數 位數=\"1\" 四捨五入=\"True\" 無條件捨去=\"False\" 無條件進位=\"False\" /><畢業成績計算位數 位數=\"1\" 四捨五入=\"True\" 無條件捨去=\"False\" 無條件進位=\"False\" /></各項成績計算位數><分項成績計算項目><體育 併入學期學業成績=\"True\" 計算成績=\"False\" /><國防通識 併入學期學業成績=\"True\" 計算成績=\"False\" /><健康與護理 併入學期學業成績=\"True\" 計算成績=\"False\" /><實習科目 併入學期學業成績=\"True\" 計算成績=\"False\" /></分項成績計算項目><延修及重讀成績處理規則><延修成績 延修成績登錄至各修課學年度學期=\"True\" 所有延修成績合併登錄至同一延修學期=\"False\" 開始年級=\"4\" /><重讀成績 擇優採計成績=\"True\" /></延修及重讀成績處理規則><學年調整成績 不登錄學年調整成績=\"False\" 不重新計算學期分項成績=\"False\" 以六十分登錄=\"True\" 以學生及格標準登錄=\"False\" 重新計算學期分項成績=\"True\" /><重修成績 登錄至原學期=\"True\" /><及格標準><學生類別 一年級及格標準=\"60\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"60\" 二年級補考標準=\"0\" 四年級及格標準=\"60\" 四年級補考標準=\"0\" 類別=\"預設\" /><學生類別 一年級及格標準=\"40\" 一年級補考標
//準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"原住民生\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"派外人員子女\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"退伍軍人\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"僑生\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"蒙藏學生\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"外國學生\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"重大災害地區學生\" /><學生類別 一年級及格標準=\"60\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"60\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\
//" 類別=\"身心障礙生\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"境外優秀科學技術人才子女\" /><學生類別 一年級及格標準=\"40\" 一年級補考標準=\"0\" 三年級及格標準=\"50\" 三年級補考標準=\"0\" 二年級及格標準=\"40\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"體育績優生\" /><學生類別 一年級及格標準=\"50\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"50\" 二年級補考標準=\"0\" 四年級及格標準=\"\" 四年級補考標準=\"\" 類別=\"技藝技能甄審學生\" /></及格標準><畢業成績計算規則>學期分項成績平均</畢業成績計算規則><德行成績畢業判斷規則>每學年德行成績均及格</德行成績畢業判斷規則><畢業學分數><學科累計總學分數>160</學科累計總學分數><必修學分數>120</必修學分數><選修學分數>40</選修學分數></畢業學分數><核心科目表 /><學期科目成績計算至學年科目成績規則 進校上下學期及格規則=\"False\" />"
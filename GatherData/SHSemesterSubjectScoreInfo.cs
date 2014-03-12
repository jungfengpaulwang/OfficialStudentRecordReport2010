using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHSchool.Data;
using System.Xml.Linq;
using K12.Data;

namespace OfficialStudentRecordReport2010
{
    #region 課程成績資訊之 ValueObject
    /// <summary>
    /// 課程成績資訊之 ValueObject(未來將包含課程規劃中的內容)
    /// </summary>
    public class SHSubjectSemesterScoreInfo
    {
        public SHSubjectSemesterScoreInfo() { }

        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { set; get; }      
        /// <summary>
        /// 學年度
        /// </summary>
        public int SchoolYear { set; get; }   
        /// <summary>
        /// 學期
        /// </summary>      
        public int Semester { set; get; }   
        /// <summary>
        /// 分項類別：學業, 體育, 實習科目, 國防通識, 健康與護理, 專業科目
        /// </summary>
        public string Entry { set; get; }
        /// <summary>
        /// 科目名稱
        /// </summary>  
        public string SubjectName { set; get; }
        /// <summary>
        /// 成績年級
        /// </summary>
        public int? GradeYear { set; get; }
        /// <summary>
        /// 級別
        /// </summary>
        public int? Level { set; get; }                     
        /// <summary>
        /// 原始成績
        /// </summary>
        public decimal? Score { set; get; }                 
        /// <summary>
        /// 補考成績
        /// </summary>
        public decimal? ReExamScore { set; get; }           
        /// <summary>
        /// 重修成績
        /// </summary>
        public decimal? ReCourseScore { set; get; }         
        /// <summary>
        /// 手調成績
        /// </summary>
        public decimal? ManualScore { set; get; }           
        /// <summary>
        /// 計算學年成績後，異動之學期成績
        /// </summary>
        public decimal? SchoolYearAdjustScore { set; get; } 
        /// <summary>
        /// 修習課程之學分
        /// </summary>
        public decimal? Credit { set; get; }         
        //public decimal? ProgramCredit { set; get; }         //  課程規劃之學分
        /// <summary>
        /// 修習課程之校/部訂
        /// </summary>
        public string RequiredBy { set; get; }       
        //public string ProgramRequiredBy { set; get; }       //  課程規劃之校/部訂
        /// <summary>
        /// 修習課程之必/選修
        /// </summary>
        public bool? Required { set; get; }           
        //public bool ProgramRequired { set; get; }           //  課程規劃之必/選修
        //public string ProgramEntry { set; get; }       //  課程規劃之分項類別        
        /// <summary>
        /// 是否為核心科目
        /// </summary>
        public bool CoreSubject { set; get; }
        /// <summary>
        /// 是否為專業科目
        /// </summary>
        public bool ProSubject { set; get; }   
        /// <summary>
        /// 不計學分
        /// </summary>
        public bool? NotIncludedInCredit { set; get; }       
        /// <summary>
        /// 不計分
        /// </summary>
        public bool? NotIncludedInCalc { set; get; }         
        /// <summary>
        /// 註記
        /// </summary>
        public string Comment { set; get; }
        /// <summary>
        /// 是否取得學分
        /// </summary>
        public bool? Pass { set; get; }        
    }
    #endregion

    #region 取得學生學期科目成績內容之物件
    public class SHSemesterSubjectScoreInfo : IDisposable
    {
        //  學生科目學期成績資料
        private Dictionary<string, SHSubjectSemesterScoreInfo> _PersonalSubjectSemesterScoreInfo;
        //  學生學期科目成績資料
        private Dictionary<string, List<SHSubjectSemesterScoreInfo>> _PersonalSemesterSubjectScoreInfo;
        //  學生所有學期科目成績資料
        private Dictionary<string, Dictionary<int, List<KeyValuePair<int, int>>>> _PersonalSemesterSubjectScoreHistoryInfo;
        //  學生所有學期科目成績資料
        private Dictionary<string, List<SHSubjectSemesterScoreInfo>> _PersonalAllSemesterSubjectScoreInfo;
        //  後期中等教育核心科目標示
        private List<string> _CoreSubjectTable;
        //  核心科目表
        private List<SHSubjectTableRecord> _SubjectTable;
        //  單一學生重讀學年
        private Dictionary<string, List<int>> _PersonalReReadSchoolYear;
        
        public List<SHSubjectTableRecord> GetSubjectTable()
        {
            return _SubjectTable;
        }

        public void Dispose()
        {
            _PersonalSubjectSemesterScoreInfo.GetEnumerator().Dispose();
            _PersonalSemesterSubjectScoreInfo.GetEnumerator().Dispose();
            _PersonalSemesterSubjectScoreHistoryInfo.GetEnumerator().Dispose();
            _PersonalAllSemesterSubjectScoreInfo.GetEnumerator().Dispose();
            _CoreSubjectTable.GetEnumerator().Dispose();
            _SubjectTable.GetEnumerator().Dispose();
            _PersonalReReadSchoolYear.GetEnumerator().Dispose();

            GC.SuppressFinalize(this);
        }

        public enum SubjectSemesterScoreType
        {
            由學年與學期取得之學期科目, 由上學期科目取得之下學期科目
        }
        
        public enum RequiredBy
        {
            部, 校
        }

        public enum Entry
        {
            學業, 體育, 實習科目, 國防通識, 健康與護理, 專業科目
        }

        public SHSemesterSubjectScoreInfo(List<SHStudentRecord> pStudents)
        {
            _PersonalSubjectSemesterScoreInfo = new Dictionary<string, SHSubjectSemesterScoreInfo>();
            _PersonalSemesterSubjectScoreInfo = new Dictionary<string, List<SHSubjectSemesterScoreInfo>>();
            _PersonalSemesterSubjectScoreHistoryInfo = new Dictionary<string, Dictionary<int, List<KeyValuePair<int, int>>>>();
            _PersonalAllSemesterSubjectScoreInfo = new Dictionary<string, List<SHSubjectSemesterScoreInfo>>();
            _CoreSubjectTable = new List<string>();
            _PersonalReReadSchoolYear = new Dictionary<string, List<int>>();

            ProduceData(pStudents);
        }

        public Dictionary<int, List<KeyValuePair<int, int>>> GetPersonalSemesterSubjectScoreHistoryInfo(string pStudentID)
        {
            if (_PersonalSemesterSubjectScoreHistoryInfo.ContainsKey(pStudentID))
                return _PersonalSemesterSubjectScoreHistoryInfo[pStudentID];
            else
                return new Dictionary<int,List<KeyValuePair<int,int>>>();
        }

        public List<int> GetPersonalReReadSchoolYear(string pStudentID)
        {
            if (_PersonalReReadSchoolYear.ContainsKey(pStudentID))
                return _PersonalReReadSchoolYear[pStudentID];
            else
                return new List<int>();
        }

        #region 取得累計學分
        /// <summary>
        /// 累計實得學分
        /// </summary>
        /// <param name="pStudentID">學生系統編號</param>
        /// <param name="pSchoolYear">學年度</param>
        /// <param name="pSemester">學期</param>
        /// <param name="pRequired">必/選修</param>
        /// <param name="pRequiredBy">校/部訂</param>
        /// <param name="pEntry">分項類別(學業、體育、實習科目、國防通識、健康與護理、專業科目)，不依據此項目統計時，傳入 null</param>
        /// <returns>學分數</returns>
        public decimal GetAccumulatedAcquiredCredit(string pStudentID, List<int> pSchoolYear, List<int> pSemester, bool pRequired, RequiredBy pRequiredBy, Entry pEntry)
        {
            decimal? credit = 0;
            foreach (int schoolYear in pSchoolYear)
            {
                foreach (int semester in pSemester)
                {
                    string key = pStudentID + "_" + pSchoolYear.ToString() + "_" + pSemester;
                    foreach (SHSubjectSemesterScoreInfo shsssi in _PersonalSemesterSubjectScoreInfo[key])
                    {
                        //if (shsssi.NotIncludedInCredit)
                        //    continue;

                        //if (!shsssi.Pass)
                        //    continue;

                        if (shsssi.Entry != pEntry.ToString())
                            continue;

                        if (shsssi.RequiredBy != pRequiredBy.ToString())
                            continue;

                        if (shsssi.Required != pRequired)
                            continue;

                        credit += ((shsssi.Credit.HasValue) ? shsssi.Credit.Value : 0);
                    }
                }
            }
            return ((credit.HasValue) ? credit.Value : 0);
        }

        /// <summary>
        /// 累計應得學分
        /// </summary>
        /// <param name="pStudentID">學生系統編號</param>
        /// <param name="pSchoolYear">學年度</param>
        /// <param name="pSemester">學期</param>
        /// <param name="pRequired">必/選修</param>
        /// <param name="pRequiredBy">校/部訂</param>
        /// <param name="pSubjectType">分項類別(學業、體育、實習科目、國防通識、健康與護理、專業科目)，不依據此項目統計時，傳入 null</param>
        /// <returns>學分數</returns>
        public decimal GetAccumulatedDeservedCredit(string pStudentID, List<int> pSchoolYear, List<int> pSemester, bool pRequired, RequiredBy pRequiredBy, Entry pEntry)
        {            
            decimal? credit = 0;
            foreach(int schoolYear in pSchoolYear)
            {
                foreach(int semester in pSemester)
                {
                    string key = pStudentID + "_" + pSchoolYear.ToString() + "_" + pSemester;
                    foreach (SHSubjectSemesterScoreInfo shsssi in _PersonalSemesterSubjectScoreInfo[key])
                    {
                        //if (shsssi.NotIncludedInCredit)
                        //    continue;

                        if (shsssi.Entry != pEntry.ToString())
                            continue;

                        if (shsssi.RequiredBy != pRequiredBy.ToString())
                            continue;

                        if (shsssi.Required != pRequired)
                            continue;

                        credit += ((shsssi.Credit.HasValue) ? shsssi.Credit.Value : 0);                        
                    }
                }
            }
            return ((credit.HasValue) ? credit.Value : 0);
        }

        #endregion

        public List<SHSubjectSemesterScoreInfo> SortSHSubjectSemesterScore(List<SHSubjectSemesterScoreInfo> firstSemesterSubjectScoreInfo, List<SHSubjectSemesterScoreInfo> secondSemesterSubjectScoreInfo)
        {
            SHSubjectSemesterScoreInfo[] byteSecondSemesterSubjectScoreInfo = new List<SHSubjectSemesterScoreInfo>().ToArray();
            if (secondSemesterSubjectScoreInfo != null && secondSemesterSubjectScoreInfo.Count>0)
                byteSecondSemesterSubjectScoreInfo = secondSemesterSubjectScoreInfo.ToArray();

            List<SHSubjectSemesterScoreInfo> reSortedSecondSemesterSubjectScoreInfo = new List<SHSubjectSemesterScoreInfo>();

            if ((firstSemesterSubjectScoreInfo == null || firstSemesterSubjectScoreInfo.Count == 0) && (secondSemesterSubjectScoreInfo == null || secondSemesterSubjectScoreInfo.Count == 0))
                return new List<SHSubjectSemesterScoreInfo>();

            if (secondSemesterSubjectScoreInfo == null || secondSemesterSubjectScoreInfo.Count == 0)
                return firstSemesterSubjectScoreInfo;

            if (firstSemesterSubjectScoreInfo == null || firstSemesterSubjectScoreInfo.Count == 0)
                return secondSemesterSubjectScoreInfo;

            int j = 0;
            SHSubjectSemesterScoreInfo s2 = new SHSubjectSemesterScoreInfo();
            foreach (SHSubjectSemesterScoreInfo s1 in firstSemesterSubjectScoreInfo)
            {
                bool found = false;
                for (int i = 0; i < (byteSecondSemesterSubjectScoreInfo.Length - j); i++)
                {
                    s2 = (byteSecondSemesterSubjectScoreInfo.GetValue(i) as SHSubjectSemesterScoreInfo);

                    if ((s1.SubjectName.Trim() == s2.SubjectName.Trim()) && (s1.GradeYear.HasValue && s2.GradeYear.HasValue && s1.GradeYear == s2.GradeYear))
                    {
                        found = true;

                        reSortedSecondSemesterSubjectScoreInfo.Add(s2);
                        byteSecondSemesterSubjectScoreInfo.SetValue(byteSecondSemesterSubjectScoreInfo.GetValue((byteSecondSemesterSubjectScoreInfo.Length - 1 - j)), i);
                        byteSecondSemesterSubjectScoreInfo.SetValue(s2, (byteSecondSemesterSubjectScoreInfo.Length - 1 - j));

                        j++;
                    }
                }
                if (!found)
                {
                    SHSubjectSemesterScoreInfo _SHSubjectSemesterScoreInfo = new SHSubjectSemesterScoreInfo();

                    _SHSubjectSemesterScoreInfo.SubjectName = string.Empty;
                    _SHSubjectSemesterScoreInfo.StudentID = s2.StudentID;
                    _SHSubjectSemesterScoreInfo.Semester = s2.Semester;
                    _SHSubjectSemesterScoreInfo.Score = null;
                    _SHSubjectSemesterScoreInfo.SchoolYearAdjustScore = null;
                    _SHSubjectSemesterScoreInfo.SchoolYear = s2.SchoolYear;
                    _SHSubjectSemesterScoreInfo.RequiredBy = string.Empty;
                    _SHSubjectSemesterScoreInfo.Required = null;
                    _SHSubjectSemesterScoreInfo.ReExamScore = null;
                    _SHSubjectSemesterScoreInfo.ReCourseScore = null;
                    _SHSubjectSemesterScoreInfo.Pass = null;
                    _SHSubjectSemesterScoreInfo.NotIncludedInCredit = null;
                    _SHSubjectSemesterScoreInfo.NotIncludedInCalc = null;
                    _SHSubjectSemesterScoreInfo.ManualScore = null;
                    _SHSubjectSemesterScoreInfo.Level = null;
                    _SHSubjectSemesterScoreInfo.Entry = string.Empty;
                    _SHSubjectSemesterScoreInfo.Credit = null;
                    _SHSubjectSemesterScoreInfo.CoreSubject = false;
                    _SHSubjectSemesterScoreInfo.Comment = string.Empty;
                    _SHSubjectSemesterScoreInfo.GradeYear = null;
                    _SHSubjectSemesterScoreInfo.ProSubject = false;

                    reSortedSecondSemesterSubjectScoreInfo.Add(_SHSubjectSemesterScoreInfo);
                }
            }
            for (int i = 0; i < (byteSecondSemesterSubjectScoreInfo.Length - j); i++)
            {
                reSortedSecondSemesterSubjectScoreInfo.Add(byteSecondSemesterSubjectScoreInfo[i]);
            }
            return reSortedSecondSemesterSubjectScoreInfo;

        }
        
        #region 取得個人學期科目成績資訊
        /// <summary>
        /// 個人學期科目成績資訊
        /// </summary>
        /// <param name="pStudentID">學生系統編號</param>
        /// <param name="pSchoolYear">學年度</param>
        /// <param name="pSemester">學期</param>
        /// <returns>個人學期科目成績資訊</returns>  
        public List<SHSubjectSemesterScoreInfo> GetPersonalSemesterSubjectScoreInfo(string pStudentID, int pSchoolYear1, int pSemester1, SubjectSemesterScoreType pSubjectSemesterScoreType, int pSchoolYear2, int pSemester2)
        {
            string key = string.Empty;

            if (pSubjectSemesterScoreType == SubjectSemesterScoreType.由學年與學期取得之學期科目)
            {
                key = pStudentID + "_" + pSchoolYear1.ToString() + "_" + pSemester1.ToString();

                if (_PersonalSemesterSubjectScoreInfo.ContainsKey(key))
                {
                    _PersonalSemesterSubjectScoreInfo[key].Sort(Util.SortBySubjectName);
                    return _PersonalSemesterSubjectScoreInfo[key];
                }
                else
                    return null;
            }
            else if (pSubjectSemesterScoreType == SubjectSemesterScoreType.由上學期科目取得之下學期科目)
            {
                return GetPersonal2ndSemesterSubjectScoreInfo(pStudentID, pSchoolYear1, pSchoolYear2, pSemester1, pSemester2);
            }
            return null;
        }
        #endregion   

        #region 由上學期科目取得下學期科目
        /// <summary>
        /// 由上學期科目取得下學期科目
        /// </summary>
        /// <param name="pStudentID">學生系統編號</param>
        /// <param name="pSchoolYear">學年度</param>
        /// <returns>個人與上學期對齊之下學期科目成績資訊</returns>
        public List<SHSubjectSemesterScoreInfo> GetPersonal2ndSemesterSubjectScoreInfo(string pStudentID, int pSchoolYear1, int pSchoolYear2, int pSemester1, int pSemester2)
        {
            string keyFirstSemester = pStudentID + "_" + pSchoolYear1.ToString() + "_" + pSemester1.ToString();
            string keySecondSemester = pStudentID + "_" + pSchoolYear2.ToString() + "_" + pSemester2.ToString();

            if (!_PersonalSemesterSubjectScoreInfo.ContainsKey(keyFirstSemester))
                return null;

            if (!_PersonalSemesterSubjectScoreInfo.ContainsKey(keySecondSemester))
                return null;
            
            List<SHSubjectSemesterScoreInfo> firstSemesterSubjectScoreInfo = _PersonalSemesterSubjectScoreInfo[keyFirstSemester];
            SHSubjectSemesterScoreInfo[] secondSemesterSubjectScoreInfo = _PersonalSemesterSubjectScoreInfo[keySecondSemester].ToArray();
            List<SHSubjectSemesterScoreInfo> reSortedSecondSemesterSubjectScoreInfo = new List<SHSubjectSemesterScoreInfo>();

            if (secondSemesterSubjectScoreInfo.Length == 0)
                return null;

            int j = 0;
            foreach (SHSubjectSemesterScoreInfo s1 in firstSemesterSubjectScoreInfo)
            {
                bool found = false;
                SHSubjectSemesterScoreInfo s2 = new SHSubjectSemesterScoreInfo();
                for (int i = 0; i < (secondSemesterSubjectScoreInfo.Length - j); i++)
                {
                    s2 = (secondSemesterSubjectScoreInfo.GetValue(i) as SHSubjectSemesterScoreInfo);

                    if ((s1.SubjectName.Trim() == s2.SubjectName.Trim()) && (s1.GradeYear.HasValue && s2.GradeYear.HasValue && s1.GradeYear == s2.GradeYear))
                    {
                        found = true;

                        reSortedSecondSemesterSubjectScoreInfo.Add(s2);
                        secondSemesterSubjectScoreInfo.SetValue(secondSemesterSubjectScoreInfo.GetValue((secondSemesterSubjectScoreInfo.Length - 1 - j)), i);
                        secondSemesterSubjectScoreInfo.SetValue(s2, (secondSemesterSubjectScoreInfo.Length - 1 - j));

                        j++;
                    }
                }
                if (!found)
                {
                    SHSubjectSemesterScoreInfo _SHSubjectSemesterScoreInfo = new SHSubjectSemesterScoreInfo();

                    _SHSubjectSemesterScoreInfo.SubjectName = string.Empty;
                    _SHSubjectSemesterScoreInfo.StudentID = pStudentID;
                    _SHSubjectSemesterScoreInfo.Semester = s2.Semester;
                    _SHSubjectSemesterScoreInfo.Score = null;
                    _SHSubjectSemesterScoreInfo.SchoolYearAdjustScore = null;
                    _SHSubjectSemesterScoreInfo.SchoolYear = s2.SchoolYear;
                    _SHSubjectSemesterScoreInfo.RequiredBy = string.Empty;
                    _SHSubjectSemesterScoreInfo.Required = null;
                    _SHSubjectSemesterScoreInfo.ReExamScore = null;
                    _SHSubjectSemesterScoreInfo.ReCourseScore = null;
                    _SHSubjectSemesterScoreInfo.Pass = null;
                    _SHSubjectSemesterScoreInfo.NotIncludedInCredit = null;
                    _SHSubjectSemesterScoreInfo.NotIncludedInCalc = null;
                    _SHSubjectSemesterScoreInfo.ManualScore = null;
                    _SHSubjectSemesterScoreInfo.Level = null;
                    _SHSubjectSemesterScoreInfo.Entry = string.Empty;
                    _SHSubjectSemesterScoreInfo.Credit = null;
                    _SHSubjectSemesterScoreInfo.CoreSubject = false;
                    _SHSubjectSemesterScoreInfo.Comment = string.Empty;
                    _SHSubjectSemesterScoreInfo.GradeYear = null;
                    _SHSubjectSemesterScoreInfo.ProSubject = false;

                    reSortedSecondSemesterSubjectScoreInfo.Add(s2);
                }
            }
            for (int i = 0; i < (secondSemesterSubjectScoreInfo.Length-j); i++)
            {
                reSortedSecondSemesterSubjectScoreInfo.Add(secondSemesterSubjectScoreInfo[i]);
            }
            return reSortedSecondSemesterSubjectScoreInfo;
        }

        #endregion

        #region 取得個人所有學期科目成績資訊
        public List<SHSubjectSemesterScoreInfo> GetPersonalAllSemesterSubjectScoreInfo(string pStudentID)
        {
            if (_PersonalAllSemesterSubjectScoreInfo.ContainsKey(pStudentID))
                return _PersonalAllSemesterSubjectScoreInfo[pStudentID];
            else
                return new List<SHSubjectSemesterScoreInfo>();
        }
        #endregion

        public List<SHSubjectSemesterScoreInfo> GetPersonalSemesterSubjectScoreInfo(string pStudentID, int pSchoolYear, int pSemester)
        {
            string key = string.Empty;

            key = pStudentID + "_" + pSchoolYear.ToString() + "_" + pSemester.ToString();
            if (_PersonalSemesterSubjectScoreInfo.ContainsKey(key))
            {
                _PersonalSemesterSubjectScoreInfo[key].Sort(Util.SortBySubjectName);
                return _PersonalSemesterSubjectScoreInfo[key];
            }
            else
                return new List<SHSubjectSemesterScoreInfo>();
        }
        
        //  判讀學生修習科目是否為核心科目
        private bool IsCoreSubject(string pSubjectName, int? pLevel)
        {
            string key = pSubjectName + ((pLevel.HasValue) ? "_" + pLevel.Value.ToString() : "");
            bool youGotIt = false;

            _CoreSubjectTable.ForEach(delegate(String name)
            {
                if (!String.IsNullOrEmpty(name) && key.IndexOf(name) >= 0)
                    youGotIt = true;
            });

            return youGotIt;
        }

        //  收集學生修習科目學期成績資料
        private void ProducePersonalSubjectSemesterScoreInfo(Dictionary<string, SHSubjectSemesterScoreInfo> pPersonalSubjectSemesterScoreInfo, SHSchool.Data.SHSubjectScore pSubjectScore, int? pGradeYear)
        {
            string key = pSubjectScore.RefStudentID + "_" + pSubjectScore.SchoolYear.ToString() + "_" + pSubjectScore.Semester.ToString() + "_" + pSubjectScore.Subject + "_" + ((pSubjectScore.Level.HasValue) ? pSubjectScore.Level.ToString() : "");

            SHSubjectSemesterScoreInfo SHSSSI = pPersonalSubjectSemesterScoreInfo[key];

            SHSSSI.StudentID = pSubjectScore.RefStudentID;
            SHSSSI.SchoolYear = pSubjectScore.SchoolYear;
            SHSSSI.Semester = pSubjectScore.Semester;
            SHSSSI.Entry = pSubjectScore.Entry;
            SHSSSI.SubjectName = pSubjectScore.Subject;
            SHSSSI.Level = pSubjectScore.Level;
            SHSSSI.Score = pSubjectScore.Score;
            SHSSSI.ReExamScore = pSubjectScore.ScoreReExam;
            SHSSSI.ReCourseScore = pSubjectScore.ScoreReCourse;
            SHSSSI.ManualScore = pSubjectScore.ScoreBetter;
            SHSSSI.SchoolYearAdjustScore = pSubjectScore.ScoreSchoolYearAdjust;
            SHSSSI.Credit = pSubjectScore.Credit;
            //SHSSSI.ProgramCredit = GetPersonalSubjectCreditUsingProgramPlan(_StudentProgramPlan[pSubjectScore.RefStudentID], pSubjectScore.Subject, pSubjectScore.Level);
            SHSSSI.RequiredBy = pSubjectScore.RequiredBy;
            //SHSSSI.ProgramRequiredBy = string.Empty;
            SHSSSI.Required = pSubjectScore.Required;
            //SHSSSI.ProgramRequired = ;
            //SHSSSI.ProgramBranchType = string.Empty;
            SHSSSI.CoreSubject = IsCoreSubject(pSubjectScore.Subject, pSubjectScore.Level);
            SHSSSI.NotIncludedInCredit = pSubjectScore.NotIncludedInCredit;
            SHSSSI.NotIncludedInCalc = pSubjectScore.NotIncludedInCalc;
            SHSSSI.Comment = pSubjectScore.Comment;
            SHSSSI.Pass = pSubjectScore.Pass;
            SHSSSI.GradeYear = pGradeYear;

            if (!_PersonalAllSemesterSubjectScoreInfo.ContainsKey(pSubjectScore.RefStudentID))
                _PersonalAllSemesterSubjectScoreInfo.Add(pSubjectScore.RefStudentID, new List<SHSubjectSemesterScoreInfo>());

            _PersonalAllSemesterSubjectScoreInfo[pSubjectScore.RefStudentID].Add(SHSSSI);
        }

        //  初始化所需資料
        private void ProduceData(List<SHStudentRecord> pStudents)
        {
            IEnumerable<string> pStudentIDs = pStudents.Select(x => x.ID);
            
            //  先快取班級資料，否則在查詢學生所屬班級的成績計算規則資料時會降低系統效能
            IEnumerable<string> classiDs = pStudents.Select(x => x.RefClassID);
            SHClass.SelectByIDs(classiDs);


            //  後期中等教育核心科目標示
           _SubjectTable = SHSubjectTable.Select(null, null, null);
            foreach (SHSubjectTableRecord sr in _SubjectTable)
            {
                if (sr.Catalog.Equals("核心科目表") && sr.Name.Equals("後期中等教育核心課程"))
                {
                    foreach (SHSubjectTableSubject ss in sr.Subjects)
                    {
                        string key = string.Empty;
                        if (ss.Levels.Count == 0)
                        {
                            key = ss.Name;
                            if (!_CoreSubjectTable.Contains(key))
                                _CoreSubjectTable.Add(key);
                        }
                        else
                        {
                            foreach (int a in ss.Levels)
                            {
                                key = ss.Name + "_" + a.ToString();
                                if (!_CoreSubjectTable.Contains(key))
                                    _CoreSubjectTable.Add(key);
                            }
                        }
                    }
                }
            }

            //  學生修習科目學期成績資料
            List<SHSemesterScoreRecord> _SHSemesterScore = SHSemesterScore.SelectByStudentIDs(pStudentIDs,false);

            if (_SHSemesterScore.Count > 0)
                _SHSemesterScore = _SHSemesterScore.OrderByDescending(x => x.SchoolYear).ThenBy(x=>x.Semester).Select(x => x).ToList();

            //Dictionary<string, List<SHSemesterScoreRecord>> dReReading = new Dictionary<string, List<SHSemesterScoreRecord>>();

            //  收集學生修習科目學期成績資料
            foreach (SHSemesterScoreRecord sr in _SHSemesterScore)
            {
                //  重讀要檢查所有科目成績資料，若「科目名稱+級別+成績年級」有2筆以上，表示該科目成績資料屬重讀
                //  若已存在同年級成績，且成績所屬學年度不同，則移除前學年度之所有學期成績
                //  重修要檢查所有科目成績資料，若「科目名稱+級別」有2筆以上，表示該科目成績資料屬重修(成績計算時才用的到，此時判斷將造成資料謬誤)
                //string keyReReadingSR = sr.RefStudentID + "_" + sr.GradeYear.ToString();
                //if (!dReReading.ContainsKey(keyReReadingSR))
                //{
                //    dReReading.Add(keyReReadingSR, new List<SHSemesterScoreRecord>() { sr });
                //}
                //else
                //{
                //    if (sr.SchoolYear == dReReading[keyReReadingSR][0].SchoolYear)
                //        dReReading[keyReReadingSR].Add(sr);
                //    else
                //    {
                //        if (sr.SchoolYear > dReReading[keyReReadingSR][0].SchoolYear)
                //        {
                //            for (int i = 0; i < dReReading[keyReReadingSR].Count; i++)
                //            {
                //                _PersonalSemesterSubjectScoreInfo.Remove(sr.RefStudentID + "_" + dReReading[keyReReadingSR][0].SchoolYear.ToString() + "_" + dReReading[keyReReadingSR][0].Semester.ToString());
                //                KeyValuePair<int, int> schoolYearSemesterPair = new KeyValuePair<int, int>(dReReading[keyReReadingSR][0].SchoolYear, dReReading[keyReReadingSR][0].Semester);
                //                if (_PersonalSemesterSubjectScoreHistoryInfo.ContainsKey(sr.RefStudentID))
                //                    if (_PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID].ContainsKey(schoolYearSemesterPair))
                //                        _PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID].Remove(schoolYearSemesterPair);
                //            }
                //        }
                //        else
                //            continue;
                //    }
                //}

                if (!_PersonalSemesterSubjectScoreHistoryInfo.ContainsKey(sr.RefStudentID))
                    _PersonalSemesterSubjectScoreHistoryInfo.Add(sr.RefStudentID, new Dictionary<int, List<KeyValuePair<int, int>>>());

                if (!_PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID].ContainsKey(sr.GradeYear))
                {
                    _PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID].Add(sr.GradeYear, new List<KeyValuePair<int,int>>());
                    _PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID][sr.GradeYear].Add(new KeyValuePair<int,int>(sr.SchoolYear, sr.Semester));
                }
                else
                {
                    if (_PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID][sr.GradeYear].Where(x=>(x.Value == sr.Semester)).Count()==0)
                    {
                        _PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID][sr.GradeYear].Add(new KeyValuePair<int, int>(sr.SchoolYear, sr.Semester));
                    }
                    else
                    {
                        if (!_PersonalReReadSchoolYear.ContainsKey(sr.RefStudentID))
                            _PersonalReReadSchoolYear.Add(sr.RefStudentID, new List<int>());

                        _PersonalReReadSchoolYear[sr.RefStudentID].Add(sr.SchoolYear);
                    }
                }

                //  重讀成績不處理
                if (_PersonalReReadSchoolYear.ContainsKey(sr.RefStudentID) && _PersonalReReadSchoolYear[sr.RefStudentID].Contains(sr.SchoolYear))
                    continue;

                foreach(SHSchool.Data.SHSubjectScore ss in sr.Subjects.Values)
                {
                    //  不計學分不處理
                    if (ss.NotIncludedInCredit)
                        continue;

                    //KeyValuePair<int, int> schoolYearSemesterPair = new KeyValuePair<int, int>(ss.SchoolYear, ss.Semester);

                    //_PersonalSemesterSubjectScoreHistoryInfo[sr.RefStudentID][sr.GradeYear].Add(schoolYearSemesterPair);

                    string keySingleSubject = sr.RefStudentID + "_" + sr.SchoolYear.ToString() + "_" + sr.Semester.ToString() + "_" + ss.Subject + "_" + ((ss.Level.HasValue) ? ss.Level.ToString() : "");
                    string keySemesterSubject = sr.RefStudentID + "_" + sr.SchoolYear.ToString() + "_" + sr.Semester.ToString();

                    if (!_PersonalSubjectSemesterScoreInfo.ContainsKey(keySingleSubject))
                        _PersonalSubjectSemesterScoreInfo.Add(keySingleSubject, new SHSubjectSemesterScoreInfo());
                    
                    ProducePersonalSubjectSemesterScoreInfo(_PersonalSubjectSemesterScoreInfo, ss, sr.GradeYear);

                    //if (!_PersonalSemesterSubjectScoreHistoryInfo.ContainsKey(ss.RefStudentID))
                    //    _PersonalSemesterSubjectScoreHistoryInfo.Add(ss.RefStudentID, new Dictionary<KeyValuePair<int, int>, int>());

                    //if (!_PersonalSemesterSubjectScoreHistoryInfo[ss.RefStudentID].ContainsKey(schoolYearSemesterPair))
                    //    _PersonalSemesterSubjectScoreHistoryInfo[ss.RefStudentID].Add(schoolYearSemesterPair, sr.GradeYear);
                    
                    if (!_PersonalSemesterSubjectScoreInfo.ContainsKey(keySemesterSubject))
                        _PersonalSemesterSubjectScoreInfo.Add(keySemesterSubject, new List<SHSubjectSemesterScoreInfo>());

                    _PersonalSemesterSubjectScoreInfo[keySemesterSubject].Add(_PersonalSubjectSemesterScoreInfo[keySingleSubject]);
                }
            }
            //Dictionary<string, Dictionary<KeyValuePair<int, int>, int>> dd = new Dictionary<string, Dictionary<KeyValuePair<int, int>, int>>();
            //foreach (string studentID in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
            //{
            //    var _SemesterHistoryItems = from pair in _PersonalSemesterSubjectScoreHistoryInfo[studentID] orderby pair.Value descending, pair.Key.Key descending, pair.Key.Value ascending select pair;
            //    Dictionary<KeyValuePair<int, int>, int> d = new Dictionary<KeyValuePair<int, int>, int>();
            //    int yIndex = 0;
            //    int year = 0;
            //    foreach (KeyValuePair<KeyValuePair<int, int>, int> sh in _SemesterHistoryItems)
            //    {
            //        if (yIndex != sh.Value)
            //        {
            //            d.Add(new KeyValuePair<int, int>(sh.Key.Key, sh.Key.Value), sh.Value);
            //            yIndex = sh.Value;
            //            year = sh.Key.Key;
            //        }
            //        else 
            //        {
            //            if (year == sh.Key.Key)
            //            {
            //                d.Add(new KeyValuePair<int, int>(sh.Key.Key, sh.Key.Value), sh.Value);
            //                year = sh.Key.Key;
            //            }
            //        }
            //    }
            //    dd.Add(studentID, d);
            //}
            //_PersonalSemesterSubjectScoreHistoryInfo = dd;
            _SHSemesterScore = null;    
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
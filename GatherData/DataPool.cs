using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using FISCA.Presentation;
using SHSchool.Data;

namespace OfficialStudentRecordReport2010
{
    public class DataPool : IDisposable
    {   
        private List<SHSchool.Data.SHStudentRecord> _Students;                                                //所有學生
        private Dictionary<string, SHSchool.Data.SHStudentRecord> _StudentRecord;                             //所有學生
        private SHYearSubjectScoreInfo _SHYearSubjectScoreInfo;                                               //所有學生的學年科目成績
        private SHSemesterSubjectScoreInfo _SHSemesterSubjectScoreInfo;                                       //所有學生的學期科目成績
        private Dictionary<string, SHSchool.Data.SHClassRecord> _Class;                                       //單一學生的班級
        private Dictionary<string, SHSchool.Data.SHDepartmentRecord> _Department;                             //單一學生的科別
        private Dictionary<string, List<SHSchool.Data.SHStudentTagRecord>> _StudentTag;                       //單一學生的類別資訊
        private Dictionary<string, string> _GraduatePhoto;                                                    //單一學生的畢業照
        private Dictionary<string, string> _FreshmanPhoto;                                                    //單一學生的入學照
        private Dictionary<string, SHSchool.Data.SHParentRecord> _SHParentRecord;                             //單一學生的監護人(包含父、母親)
        private Dictionary<string, SHSchool.Data.SHPhoneRecord> _SHPhoneRecord;     //單一學生的電話(包含通訊、連絡、手機、電話1、電話2、電話3)
        private Dictionary<string, SHSchool.Data.SHAddressRecord> _SHAddressRecord; //單一學生的地址(包含通訊、連絡、地址1、地址2、地址3)
        private Dictionary<string, List<SHSchool.Data.SHSemesterEntryScoreRecord>> _SHSemesterEntryScore;     //單一學生的學業成績
        private Dictionary<string, List<SHSchool.Data.SHSchoolYearEntryScoreRecord>> _SHYearEntryScore;       //單一學生的學年學業成績
        private Dictionary<string, List<SHSchool.Data.SHMoralScoreRecord>> _SHMoralScoreRecord;               //單一學生的德行評量
        private Dictionary<string, SHSchool.Data.SHGradScoreRecord> _SHGradScoreRecord;                       //單一學生的畢業成績
        private MeritDeMeritInfo _MeritDeMeritInfo;                                                           //所有學生的獎懲
        private Dictionary<string, List<SHSchool.Data.SHUpdateRecordRecord>> _SHUpdateRecord;                 //單一學生的異動 
        private Dictionary<string, SHSchool.Data.SHScoreCalcRuleRecord> _PersonalSHScoreCalcRuleRecordInfo;   //成績計算規則  
        private Dictionary<string, List<string>> _PersonalStatusMappingInfo;                                  //學籍身份對照表          
        private List<string> _CoreCourseTable;                                                                //綜合高中學程核心課程標示 
        private Dictionary<string, SHLeaveInfoRecord> _PersonalSHLeaveInfo;                                   //單一學生的畢業及離校資訊
        //  單一學生重讀學年
        private Dictionary<string, List<int>> _PersonalReReadSchoolYear;
        //  單一學生學期分項重讀學年
        private Dictionary<string, List<int>> _PersonalSemesterEntryReReadSchoolYear;
        //  單一學生學年分項重讀學年
        private Dictionary<string, List<int>> _PersonalYearEntryReReadSchoolYear;
                
        public DataPool(IEnumerable<string> pStudentIDs)
        {
            InitializeData(pStudentIDs);
        }
        
        public DataPool(KeyValuePair<string, List<string>> kv)
        {
            if (kv.Key.ToUpper() == "CLASS")
            {
                _Students = SHStudent.SelectByClassIDs(kv.Value).Where(x=>x.Status == K12.Data.StudentRecord.StudentStatus.一般).ToList();
                InitializeData(_Students.Select(x => x.ID));
            }
            else if (kv.Key.ToUpper() == "STUDENT")
            {
                _Students = SHStudent.SelectByIDs(kv.Value);
                InitializeData(kv.Value);                
            }
            else
                throw new Exception("傳入錯誤的格式。");
        }
        
        private void InitializeData(IEnumerable<string> pStudentIDs) 
        {
            if (pStudentIDs.Count() == 0)
                throw new Exception("未傳入學生資料。");

            _PersonalReReadSchoolYear = new Dictionary<string, List<int>>();
            _PersonalSemesterEntryReReadSchoolYear = new Dictionary<string, List<int>>();
            _PersonalYearEntryReReadSchoolYear = new Dictionary<string, List<int>>();

            //  學年科目成績
            _SHYearSubjectScoreInfo = new SHYearSubjectScoreInfo(_Students);
            //  學期科目成績
            _SHSemesterSubjectScoreInfo = new SHSemesterSubjectScoreInfo(_Students);

            //  畢業及離校資訊
            _PersonalSHLeaveInfo = SHSchool.Data.SHLeaveInfo.SelectByStudentIDs(pStudentIDs).ToDictionary(x=>x.RefStudentID);

            //  學生
            _StudentRecord = new Dictionary<string, SHSchool.Data.SHStudentRecord>();
            foreach(SHSchool.Data.SHStudentRecord student in _Students)
            {
                _StudentRecord.Add(student.ID, student);
            }

            //  班級
            IEnumerable<string> classiDs = _Students.Select(x => x.RefClassID);
            Dictionary<string, SHSchool.Data.SHClassRecord> _SHClass = new Dictionary<string,SHSchool.Data.SHClassRecord>();
            _Class = new Dictionary<string, SHSchool.Data.SHClassRecord>();

            foreach(SHSchool.Data.SHClassRecord clazz in SHSchool.Data.SHClass.SelectByIDs(classiDs))
            {
                if (!_SHClass.ContainsKey(clazz.ID))
                    _SHClass.Add(clazz.ID, clazz);
            }
            foreach (SHSchool.Data.SHStudentRecord pStudent in _Students)
            {
                if (_SHClass.ContainsKey(pStudent.RefClassID))
                    if (!_Class.ContainsKey(pStudent.ID))
                        _Class.Add(pStudent.ID, _SHClass[pStudent.RefClassID]);
            }
            classiDs = null;
            //  科別
            List<string> _DepartmentIDs = new List<string>();
            Dictionary<string, string> _StudentDepartmentIDs = new Dictionary<string, string>();
            _Department = new Dictionary<string,SHSchool.Data.SHDepartmentRecord>();
            foreach (SHSchool.Data.SHStudentRecord pStudent in _Students)
            {
                if (pStudent.DepartmentID != "")
                {
                    _StudentDepartmentIDs.Add(pStudent.ID, pStudent.DepartmentID);

                    if (!_DepartmentIDs.Contains(pStudent.DepartmentID))
                        _DepartmentIDs.Add(pStudent.DepartmentID);
                }
                else 
                {
                    if (pStudent.RefClassID == "")
                        continue;
                    
                    if (_SHClass[pStudent.RefClassID] == null)
                        continue;
                        
                    if (_SHClass[pStudent.RefClassID].RefDepartmentID != "")
                    {
                        _StudentDepartmentIDs.Add(pStudent.ID, _SHClass[pStudent.RefClassID].RefDepartmentID);

                        if (!_DepartmentIDs.Contains(pStudent.DepartmentID))
                            _DepartmentIDs.Add(pStudent.DepartmentID);
                    }
                }
            }
            Dictionary<string, SHSchool.Data.SHDepartmentRecord> _DepartmentRecords = new Dictionary<string, SHDepartmentRecord>();
            List<SHSchool.Data.SHDepartmentRecord> allDepartmentRecords = SHSchool.Data.SHDepartment.SelectAll();
            foreach (SHSchool.Data.SHDepartmentRecord record in allDepartmentRecords)
            {
                if (_DepartmentIDs.Contains(record.ID))
                    _DepartmentRecords.Add(record.ID, record);
            }
                //SHSchool.Data.SHDepartment.SelectByIDs(_DepartmentIDs).ToDictionary(x=>x.ID);
            foreach (SHSchool.Data.SHStudentRecord pStudent in _Students)
            {
                if (pStudent.DepartmentID != "")
                {
                    if (_StudentDepartmentIDs.ContainsKey(pStudent.ID))
                        _Department.Add(pStudent.ID, _DepartmentRecords[_StudentDepartmentIDs[pStudent.ID]]);
                }
                else
                {
                    if (pStudent.RefClassID == "")
                        continue;
                    
                    if (_SHClass[pStudent.RefClassID] == null)
                        continue;
                        
                    if (_SHClass[pStudent.RefClassID].RefDepartmentID != "")
                        _Department.Add(pStudent.ID, _DepartmentRecords[_SHClass[pStudent.RefClassID].RefDepartmentID]);
                }
            }
            _SHClass = null;
            _DepartmentIDs = null;
            _StudentDepartmentIDs = null;
            //  畢業照
            _GraduatePhoto = K12.Data.Photo.SelectGraduatePhoto(pStudentIDs);

            //  入學照
            _FreshmanPhoto = K12.Data.Photo.SelectFreshmanPhoto(pStudentIDs);

            //  監護人
            List<SHSchool.Data.SHParentRecord> _Parents = SHSchool.Data.SHParent.SelectByStudentIDs(pStudentIDs);
            _SHParentRecord = new Dictionary<string,SHSchool.Data.SHParentRecord>();
            foreach(SHSchool.Data.SHParentRecord pr in _Parents)
            {
                _SHParentRecord.Add(pr.RefStudentID, pr);
            }
            _Parents = null;
            //  電話
            List<SHSchool.Data.SHPhoneRecord> _Phone = SHSchool.Data.SHPhone.SelectByStudentIDs(pStudentIDs);
            _SHPhoneRecord = new Dictionary<string,SHSchool.Data.SHPhoneRecord>();
            foreach (SHSchool.Data.SHPhoneRecord pr in _Phone)
            {
                _SHPhoneRecord.Add(pr.RefStudentID, pr);
            }
            _Phone = null;
            //  地址
            List<SHSchool.Data.SHAddressRecord> _Address = SHSchool.Data.SHAddress.SelectByStudentIDs(pStudentIDs);
            _SHAddressRecord = new Dictionary<string,SHSchool.Data.SHAddressRecord>();
            foreach (SHSchool.Data.SHAddressRecord pr in _Address)
            {
                _SHAddressRecord.Add(pr.RefStudentID, pr);
            }
            _Address = null;
            //  學業成績
            List<SHSchool.Data.SHSemesterEntryScoreRecord> _SHSemesterEntryScores = SHSchool.Data.SHSemesterEntryScore.Select(null, pStudentIDs, "學習", null, false);
            _SHSemesterEntryScore = new Dictionary<string,List<SHSchool.Data.SHSemesterEntryScoreRecord>>();
            foreach (SHSchool.Data.SHSemesterEntryScoreRecord sr in _SHSemesterEntryScores)
            {
                if (!_SHSemesterEntryScore.ContainsKey(sr.RefStudentID))
                    _SHSemesterEntryScore.Add(sr.RefStudentID, new List<SHSchool.Data.SHSemesterEntryScoreRecord>());

                _SHSemesterEntryScore[sr.RefStudentID].Add(sr);
            }
            _SHSemesterEntryScores = null;
            //  學年學業成績
            List<SHSchool.Data.SHSchoolYearEntryScoreRecord> _SHYearEntryScores = SHSchool.Data.SHSchoolYearEntryScore.Select(null, pStudentIDs, "學習", null);
            _SHYearEntryScore = new Dictionary<string, List<SHSchool.Data.SHSchoolYearEntryScoreRecord>>();
            foreach (SHSchool.Data.SHSchoolYearEntryScoreRecord sr in _SHYearEntryScores)
            {
                if (!_SHYearEntryScore.ContainsKey(sr.RefStudentID))
                    _SHYearEntryScore.Add(sr.RefStudentID, new List<SHSchool.Data.SHSchoolYearEntryScoreRecord>());

                _SHYearEntryScore[sr.RefStudentID].Add(sr);
            }
            _SHYearEntryScores = null;
            //  德行評量
            List<SHSchool.Data.SHMoralScoreRecord> _SHMoralScoreRecords = SHSchool.Data.SHMoralScore.SelectByStudentIDs(pStudentIDs);
            Dictionary<string, List<SHSchool.Data.SHMoralScoreRecord>> sHMoralScoreRecord = new Dictionary<string, List<SHSchool.Data.SHMoralScoreRecord>>();
            _SHMoralScoreRecord = new Dictionary<string,List<SHSchool.Data.SHMoralScoreRecord>>();
            foreach (SHSchool.Data.SHMoralScoreRecord sr in _SHMoralScoreRecords)
            {
                if (!_SHMoralScoreRecord.ContainsKey(sr.RefStudentID))
                    _SHMoralScoreRecord.Add(sr.RefStudentID, new List<SHSchool.Data.SHMoralScoreRecord>());

                _SHMoralScoreRecord[sr.RefStudentID].Add(sr);
            }
            _SHMoralScoreRecords = null;
            //  畢業成績
            List<SHSchool.Data.SHGradScoreRecord> _SHGradScoreRecords = SHSchool.Data.SHGradScore.SelectByIDs(pStudentIDs);
            _SHGradScoreRecord = new Dictionary<string,SHSchool.Data.SHGradScoreRecord>();
            foreach (SHSchool.Data.SHGradScoreRecord sr in _SHGradScoreRecords)
            {
                _SHGradScoreRecord.Add(sr.RefStudentID, sr);
            }
            _SHGradScoreRecords = null;
            //  獎懲
            _MeritDeMeritInfo = new MeritDeMeritInfo(_Students);
            //  異動
            List<SHSchool.Data.SHUpdateRecordRecord> _Shurr = SHSchool.Data.SHUpdateRecord.SelectByStudentIDs(pStudentIDs);
            _SHUpdateRecord = new Dictionary<string, List<SHSchool.Data.SHUpdateRecordRecord>>();
            foreach (SHSchool.Data.SHUpdateRecordRecord sr in _Shurr)
            {
                if (!_SHUpdateRecord.ContainsKey(sr.StudentID))
                    _SHUpdateRecord.Add(sr.StudentID, new List<SHSchool.Data.SHUpdateRecordRecord>());

                _SHUpdateRecord[sr.StudentID].Add(sr);
            }
            _Shurr = null;
            //  成績計算規則 
            _PersonalSHScoreCalcRuleRecordInfo = new Dictionary<string, SHSchool.Data.SHScoreCalcRuleRecord>();
            Dictionary<string, SHSchool.Data.SHScoreCalcRuleRecord> _SHScoreCalcRule = SHSchool.Data.SHScoreCalcRule.SelectAll().ToDictionary(x => x.ID);
            foreach (SHSchool.Data.SHStudentRecord student in _Students)
            {
                string _ScoreCalcRuleID = string.Empty;

                if (!string.IsNullOrEmpty(student.OverrideScoreCalcRuleID))
                {
                    _ScoreCalcRuleID = student.OverrideScoreCalcRuleID;
                }
                else
                {
                    SHSchool.Data.SHClassRecord clazz = this.GetClass(student.ID);

                    if (clazz == null)
                        continue;
                    else
                    {
                        if (string.IsNullOrEmpty(clazz.RefScoreCalcRuleID))
                            continue;
                        else
                            _ScoreCalcRuleID = clazz.RefScoreCalcRuleID;
                    }
                }
                if (_SHScoreCalcRule.ContainsKey(_ScoreCalcRuleID))
                    _PersonalSHScoreCalcRuleRecordInfo.Add(student.ID, _SHScoreCalcRule[_ScoreCalcRuleID]);

            }
            //  身份類別
            _StudentTag = new Dictionary<string, List<SHSchool.Data.SHStudentTagRecord>>();
            List<SHSchool.Data.SHStudentTagRecord> _StudentTags = SHSchool.Data.SHStudentTag.SelectByStudentIDs(pStudentIDs);
            foreach (SHSchool.Data.SHStudentTagRecord st in _StudentTags)
            {
                if (!_StudentTag.ContainsKey(st.RefStudentID))
                    _StudentTag.Add(st.RefStudentID, new List<SHSchool.Data.SHStudentTagRecord>());

                _StudentTag[st.RefStudentID].Add(st);
            }
            _StudentTags = null;
            //  學籍身份對照表
            _PersonalStatusMappingInfo = new Dictionary<string, List<string>>();
            List<SHSchool.Data.SHPermrecStatusMappingInfo> _StatusMappingInfo = SHSchool.Data.SHPermrecStatusMapping.SelectAll();
            foreach (SHSchool.Data.SHPermrecStatusMappingInfo sm in _StatusMappingInfo)
                _PersonalStatusMappingInfo.Add(sm.Name, sm.TagFullNames);

            _StatusMappingInfo = null;

            //  綜合高中學程核心課程標示
            _CoreCourseTable = new List<string>();
            List<SHSubjectTableRecord> _CourseTable = SHSubjectTable.Select(null, null, null);
            foreach (SHSubjectTableRecord sr in _CourseTable)
            {
                foreach (SHSubjectTableSubject ss in sr.Subjects)
                {
                    if (ss.IsCore.HasValue)
                    {
                        if (ss.IsCore.Value == false)
                            continue;
                    }
                    else
                        continue;

                    string key = string.Empty;
                    if (ss.Levels.Count == 0)
                    {
                        key = sr.Name + "_" + ss.Name;
                        if (!_CoreCourseTable.Contains(key))
                            _CoreCourseTable.Add(key);
                    }
                    else
                    {
                        foreach (int a in ss.Levels)
                        {
                            key = sr.Name + "_" + ss.Name + "_" + a.ToString();
                            if (!_CoreCourseTable.Contains(key))
                                _CoreCourseTable.Add(key);
                        }
                    }
                }
            }
            _CourseTable = null;
        }

        public SHLeaveInfoRecord GetPersonalSHLeaveInfo(string pStudentID)
        {
            if (_PersonalSHLeaveInfo == null)
                return null;
            else if (_PersonalSHLeaveInfo.ContainsKey(pStudentID))
                return _PersonalSHLeaveInfo[pStudentID];
            else
                return null;
        }


        /// <summary>
        /// 畢業規定總學分
        /// </summary>
        /// <param name="pStudentID">學生系統編號</param>
        /// <returns>學分</returns>
        public string GetGraduationDeservedCredit(string pStudentID)
        {
            string GraduationDeservedCredit = string.Empty;

            if (!_PersonalSHScoreCalcRuleRecordInfo.ContainsKey(pStudentID))
                return "";
            else
            {
                //  <畢業學分數><學科累計總學分數>160</學科累計總學分數><必修學分數>120</必修學分數><選修學分數>40</選修學分數></畢業學分數>
                XDocument doc = XDocument.Parse("<root>" + _PersonalSHScoreCalcRuleRecordInfo[pStudentID].Content.InnerXml + "</root>");

                if (doc.Document == null)
                    return "";

                if (doc.Document.Element("root").Element("畢業學分數") == null)
                    return "";

                if (doc.Document.Element("root").Element("畢業學分數").Element("學科累計總學分數") == null)
                    return "";

                GraduationDeservedCredit = doc.Document.Element("root").Element("畢業學分數").Element("學科累計總學分數").Value;

                return GraduationDeservedCredit;
            }
        }

        /****<學生類別 一年級及格標準=\"60\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"60\" 二年級補考標準=\"0\" 四年級及格標準=\"60\" 四年級補考標準=\"0\" 類別=\"預設\" />
         * 比對學生成績身份，若比對不到則及格標準使用「預設」
         ****/
        public decimal GetPassingStandard(string pStudentID, int? pSchoolYear)
        {
            pSchoolYear = pSchoolYear.HasValue ? pSchoolYear.Value : 0;
            decimal passingStandard = 60;
            decimal tmpPassingStandard = 60;
            //bool youGotIt = false;
            List<SHSchool.Data.SHStudentTagRecord> pStudentTag = GetStudentTag(pStudentID);

            if (pStudentTag == null)
                return passingStandard;

            List<string> pStudentTagFullName = new List<string>();
            List<string> pStudentTagPrefix = new List<string>();
            foreach (SHSchool.Data.SHStudentTagRecord sr in pStudentTag)
            {
                pStudentTagFullName.Add(sr.FullName);
                pStudentTagPrefix.Add(sr.Prefix);
            }

            if (!_PersonalSHScoreCalcRuleRecordInfo.ContainsKey(pStudentID))
                return passingStandard;

            XDocument doc = XDocument.Parse("<root>" + _PersonalSHScoreCalcRuleRecordInfo[pStudentID].Content.InnerXml + "</root>");
            
            if (doc.Document == null)
                return passingStandard;

            if (doc.Document.Element("root").Element("及格標準") == null)
                return passingStandard;

            if (doc.Document.Element("root").Element("及格標準").Elements("學生類別") == null)
                return passingStandard;

            foreach (XElement x in doc.Document.Element("root").Element("及格標準").Elements("學生類別"))
            {
                if (x.Attribute("類別") != null)
                {
                    if (pStudentTagFullName.Contains(x.Attribute("類別").Value) || pStudentTagPrefix.Contains(x.Attribute("類別").Value))
                    {
                        if (x.Attribute(NumericToChinese(pSchoolYear) + "年級及格標準") != null)
                        {
                            bool result = decimal.TryParse(x.Attribute(NumericToChinese(pSchoolYear) + "年級及格標準").Value, out tmpPassingStandard);
                            if (result)
                            {
                                if (tmpPassingStandard < passingStandard)
                                {
                                    //youGotIt = true;
                                    passingStandard = tmpPassingStandard;
                                }
                            }
                        }
                    }
                }
            }

            //if (!youGotIt)
            //{
            try
            {
                var elem = doc.XPathSelectElement("root/及格標準/學生類別[@類別='預設']");
                bool result = decimal.TryParse(elem.Attribute(NumericToChinese(pSchoolYear) + "年級及格標準").Value, out tmpPassingStandard);
                if (result)
                {
                    if (passingStandard > tmpPassingStandard)
                        passingStandard = tmpPassingStandard;
                }
            }
            catch
            {
                return passingStandard;
            }
            //}

            return passingStandard;
        }

        public string TransferStudentTagToIdentity(string pStudentID)
        {
            string identity = "一般生";
            string identityTmp = string.Empty;

            List<SHSchool.Data.SHStudentTagRecord> pStudentTag = GetStudentTag(pStudentID);

            if (pStudentTag == null)
                return identity;

            List<string> pStudentTags = new List<string>();
            foreach (SHSchool.Data.SHStudentTagRecord sr in pStudentTag)
                pStudentTags.Add(sr.FullName);

            foreach (KeyValuePair<string, List<string>> kv in _PersonalStatusMappingInfo)
            {
                foreach (string s in pStudentTags)
                {
                    if (kv.Value.Contains(s))
                        identityTmp += kv.Key + ",";
                }
            }

            if (identityTmp.Length > 0)
                return (identityTmp.EndsWith(",") ? identityTmp.Remove(identityTmp.Length - 1, 1) : identityTmp);
            else
                return identity;
        }

        public char NumericToChinese(int? pNumber)
        {
            int qNumber = pNumber.HasValue ? pNumber.Value : 0;
            if (pNumber > 10 || pNumber < 0)
                throw new Exception("轉換的數字必須介於 0~10");

            char[] number = new char[] {'○', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'};

            return number[qNumber];
        }

        /***********************************
         * 學生之學年學業成績符合下列各款規定之一者，准予升級：
         *   一、各科目學年成績均及格。
         *   二、學年成績不及格之各科目每週教學總節數，未超過修習各科目每週教學總節數二分之一，且不及格之科目成績無零分，而學年總平均成績及格。
         *   三、學年成績不及格之科目未超過修習全部科目二分之一，且其不及格之科目成績無零分，而學年總平均成績及格。
         ***********************************/
        public bool RetainInTheSameGrade(string pStudentID, KeyValuePair<int, int> kv)
        {
            string key = pStudentID + "_" + kv.Value;
            decimal? totalCredit = 0;
            decimal? notPassTotalCredit = 0;
            int totalSubjectAmount = 0;
            int notPassTotalSubjectAmount = 0;
            bool hasZeroScoreSubject = false;
            decimal yearEntryScore = GetYearEntryScore(pStudentID, kv.Value, SHYearSubjectScoreInfo.Entry.學業);

            List<SHSubjectYearScoreInfo> _YearSubjectScoreInfo = _SHYearSubjectScoreInfo.GetPersonalYearSubjectScoreInfo(pStudentID, kv.Value);

            foreach (SHSubjectYearScoreInfo ss in _YearSubjectScoreInfo)
            {
                //  所有科目每週教學總節數(ischool 無課時，故取學分)
                totalCredit += (ss.Credit.HasValue ? ss.Credit.Value : 0);
                //  所有科目數
                totalSubjectAmount++;

                if ((ss.Score.HasValue ? ss.Score.Value : 0) < this.GetPassingStandard(pStudentID, ss.GradeYear))
                {
                    notPassTotalCredit += (ss.Credit.HasValue ? ss.Credit.Value : 0);
                    notPassTotalSubjectAmount += 1;
                }

                if (ss.Score.HasValue && ss.Score.Value == 0)
                    hasZeroScoreSubject = true;
            }

            //  各科目學年成績均及格：升級
            if (notPassTotalSubjectAmount == 0)
                return false;

            //  剛好 1/2：升級
            //  學年成績不及格之各科目每週教學總節數，未超過修習各科目每週教學總節數二分之一，且不及格之科目成績無零分，而學年總平均成績及格
            if (((notPassTotalCredit * 2) <= totalCredit) && (!hasZeroScoreSubject) && (GetYearEntryScore(pStudentID, kv.Value, SHYearSubjectScoreInfo.Entry.學業) >= (GetPassingStandard(pStudentID, kv.Key))))
                return false;

            //  剛好 1/2：升級
            //  學年成績不及格之科目未超過修習全部科目二分之一，且其不及格之科目成績無零分，而學年總平均成績及格
            if (((notPassTotalSubjectAmount * 2) <= totalSubjectAmount) && (!hasZeroScoreSubject) && (GetYearEntryScore(pStudentID, kv.Value, SHYearSubjectScoreInfo.Entry.學業) >= (GetPassingStandard(pStudentID, kv.Key))))
                return false;

            return true;
        }

        public decimal GetProSubjectAccquiredCredit(string pStudentID)
        {
            decimal proSubjectAccumulatedCredit = 0;

            if (!_PersonalSHScoreCalcRuleRecordInfo.ContainsKey(pStudentID))
                return proSubjectAccumulatedCredit;

            XDocument doc = XDocument.Parse("<root>" + _PersonalSHScoreCalcRuleRecordInfo[pStudentID].Content.InnerXml + "</root>");

            if (doc.Document == null)
                return proSubjectAccumulatedCredit;

            if (doc.Document.Element("root").Element("核心科目表") == null)
                return proSubjectAccumulatedCredit;

            if (doc.Document.Element("root").Element("核心科目表").Element("專業科目表") == null)
                return proSubjectAccumulatedCredit;

            List<SHSubjectTableRecord> shSubjectTable = this._SHSemesterSubjectScoreInfo.GetSubjectTable();
            List<string> proSubjects = new List<string>();
            foreach (SHSubjectTableRecord sr in shSubjectTable)
            {
                if (sr.Name == doc.Document.Element("root").Element("核心科目表").Element("專業科目表").Value)
                {
                    foreach (SHSubjectTableSubject ss in sr.Subjects)
                    {
                        if (ss.Levels.Count == 0)
                        {
                            if (!proSubjects.Contains(ss.Name))
                                proSubjects.Add(ss.Name);
                        }
                        else
                        {
                            foreach (int s in ss.Levels)
                            {
                                if (!proSubjects.Contains(ss.Name + "_" + s.ToString()))
                                    proSubjects.Add(ss.Name + "_" + s.ToString());
                            }
                        }
                    }
                }
            }

            List<SHSubjectSemesterScoreInfo> shSubjectSemesterScoreInfo = this._SHSemesterSubjectScoreInfo.GetPersonalAllSemesterSubjectScoreInfo(pStudentID);

            //  學期歷程
            Dictionary<int, List<KeyValuePair<int, int>>> _PersonalSemesterSubjectScoreHistoryInfo = this.GetPersonalSemesterSubjectScoreHistoryInfo(pStudentID);
            List<KeyValuePair<int, int>> kvs = new List<KeyValuePair<int, int>>();
            foreach (int index in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                _PersonalSemesterSubjectScoreHistoryInfo[index].ForEach(x => kvs.Add(x));
            if (shSubjectSemesterScoreInfo != null)
            {
                foreach (SHSubjectSemesterScoreInfo ss in shSubjectSemesterScoreInfo)
                {
                    if (kvs.Where(y => (y.Key == ss.SchoolYear && y.Value == ss.Semester)).Count() > 0)
                    {
                        string content = ss.SubjectName + (ss.Level.HasValue ? "_" + ss.Level.Value.ToString() : "");

                        if (proSubjects.Contains(content))
                            if (ss.Credit.HasValue && ss.Pass.HasValue && ss.Pass.Value)
                                proSubjectAccumulatedCredit += ss.Credit.Value;
                    }
                }
            }

            return proSubjectAccumulatedCredit;
        }

        public decimal GetProSubjectAccumulatedCredit(string pStudentID)
        {
            decimal proSubjectAccumulatedCredit = 0;

            if (!_PersonalSHScoreCalcRuleRecordInfo.ContainsKey(pStudentID))
                return proSubjectAccumulatedCredit;

            XDocument doc = XDocument.Parse("<root>" + _PersonalSHScoreCalcRuleRecordInfo[pStudentID].Content.InnerXml + "</root>");

            if (doc.Document == null)
                return proSubjectAccumulatedCredit;

            if (doc.Document.Element("root").Element("核心科目表") == null)
                return proSubjectAccumulatedCredit;

            if (doc.Document.Element("root").Element("核心科目表").Element("專業科目表") == null)
                return proSubjectAccumulatedCredit;

            List<SHSubjectTableRecord> shSubjectTable = this._SHSemesterSubjectScoreInfo.GetSubjectTable();
            List<string> proSubjects = new List<string>();
            foreach(SHSubjectTableRecord sr in shSubjectTable)
            {
                if (sr.Name == doc.Document.Element("root").Element("核心科目表").Element("專業科目表").Value)
                {
                    foreach (SHSubjectTableSubject ss in sr.Subjects)
                    {
                        if (ss.Levels.Count == 0)
                        {
                            if (!proSubjects.Contains(ss.Name))
                                proSubjects.Add(ss.Name);
                        }
                        else
                        {
                            foreach (int s in ss.Levels)
                            {
                                if (!proSubjects.Contains(ss.Name + "_" + s.ToString()))
                                    proSubjects.Add(ss.Name + "_" + s.ToString());
                            }
                        }
                    }
                }
            }

            List<SHSubjectSemesterScoreInfo> shSubjectSemesterScoreInfo = this._SHSemesterSubjectScoreInfo.GetPersonalAllSemesterSubjectScoreInfo(pStudentID);
            //  學期歷程
            Dictionary<int, List<KeyValuePair<int, int>>> _PersonalSemesterSubjectScoreHistoryInfo = this.GetPersonalSemesterSubjectScoreHistoryInfo(pStudentID);
            List<KeyValuePair<int, int>> kvs = new List<KeyValuePair<int, int>>();
            foreach (int index in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                _PersonalSemesterSubjectScoreHistoryInfo[index].ForEach(x => kvs.Add(x));
            if (shSubjectSemesterScoreInfo != null)
            {
                foreach (SHSubjectSemesterScoreInfo ss in shSubjectSemesterScoreInfo)
                {
                    if (kvs.Where(y => (y.Key == ss.SchoolYear && y.Value == ss.Semester)).Count() > 0)
                    {
                        string content = ss.SubjectName + (ss.Level.HasValue ? "_" + ss.Level.Value.ToString() : "");

                        if (proSubjects.Contains(content))
                            if (ss.Credit.HasValue)
                                proSubjectAccumulatedCredit += ss.Credit.Value;
                    }
                }
            }
            
            return proSubjectAccumulatedCredit;
        }

        public List<SHSchool.Data.SHStudentTagRecord> GetStudentTag(string pStudentID)
        {
            if (_StudentTag.ContainsKey(pStudentID))
                return _StudentTag[pStudentID];
            else
                return null;
        }

        public List<SHSchool.Data.SHStudentRecord> GetSelectStudents(List<string> pStudentIDs)
        {
            return _Students.Where(x => pStudentIDs.Contains(x.ID)).Select(x=>x).ToList();
        }

        public List<SHSchool.Data.SHStudentRecord> GetAllStudents()
        {
            return _Students; 
        }

        public SHSchool.Data.SHStudentRecord GetStudent(string pStudentID)
        {
            if (_StudentRecord.ContainsKey(pStudentID))
                return _StudentRecord[pStudentID];
            else
                return null;
        }

        public SHSchool.Data.SHClassRecord GetClass(string pStudentID)
        {
            if (_Class.ContainsKey(pStudentID))
                return _Class[pStudentID];
            else
                return null;
        }

        public SHSchool.Data.SHDepartmentRecord GetDepartment(string pStudentID)
        {
            if (_Department.ContainsKey(pStudentID))
                return _Department[pStudentID];
            else
                return null;
        }

        public string GetGraduatePhoto(string pStudentID)
        {
            if (_GraduatePhoto.ContainsKey(pStudentID))
                return _GraduatePhoto[pStudentID];
            else
                return string.Empty;
        }
        
        //  判讀學生修習科目是否為核心課程
        public bool IsCoreCourse(string pStudentID, string pSubjectName, int? pLevel)
        {
            string key = string.Empty;
            SHDepartmentRecord dept = _Department.ContainsKey(pStudentID) ? _Department[pStudentID] : null;

            if (dept == null)
            {
                SHLeaveInfoRecord record = this.GetPersonalSHLeaveInfo(pStudentID);
                if (record == null)
                    return false;
                else
                {
                    if (string.IsNullOrWhiteSpace(record.DepartmentName))
                        return false;

                    string[] SubDepartments = record.DepartmentName.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                    if (SubDepartments.Count() == 0)
                        return false;
                    else
                        key = SubDepartments[SubDepartments.Count() - 1] + "_" + pSubjectName + (((pLevel.HasValue) ? "_" + pLevel.Value.ToString() : ""));
                }
            }
            else
                key = dept.SubDepartment + "_" + pSubjectName + (((pLevel.HasValue) ? "_" + pLevel.Value.ToString() : ""));

            bool youGotIt = false;

            _CoreCourseTable.ForEach(delegate(String name)
            {
                if (!String.IsNullOrEmpty(name) && key.IndexOf(name) >= 0)
                    youGotIt = true;
            });

            return youGotIt;
        }

        public string GetFreshmanPhoto(string pStudentID)
        {
            if (_FreshmanPhoto.ContainsKey(pStudentID))
                return _FreshmanPhoto[pStudentID];
            else
                return string.Empty;
        }

        public SHSchool.Data.SHParentRecord GetParent(string pStudentID)
        {
            if (_SHParentRecord.ContainsKey(pStudentID))
                return _SHParentRecord[pStudentID];
            else
                return null;
        }

        public SHSchool.Data.SHPhoneRecord GetPhone(string pStudentID)
        {
            if (_SHPhoneRecord.ContainsKey(pStudentID))
                return _SHPhoneRecord[pStudentID];
            else
                return null;
        }

        public SHSchool.Data.SHAddressRecord GetAddress(string pStudentID)
        {
            if (_SHAddressRecord.ContainsKey(pStudentID))
                return _SHAddressRecord[pStudentID];
            else
                return null;
        }

        public List<int> GetPersonalSemesterEntryReReadSchoolYear(string pStudentID)
        {
            if (_PersonalSemesterEntryReReadSchoolYear.ContainsKey(pStudentID))
                return _PersonalSemesterEntryReReadSchoolYear[pStudentID];
            else
                return new List<int>();
        }

        public List<int> GetPersonalYearEntryReReadSchoolYear(string pStudentID)
        {
            if (_PersonalYearEntryReReadSchoolYear.ContainsKey(pStudentID))
                return _PersonalYearEntryReReadSchoolYear[pStudentID];
            else
                return new List<int>();
        }

        public Dictionary<int, int> GetSchoolYearHistory(string pStudentID)
        {
            Dictionary<int, int> schoolYearHistory = new Dictionary<int, int>();

            Dictionary<int, List<KeyValuePair<int, int>>> _PersonalSemesterSubjectScoreHistoryInfo = this.GetPersonalSemesterSubjectScoreHistoryInfo(pStudentID);
            Dictionary<int, int> _PersonalYearSubjectScoreHistoryInfo = this.GetPersonalYearSubjectScoreHistoryInfo(pStudentID);
            Dictionary<int, List<KeyValuePair<int, int>>> _PersonalSemesterEntryScoreHistory = this.GetSemesterEntryScoreHistory(pStudentID);
            Dictionary<int, int> _PersonalYearEntryScoreHistory = this.GetYearEntryScoreHistory(pStudentID);

            if (_PersonalSemesterSubjectScoreHistoryInfo != null && _PersonalSemesterSubjectScoreHistoryInfo.Count > 0)
            {
                foreach (int key in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                {
                    KeyValuePair<int, int> kv = _PersonalSemesterSubjectScoreHistoryInfo[key].OrderByDescending(x=>x.Key).ElementAt(0);
                    if (!schoolYearHistory.ContainsKey(kv.Key))
                        schoolYearHistory.Add(kv.Key, key);
                }
            }
            if (_PersonalSemesterEntryScoreHistory != null && _PersonalSemesterEntryScoreHistory.Count > 0)
            {
                foreach (int key in _PersonalSemesterEntryScoreHistory.Keys)
                {
                    KeyValuePair<int, int> kv = _PersonalSemesterEntryScoreHistory[key].OrderByDescending(x => x.Key).ElementAt(0);
                    if (!schoolYearHistory.ContainsKey(kv.Key))
                        schoolYearHistory.Add(kv.Key, key);
                }
            }
            if (_PersonalYearSubjectScoreHistoryInfo != null && _PersonalYearSubjectScoreHistoryInfo.Count > 0)
            {
                foreach (int key in _PersonalYearSubjectScoreHistoryInfo.Keys)
                {
                    if (!schoolYearHistory.ContainsKey(_PersonalYearSubjectScoreHistoryInfo[key]))
                        schoolYearHistory.Add(_PersonalYearSubjectScoreHistoryInfo[key], key);
                }
            }
            if (_PersonalYearEntryScoreHistory != null && _PersonalYearEntryScoreHistory.Count > 0)
            {
                foreach (int key in _PersonalYearEntryScoreHistory.Keys)
                {
                    if (!schoolYearHistory.ContainsKey(_PersonalYearEntryScoreHistory[key]))
                        schoolYearHistory.Add(_PersonalYearEntryScoreHistory[key], key);
                }
            }

            return schoolYearHistory;
        }

        public List<int> GetReReadSchoolYear(string pStudentID)
        {
            List<int> lstReReadSchoolYear = new List<int>();

            List<int> lstSemesterReReadSchoolYear = this.GetPersonalSemesterReReadSchoolYear(pStudentID);
            List<int> lstYearReReadSchoolYear = this.GetPersonalYearReReadSchoolYear(pStudentID);
            List<int> lstSemesterEntryReReadSchoolYear = this.GetPersonalSemesterEntryReReadSchoolYear(pStudentID);
            List<int> lstYearEntryReReadSchoolYear = this.GetPersonalYearEntryReReadSchoolYear(pStudentID);

            lstReReadSchoolYear = lstSemesterReReadSchoolYear.Intersect(lstSemesterEntryReReadSchoolYear).Intersect(lstYearReReadSchoolYear).Intersect(lstYearEntryReReadSchoolYear).ToList();

            //lstSemesterReReadSchoolYear.ForEach((x)=>
            //{
            //    if (!lstReReadSchoolYear.Contains(x))
            //        lstReReadSchoolYear.Add(x);
            //});
            //lstYearReReadSchoolYear.ForEach((x) =>
            //{
            //    if (!lstReReadSchoolYear.Contains(x))
            //        lstReReadSchoolYear.Add(x);
            //});
            //lstSemesterEntryReReadSchoolYear.ForEach((x) =>
            //{
            //    if (!lstReReadSchoolYear.Contains(x))
            //        lstReReadSchoolYear.Add(x);
            //});
            //lstYearEntryReReadSchoolYear.ForEach((x) =>
            //{
            //    if (!lstReReadSchoolYear.Contains(x))
            //        lstReReadSchoolYear.Add(x);
            //});

            return lstReReadSchoolYear;
        }

        public Dictionary<int, List<KeyValuePair<int, int>>> GetSemesterEntryScoreHistory(string pStudentID)
        {
            List<SHSemesterEntryScoreRecord> _SHSemesterEntryScoreRecords = this.GetSemesterEntryScoreInfo(pStudentID).OrderByDescending(x => x.SchoolYear).Select(x => x).ToList();

            Dictionary<int, List<KeyValuePair<int, int>>> _SemesterEntryScoreHistory = new Dictionary<int, List<KeyValuePair<int, int>>>();
            foreach (SHSemesterEntryScoreRecord _SHSemesterEntryScoreRecord in _SHSemesterEntryScoreRecords)
            {
                if (!_SemesterEntryScoreHistory.ContainsKey(_SHSemesterEntryScoreRecord.GradeYear))
                {
                    _SemesterEntryScoreHistory.Add(_SHSemesterEntryScoreRecord.GradeYear, new List<KeyValuePair<int, int>>());

                    _SemesterEntryScoreHistory[_SHSemesterEntryScoreRecord.GradeYear].Add(new KeyValuePair<int, int>(_SHSemesterEntryScoreRecord.SchoolYear, _SHSemesterEntryScoreRecord.Semester));
                }
                else
                {
                    if (_SemesterEntryScoreHistory[_SHSemesterEntryScoreRecord.GradeYear].Where(x => (x.Value == _SHSemesterEntryScoreRecord.Semester)).Count() == 0)
                    {
                        _SemesterEntryScoreHistory[_SHSemesterEntryScoreRecord.GradeYear].Add(new KeyValuePair<int, int>(_SHSemesterEntryScoreRecord.SchoolYear, _SHSemesterEntryScoreRecord.Semester));
                    }
                    else
                    {
                        if (!_PersonalSemesterEntryReReadSchoolYear.ContainsKey(pStudentID))
                            _PersonalSemesterEntryReReadSchoolYear.Add(pStudentID, new List<int>());

                        _PersonalSemesterEntryReReadSchoolYear[pStudentID].Add(_SHSemesterEntryScoreRecord.SchoolYear);
                    }
                }
            }

            return _SemesterEntryScoreHistory;
        }

        public Dictionary<int, int> GetYearEntryScoreHistory(string pStudentID)
        {
            List<SHSchool.Data.SHSchoolYearEntryScoreRecord> _SHYearEntryScoreRecords = this.GetYearEntryScoreInfo(pStudentID).OrderByDescending(x => x.SchoolYear).Select(x => x).ToList();

            Dictionary<int, int> _YearEntryScoreHistory = new Dictionary<int, int>();
            foreach (SHSchoolYearEntryScoreRecord _SHYearEntryScoreRecord in _SHYearEntryScoreRecords)
            {
                if (!_YearEntryScoreHistory.ContainsKey(_SHYearEntryScoreRecord.GradeYear))
                    _YearEntryScoreHistory.Add(_SHYearEntryScoreRecord.GradeYear, _SHYearEntryScoreRecord.SchoolYear);
                else
                {
                    if (!this._PersonalYearEntryReReadSchoolYear.ContainsKey(pStudentID))
                        _PersonalYearEntryReReadSchoolYear.Add(pStudentID, new List<int>());

                    _PersonalYearEntryReReadSchoolYear[pStudentID].Add(_SHYearEntryScoreRecord.SchoolYear);
                }
            }

            return _YearEntryScoreHistory;
        }

        public List<int> GetYearEntryScoreSchoolYearList(string pStudentID)
        {
            return this.GetYearEntryScoreInfo(pStudentID).Select(x => x.SchoolYear).Distinct().ToList();
        }

        public List<SHSchool.Data.SHSemesterEntryScoreRecord> GetSemesterEntryScoreInfo(string pStudentID)
        {
            if (_SHSemesterEntryScore.ContainsKey(pStudentID))
                return _SHSemesterEntryScore[pStudentID];
            else
                return new List<SHSemesterEntryScoreRecord>();
        }

        public List<SHSchool.Data.SHSchoolYearEntryScoreRecord> GetYearEntryScoreInfo(string pStudentID)
        {
            if (_SHYearEntryScore.ContainsKey(pStudentID))
                return _SHYearEntryScore[pStudentID];
            else
                return new List<SHSchoolYearEntryScoreRecord>();
        }

        public decimal GetYearEntryScore(string pStudentID, int pSchoolYear, SHYearSubjectScoreInfo.Entry pEntry)
        {
            decimal score = 0;

            if (_SHYearEntryScore.ContainsKey(pStudentID))
            {
                var _SHSchoolYearEntryScore = _SHYearEntryScore[pStudentID].Where(x => x.SchoolYear == pSchoolYear && x.RefStudentID == pStudentID && x.Scores.ContainsKey(pEntry.ToString()));

                foreach (SHSchool.Data.SHSchoolYearEntryScoreRecord sr in _SHSchoolYearEntryScore)
                    score = sr.Scores[pEntry.ToString()];
            }
            return score;
        }

        public List<KeyValuePair<int, int>> GetMoralScoreList(string pStudentID)
        {
            List<SHSchool.Data.SHMoralScoreRecord> _SHMoralScoreRecords = this.GetMoralScore(pStudentID).OrderByDescending(x => x.SchoolYear).Select(x => x).ToList();

            List<KeyValuePair<int, int>> _MoralScoreList = new List<KeyValuePair<int, int>>();
            foreach (SHSchool.Data.SHMoralScoreRecord _SHMoralScoreRecord in _SHMoralScoreRecords)
            {
                if (_MoralScoreList.Where(x => (x.Key == _SHMoralScoreRecord.SchoolYear && x.Value == _SHMoralScoreRecord.Semester)).Count() == 0)
                    _MoralScoreList.Add(new KeyValuePair<int, int>(_SHMoralScoreRecord.SchoolYear, _SHMoralScoreRecord.Semester));
            }

            return _MoralScoreList;
        }

        public List<SHSchool.Data.SHMoralScoreRecord> GetMoralScore(string pStudentID)
        {
            if (_SHMoralScoreRecord.ContainsKey(pStudentID))
                return _SHMoralScoreRecord[pStudentID];
            else
                return null;
        }

        public SHSchool.Data.SHGradScoreRecord GetGradScore(string pStudentID)
        {
            if (_SHGradScoreRecord.ContainsKey(pStudentID))
                return _SHGradScoreRecord[pStudentID];
            else
                return null;
        }

        public MeritDeMeritInfo GetMeritDeMeritInfo(string pStudentID)
        {
            return _MeritDeMeritInfo;
        }

        public List<SHSchool.Data.SHUpdateRecordRecord> GetUpdateRecord(string pStudentID)
        {
            if (_SHUpdateRecord.ContainsKey(pStudentID))
                return _SHUpdateRecord[pStudentID];
            else
                return null;
        }

        public List<SHSubjectSemesterScoreInfo> GetPersonalSemesterSubjectScoreInfo(string pStudentID, int pSchoolYear1, int pSemester1, SHSemesterSubjectScoreInfo.SubjectSemesterScoreType pSubjectSemesterScoreType, int pSchoolYear2, int pSemester2)
        {
            return _SHSemesterSubjectScoreInfo.GetPersonalSemesterSubjectScoreInfo(pStudentID, pSchoolYear1, pSemester1, pSubjectSemesterScoreType, pSchoolYear2, pSemester2);
        }

        public List<SHSubjectSemesterScoreInfo> GetPersonalSemesterSubjectScoreInfo(string pStudentID, int pSchoolYear, int pSemester)
        {
            return _SHSemesterSubjectScoreInfo.GetPersonalSemesterSubjectScoreInfo(pStudentID, pSchoolYear, pSemester);
        }

        public List<SHSubjectSemesterScoreInfo> GetPersonalAllSemesterSubjectScoreInfo(string pStudentID)
        {
            return _SHSemesterSubjectScoreInfo.GetPersonalAllSemesterSubjectScoreInfo(pStudentID);
        }

        public List<SHSubjectYearScoreInfo> GetPersonalYearSubjectScoreInfo(Dictionary<int, List<SHSubjectSemesterScoreInfo>> dicSHSubjectSemesterScoreInfos)
        {
            return _SHYearSubjectScoreInfo.GetPersonalYearSubjectScoreInfo(dicSHSubjectSemesterScoreInfos);
        }

        public List<SHSubjectYearScoreInfo> GetPersonalYearSubjectScoreInfo(string pStudentID, int pSchoolYear)
        {
            return _SHYearSubjectScoreInfo.GetPersonalYearSubjectScoreInfo(pStudentID, pSchoolYear);
        }

        public Dictionary<int, List<KeyValuePair<int, int>>> GetPersonalSemesterSubjectScoreHistoryInfo(string pStudentID)
        {
            return _SHSemesterSubjectScoreInfo.GetPersonalSemesterSubjectScoreHistoryInfo(pStudentID);
        }

        public List<int> GetPersonalSemesterReReadSchoolYear(string pStudentID)
        {
            return _SHSemesterSubjectScoreInfo.GetPersonalReReadSchoolYear(pStudentID);
        }

        public List<int> GetPersonalYearReReadSchoolYear(string pStudentID)
        {
            return _SHYearSubjectScoreInfo.GetPersonalReReadSchoolYear(pStudentID);
        }

        public Dictionary<int, int> GetPersonalYearSubjectScoreHistoryInfo(string pStudentID)
        {
            return _SHYearSubjectScoreInfo.GetPersonalYearSubjectScoreHistoryInfo(pStudentID);
        }

        public bool NotExceedThreeMajorDemerits(string pStudentID)
        {
            return _MeritDeMeritInfo.NotExceedThreeMajorDemerits(pStudentID);
        }

        public List<int> GetPersonalReReadSchoolYear(string pStudentID)
        {
            return _SHYearSubjectScoreInfo.GetPersonalReReadSchoolYear(pStudentID);
        }

        public List<SHSubjectSemesterScoreInfo> SortSHSubjectSemesterScore(List<SHSubjectSemesterScoreInfo> firstSemesterSubjectScoreInfo, List<SHSubjectSemesterScoreInfo> secondSemesterSubjectScoreInfo)
        {
            return _SHSemesterSubjectScoreInfo.SortSHSubjectSemesterScore(firstSemesterSubjectScoreInfo, secondSemesterSubjectScoreInfo);
        }

        #region IDisposable 成員

        public void Dispose()
        {
            _Students.GetEnumerator().Dispose();
            _StudentRecord.GetEnumerator().Dispose();
            _SHYearSubjectScoreInfo.Dispose();
            _SHSemesterSubjectScoreInfo.Dispose();

            _Class.GetEnumerator().Dispose();
            _Department.GetEnumerator().Dispose();
            _StudentTag.GetEnumerator().Dispose();
            _GraduatePhoto.GetEnumerator().Dispose();
            _FreshmanPhoto.GetEnumerator().Dispose();
            _SHParentRecord.GetEnumerator().Dispose();
            _SHPhoneRecord.GetEnumerator().Dispose();
            _SHAddressRecord.GetEnumerator().Dispose();
            _SHSemesterEntryScore.GetEnumerator().Dispose();
            _SHYearEntryScore.GetEnumerator().Dispose();
            _SHMoralScoreRecord.GetEnumerator().Dispose();
            _SHGradScoreRecord.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _Students.GetEnumerator().Dispose();
            _MeritDeMeritInfo.Dispose();
            _SHUpdateRecord.GetEnumerator().Dispose();
            _PersonalSHScoreCalcRuleRecordInfo.GetEnumerator().Dispose();
            _PersonalStatusMappingInfo.GetEnumerator().Dispose();
            _CoreCourseTable.GetEnumerator().Dispose();
            _PersonalSHLeaveInfo.GetEnumerator().Dispose();

            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
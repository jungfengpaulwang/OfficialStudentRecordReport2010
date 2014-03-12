using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml.Linq;
using K12.Data;
using ReportHelper;
using SHSchool.Data;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml;
//using System.IO;

namespace OfficialStudentRecordReport2010
{
    public class ReportDataFormatTransfer : IDisposable
    {
        //MemoryStream selectedTemplate;  //  學籍表樣版
        int templateNumber;             //  樣版代碼：1，2-->高中。3，4-->高職。5，6-->進校
        string text1;                   //  學籍表字
        string text2;                   //  學籍表號
        string text3;                   //  學籍表輸出路徑，暫時不用
        int custodian;                  //  監護人代碼：0-->監護人。1-->父親。2-->母親
        int address;                    //  地址代碼：0-->戶籍。1-->聯絡。
        int phone;                      //  電話代碼：0-->戶籍。1-->聯絡。
        string coreSubjectSign;         //  核心科目識別符號
        string coreCourseSign;          //  核心課程識別符號
        string resitSign;               //  補考成績識別符號
        string retakeSign;              //  重修成績識別符號
        string failedSign;              //  不及格成績識別符號
        string schoolYearAdjustSign;    //  學年調整成績識別符號
        string manualAdjustSign;        //  手動調整成績識別符號
        DataPool dataPool;              //  所有學生的資料
        int reportType;              //  學籍表存檔選項：1-->每班一個檔案。2-->每個學生一個檔案。3-->每100個學生一個檔案。
        int textScoreOption;            //  文字評量選項：1-->導師評語。2-->文字評量。3-->2者皆要
        IEnumerable<SHStudentRecord> students;

        public List<DataSet> TransferFormat(object[] bgwObject)
        {
            //selectedTemplate = (MemoryStream)bgwObject[0];
            templateNumber = (int)bgwObject[1];
            text1 = (string)bgwObject[2];
            text2 = (string)bgwObject[3];
            text3 = (string)bgwObject[4];
            custodian = (int)bgwObject[5];
            address = (int)bgwObject[6];
            phone = (int)bgwObject[7];
            coreSubjectSign = (string)bgwObject[8];
            coreCourseSign = (string)bgwObject[9];
            resitSign = (string)bgwObject[10];
            retakeSign = (string)bgwObject[11];
            failedSign = (string)bgwObject[12];
            schoolYearAdjustSign = (string)bgwObject[13];
            manualAdjustSign = (string)bgwObject[14];
            dataPool = (DataPool)bgwObject[15];
            reportType = (int)bgwObject[16];
            textScoreOption = (int)bgwObject[17];
            students = (IEnumerable<SHStudentRecord>)bgwObject[18];

            return Transfer();
        }

        public void Dispose()
        {
            dataPool.Dispose();
            students = null;
        }

        //  科目學期成績資料
        private void ProduceSHSemesterScoreData(SHStudentRecord pStudent, DataSet schoolRollTable, int yIndex, int pSchoolYear, int pSemester, DataPool dataPool, List<SHSubjectSemesterScoreInfo> pSHSubjectSemesterScoreInfo)
        {
            if (pSHSubjectSemesterScoreInfo == null)
                return;

            string prefix = "Y" + yIndex + "S" + pSemester;

            //  後期中等教育核心科目標示
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => (x.CoreSubject ? coreSubjectSign : "")).ToDataTable(prefix + "CoreSubjectSign", "核心科目標示"));
            //  綜合高中學程核心課程標示
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => ((dataPool.IsCoreCourse(x.StudentID, x.SubjectName, x.Level)) ? coreCourseSign : "")).ToDataTable(prefix + "CoreCourseSign", "核心課程標示"));
            //  科目名稱
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.SubjectName).ToDataTable(prefix + "Subject", "科目名稱"));
            //  學分
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.Credit).ToDataTable(prefix + "Credit", "學分"));

            //_CoreCourseTable.ForEach(delegate(String name)
            //{
            //    if (!String.IsNullOrEmpty(name) && key.IndexOf(name) >= 0)
            //        youGotIt = true;
            //});

            //  必選修
            Dictionary<bool?, string> d = new Dictionary<bool?, string>();
            d.Add(true, "必");
            d.Add(false, "選");
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.Required).ConvertToString(d).ToDataTable(prefix + "Required", "必/選修"));
            //  校/部訂
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => (x.RequiredBy.Length == 0 ? "" : x.RequiredBy.Substring(0, 1))).ToDataTable(prefix + "RequiredBy", "校/部訂"));
            //  級別
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.Level).ToDataTable(prefix + "Level", "級別"));
            //  羅馬數字級別
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => Util.NumberToRomanChar((x.Level.HasValue ? x.Level.Value : 0))).ToDataTable(prefix + "RomanLevel", "羅馬數字級別"));
            // 原始成績
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.Score).ToDataTable(prefix + "Score", "原始成績"));
            // 補考成績
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.ReExamScore).ToDataTable(prefix + "ReExamScore", "補考成績"));
            // 手動調整成績
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.ManualScore).ToDataTable(prefix + "ManualScore", "手動調整成績"));
            // 學年擇優成績
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.SchoolYearAdjustScore).ToDataTable(prefix + "SchoolYearAdjustScore", "學年擇優成績"));
            // 重修成績
            schoolRollTable.Tables.Add(pSHSubjectSemesterScoreInfo.Select(x => x.ReCourseScore).ToDataTable(prefix + "ReCourseScore", "重修成績"));
            // 成績加符號標示
            List<string> ScoreWithSign = new List<string>();
            // 擇優成績
            List<string> ScoreBetter = new List<string>();
            // 未取得學分標示
            List<string> ScoreNoPass = new List<string>();
            foreach (SHSubjectSemesterScoreInfo ss in pSHSubjectSemesterScoreInfo)
            {
                //  不計學分不列舉，此邏輯為 ischool 之現況(收集資料時已篩汰)
                //  不評分要列舉
                //  所有分項要列舉(不判斷科目是否屬於「學業」之分項)

                //if (ss.NotIncludedInCredit.HasValue && ss.NotIncludedInCredit.Value)
                //    continue;

                decimal? scoreBetter = null;
                string scoreSign = string.Empty;
                string scoreWithSign = string.Empty;

                if (ss.Score.HasValue)
                    scoreBetter = ss.Score.Value;

                //  手動調整成績
                if (ss.ManualScore.HasValue)
                {
                    if (scoreBetter.HasValue)
                    {
                        if (ss.ManualScore.Value > scoreBetter.Value)
                        {
                            scoreBetter = ss.ManualScore.Value;
                            scoreSign = manualAdjustSign;
                        }
                    }
                    else
                    {
                        scoreBetter = ss.ManualScore.Value;
                        scoreSign = manualAdjustSign;
                    }
                }

                //  學年擇優成績
                if (ss.SchoolYearAdjustScore.HasValue)
                {
                    if (scoreBetter.HasValue)
                    {
                        if (ss.SchoolYearAdjustScore.Value > scoreBetter.Value)
                        {
                            scoreBetter = ss.SchoolYearAdjustScore.Value;
                            scoreSign = schoolYearAdjustSign;
                        }
                    }
                    else
                    {
                        scoreBetter = ss.SchoolYearAdjustScore.Value;
                        scoreSign = schoolYearAdjustSign;
                    }
                }

                //  補考成績
                if (ss.ReExamScore.HasValue)
                {
                    if (scoreBetter.HasValue)
                    {
                        if (ss.ReExamScore.Value > scoreBetter.Value)
                        {
                            scoreBetter = ss.ReExamScore.Value;
                            scoreSign = resitSign;
                        }
                    }
                    else
                    {
                        scoreBetter = ss.ReExamScore.Value;
                        scoreSign = resitSign;
                    }
                }

                scoreWithSign = scoreSign + (scoreBetter.HasValue ? scoreBetter.Value.ToString() : "");

                //  不及格標示
                if (ss.Pass.HasValue)
                {
                    if (!ss.Pass.Value)
                        ScoreNoPass.Add(failedSign);
                    else
                        ScoreNoPass.Add("");
                }
                else
                    ScoreNoPass.Add("");

                //  重修成績
                if (ss.ReCourseScore.HasValue)
                    scoreWithSign += retakeSign + ss.ReCourseScore.Value;

                ScoreWithSign.Add(scoreWithSign);
                ScoreBetter.Add(scoreBetter.HasValue ? scoreBetter.Value.ToString() : string.Empty);
            }
            // 成績加符號標示
            schoolRollTable.Tables.Add(ScoreWithSign.ToDataTable(prefix + "ScoreWithSign", "成績加符號標示"));
            // 擇優成績
            schoolRollTable.Tables.Add(ScoreBetter.ToDataTable(prefix + "ScoreBetter", "擇優成績"));
            // 未取得學分標示
            schoolRollTable.Tables.Add(ScoreNoPass.ToDataTable(prefix + "ScoreNoPass", "未取得學分標示"));
        }

        //  科目學年成績資料
        private void ProduceSHYearScoreData(SHStudentRecord pStudent, DataSet schoolRollTable, int yIndex, List<SHSubjectYearScoreInfo> pSHSubjectYearScoreInfo)
        {
            if (pSHSubjectYearScoreInfo == null)
                return;

            if (pSHSubjectYearScoreInfo.Count == 0)
                return;

            string prefix = "Y" + yIndex.ToString();
            //  課時
            schoolRollTable.Tables.Add(pSHSubjectYearScoreInfo.Select(x => x.Hour).ToDataTable(prefix + "Hour", "課時"));
            //  學分
            schoolRollTable.Tables.Add(pSHSubjectYearScoreInfo.Select(x => (x.Credit.HasValue ? x.Credit.Value.ToString() : "")).ToDataTable(prefix + "Credit", "學分"));
            //  原始成績
            schoolRollTable.Tables.Add(pSHSubjectYearScoreInfo.Select(x => (x.Score.HasValue ? x.Score.Value.ToString() : "")).ToDataTable(prefix + "ScoreWithSign", "原始成績"));
            //  學年實得學分
            schoolRollTable.Tables.Add(pSHSubjectYearScoreInfo.Sum(x => x.Credit).ToDataTable(prefix + "AcquiredCredit", "學年實得學分"));
        }

        /// <summary>
        /// 範例程式：收集學籍表相關資料。請注意：請儘量減少取得資料的 IO
        /// </summary>
        /// <param name="studentID">學生系統編號</param>
        /// <returns>個人學籍表相關資料</returns>
        private DataSet GetSchoolRollTable(SHStudentRecord student, DataPool dataPool)
        {
            //  此涵式的區域變數
            //  學年歷程
            Dictionary<int, int> _PersonalYearSubjectScoreHistoryInfo = dataPool.GetPersonalYearSubjectScoreHistoryInfo(student.ID);
            //  學期歷程
            Dictionary<int, List<KeyValuePair<int, int>>> _PersonalSemesterSubjectScoreHistoryInfo = dataPool.GetPersonalSemesterSubjectScoreHistoryInfo(student.ID);
            //  學生修習科目學期成績資料
            List<SHSubjectSemesterScoreInfo> _SHSemesterScore = dataPool.GetPersonalAllSemesterSubjectScoreInfo(student.ID);

            DataSet schoolRollTable = new DataSet("DataSection");

            #region 學籍表字號

            //  字
            schoolRollTable.Tables.Add(text1.ToDataTable("AuthorityCode", "證明書字"));

            //  號
            if (!string.IsNullOrEmpty(text2))
            {
                schoolRollTable.Tables.Add(text2.ToDataTable("SerialNumber", "證明書流水號"));

                text2 = ((Convert.ToInt32(text2) + 1)).ToString("d" + text2.Length);
            }

            #endregion

            # region 學校、學生、監護人基本資料

            // 校名
            schoolRollTable.Tables.Add(School.ChineseName.ToDataTable("SchoolName", "校名"));

            // 學校代碼
            schoolRollTable.Tables.Add(School.Code.ToDataTable("SchoolCode", "學校代碼"));

            // 學生基本資料：姓名
            schoolRollTable.Tables.Add(student.Name.ToDataTable("StudentName", "學生姓名"));

            // 學生基本資料：性別
            schoolRollTable.Tables.Add(student.Gender.ToDataTable("Gender", "性別"));

            // 學生基本資料：出生日期
            string birthday = student.Birthday.HasValue ? student.Birthday.Value.ToShortDateString() : string.Empty;
            if (!String.IsNullOrEmpty(birthday))
            {
                if (Convert.ToDateTime(birthday).Year.ToString().Length == 4)
                    birthday = (Convert.ToDateTime(birthday).Year - 1911).ToString() + "/" + Convert.ToDateTime(birthday).Month.ToString() + "/" + Convert.ToDateTime(birthday).Day.ToString();
            }
            schoolRollTable.Tables.Add(birthday.ToDataTable("Birthday", "出生日期"));
            // 學生基本資料：學號
            schoolRollTable.Tables.Add(student.StudentNumber.ToDataTable("StudentNumber", "學號"));

            // 學生基本資料：座號
            schoolRollTable.Tables.Add(student.SeatNo.ToString().ToDataTable("SeatNo", "座號"));

            // 學生基本資料：成績身份
            schoolRollTable.Tables.Add(dataPool.TransferStudentTagToIdentity(student.ID).ToDataTable("EvaluationIdentity", "成績身份"));

            string pDepartmentName = string.Empty;
            string pDepartmentCode = string.Empty;
            if (dataPool.GetDepartment(student.ID) != null)
            {
                pDepartmentName = dataPool.GetDepartment(student.ID).Name;
                pDepartmentCode = dataPool.GetDepartment(student.ID).Code;
            }
            else
            {
                SHLeaveInfoRecord sr = dataPool.GetPersonalSHLeaveInfo(student.ID);
                if (sr != null)
                    pDepartmentName = (sr.DepartmentName == null ? string.Empty : sr.DepartmentName);
            }

            // 學生基本資料：科別
            schoolRollTable.Tables.Add(pDepartmentName.ToDataTable("DepartmentName", "科別"));

            // 學生基本資料：科別代碼
            schoolRollTable.Tables.Add(pDepartmentCode.ToDataTable("DepartmentCode", "科別代碼"));

            // 學生基本資料：身份證字號
            schoolRollTable.Tables.Add(student.IDNumber.ToDataTable("IDNumber", "身份證字號"));

            // 學生基本資料：班級          
            SHClassRecord clazz = dataPool.GetClass(student.ID);
            schoolRollTable.Tables.Add((clazz == null ? "" : clazz.Name).ToDataTable("ClassName", "班級"));

            // 學生基本資料：畢業照 => ischool 的格式為 Base64String，須轉為 ReportHelper 的格式：Byte[]
            // 如果沒有畢業照片，則用入學照片
            string graduatePhoto = dataPool.GetGraduatePhoto(student.ID);
            string freshmanPhoto = dataPool.GetFreshmanPhoto(student.ID);

            if (string.IsNullOrEmpty(graduatePhoto))
                graduatePhoto = freshmanPhoto;

            schoolRollTable.Tables.Add(graduatePhoto.FromBase64StringToByte().ToDataTable("GraduatePhoto", "畢業照", Type.GetType("System.Byte[]")));

            // 監護人判讀
            SHParentRecord parent = dataPool.GetParent(student.ID);
            string strCustodian = string.Empty;
            string strCustodianJob = string.Empty;
            string strCustodianRelationship = string.Empty;
            if (parent != null)
            {
                if (custodian == 0)
                {
                    strCustodian = parent.CustodianName;
                    strCustodianJob = parent.CustodianJob;
                    strCustodianRelationship = parent.CustodianRelationship;
                }

                if (custodian == 1)
                {
                    strCustodian = parent.FatherName;
                    strCustodianJob = parent.FatherJob;
                    strCustodianRelationship = "父親";
                }

                if (custodian == 2)
                {
                    strCustodian = parent.MotherName;
                    strCustodianJob = parent.MotherJob;
                    strCustodianRelationship = "母親";
                }
            }
            // 監護人基本資料：監護人姓名
            schoolRollTable.Tables.Add(strCustodian.ToDataTable("CustodianName", "監護人姓名"));

            // 監護人基本資料：職業
            schoolRollTable.Tables.Add(strCustodianJob.ToDataTable("CustodianJob", "監護人職業"));

            // 監護人基本資料：稱謂
            schoolRollTable.Tables.Add(strCustodianRelationship.ToDataTable("CustodianRelationship", "稱謂"));

            // 學生基本資料：戶籍電話
            SHPhoneRecord oPhone = dataPool.GetPhone(student.ID);
            string strPhone = string.Empty;
            if (oPhone != null)
            {
                if (phone == 0)
                    strPhone = oPhone.Permanent;

                if (phone == 1)
                    strPhone = oPhone.Contact;
            }
            schoolRollTable.Tables.Add(strPhone.ToDataTable("Phone", "電話"));

            // 學生基本資料：戶籍地址
            SHAddressRecord oAddress = dataPool.GetAddress(student.ID);
            string strAddress = string.Empty;
            if (oAddress != null)
            {
                if (address == 0)
                    strAddress = oAddress.PermanentAddress.Trim();

                if (address == 1)
                    strAddress = oAddress.MailingAddress.Trim();
            }
            schoolRollTable.Tables.Add(strAddress.ToDataTable("Address", "地址"));

            # endregion

            # region 學期(年)成績資料

            if (_PersonalSemesterSubjectScoreHistoryInfo != null && _PersonalSemesterSubjectScoreHistoryInfo.Count > 0)
            {
                int yIndex = 0;
                List<int> schoolYear = new List<int>();
                foreach (int gradeYear in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                {
                    yIndex = gradeYear;
                    //  復學生有可能在相同的成績年級中有多筆學年度資料
                    List<KeyValuePair<int, int>> allSchoolYearSemesterPair = _PersonalSemesterSubjectScoreHistoryInfo[yIndex];

                    if (allSchoolYearSemesterPair != null && allSchoolYearSemesterPair.Count>0)
                        schoolYear = allSchoolYearSemesterPair.Select(x => x.Key).Distinct().ToList();

                    string sYear = string.Empty;
                    foreach (int s in schoolYear.OrderBy(x=>x))
                    {
                        sYear += s.ToString() + "/";
                    }
                    sYear = sYear.Substring(0, sYear.Length - 1);
                    schoolRollTable.Tables.Add(sYear.ToDataTable("Y" + yIndex.ToString() + "SchoolYear", "學年度"));

                    Dictionary<int, List<SHSubjectSemesterScoreInfo>> dicSHSubjectSemesterScoreInfos = new Dictionary<int, List<SHSubjectSemesterScoreInfo>>();

                    foreach (KeyValuePair<int, int> kv in allSchoolYearSemesterPair)
                    {
                        if (!dicSHSubjectSemesterScoreInfos.ContainsKey(kv.Value))
                            dicSHSubjectSemesterScoreInfos.Add(kv.Value, new List<SHSubjectSemesterScoreInfo>());

                        dicSHSubjectSemesterScoreInfos[kv.Value] = dataPool.GetPersonalSemesterSubjectScoreInfo(student.ID, kv.Key, kv.Value);
                    }
                    //  學年學分制排序上下學期與學年成績
                    if (allSchoolYearSemesterPair.Count > 1 && templateNumber != 3 && templateNumber != 4)
                    {
                        dicSHSubjectSemesterScoreInfos[2] = dataPool.SortSHSubjectSemesterScore(dicSHSubjectSemesterScoreInfos[1], dicSHSubjectSemesterScoreInfos[2]);
                    }
                    foreach (KeyValuePair<int, int> kv in allSchoolYearSemesterPair)
                    {
                        ProduceSHSemesterScoreData(student, schoolRollTable, yIndex, kv.Key, kv.Value, dataPool, dicSHSubjectSemesterScoreInfos[kv.Value]);
                    }

                    //  學年成績
                    List<SHSubjectYearScoreInfo> pSHSubjectYearScoreInfo = dataPool.GetPersonalYearSubjectScoreInfo(dicSHSubjectSemesterScoreInfos);
                    ProduceSHYearScoreData(student, schoolRollTable, yIndex, pSHSubjectYearScoreInfo);   
                }
                foreach (int gradeYear in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                {
                    yIndex = gradeYear;
                    string prefix = "Y" + yIndex + "S";
                    List<string> mergeSubjects = new List<string>();

                    if (schoolRollTable.Tables.Contains(prefix + "1Subject") || schoolRollTable.Tables.Contains(prefix + "2Subject"))
                    {
                        string mergeSubject = string.Empty;

                        int j = 0;
                        if (schoolRollTable.Tables.Contains(prefix + "1Subject"))
                        {
                            foreach (DataRow dr in schoolRollTable.Tables[prefix + "1Subject"].Rows)
                            {
                                string s1Level = string.Empty;
                                string s2Level = string.Empty;
                                if (schoolRollTable.Tables.Contains(prefix + "2Level"))
                                    s2Level = (schoolRollTable.Tables[prefix + "2Level"].Rows.Count < (j + 1)) ? "" : schoolRollTable.Tables[prefix + "2Level"].Rows[j]["級別"].ToString();
                                if (schoolRollTable.Tables.Contains(prefix + "1Level"))
                                    s1Level = (schoolRollTable.Tables[prefix + "1Level"].Rows.Count < (j + 1)) ? "" : schoolRollTable.Tables[prefix + "1Level"].Rows[j]["級別"].ToString();

                                string subjectName = dr["科目名稱"].ToString();

                                if (string.IsNullOrWhiteSpace(subjectName))
                                {
                                    int rowIndex = dr.Table.Rows.IndexOf(dr);
                                    if (rowIndex > 0)
                                        subjectName = schoolRollTable.Tables[prefix + "2Subject"].Rows[rowIndex]["科目名稱"].ToString();
                                }

                                mergeSubject = subjectName + " " + (string.IsNullOrWhiteSpace(Util.NumberToRomanChar(s1Level)) ? Util.NumberToRomanChar(s1Level) : Util.NumberToRomanChar(s1Level) + ",") + Util.NumberToRomanChar(s2Level);

                                if (mergeSubject.EndsWith(","))
                                    mergeSubject = mergeSubject.Substring(0, mergeSubject.Length - 1);

                                mergeSubjects.Add(mergeSubject);
                                j++;
                            }
                        }
                        else
                        {
                            if (schoolRollTable.Tables.Contains(prefix + "2Subject"))
                            {
                                for (int k = j; k < schoolRollTable.Tables[prefix + "2Subject"].Rows.Count; k++)
                                {
                                    mergeSubject = schoolRollTable.Tables[prefix + "2Subject"].Rows[k]["科目名稱"].ToString() + " " + Util.NumberToRomanChar(schoolRollTable.Tables[prefix + "2Level"].Rows[k]["級別"].ToString());

                                    mergeSubjects.Add(mergeSubject);
                                }
                            }
                        }
                        if (schoolRollTable.Tables.Contains(prefix + "1Subject") && schoolRollTable.Tables.Contains(prefix + "2Subject"))
                        {
                            if (schoolRollTable.Tables[prefix + "2Subject"].Rows.Count > schoolRollTable.Tables[prefix + "1Subject"].Rows.Count)
                            {
                                for (int k = j; k < schoolRollTable.Tables[prefix + "2Subject"].Rows.Count; k++)
                                {
                                    mergeSubject = schoolRollTable.Tables[prefix + "2Subject"].Rows[k]["科目名稱"].ToString() + " " + Util.NumberToRomanChar(schoolRollTable.Tables[prefix + "2Level"].Rows[k]["級別"].ToString());

                                    mergeSubjects.Add(mergeSubject);
                                }
                            }
                        }
                    }
                    schoolRollTable.Tables.Add(mergeSubjects.ToDataTable("Y" + yIndex.ToString() + "MergedSubject", "合併級別科目名稱"));
                }
            }
            #endregion

            #region 學期學分累計資料

            if (_PersonalSemesterSubjectScoreHistoryInfo != null && _PersonalSemesterSubjectScoreHistoryInfo.Count > 0)
            {
                decimal? AccumulatedTotalCredit = 0;
                decimal? AccumulatedAcquiredCredit = 0;
                foreach (int gradeYear in _PersonalSemesterSubjectScoreHistoryInfo.Keys.OrderBy(x=>x))
                {
                    decimal? TotalCredit = 0;
                    decimal? AcquiredCredit = 0;

                    foreach (KeyValuePair<int, int> sh in _PersonalSemesterSubjectScoreHistoryInfo[gradeYear])
                    {
                        TotalCredit += _SHSemesterScore.Where(x => (x.SchoolYear == sh.Key && x.Semester == sh.Value)).Where(x => x.Credit != null).Sum(x => x.Credit);
                        AcquiredCredit += _SHSemesterScore.Where(x => (x.SchoolYear == sh.Key && x.Semester == sh.Value)).Where(x => x.Pass.HasValue && x.Pass.Value).Sum(x => x.Credit);

                        AccumulatedTotalCredit += _SHSemesterScore.Where(x => (x.SchoolYear == sh.Key && x.Semester == sh.Value)).Where(x => x.Credit != null).Sum(x => x.Credit);
                        AccumulatedAcquiredCredit += _SHSemesterScore.Where(x => (x.SchoolYear == sh.Key && x.Semester == sh.Value)).Where(x => x.Pass.HasValue && x.Pass.Value).Sum(x => x.Credit);

                        string prefix = "Y" + gradeYear.ToString() + "S" + sh.Value;

                        schoolRollTable.Tables.Add(_SHSemesterScore.Where(x => (x.SchoolYear == sh.Key && x.Semester == sh.Value)).Where(x => x.Credit != null).Sum(x => x.Credit).ToDataTable(prefix + "TotalCredit", "學期應得學分"));

                        schoolRollTable.Tables.Add(_SHSemesterScore.Where(x => (x.SchoolYear == sh.Key && x.Semester == sh.Value)).Where(x => x.Pass.HasValue && x.Pass.Value).Sum(x => x.Credit).ToDataTable(prefix + "AcquiredCredit", "學期實得學分"));

                        schoolRollTable.Tables.Add(_SHSemesterScore.Where(x => (x.SchoolYear < sh.Key) || ((x.SchoolYear == sh.Key) && (x.Semester <= sh.Value))).Sum(x => x.Credit).ToDataTable(prefix + "AccumulatedTotalCredit", "學期累計應得學分"));

                        schoolRollTable.Tables.Add(_SHSemesterScore.Where(x => (x.SchoolYear < sh.Key) || ((x.SchoolYear == sh.Key) && (x.Semester <= sh.Value))).Where(x => x.Pass.HasValue && x.Pass.Value).Sum(x => x.Credit).ToDataTable(prefix + "AccumulatedAcquiredCredit", "學期累計實得學分"));

                        if (schoolRollTable.Tables.Contains("Y" + gradeYear.ToString() + "TotalCredit"))
                            schoolRollTable.Tables.Remove("Y" + gradeYear.ToString() + "TotalCredit");

                        schoolRollTable.Tables.Add(TotalCredit.ToDataTable("Y" + gradeYear.ToString() + "TotalCredit", "學年應得學分"));

                        if (schoolRollTable.Tables.Contains("Y" + gradeYear.ToString() + "AcquiredCredit"))
                            schoolRollTable.Tables.Remove("Y" + gradeYear.ToString() + "AcquiredCredit");
                            
                        schoolRollTable.Tables.Add(AcquiredCredit.ToDataTable("Y" + gradeYear.ToString() + "AcquiredCredit", "學年實得學分"));

                        if (schoolRollTable.Tables.Contains("Y" + gradeYear.ToString() + "AccumulatedTotalCredit"))
                            schoolRollTable.Tables.Remove("Y" + gradeYear.ToString() + "AccumulatedTotalCredit");

                        schoolRollTable.Tables.Add(AccumulatedTotalCredit.ToDataTable("Y" + gradeYear.ToString() + "AccumulatedTotalCredit", "學年累計應得學分"));

                        if (schoolRollTable.Tables.Contains("Y" + gradeYear.ToString() + "AccumulatedAcquiredCredit"))
                            schoolRollTable.Tables.Remove("Y" + gradeYear.ToString() + "AccumulatedAcquiredCredit");

                        schoolRollTable.Tables.Add(AccumulatedAcquiredCredit.ToDataTable("Y" + gradeYear.ToString() + "AccumulatedAcquiredCredit", "學年累計實得學分"));
                    }
                }
            }
            #endregion

            #region 學期學業成績資料

            List<SHSemesterEntryScoreRecord> _SHSemesterEntryScoreRecords = dataPool.GetSemesterEntryScoreInfo(student.ID);
            Dictionary<int, List<KeyValuePair<int, int>>> _SHSemesterEntryScoreHistory = dataPool.GetSemesterEntryScoreHistory(student.ID);

            if (_SHSemesterEntryScoreRecords.Count > 0 && _SHSemesterEntryScoreHistory.Count > 0)
            {
                string prefix = string.Empty;
                int yIndex = 0;
                int sIndex = 0;

                foreach (SHSemesterEntryScoreRecord sr in _SHSemesterEntryScoreRecords)
                {
                    yIndex = sr.GradeYear;
                    if (!_SHSemesterEntryScoreHistory.ContainsKey(yIndex))
                        continue;

                    if (_SHSemesterEntryScoreHistory[yIndex].Where(x => (x.Key == sr.SchoolYear && x.Value == sr.Semester)).Count() == 0)
                        continue;

                    sIndex = sr.Semester;

                    prefix = "Y" + yIndex + "S" + sIndex;

                    // 學業
                    if (sr.Scores.ContainsKey("學業"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["學業"].ToDataTable(prefix + "AcademicScore", "學業成績"));

                        SHRankingInfo rank = sr.ClassRating.Find(x => x.Name.Equals("學業"));
                        schoolRollTable.Tables.Add((rank == null ? "" : rank.Ranking.ToString()).ToDataTable(prefix + "AcademicClassRank", "學業成績班級排名"));

                        rank = sr.DeptRating.Find(x => x.Name.Equals("學業"));
                        schoolRollTable.Tables.Add((rank == null ? "" : rank.Ranking.ToString()).ToDataTable(prefix + "AcademicDepartmentRank", "學業成績科排名"));

                        rank = sr.YearRating.Find(x => x.Name.Equals("學業"));
                        schoolRollTable.Tables.Add((rank == null ? "" : rank.Ranking.ToString()).ToDataTable(prefix + "AcademicGradeRank", "學業成績年級排名"));
                    }
                    //  體育
                    if (sr.Scores.ContainsKey("體育"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["體育"].ToDataTable(prefix + "PhysicalScore", "體育成績"));
                    }
                    //  健康與護理
                    if (sr.Scores.ContainsKey("健康與護理"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["健康與護理"].ToDataTable(prefix + "HealthScore", "健康與護理成績"));
                    }
                    //  國防通識
                    if (sr.Scores.ContainsKey("國防通識"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["國防通識"].ToDataTable(prefix + "NationalDefenseScore", "國防通識成績"));
                    }
                    //  實習科目
                    if (sr.Scores.ContainsKey("實習科目"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["實習科目"].ToDataTable(prefix + "PracticeScore", "實習科目成績"));
                    }
                    //  專業科目
                    if (sr.Scores.ContainsKey("專業科目"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["專業科目"].ToDataTable(prefix + "SpecializationScore", "專業科目成績"));
                    }
                }
            }
            #endregion

            #region 學年學業成績資料

            List<SHSchoolYearEntryScoreRecord> _SHYearEntryScoreRecords = dataPool.GetYearEntryScoreInfo(student.ID);
            Dictionary<int, int> _SHYearEntryScoreHistory = dataPool.GetYearEntryScoreHistory(student.ID);
            if (_SHYearEntryScoreRecords.Count > 0 && _SHYearEntryScoreHistory.Count > 0)
            {
                string prefix = string.Empty;
                int yIndex = 0;

                foreach (SHSchoolYearEntryScoreRecord sr in _SHYearEntryScoreRecords)
                {
                    if (!_SHYearEntryScoreHistory.ContainsValue(sr.SchoolYear))
                        continue;

                    yIndex = sr.GradeYear;

                    prefix = "Y" + yIndex.ToString();

                    // 學業
                    if (sr.Scores.ContainsKey("學業"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["學業"].ToDataTable(prefix + "AcademicScore", "學業成績"));

                        SHRankingInfo rank = sr.ClassRating.Find(x => x.Name.Equals("學業"));
                        schoolRollTable.Tables.Add((rank == null ? "" : rank.Ranking.ToString()).ToDataTable(prefix + "AcademicClassRank", "學業成績班級排名"));

                        rank = sr.DeptRating.Find(x => x.Name.Equals("學業"));
                        schoolRollTable.Tables.Add((rank == null ? "" : rank.Ranking.ToString()).ToDataTable(prefix + "AcademicDepartmentRank", "學業成績科排名"));

                        rank = sr.YearRating.Find(x => x.Name.Equals("學業"));
                        schoolRollTable.Tables.Add((rank == null ? "" : rank.Ranking.ToString()).ToDataTable(prefix + "AcademicGradeRank", "學業成績年級排名"));
                    }
                    //  體育
                    if (sr.Scores.ContainsKey("體育"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["體育"].ToDataTable(prefix + "PhysicalScore", "體育成績"));
                    }
                    //  健康與護理
                    if (sr.Scores.ContainsKey("健康與護理"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["健康與護理"].ToDataTable(prefix + "HealthScore", "健康與護理成績"));
                    }
                    //  國防通識
                    if (sr.Scores.ContainsKey("國防通識"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["國防通識"].ToDataTable(prefix + "NationalDefenseScore", "國防通識成績"));
                    }
                    //  實習科目
                    if (sr.Scores.ContainsKey("實習科目"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["實習科目"].ToDataTable(prefix + "PracticeScore", "實習科目成績"));
                    }
                    //  專業科目
                    if (sr.Scores.ContainsKey("專業科目"))
                    {
                        schoolRollTable.Tables.Add(sr.Scores["專業科目"].ToDataTable(prefix + "SpecializationScore", "專業科目成績"));
                    }
                }
            }

            #endregion

            #region 德行評量(導師評語)資料

            List<SHMoralScoreRecord> _SHMoralScoreRecords = dataPool.GetMoralScore(student.ID);
            if (_SHMoralScoreRecords != null && _SHMoralScoreRecords.Count > 0)
                _SHMoralScoreRecords = _SHMoralScoreRecords.OrderBy(x => x.SchoolYear).ThenBy(x => x.Semester).ToList();

            //  德行評量沒有成績年級，採用「科目學期成績」的學期歷程
            if (_PersonalSemesterSubjectScoreHistoryInfo != null && _PersonalSemesterSubjectScoreHistoryInfo.Count > 0 && _SHMoralScoreRecords != null && _SHMoralScoreRecords.Count > 0)
            {
                foreach (SHMoralScoreRecord sr in _SHMoralScoreRecords)
                {
                    int yIndex = 0;
                    foreach (int gradeYear in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                    {
                        List<KeyValuePair<int, int>> kvs = _PersonalSemesterSubjectScoreHistoryInfo[gradeYear];
                        if (kvs.Count > 0 && kvs.Where(x => (x.Key == sr.SchoolYear && x.Value == sr.Semester)).Count() > 0)
                        {
                            yIndex = gradeYear;
                            break;
                        }
                    }
                    if (yIndex == 0)
                        continue;

                    string moralScore = string.Empty;
                    int sIndex = sr.Semester;

                    string prefix = "Y" + yIndex + "S" + sIndex;

                    //  德行評量選項：1-->導師評語。2-->文字評量。3-->2者皆要
                    //  導師評語
                    if (textScoreOption == 1 || textScoreOption == 3)
                        if (sr.Comment.Trim() != "")
                            moralScore = sr.Comment + ";";

                    if (textScoreOption == 2 || textScoreOption == 3)
                    {
                        //  文字評量
                        XDocument doc = XDocument.Parse("<root>" + sr.TextScore.InnerXml + "</root>");
                        foreach (XElement e in doc.Document.Descendants("Morality"))
                        {
                            if (e.Value.ToString().Trim() == "")
                                continue;

                            moralScore += e.Attribute("Face").Value + ":" + e.Value + ";";
                        }
                    }

                    if (moralScore.EndsWith(";"))
                        moralScore = moralScore.Substring(0, moralScore.Length - 1);

                    if (schoolRollTable.Tables.Contains(prefix + "MoralScore"))
                        schoolRollTable.Tables.Remove(prefix + "MoralScore");

                    schoolRollTable.Tables.Add(moralScore.ToDataTable(prefix + "MoralScore", "德行評量"));
                }
            }
            
            #endregion

            #region 畢業成績資訊(含學業成績、實習成績、畢業規定總學分，其餘累計學分於次項統計)
            SHGradScoreRecord srs = dataPool.GetGradScore(student.ID);
            if (srs != null)
            {
                if (srs.Entries.Count > 0)
                {
                    //  學業成績
                    if (srs.Entries.ContainsKey("學業"))
                        schoolRollTable.Tables.Add(srs.Entries["學業"].Score.ToString().ToDataTable("GraduationAcademicScore", "學業成績"));
                    //  實習成績
                    if (srs.Entries.ContainsKey("實習科目"))
                        schoolRollTable.Tables.Add(srs.Entries["實習科目"].Score.ToString().ToDataTable("GraduationPracticeScore", "實習成績"));
                }
            }
            schoolRollTable.Tables.Add(dataPool.GetGraduationDeservedCredit(student.ID).ToDataTable("GraduationDeservedCredit", "畢業規定總學分"));
            #endregion

            #region 畢業成績資訊(累計學分)

            if (_SHSemesterScore != null && _SHSemesterScore.Count > 0 && _PersonalSemesterSubjectScoreHistoryInfo != null && _PersonalSemesterSubjectScoreHistoryInfo.Count > 0)
            {
                List<KeyValuePair<int, int>> kvs = new List<KeyValuePair<int, int>>();
                foreach (int index in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                    _PersonalSemesterSubjectScoreHistoryInfo[index].ForEach(x => kvs.Add(x));

                decimal? graduationDecidesRequiredAccumulatedCredit = 0;
                decimal? graduationDecidesRequiredAcquiredCredit = 0;
                decimal? graduationSchoolRequiredAcquiredCredit = 0;
                decimal? graduationOptionalAcquiredCredit = 0;
                decimal? graduationSchoolRequiredAccumulatedCredit = 0;
                decimal? graduationOptionalAccumulatedCredit = 0;
                decimal? graduationAcquiredCredit = 0;

                foreach(SHSubjectSemesterScoreInfo x in _SHSemesterScore)
                {
                    if (kvs.Where(y => (y.Key == x.SchoolYear && y.Value == x.Semester)).Count() > 0)
                    {
                        if (!String.IsNullOrEmpty(x.RequiredBy) && x.RequiredBy.Substring(0, 1) == "部" && x.Required.HasValue && x.Required == true)
                            graduationDecidesRequiredAccumulatedCredit += (x.Credit.HasValue ? x.Credit.Value : 0);

                        if (!String.IsNullOrEmpty(x.RequiredBy) && x.RequiredBy.Substring(0, 1) == "部" && x.Required.HasValue && x.Required == true && ((x.Pass.HasValue ? x.Pass.Value : false) == true))
                            graduationDecidesRequiredAcquiredCredit += (x.Credit.HasValue ? x.Credit.Value : 0);

                        if (!String.IsNullOrEmpty(x.RequiredBy) && x.RequiredBy.Substring(0, 1) == "校" && x.Required.HasValue && x.Required == true && ((x.Pass.HasValue ? x.Pass.Value : false) == true))
                            graduationSchoolRequiredAcquiredCredit += (x.Credit.HasValue ? x.Credit.Value : 0);

                        if (!String.IsNullOrEmpty(x.RequiredBy) && x.RequiredBy.Substring(0, 1) == "校" && x.Required.HasValue && x.Required == false && ((x.Pass.HasValue ? x.Pass.Value : false) == true))
                            graduationOptionalAcquiredCredit += (x.Credit.HasValue ? x.Credit.Value : 0);

                        if (!String.IsNullOrEmpty(x.RequiredBy) && x.RequiredBy.Substring(0, 1) == "校" && x.Required.HasValue && x.Required == true)
                            graduationSchoolRequiredAccumulatedCredit += (x.Credit.HasValue ? x.Credit.Value : 0);

                        if (!String.IsNullOrEmpty(x.RequiredBy) && x.RequiredBy.Substring(0, 1) == "校" && x.Required.HasValue && x.Required == false)
                            graduationOptionalAccumulatedCredit += (x.Credit.HasValue ? x.Credit.Value : 0);

                        if (((x.Pass.HasValue ? x.Pass.Value : false) == true))
                            graduationAcquiredCredit += (x.Credit.HasValue ? x.Credit.Value : 0);

                        //schoolRollTable.Tables.Add(_SHSemesterScore.Where(x => !String.IsNullOrEmpty(x.RequiredBy)).Where(x => x.RequiredBy.Substring(0, 1) == "校").Where(x => x.Required.HasValue).Where(x => x.Required == true).Sum(x => x.Credit).ToDataTable("GraduationSchoolRequiredAcquiredCredit", "校訂必修實得學分"));
                        //schoolRollTable.Tables.Add(_SHSemesterScore.Where(x => !String.IsNullOrEmpty(x.RequiredBy)).Where(x => x.RequiredBy.Substring(0, 1) == "校").Where(x => x.Required.HasValue).Where(x => x.Required == false).Sum(x => x.Credit).ToDataTable("GraduationOptionalAcquiredCredit", "校訂選修實得學分"));
                        //schoolRollTable.Tables.Add(_SHSemesterScore.Where(x => x.Pass.HasValue).Where(x => (bool)x.Pass).Sum(x => x.Credit).ToDataTable("GraduationAcquiredCredit", "畢業獲得總學分"));
                        //_SHSemesterScore.Where(x => !String.IsNullOrEmpty(x.RequiredBy)).Where(x => x.RequiredBy.Substring(0, 1) == "部").Where(x => x.Required.HasValue).Where(x => x.Required == true).Sum(x => x.Credit);
                        //decimal? graduationDecidesRequiredAcquiredCredit = _SHSemesterScore.Where(x => !String.IsNullOrEmpty(x.RequiredBy)).Where(x => x.RequiredBy.Substring(0, 1) == "部").Where(x => x.Required.HasValue).Where(x => x.Required == true).Where(x => ((x.Pass.HasValue ? x.Pass.Value : false) == true)).Sum(x => x.Credit);
                    }
                }

                if (graduationDecidesRequiredAccumulatedCredit.HasValue)
                    schoolRollTable.Tables.Add(graduationDecidesRequiredAccumulatedCredit.ToDataTable("GraduationDecidesRequiredAccumulatedCredit", "部定必修應得學分"));

                if (graduationDecidesRequiredAcquiredCredit.HasValue)
                    schoolRollTable.Tables.Add(graduationDecidesRequiredAcquiredCredit.ToDataTable("GraduationDecidesRequiredAcquiredCredit", "部定必修實得學分"));

                if (graduationDecidesRequiredAccumulatedCredit.HasValue)
                {
                    if (graduationDecidesRequiredAcquiredCredit.HasValue && graduationDecidesRequiredAccumulatedCredit.Value != 0M)
                        schoolRollTable.Tables.Add((Math.Round(((graduationDecidesRequiredAcquiredCredit.Value * 100) / graduationDecidesRequiredAccumulatedCredit.Value), 1).ToString() + "%").ToDataTable("GraduationDecidesRequiredCreditPassingRate", "部定必修及格率"));
                    //schoolRollTable.Tables.Add((((graduationDecidesRequiredAcquiredCredit.Value) / graduationDecidesRequiredAccumulatedCredit.Value)).ToDataTable("GraduationDecidesRequiredCreditPassingRate", "部定必修及格率"));
                    else
                        schoolRollTable.Tables.Add("".ToDataTable("GraduationDecidesRequiredCreditPassingRate", "部定必修及格率"));
                }

                if (graduationSchoolRequiredAcquiredCredit.HasValue)
                    schoolRollTable.Tables.Add(graduationSchoolRequiredAcquiredCredit.ToDataTable("GraduationSchoolRequiredAcquiredCredit", "校定必修實得學分"));

                if (graduationOptionalAcquiredCredit.HasValue)
                    schoolRollTable.Tables.Add(graduationOptionalAcquiredCredit.ToDataTable("GraduationOptionalAcquiredCredit", "校定選修實得學分"));

                if (graduationSchoolRequiredAccumulatedCredit.HasValue)
                    schoolRollTable.Tables.Add(graduationSchoolRequiredAccumulatedCredit.ToDataTable("GraduationSchoolRequiredAccumulatedCredit", "校定必修應得學分"));

                if (graduationOptionalAccumulatedCredit.HasValue)
                    schoolRollTable.Tables.Add(graduationOptionalAccumulatedCredit.ToDataTable("GraduationOptionalAccumulatedCredit", "校定選修應得學分"));

                if (graduationAcquiredCredit.HasValue)
                    schoolRollTable.Tables.Add(graduationAcquiredCredit.ToDataTable("GraduationAcquiredCredit", "畢業獲得總學分"));
            }

            #endregion

            #region 專業(實習)科目及格學分與累計學分
            //  專業科目及格學分
            schoolRollTable.Tables.Add(dataPool.GetProSubjectAccquiredCredit(student.ID).ToDataTable("GraduationSpecializationAcquiredCredit", "專業科目及格學分"));
            //  專業科目累計學分
            schoolRollTable.Tables.Add(dataPool.GetProSubjectAccumulatedCredit(student.ID).ToDataTable("GraduationSpecializationAccumulatedCredit", "專業科目累計學分"));
            //  實習科目累計學分與及格學分
            List<SHSubjectSemesterScoreInfo> _SHSubjectSemesterScoreInfo = dataPool.GetPersonalAllSemesterSubjectScoreInfo(student.ID);
            if (_SHSubjectSemesterScoreInfo != null)
            {
                List<KeyValuePair<int, int>> kvs = new List<KeyValuePair<int, int>>();
                foreach (int index in _PersonalSemesterSubjectScoreHistoryInfo.Keys)
                    _PersonalSemesterSubjectScoreHistoryInfo[index].ForEach(x => kvs.Add(x));
                decimal proAccumulatedCredit = 0;
                decimal proAccquiredCredit = 0;
                foreach (SHSubjectSemesterScoreInfo sss in _SHSubjectSemesterScoreInfo)
                {
                    if (kvs.Where(y => (y.Key == sss.SchoolYear && y.Value == sss.Semester)).Count() > 0 && sss.Entry == "實習科目")
                    {
                        if (sss.Credit.HasValue && sss.Pass.HasValue && sss.Pass.Value)
                            if (sss.Credit.HasValue)
                                proAccquiredCredit += sss.Credit.Value;

                        if (sss.Credit.HasValue)
                            proAccumulatedCredit += sss.Credit.Value;
                    }
                }
                schoolRollTable.Tables.Add(proAccquiredCredit.ToDataTable("GraduationPracticeAcquiredCredit", "實習科目及格學分"));
                schoolRollTable.Tables.Add(proAccumulatedCredit.ToDataTable("GraduationPracticeAccumulatedCredit", "實習科目累計學分"));
            }
            #endregion

            #region 獎懲相抵未滿三大過判讀資料

            schoolRollTable.Tables.Add((dataPool.NotExceedThreeMajorDemerits(student.ID) ? "是" : "否").ToDataTable("NotExceedThreeMajorDemerits", "獎懲相抵未滿三大過"));

            #endregion

            #region 學年重讀或升級判讀資料

            if (_PersonalYearSubjectScoreHistoryInfo != null)
            {
                foreach (KeyValuePair<int, int> kv in _PersonalYearSubjectScoreHistoryInfo)
                {
                    bool retainInTheSameGrade = dataPool.RetainInTheSameGrade(student.ID, kv);

                    schoolRollTable.Tables.Add((retainInTheSameGrade ? "重讀" : "升級").ToDataTable("Y" + kv.Key.ToString() + "RetainInTheSameGrade", "升級或應重讀判斷"));
                }
            }

            #endregion

            # region 學生異動資料
            List<SHUpdateRecordRecord> shurrs = dataPool.GetUpdateRecord(student.ID);

            if (shurrs != null && shurrs.Count > 0)
            {
                //   入學異動(新生與他校轉入依據異動日期遞減排序，取第一筆)//  入學異動
                //  異動代碼首碼為「0」者：新生異動
                //  異動代碼前3碼為「101」且為進校學籍表：它校轉入
                //  異動代碼前3碼為「111」且為日校、高職學籍表：它校轉入
                var enrollUpdateRecords = from updateRecords in shurrs
                                          where (updateRecords.UpdateCode.Substring(0, 1) == "0" || ((templateNumber == 5 || templateNumber == 6) && updateRecords.UpdateCode.Substring(0, 3) == "101") || ((templateNumber != 5 && templateNumber != 6) && updateRecords.UpdateCode.Substring(0, 3) == "111"))
                                          orderby updateRecords.UpdateDate descending
                                          select updateRecords;

                DataTable EnrollRecordTable = new DataTable("EnrollUpdateRecord");
                EnrollRecordTable.Columns.Add("EnrollUpdateRecord");
                foreach (SHUpdateRecordRecord u in enrollUpdateRecords)
                {
                    DataRow pRow = EnrollRecordTable.NewRow();

                    string graduateSchool = (u.UpdateCode.Substring(0, 1) == "0" ? (u.GraduateSchoolLocationCode + " " + u.GraduateSchool) : u.PreviousSchool);
                    pRow["EnrollUpdateRecord"] = graduateSchool;
                    EnrollRecordTable.Rows.Add(pRow);

                    pRow = EnrollRecordTable.NewRow();
                    pRow["EnrollUpdateRecord"] = u.UpdateCode + " " + u.UpdateDescription;
                    EnrollRecordTable.Rows.Add(pRow);

                    pRow = EnrollRecordTable.NewRow();
                    pRow["EnrollUpdateRecord"] = u.ADDate + "\n" + u.ADNumber;
                    EnrollRecordTable.Rows.Add(pRow);

                    break;
                }
                schoolRollTable.Tables.Add(EnrollRecordTable);
                //  畢業異動
                var graduateUpdateRecords = from updateRecords in shurrs
                                            where (updateRecords.UpdateCode.Substring(0, 1) == "5")
                                            select updateRecords;

                DataTable GraduateRecordTable = new DataTable("GraduateUpdateRecord");
                GraduateRecordTable.Columns.Add("GraduateUpdateRecord");
                foreach (SHUpdateRecordRecord u in graduateUpdateRecords)
                {
                    DataRow pRow = GraduateRecordTable.NewRow();

                    pRow["GraduateUpdateRecord"] = u.GraduateCertificateNumber;
                    GraduateRecordTable.Rows.Add(pRow);

                    pRow = GraduateRecordTable.NewRow();
                    pRow["GraduateUpdateRecord"] = u.ADDate + " " + u.ADNumber;
                    GraduateRecordTable.Rows.Add(pRow);

                    break;
                }
                schoolRollTable.Tables.Add(GraduateRecordTable);
                //  學籍異動
                var otherUpdateRecords = from updateRecords in shurrs
                                         where (updateRecords.UpdateCode.Substring(0, 1) != "0" && updateRecords.UpdateCode.Substring(0, 1) != "5")
                                         orderby updateRecords.UpdateDate
                                         select updateRecords;

                DataTable OtherUpdateRecordTable = new DataTable("OtherUpdateRecord");
                OtherUpdateRecordTable.Columns.Add("其它異動日期");
                OtherUpdateRecordTable.Columns.Add("其它異動事項");
                OtherUpdateRecordTable.Columns.Add("其它異動文號");
                foreach (SHUpdateRecordRecord u in otherUpdateRecords)
                {
                    if ((templateNumber == 5 || templateNumber == 6) && u.UpdateCode.Substring(0, 3) == "101")
                        continue;

                    if ((templateNumber != 5 && templateNumber != 6) && u.UpdateCode.Substring(0, 3) == "111")
                        continue;

                    DataRow pRow = OtherUpdateRecordTable.NewRow();

                    pRow["其它異動日期"] = u.UpdateDate;
                    pRow["其它異動事項"] = u.UpdateCode + " " + u.UpdateDescription;
                    pRow["其它異動文號"] = u.ADDate + " " + u.ADNumber;

                    OtherUpdateRecordTable.Rows.Add(pRow);
                }
                schoolRollTable.Tables.Add(OtherUpdateRecordTable);
            }
            #endregion

            #region 報表列印日期

            schoolRollTable.Tables.Add((DateTime.Today.Year - 1911).ConvertToChineseNumber().ToDataTable("PrintYear", "報表列印日期之民國年"));
            schoolRollTable.Tables.Add(DateTime.Today.Month.ConvertToChineseNumber().ToDataTable("PrintMonth", "報表列印日期之月份"));
            schoolRollTable.Tables.Add(DateTime.Today.Day.ConvertToChineseNumber().ToDataTable("PrintDay", "報表列印日期之日"));

            #endregion

            return schoolRollTable;
        }

        /// <summary>
        /// 產生學籍表
        /// </summary>       
        private List<DataSet> Transfer()
        {
            List<DataSet> newFormat = new List<DataSet>();
            foreach (SHStudentRecord student in students)
            {
                // 轉換每一位學生的學籍表資料為樣版所需格式
                newFormat.Add(GetSchoolRollTable(student, dataPool));
            }

            return newFormat;
        }

        // Given an workbook file and custom XML content as a string, add a new custom
        // XML part to the workbook.
        //public void XLInsertCustomXml(string fileName, string customXML)
        //{
        //    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
        //    {
        //        WorkbookPart wbPart = document.WorkbookPart;

        //        var part = wbPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        //        using (StreamWriter sw =
        //          new StreamWriter(part.GetStream(FileMode.OpenOrCreate, FileAccess.Write)))
        //        {
        //            sw.Write(customXML);
        //        }
        //        wbPart.Workbook.Save();
        //    }
        //}
    }
}
//  Path.Combine(pPath, s + ".xls")
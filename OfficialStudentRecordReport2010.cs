using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using FISCA.Presentation;
using FISCA.Permission;
using System.Data;
using SHSchool.Data;
using K12.Presentation;
using System.Windows.Forms;
using System.IO;
using FISCA.Presentation.Controls;
using Aspose.Cells;
using System.Threading;
using ReportHelper;

namespace OfficialStudentRecordReport2010
{
    class OfficialStudentRecordReport2010
    {        
        // 所有學生的學籍表原始資料
        DataPool dataPool = null;
        Catalog button_student, button_class;
        BackgroundWorker _BGWStudentRecord;
        //ManualResetEvent SyncEvent;

        string reportName = string.Empty;
        string reportPath = string.Empty;

        KeyValuePair<string, List<string>> kv;
        object[] bgwObject;

        IEnumerable<SHStudentRecord> students;

        public OfficialStudentRecordReport2010()
        {
            reportName = "學籍表(97學年度入學適用)";
            reportPath = "資料統計";

            //  學生頁籤-->成績相關報表-->學籍表(97學年度入學適用)
            button_student = RoleAclSource.Instance["學生"]["報表"];
            button_student.Add(new RibbonFeature("Student_Button_OfficialStudentRecordReport2010", "學籍表(97學年度入學適用)"));
            MotherForm.RibbonBarItems["學生", reportPath]["報表"]["成績相關報表"][reportName].Enable = UserAcl.Current["Student_Button_OfficialStudentRecordReport2010"].Executable;
            MotherForm.RibbonBarItems["學生", reportPath]["報表"]["成績相關報表"][reportName].Click += new EventHandler(button_student_OnClick);

            //  班級頁籤-->成績相關報表-->學籍表(97學年度入學適用)
            button_class = RoleAclSource.Instance["班級"]["報表"];
            button_class.Add(new RibbonFeature("Class_Button_OfficialStudentRecordReport2010", "學籍表(97學年度入學適用)"));
            MotherForm.RibbonBarItems["班級", reportPath]["報表"]["成績相關報表"][reportName].Enable = UserAcl.Current["Class_Button_OfficialStudentRecordReport2010"].Executable;
            MotherForm.RibbonBarItems["班級", reportPath]["報表"]["成績相關報表"][reportName].Click += new EventHandler(button_class_OnClick);

            button_student = null;
            button_class = null;
        }

        private void button_student_OnClick(object sender, EventArgs e)
        {
            if (NLDPanels.Student.SelectedSource.Count == 0)
            {
                MsgBox.Show("請先選取學生。");
                return;
            }

            kv = new KeyValuePair<string, List<string>>("STUDENT", NLDPanels.Student.SelectedSource);
            Clicked();
        }

        private void button_class_OnClick(object sender, EventArgs e)
        {
            if (NLDPanels.Class.SelectedSource.Count == 0)
            {
                MsgBox.Show("請先選取班級。");
                return;
            }

            kv = new KeyValuePair<string, List<string>>("CLASS", NLDPanels.Class.SelectedSource);
            Clicked();
        }
        
        private void PrepairData(FrontForm form)
        {
            int templateNumber = form.TemplateNumber;

            MemoryStream selectedTemplate = form.Template;

            string text1 = form.Text1;
            string text2 = form.Text2;
            string text3 = form.Text3;

            int custodian = form.Custodian;
            int address = form.Address;
            int phone = form.Phone;

            string coreSubjectSign = form.SignCoreSubject;
            string coreCourseSign = form.SignCoreCourse;
            string resitSign = form.SignResit;
            string retakeSign = form.SignRetake;
            string failedSign = form.SignFailed;
            string schoolYearAdjustSign = form.SignSchoolYearAdjust;
            string manualAdjustSign = form.SignManualAdjust;

            string dataType = string.Empty;

            int moralScoreType = form.MoralScoreOption;
            int saveFileType = form.SaveFileType;

            bgwObject = new object[] { 
                selectedTemplate,
                templateNumber,
                text1,
                text2,
                text3,
                custodian,
                address,
                phone,
                coreSubjectSign,
                coreCourseSign,
                resitSign,
                retakeSign,
                failedSign,
                schoolYearAdjustSign,
                manualAdjustSign,
                dataPool,
                saveFileType,
                moralScoreType,
                students
            };
        }

        private void DisableAllButton(FrontForm form)
        {
            foreach (Control ctl in form.Controls)
                ctl.Enabled = false;
        }

        private void EnableAllButton(FrontForm form)
        {
            foreach (Control ctl in form.Controls)
                ctl.Enabled = true;
        }

        private void Clicked()
        {
            FrontForm form = new FrontForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                MotherForm.RibbonBarItems["學生", reportPath]["報表"]["成績相關報表"][reportName].Enable = false;
                MotherForm.RibbonBarItems["班級", reportPath]["報表"]["成績相關報表"][reportName].Enable = false;
                DisableAllButton(form);
                //ThreadPool.QueueUserWorkItem(arg =>
                //{
                //    SyncEvent.WaitOne();
                try
                {
                    PrepairData(form);

                    _BGWStudentRecord = new BackgroundWorker();
                    _BGWStudentRecord.WorkerReportsProgress = true;
                    _BGWStudentRecord.WorkerSupportsCancellation = true; 
                    _BGWStudentRecord.DoWork += new DoWorkEventHandler(_BGWStudentRecord_DoWork);
                    _BGWStudentRecord.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWStudentRecord_RunWorkerCompleted);
                    _BGWStudentRecord.ProgressChanged += new ProgressChangedEventHandler(_BGWStudentRecord_ProgressChanged);
                    MotherForm.SetStatusBarMessage("學籍表(97學年度入學適用)產生中...");
                    _BGWStudentRecord.RunWorkerAsync();
                    
                    //
                    //SyncEvent = null;

                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);
                }
                finally
                {
                    _BGWStudentRecord.Dispose();
                    EnableAllButton(form);
                    form.Dispose();                    
                }
                //});
            }
            //else
            //    _BGWStudentRecord.CancelAsync();
        }

        private void Print()
        {
            bgwObject[15] = dataPool;
            string filePath = (string)bgwObject[4];
            int saveFileType = (int)bgwObject[16];

            IEnumerable<SHStudentRecord> allCacheStudents = dataPool.GetAllStudents();
            allCacheStudents = allCacheStudents.OrderBy(x => (String.IsNullOrEmpty(x.StudentNumber) ? "0" : x.StudentNumber));
            
            List<SHStudentRecord> noSNumStudents = new List<SHStudentRecord>();
            foreach(SHStudentRecord student in allCacheStudents)
            {
                if (String.IsNullOrEmpty(student.StudentNumber))
                    noSNumStudents.Add(student);
            }

            if (noSNumStudents.Count > 0)
            {
                string message = "下列學生沒有學號，請補齊資料再列印學籍表：\n";
                foreach(SHStudentRecord student in noSNumStudents)
                {
                    message += (student.Class == null ? "" : student.Class.Name) + student.Name + "\n";
                }
                throw new Exception(message);
            }

            IEnumerable<SHStudentRecord> loseClassReferenceStudents = allCacheStudents.Where(x => (x.Class == null)).Select(x => x);

            ReportDataFormatTransfer rdft = new ReportDataFormatTransfer();
            Dictionary<string, List<DataSet>> allData = new Dictionary<string,List<DataSet>>();

            //  每個班級一個檔案
            if (saveFileType == 1)
            {
                if (loseClassReferenceStudents.Count() > 0)
                {
                    string message = "下列學生沒有設定班級，無法以「每個班級儲存一個檔案」的方式產生學籍表：\n";
                    foreach (SHStudentRecord student in loseClassReferenceStudents)
                    {
                        message += student.Name + "\n";
                    }
                    throw new Exception(message);
                }
                Dictionary<string, List<SHStudentRecord>> hasClassReferenceStudents = new Dictionary<string, List<SHStudentRecord>>();

                foreach (SHStudentRecord student in allCacheStudents)
                {
                    if (student.Class != null)
                    {
                        string className = student.Class.Name;
                        if (!hasClassReferenceStudents.ContainsKey(className))
                        {
                            hasClassReferenceStudents.Add(className, new List<SHStudentRecord>());
                        }
                        hasClassReferenceStudents[className].Add(student);
                    }
                }
                //  有班級的學生每班儲存一個檔案
                int a = 0;
                foreach (string className in hasClassReferenceStudents.Keys)
                {
                    a++;
                    _BGWStudentRecord.ReportProgress((int)((a*1.0/ hasClassReferenceStudents.Keys.Count)*100));
                    var sameClassStudent = from pair in hasClassReferenceStudents[className] orderby (pair.StudentNumber == null ? 0 : Convert.ToDecimal(pair.StudentNumber)) select pair;

                    bgwObject[18] = sameClassStudent.ToList();

                    allData = new Dictionary<string, List<DataSet>>();

                    if (!allData.ContainsKey(className))
                        allData.Add(className, new List<DataSet>());

                    List<DataSet> newFormat = rdft.TransferFormat(bgwObject);
                    allData[className].AddRange(newFormat);

                    Workbook wb = Report.Produce(allData, (bgwObject[0] as MemoryStream));

                    Util.Completed(Path.Combine(filePath, className + ".xls"), wb);

                    string text2 = string.Empty;
                    if (!string.IsNullOrEmpty("" + bgwObject[3]))
                        text2 = (Convert.ToInt32(bgwObject[3]) + newFormat.Count).ToString("d" + ((string)bgwObject[3]).Length);

                    bgwObject[3] = text2;
                    wb = null;
                }
            }
            //  每個學生一個檔案
            if (saveFileType == 2)
            {
                int a = 0;
                foreach (SHStudentRecord student in allCacheStudents)
                {
                    a++;
                    _BGWStudentRecord.ReportProgress((int)((a * 1.0 / allCacheStudents.Count()) * 100));
                    bgwObject[18] = new List<SHStudentRecord>() { student };

                    allData = new Dictionary<string, List<DataSet>>();
                    string fileName = student.StudentNumber + "_" + student.Name;

                    List<DataSet> newFormat = rdft.TransferFormat(bgwObject);
                    allData.Add(fileName, newFormat);

                    Workbook wb = Report.Produce(allData, (bgwObject[0] as MemoryStream));

                    Util.Completed(Path.Combine(filePath, fileName + ".xls"), wb);

                    string text2 = string.Empty;

                    if (!string.IsNullOrEmpty(""+bgwObject[3]))
                       text2 = (Convert.ToInt32(bgwObject[3]) + newFormat.Count).ToString("d" + ((string)bgwObject[3]).Length);

                    bgwObject[3] = text2;
                    wb = null;
                }
            }
            //  每 100 個學生儲存一個檔案
            if (saveFileType == 3)
            {
                string firstSnum = string.Empty;
                string lastSnum = string.Empty;

                //decimal snum = 0;
                //var sameFileStudent = from pair in allCacheStudents orderby (decimal.TryParse(pair.StudentNumber, out snum) ? 0 : snum) select pair;

                if (allCacheStudents.Count() <= 100)
                {
                    allData = new Dictionary<string, List<DataSet>>();
                    string key = allCacheStudents.ElementAt(0).StudentNumber + "_" + allCacheStudents.ElementAt(allCacheStudents.Count() - 1).StudentNumber;
                    allData.Add(key, new List<DataSet>());

                    bgwObject[18] = allCacheStudents;

                    List<DataSet> newFormat = rdft.TransferFormat(bgwObject);
                    allData[key].AddRange(newFormat);

                    Workbook wb = Report.Produce(allData, (bgwObject[0] as MemoryStream));

                    Util.Completed(Path.Combine(filePath, key + ".xls"), wb);

                    wb = null;
                }
                else
                {
                    decimal k = Math.Ceiling((decimal)(allCacheStudents.Count() / 100.0));

                    for (int j = 0; j < k; j++)
                    {
                        int z = ((j * 100 + 100) > allCacheStudents.Count()) ? allCacheStudents.Count() : (j * 100 + 100);

                        string key = allCacheStudents.ElementAt(j * 100).StudentNumber + "_" + allCacheStudents.ElementAt(z - 1).StudentNumber;
                        allData = new Dictionary<string, List<DataSet>>();
                        allData.Add(key, new List<DataSet>());

                        List<SHStudentRecord> rangeStudent = new List<SHStudentRecord>();
                        for (int m = (j * 100); m < z; m++)
                        {
                            _BGWStudentRecord.ReportProgress((int)(((m + 1) * 1.0 / (double)allCacheStudents.Count()) * 100));
                            rangeStudent.Add(allCacheStudents.ElementAt(m));
                        }

                        bgwObject[18] = rangeStudent;
                        List<DataSet> newFormat = rdft.TransferFormat(bgwObject);
                        allData[key].AddRange(newFormat);

                        Workbook wb = Report.Produce(allData, (bgwObject[0] as MemoryStream));

                        Util.Completed(Path.Combine(filePath, key + ".xls"), wb);

                        string text2 = string.Empty;

                        if (!string.IsNullOrEmpty("" + bgwObject[3]))
                            text2 = (Convert.ToInt32(bgwObject[3]) + newFormat.Count).ToString("d" + ((string)bgwObject[3]).Length);

                        bgwObject[3] = text2;
                        wb = null;
                    }
                }
            }
            rdft.Dispose();
            rdft = null;
        }

        private void _BGWStudentRecord_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result != null && e.Result.GetType().Equals(Type.GetType("System.Exception")))
                MsgBox.Show(((System.Exception)e.Result).Message);
            else
            {
                MessageBox.Show("學籍表產生完成。", "學籍表", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                System.Diagnostics.Process.Start((string)bgwObject[4]);
            }
            //if (e.Cancelled == true)
            //{
            //    resultLabel.Text = "Canceled!";
            //}
            //else if (e.Error != null)
            //{
            //    resultLabel.Text = "Error: " + e.Error.Message;
            //}
            //else
            //{
            //    resultLabel.Text = "Done!";
            //}
            MotherForm.SetStatusBarMessage("");

            dataPool.Dispose();
            MotherForm.RibbonBarItems["學生", reportPath]["報表"]["成績相關報表"][reportName].Enable = true;
            MotherForm.RibbonBarItems["班級", reportPath]["報表"]["成績相關報表"][reportName].Enable = true;

            // 產生學籍表的過程，系統佔用之記憶體，回收之
            (bgwObject[0] as MemoryStream).Dispose();
            (bgwObject[15] as DataPool).Dispose();
            bgwObject[18] = null;
            bgwObject = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();            
        }

        private void _BGWStudentRecord_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("學籍表(97學年度入學適用)產生中...", e.ProgressPercentage);
        }

        private void _BGWStudentRecord_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                dataPool = new DataPool(kv);
                Print();
            }
            catch (Exception ex)
            {
                e.Result = new Exception(ex.Message);
            }
        }
    }
}

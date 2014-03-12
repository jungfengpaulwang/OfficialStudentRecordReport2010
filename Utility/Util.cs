using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Windows.Forms;
using System.IO;
using FISCA.Presentation.Controls;

namespace OfficialStudentRecordReport2010
{
    public class Util
    {
        public static void Completed(string inputReportName, Workbook wb)
        {
            //string reportName = inputReportName;

            //string path = Path.Combine(Application.StartupPath, "Reports");
            //if (!Directory.Exists(path))
            //    Directory.CreateDirectory(path);
            //path = Path.Combine(path, reportName + ".xls");

            //Workbook doc = wb;

            //if (File.Exists(path))
            //{
            //    int i = 1;
            //    while (true)
            //    {
            //        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
            //        if (!File.Exists(newPath))
            //        {
            //            path = newPath;
            //            break;
            //        }
            //    }
            //}

            try
            {
                wb.Save(inputReportName, FileFormatType.Excel2003);
                //System.Diagnostics.Process.Start(path);
            }
            catch
            {
                //SaveFileDialog sd = new SaveFileDialog();
                //sd.Title = "另存新檔";
                //sd.FileName = reportName + ".xls";
                //sd.Filter = "Excel 2003 相容檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                //if (sd.ShowDialog() == DialogResult.OK)
                //{
                //    try
                //    {
                //        doc.Save(path, FileFormatType.Excel2003);
                //    }
                //    catch
                //    {
                        MsgBox.Show("指定路徑無法存檔。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                //    }
                //}
            }
        }

        public static string ChineseDate(string p)
        {
            DateTime d = DateTime.Now;
            if (p != "" && DateTime.TryParse(p, out d))
            {
                return "" + (d.Year - 1911) + "/" + d.Month + "/" + d.Day;
            }
            else
                return "";
        }

        public static string NumberToRomanChar(string p)
        {
            string levelNumber;
            switch (p)
            {
                #region 對應levelNumber
                case "0":
                    levelNumber = "";
                    break;
                case "1":
                    levelNumber = "Ⅰ";
                    break;
                case "2":
                    levelNumber = "Ⅱ";
                    break;
                case "3":
                    levelNumber = "Ⅲ";
                    break;
                case "4":
                    levelNumber = "Ⅳ";
                    break;
                case "5":
                    levelNumber = "Ⅴ";
                    break;
                case "6":
                    levelNumber = "Ⅵ";
                    break;
                case "7":
                    levelNumber = "Ⅶ";
                    break;
                case "8":
                    levelNumber = "Ⅷ";
                    break;
                case "9":
                    levelNumber = "Ⅸ";
                    break;
                case "10":
                    levelNumber = "Ⅹ";
                    break;
                default:
                    levelNumber = "" + (p);
                    break;
                #endregion
            }
            return levelNumber;
        }

        public static string NumberToRomanChar(int p)
        {
            string levelNumber;
            switch (p)
            {
                #region 對應levelNumber
                case 0:
                    levelNumber = "";
                    break;
                case 1:
                    levelNumber = "Ⅰ";
                    break;
                case 2:
                    levelNumber = "Ⅱ";
                    break;
                case 3:
                    levelNumber = "Ⅲ";
                    break;
                case 4:
                    levelNumber = "Ⅳ";
                    break;
                case 5:
                    levelNumber = "Ⅴ";
                    break;
                case 6:
                    levelNumber = "Ⅵ";
                    break;
                case 7:
                    levelNumber = "Ⅶ";
                    break;
                case 8:
                    levelNumber = "Ⅷ";
                    break;
                case 9:
                    levelNumber = "Ⅸ";
                    break;
                case 10:
                    levelNumber = "Ⅹ";
                    break;
                default:
                    levelNumber = "" + (p);
                    break;
                #endregion
            }
            return levelNumber;
        }

        public static int SortBySubjectName(SHSubjectSemesterScoreInfo a, SHSubjectSemesterScoreInfo b)
        {
            string a1 = a.SubjectName.Length > 0 ? a.SubjectName.Substring(0, 1) : "";
            string b1 = b.SubjectName.Length > 0 ? b.SubjectName.Substring(0, 1) : "";
            #region 第一個字一樣的時候
            if (a1 == b1)
            {
                //if (a.SubjectName.Length > 1 && b.SubjectName.Length > 1)
                //    return SortBySubjectName(a, b);
                //else
                    return a.SubjectName.Length.CompareTo(b.SubjectName.Length);
            }
            #endregion
            #region 第一個字不同，分別取得在設定順序中的數字，如果都不在設定順序中就用單純字串比較
            int ai = getIntForSubject(a1), bi = getIntForSubject(b1);
            if (ai > 0 || bi > 0)
                return ai.CompareTo(bi);
            else
                return a1.CompareTo(b1);
            #endregion
        }

        public static int getIntForSubject(string a1)
        {
            List<string> list = new List<string>();
            list.AddRange(new string[] { "國", "英", "數", "物", "化", "生", "基", "歷", "地", "公", "文", "礎", "世" });

            int x = list.IndexOf(a1);
            if (x < 0)
                return list.Count;
            else
                return x;
        }

        public static int SortByEntryName(string a, string b)
        {
            int ai = getIntForEntry(a), bi = getIntForEntry(b);
            if (ai > 0 || bi > 0)
                return ai.CompareTo(bi);
            else
                return a.CompareTo(b);
        }

        public static int getIntForEntry(string a1)
        {
            List<string> list = new List<string>();
            list.AddRange(new string[] { "學業", "學業成績名次", "實習科目", "體育", "國防通識", "健康與護理", "德行" });

            int x = list.IndexOf(a1);
            if (x < 0)
                return list.Count;
            else
                return x;
        }

        //public static int SortSubjectByLevel(SemesterSubjectScoreInfo x, SemesterSubjectScoreInfo y)
        //{
        //    if (x.Subject != y.Subject)
        //        return 0;
        //    else
        //    {
        //        int a, b;
        //        int.TryParse(x.Level, out a);
        //        int.TryParse(y.Level, out b);
        //        return a.CompareTo(b);
        //    }
        //}
    }
}

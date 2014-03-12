using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using System.Windows.Forms;

namespace OfficialStudentRecordReport2010
{
    public static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        //[MainMethod()]
        [FISCA.MainMethod("學籍表(97學年度入學適用)")]
        static public void Main()
        {
            OfficialStudentRecordReport2010 report = new OfficialStudentRecordReport2010();
        }
    }
}

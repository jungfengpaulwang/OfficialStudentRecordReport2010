using System;
using System.Collections.Generic;
using System.Data;
using System.Collections;
using System.Reflection;

namespace OfficialStudentRecordReport2010
{
    public static class DataTableExtensions
    {
        public static List<string> ConvertToString(this IEnumerable<bool?> value, 

Dictionary<bool?, string> d)
        {
            List<string> newValue = new List<string>();
            foreach (bool? si in value)
            {
                if (!si.HasValue)
                {
                    newValue.Add("");
                    continue;
                }
                if (d.ContainsKey(si))
                    newValue.Add(d[(bool)si]);
                else
                    newValue.Add("");
            }
            return newValue;
        }

        public static string ConvertToChineseNumber(this int pNumber)
        {
            string mNumber = "";

            for (int i = 0; i < pNumber.ToString().Length; i++)
                mNumber += ChineseNo(Convert.ToInt32(pNumber.ToString().Substring(i, 1)));

            if (mNumber.Length == 2)
            {
                if (mNumber.EndsWith("Ｏ"))
                    mNumber = mNumber.Substring(0, 1) + "十";

                mNumber = mNumber.Substring(0, 1) + "十" + mNumber.Substring(1, 1);
                mNumber = mNumber.Replace("一十", "十");
                mNumber = mNumber.Replace("十十", "十");

                return mNumber;
            }
            else
                return mNumber;
        }
         
        public static char ChineseNo(int pNumber)
        {
            char[] mNumber = new char[] { 'Ｏ', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十' };
            return mNumber[pNumber];
        }
    }
}
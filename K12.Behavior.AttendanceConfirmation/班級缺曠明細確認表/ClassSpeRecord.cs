using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Behavior.AttendanceConfirmation
{
    /// <summary>
    /// 專門用來擺放一個要處理的班級資料
    /// </summary>
    class ClassSpeRecord
    {
        /// <summary>
        /// 缺曠明細，Key為 StudentID -> OccurDate -> Period
        /// </summary>
        public Dictionary<string, Dictionary<string, Dictionary<string, string>>> _dic { get; set; }

        /// <summary>
        /// 學生物件
        /// </summary>
        public Dictionary<string, StudentRecord> StudentRecordDic { get; set; }

        public int PeriodCount { get; set; }

        public Dictionary<string, int> ColumnIndex { get; set; }

        public ClassSpeRecord(Dictionary<string, Dictionary<string, Dictionary<string, string>>> dic)
        {
            _dic = dic;
            StudentRecordDic = new Dictionary<string, StudentRecord>();

            List<StudentRecord> SrList = K12.Data.Student.SelectByIDs(dic.Keys);

            foreach (StudentRecord each in SrList)
            {
                if (!StudentRecordDic.ContainsKey(each.ID))
                {
                    StudentRecordDic.Add(each.ID, each);
                }
            }

            SrList.Sort(new Comparison<StudentRecord>(tool.StudentComparer));

        }
    }
}

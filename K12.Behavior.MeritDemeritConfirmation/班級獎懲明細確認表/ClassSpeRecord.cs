using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Behavior.MeritDemeritConfirmation
{
    /// <summary>
    /// 專門用來擺放一個要處理的班級資料
    /// </summary>
    class ClassSpeRecord
    {
        /// <summary>
        /// 獎懲明細，Key為 StudentID -> 獎懲清單
        /// </summary>
        public Dictionary<string, MeritAndDemerit> _dic { get; set; }

        /// <summary>
        /// 學生物件
        /// </summary>
        public Dictionary<string, StudentRecord> StudentRecordDic { get; set; }

        public Dictionary<string, int> ColumnIndex { get; set; }

        public ClassSpeRecord(Dictionary<string, MeritAndDemerit> dic)
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

    class MeritAndDemerit
    {
        public List<MeritRecord> GetMerList { get; set; }

        public List<DemeritRecord> GetDemerList { get; set; }

        public MeritAndDemerit()
        {
            GetMerList = new List<MeritRecord>();
            GetDemerList = new List<DemeritRecord>();
        }
    }
}

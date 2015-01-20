using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace K12.懲戒通知單
{
    class ConfigOBJ
    {
        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StartDate { get; set; }
        /// <summary>
        /// 結束時間
        /// </summary>
        public DateTime EndDate { get; set; }
        /// <summary>
        /// 沒有資料即不印
        /// </summary>
        public bool PrintHasRecordOnly { get; set; }
        /// <summary>
        /// 範本
        /// </summary>
        public MemoryStream Template { get; set; }
        /// <summary>
        /// 寄件人姓名
        /// </summary>
        public string ReceiveName { get; set; }
        /// <summary>
        /// 寄件人地址
        /// </summary>
        public string ReceiveAddress { get; set; }
        /// <summary>
        /// 懲戒名稱
        /// </summary>
        public string ConditionName { get; set; }
        /// <summary>
        /// 懲戒支數
        /// </summary>
        public string ConditionNumber { get; set; }
        /// <summary>
        /// 發生日期 or 登錄日期
        /// </summary>
        public bool IsInsertDate { get; set; }

        /// <summary>
        /// 是否列印學生清單
        /// </summary>
        public bool PrintStudentList { get; set; }
    }
}

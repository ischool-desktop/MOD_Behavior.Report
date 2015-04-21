using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace K12.Behavior.AttendanceConfirmation
{
    class GetAbsenceDetail
    {

        /// <summary>
        /// 選取的班級資料
        /// </summary>
        public List<ClassRecord> selectedClass;

        /// <summary>
        /// 缺曠明細，Key為 ClassID -> StudentID -> OccurDate -> Period
        /// </summary>
        public Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>> allAbsenceDetail = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, string>>>>();

        /// <summary>
        /// 學生ID查詢班級ID
        /// </summary>
        public Dictionary<string, string> studentClassDict = new Dictionary<string, string>();

        /// <summary>
        /// 學生資料清單
        /// </summary>
        public List<StudentRecord> StudentRecordList = new List<StudentRecord>();

        /// <summary>
        /// 學生ID查詢學生資訊
        /// </summary>
        public Dictionary<string, StudentRecord> studentInfoDict = new Dictionary<string, StudentRecord>();

        public GetAbsenceDetail(DateTime startDate, DateTime endDate, List<string> periodList)
        {
            #region 建構子
            allAbsenceDetail.Clear();
            studentClassDict.Clear();
            StudentRecordList.Clear();
            studentInfoDict.Clear();

            AttendanceConfigData cd = new AttendanceConfigData();

            selectedClass = Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
            selectedClass = SortClassIndex.K12Data_ClassRecord(selectedClass);
            //selectedClass.Sort(new Comparison<ClassRecord>(ClassComparer));

            //所有學生ID
            List<string> allStudentID = new List<string>();

            //缺曠筆數

            //紀錄每一個 Column 的 Index
            Dictionary<string, int> columnTable = new Dictionary<string, int>();

            //建立學生班級對照表
            foreach (ClassRecord aClass in selectedClass)
            {
                foreach (StudentRecord student in aClass.Students)
                {
                    if (student.Status == StudentRecord.StudentStatus.刪除 || student.Status == StudentRecord.StudentStatus.畢業或離校)
                        continue;

                    StudentRecordList.Add(student);
                    allStudentID.Add(student.ID);
                    studentClassDict.Add(student.ID, aClass.ID);
                    studentInfoDict.Add(student.ID, student);
                }

                allAbsenceDetail.Add(aClass.ID, new Dictionary<string, Dictionary<string, Dictionary<string, string>>>());
            }

            //取得相關資料
            List<AttendanceRecord> GetAttList = Attendance.SelectByDate(StudentRecordList, startDate, endDate);

            //處理資料
            foreach (AttendanceRecord var in GetAttList)
            {
                string studentID = var.RefStudentID;
                string occurDate = tool.HowManyWeek(var.OccurDate);
                string classID = studentClassDict[studentID];

                if (!allAbsenceDetail.ContainsKey(classID))
                    allAbsenceDetail.Add(classID, new Dictionary<string, Dictionary<string, Dictionary<string, string>>>());

                if (!allAbsenceDetail[classID].ContainsKey(studentID))
                    allAbsenceDetail[classID].Add(studentID, new Dictionary<string, Dictionary<string, string>>());

                if (!allAbsenceDetail[classID][studentID].ContainsKey(occurDate))
                    allAbsenceDetail[classID][studentID].Add(occurDate, new Dictionary<string, string>());

                foreach (AttendancePeriod period in var.PeriodDetail)
                {
                    //與設定檔對照,如果沒有就不列入列印資料
                    if (!cd.Setup_AttendanceList.Contains(period.AbsenceType))
                        continue;
                    //節次
                    string innertext = period.Period;
                    //缺曠別
                    string absence = period.AbsenceType;

                    if (!allAbsenceDetail[classID][studentID][occurDate].ContainsKey(innertext))
                        allAbsenceDetail[classID][studentID][occurDate].Add(innertext, absence);
                }
            }


            #endregion
        }

        public StudentRecord GetStudentByID(string id)
        {
            if (studentInfoDict.ContainsKey(id))
            {
                return studentInfoDict[id];
            }
            else
            {
                return null;
            }
        }

        public static int ClassComparer(ClassRecord x, ClassRecord y)
        {
            string xx = x.Name;
            string yy = y.Name;

            return xx.CompareTo(yy);
        }
    }
}

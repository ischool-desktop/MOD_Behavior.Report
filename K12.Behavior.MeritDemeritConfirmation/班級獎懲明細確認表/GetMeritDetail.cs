using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Behavior.MeritDemeritConfirmation
{
    class GetMeritDetail
    {
        /// <summary>
        /// 選取的班級資料
        /// </summary>
        public List<ClassRecord> selectedClass;

        /// <summary>
        /// 缺曠明細，Key為 ClassID -> StudentID -> 獎勵內容
        /// </summary>
        public Dictionary<string, Dictionary<string, MeritAndDemerit>> allAbsenceDetail = new Dictionary<string, Dictionary<string, MeritAndDemerit>>();

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



        public GetMeritDetail(DateTime startDate, DateTime endDate)
        {
            allAbsenceDetail.Clear();
            studentClassDict.Clear();
            StudentRecordList.Clear();
            studentInfoDict.Clear();

            MeritConfigData mcd = new MeritConfigData();

            selectedClass = Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
            selectedClass = SortClassIndex.K12Data_ClassRecord(selectedClass);

            //所有學生ID
            List<string> allStudentID = new List<string>();


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

                allAbsenceDetail.Add(aClass.ID, new Dictionary<string, MeritAndDemerit>());
            }


            //取得獎勵相關資料
            List<MeritRecord> GetMerList = Merit.SelectByOccurDate(allStudentID, startDate, endDate);


            //取得懲戒相關資料
            List<DemeritRecord> GetDemerList = Demerit.SelectByOccurDate(allStudentID, startDate, endDate);


            //處理資料
            foreach (MeritRecord var in GetMerList)
            {
                string studentID = var.RefStudentID;
                string occurDate = tool.HowManyWeek(var.OccurDate);
                string classID = studentClassDict[studentID];


                if (!allAbsenceDetail.ContainsKey(classID))
                    allAbsenceDetail.Add(classID, new Dictionary<string, MeritAndDemerit>());

                if (!allAbsenceDetail[classID].ContainsKey(studentID))
                    allAbsenceDetail[classID].Add(studentID, new MeritAndDemerit());

                MeritAndDemerit mad = allAbsenceDetail[classID][studentID];

                mad.GetMerList.Add(var);
            }

            //處理資料
            foreach (DemeritRecord var in GetDemerList)
            {
                if (var.Cleared == "是")
                    continue;

                string studentID = var.RefStudentID;
                string occurDate = tool.HowManyWeek(var.OccurDate);
                string classID = studentClassDict[studentID];


                if (!allAbsenceDetail.ContainsKey(classID))
                    allAbsenceDetail.Add(classID, new Dictionary<string, MeritAndDemerit>());

                if (!allAbsenceDetail[classID].ContainsKey(studentID))
                    allAbsenceDetail[classID].Add(studentID, new MeritAndDemerit());

                MeritAndDemerit mad = allAbsenceDetail[classID][studentID];

                mad.GetDemerList.Add(var);
            }
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

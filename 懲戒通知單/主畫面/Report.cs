using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
using K12.Data;
using Framework.Feature;
using Aspose.Words.Reporting;

namespace K12.懲戒通知單2013
{
    internal class Report : IReport
    {

        private int DemeritAB; // 1 大過 等於 3 小過
        private int DemeritBC; // 1 小過 等於 3 警告
        private int MaxDemerit; //最小懲戒單位值
        private BackgroundWorker _BGWDisciplineNotification;

        private ConfigOBJ obj; //所有列印設定資訊

        private List<StudentRecord> SelectedStudents { get; set; }

        string entityName;

        bool PrintUpdateStudentFile = false;

        public Report(string _entityName)
        {
            entityName = _entityName;
        }

        public void Print()
        {
            #region IReport 成員
            DemeritDateRangeForm form = new DemeritDateRangeForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化懲戒通知單...");

                #region 建立設定檔
                obj = new ConfigOBJ();
                obj.StartDate = form.StartDate;
                obj.EndDate = form.EndDate;
                obj.PrintHasRecordOnly = form.PrintHasRecordOnly;
                obj.Template = form.Template;
                obj.ReceiveName = form.ReceiveName;
                obj.ReceiveAddress = form.ReceiveAddress;
                obj.ConditionName = form.ConditionName;
                obj.ConditionNumber = form.ConditionNumber;
                obj.IsInsertDate = form.radioButton1.Checked;
                obj.PrintStudentList = form.PrintStudentList;
                #endregion

                _BGWDisciplineNotification = new BackgroundWorker();
                _BGWDisciplineNotification.DoWork += new DoWorkEventHandler(_BGWDisciplineNotification_DoWork);
                _BGWDisciplineNotification.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.WordReport_RunWorkerCompleted);
                _BGWDisciplineNotification.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWDisciplineNotification.WorkerReportsProgress = true;
                _BGWDisciplineNotification.RunWorkerAsync();
            }
            #endregion
        }

        private void GetReduceList()
        {
            #region 取得獎懲對照表
            DSResponse dsrsp = Config.GetMDReduce();
            if (!dsrsp.HasContent)
            {
                FISCA.Presentation.Controls.MsgBox.Show("取得對照表失敗 : " + dsrsp.GetFault().Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DSXmlHelper helper = dsrsp.GetContent();

            string jb;

            jb = helper.GetText("Demerit/AB");
            if (!int.TryParse(jb, out DemeritAB))
            {
                MsgBox.Show("獎懲對照表有誤");
                return;
            }

            jb = helper.GetText("Demerit/BC");
            if (!int.TryParse(jb, out DemeritBC))
            {
                MsgBox.Show("獎懲對照表有誤");
                return;
            }

            #endregion
        }

        private void ChengeDemerit(string CondName, int CondNumber)
        {
            #region 取得最小單位值

            MaxDemerit = 0; //最小單位值

            if (CondName == "大過")
            {
                MaxDemerit = CondNumber * DemeritAB;
                MaxDemerit = MaxDemerit * DemeritBC;
            }
            else if (CondName == "小過")
            {
                MaxDemerit = CondNumber * DemeritBC;
            }
            else if (CondName == "警告")
            {
                MaxDemerit = CondNumber;
            }
            #endregion
        }

        private int GetSmallValueType(DemeritRecord record)
        {
            #region 將資料換算為最小單位值
            int AA = record.DemeritA.HasValue ? record.DemeritA.Value : 0;
            int BB = record.DemeritB.HasValue ? record.DemeritB.Value : 0;
            int CC = record.DemeritC.HasValue ? record.DemeritC.Value : 0;

            int SumNum1 = 0;
            int SumNum2 = 0;
            int SumNum3 = 0;

            if (AA != 0)
            {
                SumNum1 = AA * DemeritAB; //轉換成小過
                SumNum1 = SumNum1 * DemeritBC; //轉換成警告
            }

            if (BB != 0)
            {
                SumNum2 = BB * DemeritBC; //轉換成警告
            }

            if (CC != 0)
            {
                SumNum3 = CC; //轉換成警告
            }


            return SumNum1 + SumNum2 + SumNum3;
            #endregion
        }

        private void _BGWDisciplineNotification_DoWork(object sender, DoWorkEventArgs e)
        {
            #region Report
            if (entityName.ToLower() == "student") //學生模式
            {
                SelectedStudents = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
            }
            else if (entityName.ToLower() == "class") //班級模式
            {
                SelectedStudents = new List<StudentRecord>();
                foreach (StudentRecord each in Student.SelectByClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource))
                {
                    if (each.Status != StudentRecord.StudentStatus.一般)
                        continue;

                    SelectedStudents.Add(each);
                }
            }
            else
                throw new NotImplementedException();

            SelectedStudents.Sort(new Comparison<StudentRecord>(CommonMethods.ClassSeatNoComparer));
            #endregion

            #region 表頭

            GetReduceList(); //獎懲對照表

            string reportName = "懲戒通知單";

            //取得換算單位
            ChengeDemerit(obj.ConditionName, int.Parse(obj.ConditionNumber));

            Dictionary<string, int> MDMapping = new Dictionary<string, int>();
            MDMapping.Add("大過", 3);
            MDMapping.Add("小過", 4);
            MDMapping.Add("警告", 5);

            int flag = 2; //預設是2
            if (!string.IsNullOrEmpty(obj.ConditionName)) //如果condName不是空的(可能是大功~警告)
            {
                flag = 0; //小於3是獎勵,大於3是懲戒
            }

            MDFilter filter = new MDFilter();
            if (flag == 0)
                filter.SetCondition(MDMapping[obj.ConditionName], int.Parse(obj.ConditionNumber));

            #endregion

            #region 快取資訊

            //學生資訊
            //Dictionary<string, Dictionary<string, string>> studentInfo = new Dictionary<string, Dictionary<string, string>>();

            //獎懲累計資料
            //Dictionary<string, Dictionary<string, int>> studentDiscipline = new Dictionary<string, Dictionary<string, int>>();

            //懲戒明細
            //Dictionary<string, List<string>> studentDisciplineDetail = new Dictionary<string, List<string>>();



            //超級資訊物件
            Dictionary<string, StudentOBJ> StudentSuperOBJ = new Dictionary<string, StudentOBJ>();
            //所有學生ID
            List<string> allStudentID = new List<string>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

            //獎勵項目
            //Dictionary<string, string> meritTable = new Dictionary<string, string>();
            //meritTable.Add("大功", "A");
            //meritTable.Add("小功", "B");
            //meritTable.Add("嘉獎", "C");

            //懲戒項目
            //Dictionary<string, string> demeritTable = new Dictionary<string, string>();
            //demeritTable.Add("大過", "A");
            //demeritTable.Add("小過", "B");
            //demeritTable.Add("警告", "C");

            #endregion

            #region 依據 ClassID 建立班級學生清單
            //List<StudentRecord> classStudent = SelectedStudents;

            //加總用
            Dictionary<string, int> StudMeritSum = new Dictionary<string, int>();

            foreach (StudentRecord aStudent in SelectedStudents)
            {
                //string aStudentID = aStudent.ID;

                if (!StudentSuperOBJ.ContainsKey(aStudent.ID))
                {
                    StudentSuperOBJ.Add(aStudent.ID, new StudentOBJ());
                }

                //學生ID清單
                if (!allStudentID.Contains(aStudent.ID))
                    allStudentID.Add(aStudent.ID);

                StudentSuperOBJ[aStudent.ID].student = aStudent;
                StudentSuperOBJ[aStudent.ID].TeacherName = aStudent.Class != null ? (aStudent.Class.Teacher != null ? aStudent.Class.Teacher.Name : "") : "";
                StudentSuperOBJ[aStudent.ID].ClassName = aStudent.Class != null ? aStudent.Class.Name : "";
                StudentSuperOBJ[aStudent.ID].SeatNo = aStudent.SeatNo.HasValue ? aStudent.SeatNo.Value.ToString() : "";
                StudentSuperOBJ[aStudent.ID].StudentNumber = aStudent.StudentNumber;

            }
            #endregion

            #region 取得獎懲資料(日期區間)
            List<DemeritRecord> DemeritList = new List<DemeritRecord>();

            if (obj.IsInsertDate) //發生日期
            {
                DemeritList = Demerit.SelectByOccurDate(allStudentID, obj.StartDate, obj.EndDate);
            }
            else //登錄入期
            {
                DemeritList = Demerit.SelectByRegisterDate(allStudentID, obj.StartDate, obj.EndDate);
            }

            if (DemeritList.Count == 0)
                e.Cancel = true; //沒有懲戒資料

            //依日期排序
            DemeritList.Sort(SortDateTime);

            foreach (DemeritRecord var in DemeritList)
            {
                string occurMonthDay = var.OccurDate.Month + "/" + var.OccurDate.Day;
                string reason = var.Reason;

                if (var.MeritFlag == "1") //1是獎勵
                {
                    #region 獎勵
                    //XmlElement meritElement = (XmlElement)var.SelectSingleNode("Detail/Discipline/Merit");
                    //if (meritElement == null) continue;

                    //bool comma = false;
                    //StringBuilder detailString = new StringBuilder("");
                    //detailString.Append(occurMonthDay + " ");
                    //if (!string.IsNullOrEmpty(reason))
                    //    detailString.Append(reason + " ");

                    //foreach (string merit in meritTable.Keys)
                    //{
                    //    int tryTimes;
                    //    int times = int.TryParse(meritElement.GetAttribute(meritTable[merit]), out tryTimes) ? tryTimes : 0;
                    //    if (times > 0)
                    //    {
                    //        if (!studentDiscipline[studentID].ContainsKey("Range" + merit))
                    //            studentDiscipline[studentID].Add("Range" + merit, 0);
                    //        studentDiscipline[studentID]["Range" + merit] += times;
                    //        if (comma)
                    //            detailString.Append(",");
                    //        detailString.Append(merit + times + "次");
                    //        comma = true;
                    //    }
                    //}

                    //studentDisciplineDetail[studentID].Add(detailString.ToString());

                    #endregion
                }
                else if (var.MeritFlag == "0")
                {
                    #region 懲戒
                    if (var.Cleared != "是")
                    {

                        //當MaxDemerit比DemeritSum大就離開
                        StudentSuperOBJ[var.RefStudentID].DemeritSum += GetSmallValueType(var);

                        DemStr ds = new DemStr();
                        ds._date = occurMonthDay;

                        StringBuilder detailString = new StringBuilder();
                        detailString.Append(occurMonthDay + " "); //日期

                        if (!string.IsNullOrEmpty(reason))
                        {
                            detailString.Append(reason + " "); //事由
                            ds._value += reason + "_";
                        }

                        if (var.DemeritA != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].DemeritA += var.DemeritA.Value;
                            ds._value += "大過：" + var.DemeritA.Value.ToString() + "_";
                        }
                        if (var.DemeritB != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].DemeritB += var.DemeritB.Value;
                            ds._value += "小過：" + var.DemeritB.Value.ToString() + "_";
                        }
                        if (var.DemeritC != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].DemeritC += var.DemeritC.Value;
                            ds._value += "警告：" + var.DemeritC.Value.ToString();
                        }

                        //明細資料
                        StudentSuperOBJ[var.RefStudentID].DemeritStringList.Add(ds);
                    }
                    #endregion
                }
            }
            #endregion

            #region 取得獎懲資料(學期累計)

            List<DemeritRecord> DemeritSchoolYearList = Demerit.SelectBySchoolYearAndSemester(allStudentID, int.Parse(School.DefaultSchoolYear), int.Parse(School.DefaultSemester));

            foreach (DemeritRecord record in DemeritSchoolYearList)
            {
                //1是大,0是小,-1是等於
                //用意是學期統計止於結束時間
                if (record.Cleared != "是" && record.OccurDate.CompareTo(obj.EndDate) != 1)
                {
                    StudentSuperOBJ[record.RefStudentID].DemeritSchoolA += record.DemeritA.HasValue ? record.DemeritA.Value : 0;
                    StudentSuperOBJ[record.RefStudentID].DemeritSchoolB += record.DemeritB.HasValue ? record.DemeritB.Value : 0;
                    StudentSuperOBJ[record.RefStudentID].DemeritSchoolC += record.DemeritC.HasValue ? record.DemeritC.Value : 0;
                }
            }
            #endregion

            #region 取得學生通訊地址資料
            foreach (AddressRecord record in Address.SelectByStudentIDs(allStudentID))
            {
                if (obj.ReceiveAddress == "戶籍地址")
                {
                    if (!string.IsNullOrEmpty(record.PermanentAddress))
                        StudentSuperOBJ[record.RefStudentID].address = record.Permanent.County + record.Permanent.Town + record.Permanent.District + record.Permanent.Area + record.Permanent.Detail;

                    if (!string.IsNullOrEmpty(record.PermanentZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.PermanentZipCode;

                        if (record.PermanentZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.PermanentZipCode.Substring(0, 1);
                        if (record.PermanentZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.PermanentZipCode.Substring(1, 1);
                        if (record.PermanentZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.PermanentZipCode.Substring(2, 1);
                        if (record.PermanentZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.PermanentZipCode.Substring(3, 1);
                        if (record.PermanentZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.PermanentZipCode.Substring(4, 1);
                    }

                }
                else if (obj.ReceiveAddress == "聯絡地址")
                {
                    if (!string.IsNullOrEmpty(record.MailingAddress))
                        StudentSuperOBJ[record.RefStudentID].address = record.Mailing.County + record.Mailing.Town + record.Mailing.District + record.Mailing.Area + record.Mailing.Detail; //再處理

                    if (!string.IsNullOrEmpty(record.MailingZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.MailingZipCode;

                        if (record.MailingZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.MailingZipCode.Substring(0, 1);
                        if (record.MailingZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.MailingZipCode.Substring(1, 1);
                        if (record.MailingZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.MailingZipCode.Substring(2, 1);
                        if (record.MailingZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.MailingZipCode.Substring(3, 1);
                        if (record.MailingZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.MailingZipCode.Substring(4, 1);
                    }
                }
                else if (obj.ReceiveAddress == "其他地址")
                {
                    if (!string.IsNullOrEmpty(record.Address1Address))
                        StudentSuperOBJ[record.RefStudentID].address = record.Address1.County + record.Address1.Town + record.Address1.District + record.Address1.Area + record.Address1.Detail; //再處理

                    if (!string.IsNullOrEmpty(record.Address1ZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.Address1ZipCode;

                        if (record.Address1ZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.Address1ZipCode.Substring(0, 1);
                        if (record.Address1ZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.Address1ZipCode.Substring(1, 1);
                        if (record.Address1ZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.Address1ZipCode.Substring(2, 1);
                        if (record.Address1ZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.Address1ZipCode.Substring(3, 1);
                        if (record.Address1ZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.Address1ZipCode.Substring(4, 1);
                    }
                }
            }
            #endregion

            #region 取得學生監護人父母親資料

            List<ParentRecord> ParentList = Parent.SelectByStudentIDs(allStudentID);

            foreach (ParentRecord record in ParentList)
            {
                StudentSuperOBJ[record.RefStudentID].CustodianName = record.CustodianName;
                StudentSuperOBJ[record.RefStudentID].FatherName = record.FatherName;
                StudentSuperOBJ[record.RefStudentID].MotherName = record.MotherName;
            }
            #endregion

            #region 產生報表

            Aspose.Words.Document template = new Aspose.Words.Document(obj.Template);
            template.MailMerge.Execute(
                new string[] { "學校名稱", "學校地址", "學校電話" },
                new object[] { School.ChineseName, School.Address, School.Telephone }
                );

            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.RemoveAllChildren();

            Aspose.Words.Node sectionNode = template.Sections[0].Clone();

            //取得學生人數
            totalStudentNumber = StudentSuperOBJ.Count;

            foreach (string student in StudentSuperOBJ.Keys)
            {
                //如果沒有學生就離開
                if (obj.PrintHasRecordOnly)
                {
                    if (StudentSuperOBJ.Count == 0)
                        continue;
                }

                //過濾不需要列印的學生
                if (StudentSuperOBJ[student].DemeritSum < MaxDemerit)
                    continue;

                if (StudentSuperOBJ[student].DemeritStringList.Count == 0)
                    continue;

                Aspose.Words.Document eachDoc = new Aspose.Words.Document();
                eachDoc.RemoveAllChildren();
                eachDoc.Sections.Add(eachDoc.ImportNode(sectionNode, true));

                //合併列印的資料
                Dictionary<string, object> mapping = new Dictionary<string, object>();

                StudentOBJ eachStudentInfo = StudentSuperOBJ[student];

                //學生資料
                mapping.Add("學生姓名", eachStudentInfo.student.Name);
                mapping.Add("班級", eachStudentInfo.ClassName);
                mapping.Add("座號", eachStudentInfo.SeatNo);
                mapping.Add("學號", eachStudentInfo.StudentNumber);
                mapping.Add("導師", eachStudentInfo.TeacherName);
                mapping.Add("資料期間", obj.StartDate.ToShortDateString() + " 至 " + obj.EndDate.ToShortDateString());

                //收件人資料
                if (obj.ReceiveName == "監護人姓名")
                    mapping.Add("收件人姓名", eachStudentInfo.CustodianName);
                else if (obj.ReceiveName == "父親姓名")
                    mapping.Add("收件人姓名", eachStudentInfo.FatherName);
                else if (obj.ReceiveName == "母親姓名")
                    mapping.Add("收件人姓名", eachStudentInfo.MotherName);
                else
                    mapping.Add("收件人姓名", eachStudentInfo.student.Name);

                //收件人地址資料
                mapping.Add("收件人地址", eachStudentInfo.address);
                mapping.Add("郵遞區號", eachStudentInfo.ZipCode);
                mapping.Add("0", eachStudentInfo.ZipCode1);
                mapping.Add("1", eachStudentInfo.ZipCode2);
                mapping.Add("2", eachStudentInfo.ZipCode3);
                mapping.Add("4", eachStudentInfo.ZipCode4);
                mapping.Add("5", eachStudentInfo.ZipCode5);

                mapping.Add("學年度", School.DefaultSchoolYear);
                mapping.Add("學期", School.DefaultSemester);

                //學生獎懲累計資料
                int count;
                //mapping.Add("學期累計大功", eachStudentDiscipline.TryGetValue("Semester大功", out count) ? "" + count : "0");
                //mapping.Add("學期累計小功", eachStudentDiscipline.TryGetValue("Semester小功", out count) ? "" + count : "0");
                //mapping.Add("學期累計嘉獎", eachStudentDiscipline.TryGetValue("Semester嘉獎", out count) ? "" + count : "0");
                mapping.Add("學期累計大過", eachStudentInfo.DemeritSchoolA);
                mapping.Add("學期累計小過", eachStudentInfo.DemeritSchoolB);
                mapping.Add("學期累計警告", eachStudentInfo.DemeritSchoolC);
                //mapping.Add("本期累計大功", eachStudentDiscipline.TryGetValue("Range大功", out count) ? "" + count : "0");
                //mapping.Add("本期累計小功", eachStudentDiscipline.TryGetValue("Range小功", out count) ? "" + count : "0");
                //mapping.Add("本期累計嘉獎", eachStudentDiscipline.TryGetValue("Range嘉獎", out count) ? "" + count : "0");
                mapping.Add("本期累計大過", eachStudentInfo.DemeritA);
                mapping.Add("本期累計小過", eachStudentInfo.DemeritB);
                mapping.Add("本期累計警告", eachStudentInfo.DemeritC);

                #region 附件

                MemoryStream accessoryMemory;
                Aspose.Words.Document accessoryDoc;

                //懲戒明細
                bool IsAccessory = false;
                //合併列印的資料
                Dictionary<string, object> mappingAccessory = new Dictionary<string, object>();
                mappingAccessory.Add("學年度", School.DefaultSchoolYear);
                mappingAccessory.Add("學期", School.DefaultSemester);
                //學生資料
                mappingAccessory.Add("學生姓名", eachStudentInfo.student.Name);
                mappingAccessory.Add("班級", eachStudentInfo.ClassName);
                mappingAccessory.Add("座號", eachStudentInfo.SeatNo);
                mappingAccessory.Add("學號", eachStudentInfo.StudentNumber);
                mappingAccessory.Add("導師", eachStudentInfo.TeacherName);
                mappingAccessory.Add("資料期間", obj.StartDate.ToShortDateString() + " 至 " + obj.EndDate.ToShortDateString());

                #endregion

                int demerit1 = 1;
                object[] objectValues = new object[] { StudentSuperOBJ[student].DemeritStringList };
                foreach (DemStr demerit in StudentSuperOBJ[student].DemeritStringList)
                {
                    if (demerit1 <= 10) //資料數大於10,透過附件列印
                    {

                        mapping.Add("日期" + demerit1, demerit._date);
                        mapping.Add("內容" + demerit1, demerit._value);
                        demerit1++;
                    }
                    else
                    {
                        IsAccessory = true;
                        mappingAccessory.Add("日期" + demerit1, demerit._date);
                        mappingAccessory.Add("內容" + demerit1, demerit._value);
                        demerit1++;
                    }
                }

                string[] keys = new string[mapping.Count];
                object[] values = new object[mapping.Count];
                int i = 0;
                foreach (string key in mapping.Keys)
                {
                    keys[i] = key;
                    values[i++] = mapping[key];
                }

                eachDoc.MailMerge.CleanupOptions = Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveEmptyParagraphs;
                //eachDoc.MailMerge.FieldMergingCallback = new HandleMergeImageFieldFromBlob();
                eachDoc.MailMerge.Execute(keys, values);
                eachDoc.MailMerge.DeleteFields(); //刪除未合併之內容


                if (IsAccessory)
                {
                    #region 附件
                    accessoryMemory = new MemoryStream(Properties.Resources.懲戒通知單_附件一);
                    accessoryDoc = new Aspose.Words.Document(accessoryMemory);

                    string[] keysAccessory = new string[mappingAccessory.Count];
                    object[] valuesAccessory = new object[mappingAccessory.Count];
                    int xx = 0;
                    foreach (string key in mappingAccessory.Keys)
                    {
                        keysAccessory[xx] = key;
                        valuesAccessory[xx++] = mappingAccessory[key];
                    }

                    accessoryDoc.MailMerge.CleanupOptions = Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveEmptyParagraphs;
                    accessoryDoc.MailMerge.Execute(keysAccessory, valuesAccessory);
                    accessoryDoc.MailMerge.DeleteFields(); //刪除未合併之內容 

                    Aspose.Words.Node eachSectionaccessory = accessoryDoc.Sections[0].Clone();
                    eachDoc.Sections.Add(eachDoc.ImportNode(eachSectionaccessory, true));
                    #endregion
                }



                //加入文件
                foreach (Aspose.Words.Section each in eachDoc.Sections)
                {
                    Aspose.Words.Node eachSectionNode = each.Clone();
                    doc.Sections.Add(doc.ImportNode(eachSectionNode, true));
                }

                //回報進度
                _BGWDisciplineNotification.ReportProgress((int)(((double)currentStudentCount++ * 100.0) / (double)totalStudentNumber));
            }

            #endregion

            #region 產生學生清單

            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
            if (obj.PrintStudentList)
            {
                int CountRow = 0;
                wb.Worksheets[0].Cells[CountRow, 0].PutValue("班級");
                wb.Worksheets[0].Cells[CountRow, 1].PutValue("座號");
                wb.Worksheets[0].Cells[CountRow, 2].PutValue("學號");
                wb.Worksheets[0].Cells[CountRow, 3].PutValue("學生姓名");
                wb.Worksheets[0].Cells[CountRow, 4].PutValue("收件人姓名");
                wb.Worksheets[0].Cells[CountRow, 5].PutValue("地址");
                CountRow++;
                foreach (string each in StudentSuperOBJ.Keys)
                {
                    //如果沒有學生就離開
                    if (obj.PrintHasRecordOnly)
                    {
                        if (StudentSuperOBJ.Count == 0)
                            continue;
                    }

                    //過濾不需要列印的學生
                    if (StudentSuperOBJ[each].DemeritSum < MaxDemerit)
                        continue;

                    if (StudentSuperOBJ[each].DemeritStringList.Count == 0)
                        continue;

                    wb.Worksheets[0].Cells[CountRow, 0].PutValue(StudentSuperOBJ[each].ClassName);
                    wb.Worksheets[0].Cells[CountRow, 1].PutValue(StudentSuperOBJ[each].SeatNo);
                    wb.Worksheets[0].Cells[CountRow, 2].PutValue(StudentSuperOBJ[each].StudentNumber);
                    wb.Worksheets[0].Cells[CountRow, 3].PutValue(StudentSuperOBJ[each].student.Name);
                    //收件人資料
                    if (obj.ReceiveName == "監護人姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].CustodianName);
                    else if (obj.ReceiveName == "父親姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].FatherName);
                    else if (obj.ReceiveName == "母親姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].MotherName);
                    else
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].student.Name);

                    wb.Worksheets[0].Cells[CountRow, 5].PutValue(StudentSuperOBJ[each].ZipCode + " " + StudentSuperOBJ[each].address);
                    CountRow++;
                }
                wb.Worksheets[0].AutoFitColumns();
            }
            #endregion


            string path = Path.Combine(Application.StartupPath, "Reports");
            string path2 = Path.Combine(Application.StartupPath, "Reports");

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");
            path2 = Path.Combine(path2, reportName + "(學生清單).xlsx");
            e.Result = new object[] { reportName, path, doc, path2, obj.PrintStudentList, wb };
        }

        private int SortDateTime(DemeritRecord x, DemeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }

        class MDFilter
        {
            #region 功過條件過濾
            private int _condName = 0; //0:A, 1:B, 2:C
            private int _condNumber = 0;

            public MDFilter()
            {
            }

            public void SetCondition(int name, int number)
            {
                _condName = name;
                _condNumber = number;
            }

            public bool IsFilter(int A, int B, int C)
            {
                bool filtered = false;

                switch (_condName)
                {
                    case 5:
                    case 2:
                        if ((A + B) > 0)
                            filtered = false;
                        else if (C >= _condNumber)
                            filtered = false;
                        else
                            filtered = true;
                        break;
                    case 4:
                    case 1:
                        if (A > 0)
                            filtered = false;
                        else if (B >= _condNumber)
                            filtered = false;
                        else
                            filtered = true;
                        break;
                    case 3:
                    case 0:
                        if (A >= _condNumber)
                            filtered = false;
                        else
                            filtered = true;
                        break;
                    default:
                        break;
                }

                return filtered;
            }
            #endregion
        }
    }

    class DemStr
    {
        public string _date { get; set; }
        public string _value { get; set; }
    }

    //public class HandleMergeImageFieldFromBlob : IFieldMergingCallback
    //{
    //    void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
    //    {
    //        #region MailMerge_MergeField
    //        if (args.FieldName == "懲戒明細")
    //        {
    //            object[] objectValues = (object[])args.FieldValue;
    //            List<string> eachStudentDisciplineDetail = (List<string>)objectValues[0];

    //            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(args.Document);

    //            builder.MoveToField(args.Field, false);
    //            builder.StartTable();
    //            builder.CellFormat.ClearFormatting();
    //            builder.CellFormat.Borders.ClearFormatting();
    //            builder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;
    //            builder.CellFormat.LeftPadding = 3.0;
    //            builder.ParagraphFormat.LeftIndent = 0.0;
    //            builder.RowFormat.Height = 15.0;

    //            int rowNumber = 5; //共5行
    //            //當懲戒資料數 大於 預設提供的資料格時
    //            if (eachStudentDisciplineDetail.Count > rowNumber * 2)
    //            {
    //                rowNumber = eachStudentDisciplineDetail.Count / 2; //資料數除2 10
    //                rowNumber += eachStudentDisciplineDetail.Count % 2; //取餘數 1
    //            }

    //            //當懲戒資料數(還是) 大於 預設提供的資料格時
    //            if (eachStudentDisciplineDetail.Count > rowNumber * 2)
    //            {
    //                rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) / 2;
    //                rowNumber += (eachStudentDisciplineDetail.Count - (rowNumber * 2)) % 2;
    //            }

    //            //取得最後資料行數時
    //            for (int j = 0; j < rowNumber; j++)
    //            {
    //                builder.InsertCell();
    //                builder.CellFormat.Borders.Right.LineStyle = Aspose.Words.LineStyle.Single;
    //                builder.CellFormat.Borders.Right.Color = Color.Black;

    //                if (j < eachStudentDisciplineDetail.Count)
    //                {
    //                    builder.Write(eachStudentDisciplineDetail[j]);
    //                }
    //                builder.InsertCell();
    //                builder.CellFormat.Borders.Right.LineStyle = Aspose.Words.LineStyle.Single;
    //                builder.CellFormat.Borders.Right.Color = Color.Black;

    //                if (j + rowNumber < eachStudentDisciplineDetail.Count)
    //                {
    //                    builder.Write(eachStudentDisciplineDetail[j + rowNumber]);
    //                }
    //                builder.EndRow();
    //            }
    //            builder.EndTable();
    //            args.Text = string.Empty;
    //        }
    //        #endregion
    //    }
    //    public void ImageFieldMerging(ImageFieldMergingArgs args)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}
}

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
using Aspose.Words;
using SmartSchool.ePaper;
using System.Linq;

namespace 獎懲通知單
{
    internal class Report : IReport
    {

        private int DemeritAB; // 1 大過 等於 3 小過
        private int DemeritBC; // 1 小過 等於 3 警告
        private int MaxDemerit; //最小獎懲單位值
        private int MeritAB; // 1 大功 等於 3 小功
        private int MeritBC; // 1 小功 等於 3 嘉獎
        private int MaxMerit; //最小獎勵單位值

        private BackgroundWorker _BGWDisciplineNotification;

        private ConfigOBJ obj; //所有列印設定資訊

        private List<StudentRecord> SelectedStudents { get; set; }

        string entityName;

        /// <summary>
        /// 學生電子報表
        /// </summary>
        SmartSchool.ePaper.ElectronicPaper paperForStudent { get; set; }

        public Report(string _entityName)
        {
            entityName = _entityName;
        }

        public void Print()
        {
            #region IReport 成員
            MeritDemeritDateRangeForm form = new MeritDemeritDateRangeForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化獎懲通知單...");

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
                obj.PaperUpdate = form._cbPaper; //是否列印電子報表
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
                MsgBox.Show("懲戒對照表有誤");
                return;
            }

            jb = helper.GetText("Demerit/BC");
            if (!int.TryParse(jb, out DemeritBC))
            {
                MsgBox.Show("懲戒對照表有誤");
                return;
            }

            jb = helper.GetText("Merit/AB");
            if (!int.TryParse(jb, out MeritAB))
            {
                MsgBox.Show("獎勵對照表有誤");
                return;
            }

            jb = helper.GetText("Merit/BC");
            if (!int.TryParse(jb, out MeritBC))
            {
                MsgBox.Show("獎勵對照表有誤");
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

            MaxMerit = 0; //最小單位值

            if (CondName == "大功")
            {
                MaxMerit = CondNumber * MeritAB;
                MaxMerit = MaxMerit * MeritBC;
            }
            else if (CondName == "小功")
            {
                MaxMerit = CondNumber * MeritBC;
            }
            else if (CondName == "嘉獎")
            {
                MaxMerit = CondNumber;
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


        private int GetSmallValueTypeM(MeritRecord record)
        {
            #region 將資料換算為最小單位值
            int AA = record.MeritA.HasValue ? record.MeritA.Value : 0;
            int BB = record.MeritB.HasValue ? record.MeritB.Value : 0;
            int CC = record.MeritC.HasValue ? record.MeritC.Value : 0;

            int SumNum1 = 0;
            int SumNum2 = 0;
            int SumNum3 = 0;

            if (AA != 0)
            {
                SumNum1 = AA * MeritAB; //轉換成小功
                SumNum1 = SumNum1 * MeritBC; //轉換成嘉獎
            }

            if (BB != 0)
            {
                SumNum2 = BB * MeritBC; //轉換成嘉獎
            }

            if (CC != 0)
            {
                SumNum3 = CC; //轉換成嘉獎
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

            string reportName = "獎懲通知單";

            //取得換算單位
            ChengeDemerit(obj.ConditionName, int.Parse(obj.ConditionNumber));

            Dictionary<string, int> MDMapping = new Dictionary<string, int>();
            MDMapping.Add("大功", 3);
            MDMapping.Add("小功", 4);
            MDMapping.Add("嘉獎", 5);
            MDMapping.Add("大過", 3);
            MDMapping.Add("小過", 4);
            MDMapping.Add("警告", 5);

            int flag = 2; //預設是2
            if (!string.IsNullOrEmpty(obj.ConditionName)) //如果condName不是空的(可能是大功~警告)
            {
                flag = 0; //小於3是獎勵,大於3是獎懲
            }

            MDFilter filter = new MDFilter();
            if (flag == 0)
                filter.SetCondition(MDMapping[obj.ConditionName], int.Parse(obj.ConditionNumber));

            #endregion

            #region 快取資訊


            //超級資訊物件
            Dictionary<string, StudentOBJ> StudentSuperOBJ = new Dictionary<string, StudentOBJ>();
            //所有學生ID
            List<string> allStudentID = new List<string>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

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
            List<MeritRecord> MeritList = new List<MeritRecord>();

            if (obj.IsInsertDate) //發生日期
            {
                DemeritList = Demerit.SelectByOccurDate(allStudentID, obj.StartDate, obj.EndDate);
                MeritList = Merit.SelectByOccurDate(allStudentID, obj.StartDate, obj.EndDate);
            }
            else //登錄入期
            {
                DemeritList = Demerit.SelectByRegisterDate(allStudentID, obj.StartDate, obj.EndDate);
                MeritList = Merit.SelectByRegisterDate(allStudentID, obj.StartDate, obj.EndDate);
            }

            if (DemeritList.Count == 0 && MeritList.Count ==0)
                e.Cancel = true; //沒有獎懲資料

            //依日期排序
            DemeritList.Sort(SortDateTime);
            MeritList.Sort(SortDateTime);


            // 懲戒
            foreach (DemeritRecord var in DemeritList)
            {
                string occurMonthDay = var.OccurDate.Month + "/" + var.OccurDate.Day;
                string reason = var.Reason;

                if (var.MeritFlag == "0")
                {
                    #region 懲戒
                    if (var.Cleared != "是")
                    {

                        //當MaxDemerit比DemeritSum大就離開
                        StudentSuperOBJ[var.RefStudentID].DemeritSum += GetSmallValueType(var);

                        DemStr ds = new DemStr();
                        ds._date = occurMonthDay;

                        if (!string.IsNullOrEmpty(reason))
                        {
                            if (reason.Length > 35)
                            {
                                string newReason = reason.Remove(35);
                                ds._value += newReason + " (...已簡略)" + "_";
                                StudentSuperOBJ[var.RefStudentID].IsNewReason = true;
                            }
                            else
                            {
                                ds._value += reason + "_";
                            }
                        }

                        if (var.DemeritA != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].DemeritA += var.DemeritA.Value;
                            ds._value += "大過「" + var.DemeritA.Value.ToString() + "」";
                        }
                        if (var.DemeritB != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].DemeritB += var.DemeritB.Value;
                            ds._value += "小過「" + var.DemeritB.Value.ToString() + "」";
                        }
                        if (var.DemeritC != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].DemeritC += var.DemeritC.Value;
                            ds._value += "警告「" + var.DemeritC.Value.ToString() + "」";
                        }

                        //明細資料
                        StudentSuperOBJ[var.RefStudentID].DemeritStringList.Add(ds);
                    }
                    #endregion
                }
            }

            // 獎勵
            foreach (MeritRecord var in MeritList)
            {
                string occurMonthDay = var.OccurDate.Month + "/" + var.OccurDate.Day;
                string reason = var.Reason;

             if (var.MeritFlag == "1")
                {
                    #region 獎勵
                
                        //當MaxMerit比MeritSum大就離開
                        StudentSuperOBJ[var.RefStudentID].MeritSum += GetSmallValueTypeM(var);

                        DemStr ds = new DemStr();
                        ds._date = occurMonthDay;

                        if (!string.IsNullOrEmpty(reason))
                        {
                            if (reason.Length > 35)
                            {
                                string newReason = reason.Remove(35);
                                ds._value += newReason + " (...已簡略)" + "_";
                                StudentSuperOBJ[var.RefStudentID].IsNewReason = true;
                            }
                            else
                            {
                                ds._value += reason + "_";
                            }
                        }

                        if (var.MeritA != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].MeritA += var.MeritA.Value;
                            ds._value += "大功「" + var.MeritA.Value.ToString() + "」";
                        }
                        if (var.MeritB != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].MeritB += var.MeritB.Value;
                            ds._value += "小功「" + var.MeritB.Value.ToString() + "」";
                        }
                        if (var.MeritC != 0)
                        {
                            StudentSuperOBJ[var.RefStudentID].MeritC += var.MeritC.Value;
                            ds._value += "嘉獎「" + var.MeritC.Value.ToString() + "」";
                        }

                        //明細資料
                        StudentSuperOBJ[var.RefStudentID].MeritStringList.Add(ds);
                   
                    #endregion
                }
            }

            #endregion

            #region 取得獎懲資料(學期累計)

            List<DemeritRecord> DemeritSchoolYearList = Demerit.SelectBySchoolYearAndSemester(allStudentID, int.Parse(School.DefaultSchoolYear), int.Parse(School.DefaultSemester));
            List<MeritRecord> MeritSchoolYearList=Merit.SelectBySchoolYearAndSemester(allStudentID, int.Parse(School.DefaultSchoolYear), int.Parse(School.DefaultSemester));
            // 懲戒
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

            // 獎勵
            foreach (MeritRecord record in MeritSchoolYearList)
            {
                //1是大,0是小,-1是等於
                if (record.OccurDate.CompareTo(obj.EndDate) != 1)
                {
                    StudentSuperOBJ[record.RefStudentID].MeritSchoolA += record.MeritA.HasValue ? record.MeritA.Value : 0;
                    StudentSuperOBJ[record.RefStudentID].MeritSchoolB += record.MeritB.HasValue ? record.MeritB.Value : 0;
                    StudentSuperOBJ[record.RefStudentID].MeritSchoolC += record.MeritC.HasValue ? record.MeritC.Value : 0;
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
            paperForStudent = new SmartSchool.ePaper.ElectronicPaper("獎懲通知單_" + DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0'), School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);

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
                mapping.Add("學期累計大過", eachStudentInfo.DemeritSchoolA);
                mapping.Add("學期累計小過", eachStudentInfo.DemeritSchoolB);
                mapping.Add("學期累計警告", eachStudentInfo.DemeritSchoolC);
                mapping.Add("本期累計大過", eachStudentInfo.DemeritA);
                mapping.Add("本期累計小過", eachStudentInfo.DemeritB);
                mapping.Add("本期累計警告", eachStudentInfo.DemeritC);

                mapping.Add("學期累計大功", eachStudentInfo.MeritSchoolA);
                mapping.Add("學期累計小功", eachStudentInfo.MeritSchoolB);
                mapping.Add("學期累計嘉獎", eachStudentInfo.MeritSchoolC);
                mapping.Add("本期累計大功", eachStudentInfo.MeritA);
                mapping.Add("本期累計小功", eachStudentInfo.MeritB);
                mapping.Add("本期累計嘉獎", eachStudentInfo.MeritC);


                if (eachStudentInfo.IsNewReason)
                {
                    mapping.Add("註1", "註1:獎懲事由,如標記 (…已簡略) 表示事由內容過多,請至Web2查閱詳細內容");
                }

                #region 附件

                MemoryStream accessoryMemory;
                Aspose.Words.Document accessoryDoc;

                //獎懲附件1
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

                int Demerit1 = 1;
                //object[] objectValues = new object[] { StudentSuperOBJ[student].DemeritStringList };

                // 獎懲排序
                List<DemStr> DemStrSortList = new List<DemStr>();

                foreach (DemStr demerit in StudentSuperOBJ[student].DemeritStringList)
                    DemStrSortList.Add(demerit);

                foreach (DemStr demerit in StudentSuperOBJ[student].MeritStringList)
                    DemStrSortList.Add(demerit);

                try
                {
                    // 依日期排序
                    DemStrSortList = (from data in DemStrSortList orderby DateTime.Parse(data._date) ascending select data).ToList();
                }
                catch (Exception ex) { }

                foreach (DemStr demerit in DemStrSortList)
                {
                    if (Demerit1 <= 10) //資料數大於10,透過附件列印
                    {

                        mapping.Add("日期" + Demerit1, demerit._date);
                        mapping.Add("內容" + Demerit1, demerit._value);
                        Demerit1++;
                    }
                    else
                    {
                        IsAccessory = true;
                        mappingAccessory.Add("日期" + Demerit1, demerit._date);
                        mappingAccessory.Add("內容" + Demerit1, demerit._value);
                        Demerit1++;
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

                //如果要列印附件一
                if (IsAccessory)
                {
                    #region 附件
                    accessoryMemory = new MemoryStream(Properties.Resources.獎懲通知單_附件一);
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

                    MemoryStream stream = new MemoryStream();
                    eachDoc.Save(stream, SaveFormat.Doc);
                    paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, eachStudentInfo.student.ID));
                    #endregion
                }
                else
                {
                    MemoryStream stream = new MemoryStream();
                    eachDoc.Save(stream, SaveFormat.Doc);
                    paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, eachStudentInfo.student.ID));
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
                    if (StudentSuperOBJ[each].DemeritSum < MaxDemerit && StudentSuperOBJ[each].MeritSum < MaxMerit)
                        continue;

                    if (StudentSuperOBJ[each].DemeritStringList.Count == 0 && StudentSuperOBJ[each].MeritStringList.Count == 0)
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

            if (obj.PaperUpdate)
            {
                SmartSchool.ePaper.DispatcherProvider.Dispatch(paperForStudent);
            }

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

        private int SortDateTime(MeritRecord x, MeritRecord y)
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

}

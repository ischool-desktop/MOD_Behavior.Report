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
using Aspose.Words.Tables;
using Aspose.Words;
using SmartSchool.ePaper;

namespace K12.銷過通知單2015
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
            ClearDemeritDateRangeForm form = new ClearDemeritDateRangeForm();

            if (form.ShowDialog() == DialogResult.OK)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化銷過通知單...");

                #region 建立設定檔
                obj = new ConfigOBJ();
                obj.StartDate = form.StartDate;
                obj.EndDate = form.EndDate;
                obj.PrintHasRecordOnly = form.PrintHasRecordOnly;
                obj.Template = form.Template;
                obj.ReceiveName = form.ReceiveName;
                obj.ReceiveAddress = form.ReceiveAddress;
                obj.IsConditions = "銷過日期";
                obj.PrintStudentList = form.PrintStudentList;
                obj.PaperUpdate = form._cbPaper; //是否列印電子報表

                if (form.radioButton2.Checked)
                {
                    obj.IsConditions = "發生日期";
                }
                else if (form.radioButton3.Checked)
                {
                    obj.IsConditions = "登錄日期";
                }
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

            string reportName = "銷過通知單";

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

            //加總用
            Dictionary<string, int> StudMeritSum = new Dictionary<string, int>();

            foreach (StudentRecord aStudent in SelectedStudents)
            {
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

            if (obj.IsConditions == "銷過日期") //發生日期
            {
                DemeritList = Demerit.Select(allStudentID, null, null, null, null, obj.StartDate, obj.EndDate, null, null, null);
            }
            else if (obj.IsConditions == "發生日期")
            {
                DemeritList = Demerit.SelectByOccurDate(allStudentID, obj.StartDate, obj.EndDate);
            }
            else if (obj.IsConditions == "登錄日期")//登錄入期
            {
                DemeritList = Demerit.SelectByRegisterDate(allStudentID, obj.StartDate, obj.EndDate);
            }

            //依日期排序
            DemeritList.Sort(SortDateTime);

            foreach (DemeritRecord var in DemeritList)
            {
                string occurMonthDay = var.OccurDate.ToShortDateString();
                string reason = var.Reason;

                #region 懲戒
                if (var.Cleared == "是")
                {
                    ClearDemeritDetail cdd = new ClearDemeritDetail();


                    cdd.登錄日期 = occurMonthDay; //日期
                    cdd.懲戒事由 = reason; //事由
                    cdd.銷過事由 = var.ClearReason;
                    if (var.ClearDate.HasValue)
                    {
                        cdd.銷過日期 = var.ClearDate.Value.ToShortDateString();
                    }

                    StringBuilder detailString = new StringBuilder();
                    if (var.DemeritA != 0)
                    {
                        StudentSuperOBJ[var.RefStudentID].DemeritA += var.DemeritA.Value;
                        detailString.Append("大過：" + var.DemeritA.Value.ToString() + " ");
                    }
                    if (var.DemeritB != 0)
                    {
                        StudentSuperOBJ[var.RefStudentID].DemeritB += var.DemeritB.Value;
                        detailString.Append("小過：" + var.DemeritB.Value.ToString() + " ");
                    }
                    if (var.DemeritC != 0)
                    {
                        StudentSuperOBJ[var.RefStudentID].DemeritC += var.DemeritC.Value;
                        detailString.Append("警告：" + var.DemeritC.Value.ToString() + " ");
                    }
                    cdd.懲戒資料 = detailString.ToString();

                    StudentSuperOBJ[var.RefStudentID].DemeritStringList.Add(cdd);
                }
                #endregion

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
            paperForStudent = new SmartSchool.ePaper.ElectronicPaper("銷過通知單_" + DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0'), School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);

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

                #region 附件

                MemoryStream accessoryMemory;
                Aspose.Words.Document accessoryDoc;

                //懲戒附件1
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
                foreach (ClearDemeritDetail each in StudentSuperOBJ[student].DemeritStringList)
                {
                    if (Demerit1 <= 5) //資料數大於10,透過附件列印
                    {
                        mapping.Add("懲戒日期" + Demerit1, each.登錄日期);
                        mapping.Add("懲戒內容" + Demerit1, each.懲戒資料);
                        mapping.Add("懲戒事由" + Demerit1, each.懲戒事由);
                        mapping.Add("銷過日期" + Demerit1, each.銷過日期);
                        mapping.Add("銷過事由" + Demerit1, each.銷過事由);
                        Demerit1++;
                    }
                    else
                    {
                        IsAccessory = true;
                        mappingAccessory.Add("懲戒日期" + Demerit1, each.登錄日期);
                        mappingAccessory.Add("懲戒內容" + Demerit1, each.懲戒資料);
                        mappingAccessory.Add("懲戒事由" + Demerit1, each.懲戒事由);
                        mappingAccessory.Add("銷過日期" + Demerit1, each.銷過日期);
                        mappingAccessory.Add("銷過事由" + Demerit1, each.銷過事由);
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

                //合併列印 - 問題點
                eachDoc.MailMerge.CleanupOptions = Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveEmptyParagraphs; 
                eachDoc.MailMerge.Execute(keys, values);
                eachDoc.MailMerge.DeleteFields();

                //如果要列印附件一
                if (IsAccessory)
                {
                    #region 附件
                    accessoryMemory = new MemoryStream(Properties.Resources.銷過通知單_附件一);
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
    }
}

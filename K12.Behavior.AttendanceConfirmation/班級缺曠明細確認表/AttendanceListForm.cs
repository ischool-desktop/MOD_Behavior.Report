using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.Xml;
using Aspose.Words;
using FISCA.DSAUtil;
using System.IO;
using K12.Data;
using Aspose.Words.Reporting;

namespace K12.Behavior.AttendanceConfirmation
{
    public partial class AttendanceListForm : BaseForm
    {
        private int _sizeIndex = 0; //預設為0
        private BackgroundWorker _BGWClassStudentAbsenceDetail; //背景模式
        private Dictionary<string, List<string>> Absence = new Dictionary<string, List<string>>();

        private Document _doc;
        private GetAbsenceDetail Data;
        private List<string> periodList;
        private Dictionary<string, int> ColumnIndex;

        private AttendanceConfigData GetCD;

        public const string ConfigName = "班級缺曠明細確認表_Word";

        public AttendanceListForm()
        {
            InitializeComponent();

            dateTimeInput1.Value = DateTime.Today.AddMonths(-7);
            dateTimeInput2.Value = DateTime.Today;

            GetCD = new AttendanceConfigData();

            //預設為3日後
            dateTimeInput3.Value = DateTime.Today.AddDays(3);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (K12.Presentation.NLDPanels.Class.SelectedSource.Count == 0)
            {
                return;
            }

            this.btnSave.Enabled = false;

            //LoadPreference();

            GetCD = new AttendanceConfigData();
            GetCD.SaveAll();
            GetCD.Reset();


            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化班級學生缺曠明細(確認表)...");

            //日期起始,結束,紙張大小,列印假別,是否列印無資料班級
            object[] args = new object[] { dateTimeInput1.Value, dateTimeInput2.Value, _sizeIndex, Absence, GetCD.ClassNoData };

            _BGWClassStudentAbsenceDetail = new BackgroundWorker();
            _BGWClassStudentAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentAbsenceDetail_DoWork);
            _BGWClassStudentAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWClassStudentAbsenceDetail_RunWorkerCompleted);
            _BGWClassStudentAbsenceDetail.RunWorkerAsync(args);
        }


        void _BGWClassStudentAbsenceDetail_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.btnSave.Enabled = true;

            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 班級缺曠記錄明細(確認表) 已完成");

                    SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = "班級缺曠記錄明細(確認表).doc";
                    sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            _doc.Save(sd.FileName);
                            System.Diagnostics.Process.Start(sd.FileName);

                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            this.Enabled = true;
                            return;
                        }
                    }
                }
                else
                {

                    MsgBox.Show("列印發生錯誤:\n" + e.Error.Message);
                }
            }
            else
            {
                MsgBox.Show("報表列印作業已中止");
            }
        }

        //報表設定
        private void lblPrintSetup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            AttendanceSetup AS = new AttendanceSetup(GetCD);
            AS.ShowDialog();

            GetCD.Reset();
        }

        //節次設定
        private void lblAttendance_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SetupForm SF = new SetupForm(GetCD);
            SF.ShowDialog();

            GetCD.Reset();
        }

        //離開
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void _BGWClassStudentAbsenceDetail_DoWork(object sender, DoWorkEventArgs e)
        {
            //string reportName = "班級缺曠記錄明細";

            object[] args = e.Argument as object[];

            DateTime startDate = (DateTime)args[0];
            DateTime endDate = (DateTime)args[1];
            int size = (int)args[2];
            Dictionary<string, List<string>> absence = (Dictionary<string, List<string>>)args[3];
            string ClassNoData = (string)args[4];

            #region 快取資料

            //節次對照表
            periodList = tool.GetPeriod();

            //取得學生資料
            Data = new GetAbsenceDetail(startDate, endDate, periodList);

            #endregion

            Document template;

            if (GetCD.Setup_Mode == "false")
            {
                template = new Document(new MemoryStream(Properties.Resources.班級缺曠明細確認表範本));
            }
            else
            {
                if (GetCD.Setup_Temp == "")
                {
                    template = new Document(new MemoryStream(Properties.Resources.班級缺曠明細確認表範本));
                }
                else
                {
                    template = new Document(new MemoryStream(GetCD.Temp));
                }
            }

            //template.Sections[0].Body.Tables[0].Rows[2].Cells[0].Paragraphs[0].Runs[0].Font

            tool.FontSize = template.Sections[0].Body.Tables[0].Rows[1].Cells[0].Paragraphs[0].Runs[0].Font.Size;
            tool.FontName = template.Sections[0].Body.Tables[0].Rows[1].Cells[0].Paragraphs[0].Runs[0].Font.Name;

            tool._run = new Run(template);

            DocumentBuilder DB = new DocumentBuilder(template);

            ColumnIndex = new Dictionary<string, int>();
            ColumnIndex.Add("座號", 0);
            ColumnIndex.Add("姓名", 1);
            ColumnIndex.Add("日期", 2);

            DB.MoveToMergeField("缺曠假別");
            Cell AbsenceCell = DB.CurrentParagraph.ParentNode as Cell;


            tool.CellSplit(AbsenceCell, periodList.Count);
            int x = 3;
            foreach (string periodEace in periodList)
            {
                ColumnIndex.Add(periodEace, x);
                tool.Write(AbsenceCell, periodEace);
                AbsenceCell = tool.GetMoveRightCell(AbsenceCell, 1);
                x++;
            }

            _doc = new Document();
            _doc.Sections.Clear();

            foreach (ClassRecord classInfo in Data.selectedClass)
            {
                #region 如果該班級的內容為0,就離開
                if (ClassNoData == "True")
                {
                    int PageContinue = 0;
                    foreach (string test1 in Data.allAbsenceDetail[classInfo.ID].Keys) //班級
                    {
                        foreach (string test2 in Data.allAbsenceDetail[classInfo.ID][test1].Keys) //學生
                        {
                            PageContinue += Data.allAbsenceDetail[classInfo.ID][test1][test2].Count;
                        }
                    }

                    if (PageContinue == 0)
                        continue;
                }
                #endregion

                Document resultdoc = template.Clone(true) as Document;

                tool._run = new Run(resultdoc);

                #region MailMerge
                List<string> name = new List<string>();
                List<object> value = new List<object>();

                name.Add("班級");
                value.Add(classInfo.Name);

                name.Add("日期區間");
                value.Add(startDate.ToShortDateString() + "~" + endDate.ToShortDateString());

                name.Add("製表日期");
                value.Add(DateTime.Now.ToShortDateString());

                if (dateTimeInput3.Value != DateTime.MinValue)
                {
                    name.Add("繳回日期");
                    value.Add(tool.HowManyWeek(dateTimeInput3.Value));
                }
                else
                {
                    name.Add("繳回日期");
                    value.Add("　　　年　　　月　　　日");
                }


                ClassSpeRecord csr = new ClassSpeRecord(Data.allAbsenceDetail[classInfo.ID]);
                csr.PeriodCount = periodList.Count;
                csr.ColumnIndex = ColumnIndex;


                name.Add("Data");
                value.Add(csr);
                resultdoc.MailMerge.MergeField += MailMerge_MergeField;
                resultdoc.MailMerge.Execute(name.ToArray(), value.ToArray());

                resultdoc.MailMerge.DeleteFields();

                _doc.Sections.Add(_doc.ImportNode(resultdoc.FirstSection, true));
                #endregion

            }

        }

        void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            if (args.FieldName == "Data")
            {
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToField(args.Field, true);
                args.Field.Remove();

                ClassSpeRecord Csr = args.FieldValue as ClassSpeRecord;

                //插入於此Row之下
                Row refrow = builder.CurrentParagraph.ParentNode.ParentNode as Row;
                Cell SplieCell = tool.GetMoveRightCell(refrow.Cells[0], 3);
                tool.CellSplit(SplieCell, Csr.PeriodCount);

                //範本
                Row rowtemp = builder.CurrentParagraph.ParentNode.ParentNode.Clone(true) as Row;

                //此表格
                Table table = builder.CurrentParagraph.ParentNode.ParentNode.ParentNode as Table;

                foreach (StudentRecord stud in Csr.StudentRecordDic.Values) //取得學生ID
                {

                    #region 排一下日期

                    List<string> DatList = new List<string>();

                    foreach (string each2 in Csr._dic[stud.ID].Keys) //日期
                    {
                        DatList.Add(each2);
                    }

                    DatList.Sort(tool.sortdat);

                    #endregion


                    foreach (string each2 in DatList) //日期
                    {
                        //如果該日期為0節資料,就不列印(因為被設定檔過慮掉)
                        if (Csr._dic[stud.ID][each2].Count == 0)
                            continue;

                        tool.Write(refrow.Cells[Csr.ColumnIndex["座號"]], stud.SeatNo.HasValue ? stud.SeatNo.Value.ToString() : ""); //座號
                        tool.Write(refrow.Cells[Csr.ColumnIndex["姓名"]], stud.Name);   //姓名
                        tool.Write(refrow.Cells[Csr.ColumnIndex["日期"]], each2);   //日期

                        foreach (string each3 in Csr._dic[stud.ID][each2].Keys) //節次
                        {
                            if (Csr.ColumnIndex.ContainsKey(each3))
                            {
                                tool.Write(refrow.Cells[Csr.ColumnIndex[each3]], Csr._dic[stud.ID][each2][each3]);
                            }

                        }
                        refrow = table.InsertAfter(rowtemp.Clone(true), refrow) as Row;
                    }
                }

                refrow.Remove();
            }
        }
    }

}

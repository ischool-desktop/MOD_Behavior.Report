using Aspose.Words;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace K12.Behavior.MeritDemeritConfirmation
{
    public partial class MeritDemeritListForm : BaseForm
    {
        private int _sizeIndex = 0; //預設為0

        private BackgroundWorker _BGWClassStudentMeritDetail; //背景模式

        private GetMeritDetail Data;

        private MeritConfigData GetCD;

        private Dictionary<string, int> ColumnIndex;

        private Document _doc;

        public MeritDemeritListForm()
        {
            InitializeComponent();

            dateTimeInput1.Value = DateTime.Today.AddDays(-7);
            dateTimeInput2.Value = DateTime.Today;

            GetCD = new MeritConfigData();

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

            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化班級學生獎懲明細(確認表)...");

            //日期起始,結束,紙張大小,列印假別,是否列印無資料班級
            object[] args = new object[] { dateTimeInput1.Value, dateTimeInput2.Value, _sizeIndex, GetCD.ClassNoData };

            _BGWClassStudentMeritDetail = new BackgroundWorker();
            _BGWClassStudentMeritDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentMeritDetail_DoWork);
            _BGWClassStudentMeritDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWClassStudentMeritDetail_RunWorkerCompleted);
            _BGWClassStudentMeritDetail.RunWorkerAsync(args);
        }

        void _BGWClassStudentMeritDetail_DoWork(object sender, DoWorkEventArgs e)
        {

            object[] args = e.Argument as object[];

            DateTime startDate = (DateTime)args[0];
            DateTime endDate = (DateTime)args[1];
            int size = (int)args[2];
            string ClassNoData = (string)args[3]; //不印出無資料班級

            //取得學生資料
            Data = new GetMeritDetail(startDate, endDate);

            Document template;

            if (GetCD.Setup_Mode == "false")
            {
                template = new Document(new MemoryStream(Properties.Resources.班級獎懲明細確認表範本));
            }
            else
            {
                if (GetCD.Setup_Temp == "")
                {
                    template = new Document(new MemoryStream(Properties.Resources.班級獎懲明細確認表範本));
                }
                else
                {
                    template = new Document(new MemoryStream(GetCD.Temp));
                }
            }

            tool.FontSize = template.Sections[0].Body.Tables[0].Rows[1].Cells[0].Paragraphs[0].Runs[0].Font.Size;
            tool.FontName = template.Sections[0].Body.Tables[0].Rows[1].Cells[0].Paragraphs[0].Runs[0].Font.Name;

            tool._run = new Run(template);

            DocumentBuilder DB = new DocumentBuilder(template);

            ColumnIndex = new Dictionary<string, int>();
            ColumnIndex.Add("座號", 0);
            ColumnIndex.Add("姓名", 1);
            ColumnIndex.Add("日期", 2);
            ColumnIndex.Add("獎/懲別", 3);
            ColumnIndex.Add("事由", 4);

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
                        PageContinue += Data.allAbsenceDetail[classInfo.ID][test1].GetMerList.Count;
                        PageContinue += Data.allAbsenceDetail[classInfo.ID][test1].GetDemerList.Count;
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

        void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs args)
        {
            if (args.FieldName == "Data")
            {
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToField(args.Field, true);
                args.Field.Remove();
                ClassSpeRecord Csr = args.FieldValue as ClassSpeRecord;

                //插入於此Row之下
                Row refrow = builder.CurrentParagraph.ParentNode.ParentNode as Row;

                //範本
                Row rowtemp = builder.CurrentParagraph.ParentNode.ParentNode.Clone(true) as Row;

                //此表格
                Table table = builder.CurrentParagraph.ParentNode.ParentNode.ParentNode as Table;

                foreach (StudentRecord stud in Csr.StudentRecordDic.Values) //取得學生ID
                {

                    if (Csr._dic.ContainsKey(stud.ID))
                    {

                        Csr._dic[stud.ID].GetMerList.Sort(tool.sortdat);
                        Csr._dic[stud.ID].GetDemerList.Sort(tool.sortdat);
                       

                        foreach (MeritRecord merit in Csr._dic[stud.ID].GetMerList)
                        {
                            tool.Write(refrow.Cells[Csr.ColumnIndex["座號"]], stud.SeatNo.HasValue ? stud.SeatNo.Value.ToString() : ""); //座號
                            tool.Write(refrow.Cells[Csr.ColumnIndex["姓名"]], stud.Name);   //姓名
                            tool.Write(refrow.Cells[Csr.ColumnIndex["日期"]], merit.OccurDate.ToShortDateString());   //日期

                            tool.Write(refrow.Cells[Csr.ColumnIndex["獎/懲別"]], GetMerReason(merit));   //獎懲

                            tool.Write(refrow.Cells[Csr.ColumnIndex["事由"]], merit.Reason);   //事由

                            refrow = table.InsertAfter(rowtemp.Clone(true), refrow) as Row;
                        }


                        foreach (DemeritRecord demerit in Csr._dic[stud.ID].GetDemerList)
                        {
                            tool.Write(refrow.Cells[Csr.ColumnIndex["座號"]], stud.SeatNo.HasValue ? stud.SeatNo.Value.ToString() : ""); //座號
                            tool.Write(refrow.Cells[Csr.ColumnIndex["姓名"]], stud.Name);   //姓名
                            tool.Write(refrow.Cells[Csr.ColumnIndex["日期"]], demerit.OccurDate.ToShortDateString());   //日期

                            tool.Write(refrow.Cells[Csr.ColumnIndex["獎/懲別"]], GetDemerReason(demerit));   //日期

                            tool.Write(refrow.Cells[Csr.ColumnIndex["事由"]], demerit.Reason);   //日期


                            refrow = table.InsertAfter(rowtemp.Clone(true), refrow) as Row;
                        }
                    }
                }

                refrow.Remove();
            }
        }

        /// <summary>
        /// 取得事由字串
        /// </summary>
        private string GetMerReason(MeritRecord merit)
        {
            StringBuilder sb = new StringBuilder();

            if (merit.MeritA.HasValue)
                if (merit.MeritA.Value != 0)
                    sb.Append("大功「" + merit.MeritA.Value.ToString() + "」");

            if (merit.MeritB.HasValue)
                if (merit.MeritB.Value != 0)
                    sb.Append("小功「" + merit.MeritB.Value.ToString() + "」");

            if (merit.MeritC.HasValue)
                if (merit.MeritC.Value != 0)
                    sb.Append("嘉獎「" + merit.MeritC.Value.ToString() + "」");

            return sb.ToString();
        }

        /// <summary>
        /// 取得事由字串
        /// </summary>
        private string GetDemerReason(DemeritRecord demerit)
        {
            StringBuilder sb = new StringBuilder();

            if (demerit.DemeritA.HasValue)
                if (demerit.DemeritA.Value != 0)
                    sb.Append("大過「" + demerit.DemeritA.Value.ToString() + "」");

            if (demerit.DemeritB.HasValue)
                if (demerit.DemeritB.Value != 0)
                    sb.Append("小過「" + demerit.DemeritB.Value.ToString() + "」");

            if (demerit.DemeritC.HasValue)
                if (demerit.DemeritC.Value != 0)
                    sb.Append("警告「" + demerit.DemeritC.Value.ToString() + "」");

            return sb.ToString();
        }

        void _BGWClassStudentMeritDetail_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.btnSave.Enabled = true;

            if (e.Cancelled != true)
            {
                if (e.Error == null)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 班級獎懲記錄明細(確認表) 已完成");

                    SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = "班級獎懲記錄明細(確認表).doc";
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


        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lblPrintSetup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MeritSetup AS = new MeritSetup(GetCD);
            AS.ShowDialog();
        }
    }
}

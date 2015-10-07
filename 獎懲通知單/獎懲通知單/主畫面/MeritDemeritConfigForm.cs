﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar.Controls;
using DevComponents.DotNetBar.Rendering;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;

namespace 獎懲通知單
{
    public partial class MeritDemeritConfigForm : BaseForm
    {
        private byte[] _buffer = null;
        private string base64 = null;
        private bool _isUpload = false;
        private string _defaultTemplate; //預設範本1 + 預設範本2 + 自訂範本
        private bool _printHasRecordOnly;
        private DateRangeModeNew _mode = DateRangeModeNew.Month;
        private bool _printStudentList;

        string ConfigName = "獎懲通知單_ForK12.2013";


        public MeritDemeritConfigForm(string defaultTemplate, DateRangeModeNew mode, byte[] buffer, string name, string address, string condName, string condNumber, bool printStudentList)
        {
            InitializeComponent();

            //如果系統的Renderer是Office2007Renderer
            //同化_ClassTeacherView,_CategoryView的顏色
            if (GlobalManager.Renderer is Office2007Renderer)
            {
                ((Office2007Renderer)GlobalManager.Renderer).ColorTableChanged += new EventHandler(ScoreCalcRuleEditor_ColorTableChanged);
                SetForeColor(this);
            }

            _defaultTemplate = defaultTemplate;
            _mode = mode;
            _printStudentList = printStudentList;

            if (buffer != null)
                _buffer = buffer;

            if (defaultTemplate == "預設範本2") //如果是2才設為2
                rbDEF_2.Checked = true;
            else if (defaultTemplate == "自訂範本") //如果是自訂
                radioButton2.Checked = true;
            else
                rbDEF_1.Checked = true; //如果都不是就進入預設1

            checkBoxX2.Checked = printStudentList;

            switch (mode)
            {
                case DateRangeModeNew.Month:
                    radioButton3.Checked = true;
                    break;
                case DateRangeModeNew.Week:
                    radioButton4.Checked = true;
                    break;
                case DateRangeModeNew.Custom:
                    radioButton5.Checked = true;
                    break;
                default:
                    throw new Exception("Date Range Mode Error.");
            }

            //設定 ComboBox
            Dictionary<ComboBoxEx, string> cboBoxes = new Dictionary<ComboBoxEx, string>();
            cboBoxes.Add(comboBoxEx1, name);
            cboBoxes.Add(comboBoxEx2, address);
            cboBoxes.Add(comboBoxEx4, condName);

            foreach (ComboBoxEx var in cboBoxes.Keys)
            {
                var.SelectedIndex = 0;
                foreach (DevComponents.Editors.ComboItem item in var.Items)
                {
                    if (item.Text == cboBoxes[var])
                    {
                        var.SelectedIndex = var.Items.IndexOf(item);
                        break;
                    }
                }
            }

            //設定 NumericUpDown
            decimal tryValue;
            if (condNumber == "0")
                condNumber = "1";
            numericUpDown1.Value = (decimal.TryParse(condNumber, out tryValue)) ? tryValue : 1;
        }

        void ScoreCalcRuleEditor_ColorTableChanged(object sender, EventArgs e)
        {
            SetForeColor(this);
        }

        private void SetForeColor(Control parent)
        {
            foreach (Control var in parent.Controls)
            {
                if (var is RadioButton)
                    var.ForeColor = ((Office2007Renderer)GlobalManager.Renderer).ColorTable.CheckBoxItem.Default.Text;
                SetForeColor(var);
            }
        }

        private void rbDEF_1_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDEF_1.Checked)
            {
                //radioButton2.Checked = false;
                _defaultTemplate = "預設範本1";
            }
        }

        private void rbDEF_2_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDEF_2.Checked)
            {
                //radioButton2.Checked = false;
                _defaultTemplate = "預設範本2";
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                //radioButton1.Checked = false;
                _defaultTemplate = "自訂範本";
            }
        }

        private void checkBoxX2_CheckedChanged(object sender, EventArgs e)
        {
            _printStudentList = checkBoxX2.Checked;
        }

        private void linkDef1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "獎懲通知單_範本1.docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.獎懲通知單_住址中間版, 0, Properties.Resources.獎懲通知單_住址中間版.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkDef2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "獎懲通知單_範本2.docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.獎懲通知單_住址上方版, 0, Properties.Resources.獎懲通知單_住址上方版.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkViewGeDin_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "自訂獎懲通知單範本.docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(new MemoryStream(_buffer));
                    doc.Save(sfd.FileName, Aspose.Words.SaveFormat.Docx);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案無法儲存。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案無法開啟。" + ex.Message);
                    return;
                }
            }
        }

        private void linkUpData_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            if (_buffer == null)
            {
                MsgBox.Show("目前沒有任何範本，請重新上傳。");
                return;
            }

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇自訂的獎懲通知單範本";
            ofd.Filter = "Word檔案 (*.docx)|*.docx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                    byte[] tempBuffer = new byte[fs.Length];
                    fs.Read(tempBuffer, 0, tempBuffer.Length);
                    base64 = Convert.ToBase64String(tempBuffer);
                    _isUpload = true;
                    fs.Close();
                    MsgBox.Show("上傳成功。");
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference

            ConfigData cd = K12.Data.School.Configuration[ConfigName];
            XmlElement config = cd.GetXml("XmlData", null);

            //XmlElement config = CurrentUser.Instance.Preference["獎懲通知單"];

            if (config == null)
            {
                config = new XmlDocument().CreateElement("獎懲通知單");
            }

            config.SetAttribute("Default", _defaultTemplate);

            XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
            XmlElement mode = config.OwnerDocument.CreateElement("DateRangeMode");
            XmlElement receive = config.OwnerDocument.CreateElement("Receive");
            XmlElement conditions = config.OwnerDocument.CreateElement("Conditions");
            XmlElement PrintStudentList = config.OwnerDocument.CreateElement("PrintStudentList");

            PrintStudentList.SetAttribute("Checked", _printStudentList.ToString());
            config.ReplaceChild(PrintStudentList, config.SelectSingleNode("PrintStudentList"));

            if (_isUpload) //如果是自訂範本
            {
                customize.InnerText = base64;
                config.ReplaceChild(customize, config.SelectSingleNode("CustomizeTemplate"));
            }

            mode.InnerText = ((int)_mode).ToString();
            config.ReplaceChild(mode, config.SelectSingleNode("DateRangeMode"));

            receive.SetAttribute("Name", ((DevComponents.Editors.ComboItem)comboBoxEx1.SelectedItem).Text);
            receive.SetAttribute("Address", ((DevComponents.Editors.ComboItem)comboBoxEx2.SelectedItem).Text);
            if (config.SelectSingleNode("Receive") == null)
                config.AppendChild(receive);
            else
                config.ReplaceChild(receive, config.SelectSingleNode("Receive"));

            conditions.SetAttribute("ConditionName", ((DevComponents.Editors.ComboItem)comboBoxEx4.SelectedItem).Text);
            conditions.SetAttribute("ConditionNumber", numericUpDown1.Value.ToString());
            if (config.SelectSingleNode("Conditions") == null)
                config.AppendChild(conditions);
            else
                config.ReplaceChild(conditions, config.SelectSingleNode("Conditions"));

            cd.SetXml("XmlData", config);
            cd.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                //radioButton4.Checked = false;
                //radioButton5.Checked = false;
                _mode = DateRangeModeNew.Month;
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                //radioButton3.Checked = false;
                //radioButton5.Checked = false;
                _mode = DateRangeModeNew.Week;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                //radioButton3.Checked = false;
                //radioButton4.Checked = false;
                _mode = DateRangeModeNew.Custom;
            }
        }

        private void comboBoxEx4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxEx4.SelectedIndex == 0)
            {
                numericUpDown1.Enabled = false;
                numericUpDown1.Value = 1;
            }
            else
            {
                numericUpDown1.Enabled = true;
            }
        }
    }
}

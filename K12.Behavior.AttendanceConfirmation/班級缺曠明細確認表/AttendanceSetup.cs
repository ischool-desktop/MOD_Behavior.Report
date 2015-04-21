using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.Xml;
using System.IO;
using FISCA.Presentation.Controls;

namespace K12.Behavior.AttendanceConfirmation
{
    public partial class AttendanceSetup : BaseForm
    {
        //private string _reportName = "";
        //private bool _defaultTemplate;

        private byte[] _buffer = null;
        string base64;
        string nowTemplate = "false";

        //private MemoryStream _template = null;
        //private MemoryStream _defaultTemplate = new MemoryStream(Properties.Resources.班級缺曠明細確認表範本);



        AttendanceConfigData _CD;

        public AttendanceSetup(AttendanceConfigData CD)
        {
            InitializeComponent();

            _CD = CD;
        }


        private void AttendanceSetup_Load(object sender, EventArgs e)
        {
            if (_CD.Setup_Mode == "false") //預設範本模式,False是使用範本
            {
                radioButton1.Checked = true;
            }
            else
            {
                radioButton2.Checked = true;
            }

            if (_CD.ClassNoData == "False")
            {
                checkBoxX1.Checked = false;
            }
            else
            {
                checkBoxX1.Checked = true;
            }


            if (_CD.Temp != null)
            {
                _buffer = _CD.Temp;
                base64 = Convert.ToBase64String(_buffer);
            }
            else
            {
                base64 = "";
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked) //預設
            {
                nowTemplate = "false";
            }
            else if (radioButton2.Checked) //自定
            {
                nowTemplate = "true";
            }
            else //預設
            {
                nowTemplate = "false";
            }

            _CD.SavePrint(nowTemplate, base64, checkBoxX1.Checked.ToString());

            this.Close();

            //#region 儲存 Preference

            ////XmlElement config = CurrentUser.Instance.Preference[_reportName];
            //K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration[_reportName];

            //XmlElement config = cd.GetXml("XmlData", null);

            ////CurrentUser.Instance.Preference[_reportName] = config;

            //cd.SetXml("XmlData", config);
            //cd.Save();

            //#endregion

            //this.DialogResult = DialogResult.OK;
            //this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                radioButton2.Checked = false;
                //_defaultTemplate = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                radioButton1.Checked = false;
                //_defaultTemplate = false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "(範本)班級缺曠記錄明細(確認表).doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.班級缺曠明細確認表範本, 0, Properties.Resources.班級缺曠明細確認表範本.Length);
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

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "(自訂範本)班級缺曠記錄明細(確認表).doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    if (_buffer == null)
                    {
                        MsgBox.Show("尚無自定範本,請上傳範本!");
                        return;
                    }

                    fs.Write(_buffer, 0, _buffer.Length);
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

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "請選擇自訂範本";
            ofd.Filter = "Word檔案 (*.doc)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                    byte[] tempBuffer = new byte[fs.Length];
                    _buffer = tempBuffer;
                    fs.Read(tempBuffer, 0, tempBuffer.Length);
                    base64 = Convert.ToBase64String(tempBuffer);
                    //_isUpload = true;
                    fs.Close();
                    MsgBox.Show("上傳成功。");
                    radioButton2.Checked = true;

                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
    }
}
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

namespace K12.Behavior.MeritDemeritConfirmation
{
    public partial class MeritSetup : BaseForm
    {
        //private string _reportName = "";
        //private bool _defaultTemplate;

        private byte[] _buffer = null;
        string base64;
        string nowTemplate = "false";

        //private MemoryStream _template = null;
        //private MemoryStream _defaultTemplate = new MemoryStream(Properties.Resources.�Z�ů��m���ӽT�{��d��);



        MeritConfigData _CD;

        public MeritSetup(MeritConfigData CD)
        {
            InitializeComponent();

            _CD = CD;
        }


        private void AttendanceSetup_Load(object sender, EventArgs e)
        {
            if (_CD.Setup_Mode == "false") //�w�]�d���Ҧ�,False�O�ϥνd��
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
            if (radioButton1.Checked) //�w�]
            {
                nowTemplate = "false";
            }
            else if (radioButton2.Checked) //�۩w
            {
                nowTemplate = "true";
            }
            else //�w�]
            {
                nowTemplate = "false";
            }

            _CD.SavePrint(nowTemplate, base64, checkBoxX1.Checked.ToString());

            this.Close();

            //#region �x�s Preference

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
            sfd.Title = "�t�s�s��";
            sfd.FileName = "(�d��)�Z�ů��m�O������(�T�{��).doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.�Z�ż��g���ӽT�{��d��, 0, Properties.Resources.�Z�ż��g���ӽT�{��d��.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "(�ۭq�d��)�Z�ů��m�O������(�T�{��).doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    if (_buffer == null)
                    {
                        MsgBox.Show("�|�L�۩w�d��,�ФW�ǽd��!");
                        return;
                    }

                    if (Aspose.Words.Document.DetectFileFormat(new MemoryStream(_buffer)) == Aspose.Words.LoadFormat.Doc)
                        fs.Write(_buffer, 0, _buffer.Length);
                    else
                        fs.Write(Properties.Resources.�Z�ż��g���ӽT�{��d��, 0, Properties.Resources.�Z�ż��g���ӽT�{��d��.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "�п�ܦۭq�d��";
            ofd.Filter = "Word�ɮ� (*.doc)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Aspose.Words.Document.DetectFileFormat(ofd.FileName) == Aspose.Words.LoadFormat.Doc)
                    {
                        FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                        byte[] tempBuffer = new byte[fs.Length];
                        _buffer = tempBuffer;
                        fs.Read(tempBuffer, 0, tempBuffer.Length);
                        base64 = Convert.ToBase64String(tempBuffer);
                        //_isUpload = true;
                        fs.Close();
                        MsgBox.Show("�W�Ǧ��\�C");
                        radioButton2.Checked = true;
                    }
                    else
                        MsgBox.Show("�W���ɮ׮榡����");
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�}���ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace K12.Behavior.AttendanceConfirmation
{
    public partial class SetupForm : BaseForm
    {
        AttendanceConfigData _CD;

        public SetupForm(AttendanceConfigData RT)
        {
            InitializeComponent();

            _CD = RT;
        }

        private void SetupForm_Load(object sender, EventArgs e)
        {
            foreach (string each in _CD.AllAbsence)
            {
                ListViewItem item = new ListViewItem(each);
                if (_CD.Setup_AttendanceList.Contains(item.Text))
                {
                    item.Checked = true;
                }
                listView1.Items.Add(item);
            }

            foreach (string each in _CD.Setup_AttendanceList)
            {

            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            foreach (ListViewItem each in listView1.Items)
            {
                if (each.Checked && !list.Contains(each.Text))
                {
                    list.Add(each.Text);
                }
            }

            _CD.SaveAttendance(list);

            this.Close();

        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem each in listView1.Items)
            {
                each.Checked = checkBox1.Checked;
            }
        }
    }
}

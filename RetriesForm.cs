using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CBD1000_Demo
{
    public partial class RetriesForm : Form
    {
        public UInt16[] m_Max;

        public RetriesForm( UInt16[] max)
        {
            InitializeComponent();

            m_Max = max;

            label1.Text = "Double Feed";
            label2.Text = "Mark Mismatch";
            label3.Text = "Micr Error";
            label4.Text = "Preprinted Mismatch";
            label5.Text = "Sheets Near Empty in Feeder";

            numericUpDown1.Value = (decimal)m_Max[0];
            numericUpDown2.Value = (decimal)m_Max[1];
            numericUpDown3.Value = (decimal)m_Max[2];
            numericUpDown4.Value = (decimal)m_Max[3];
            numericUpDown5.Value = (decimal)m_Max[4];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            m_Max[0] = (ushort)numericUpDown1.Value;
            m_Max[1] = (ushort)numericUpDown2.Value;
            m_Max[2] = (ushort)numericUpDown3.Value;
            m_Max[3] = (ushort)numericUpDown4.Value;
            m_Max[4] = (ushort)numericUpDown5.Value;
        }


    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class haihuGroup : Form
    {
        public haihuGroup()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            const int DIALOG_OK = 1;
            const int DIALOG_NO = 0;

            if (textBox1.Text.Trim() == "")
            {
                MessageBox.Show("���������͂��Ă�������", "�m�F", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                sStatus = DIALOG_NO;
                return;
            }

            if (MessageBox.Show(textBox1.Text + " �ŃO���[�v������܂��B��낵���ł���", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                sStatus = DIALOG_NO;
                return;
            }

            sStatus = DIALOG_OK;
            sGroup = textBox1.Text;

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private int F_sStatus;

        public int sStatus
        {
            get { return F_sStatus; }
            set { F_sStatus = value; }
        }


        private string F_sGroup;

        public string sGroup
        {
            get { return F_sGroup; }
            set { F_sGroup = value; }
        }


    }
}
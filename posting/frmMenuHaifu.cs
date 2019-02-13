using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmMenuHaifu : Form
    {
        public frmMenuHaifu()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new frmHaifuShiji();
            frm.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new frmHaifuKanryoRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new frmHaifuShinchoku();
            frm.ShowDialog();
            this.Show();
        }

        private void frmMenuHaifu_Load(object sender, EventArgs e)
        {

        }
    }
}
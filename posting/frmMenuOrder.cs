using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmMenuOrder : Form
    {
        public frmMenuOrder()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //受注確定書登録
            Form frm = new frmOrder(0);

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //ポスティングエリア登録
            Form frm = new frmPosting();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            //ポスティングエリア登録
            Form frm = new frmTantouOrderRep();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void frmMenuOrder_Load(object sender, EventArgs e)
        {

        }

    }
}
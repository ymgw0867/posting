using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmKaikakeMenu : Form
    {
        public frmKaikakeMenu()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmKaikakeMotocho frm = new frmKaikakeMotocho();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 閉じる
            Close();
        }

        private void frmUrikakeMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmKaikakeZan frm = new frmKaikakeZan();
            frm.ShowDialog();
            this.Show();
        }
        
    }
}

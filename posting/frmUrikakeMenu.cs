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
    public partial class frmUrikakeMenu : Form
    {
        public frmUrikakeMenu()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmUrikakeMotocho frm = new frmUrikakeMotocho();
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
            frmUrikakeZan frm = new frmUrikakeZan();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmUriageOrderKbn frm = new frmUriageOrderKbn();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmUrikakeNen frm = new frmUrikakeNen();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmUriageGaichu frm = new frmUriageGaichu();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmClientbyOrderRep frm = new frmClientbyOrderRep();
            frm.ShowDialog();
            this.Show();
        }
    }
}

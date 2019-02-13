﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmPrg : Form
    {
        public frmPrg()
        {
            InitializeComponent();
        }

        private void frmPrg_Load(object sender, EventArgs e)
        {
            prgBar.Maximum = 100;
            prgBar.Minimum = 0;
            prgBar.Step = 1;
        }

        public int progressValue { get; set; }
        public int progressMax { get; set; }

        public void ProgressStep()
        {
            //prgBar.Value = progressValue;
            prgBar.PerformStep();
        }

        public void setProgressMax()
        {
            prgBar.Maximum = progressMax;
        }

        private void frmPrg_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}

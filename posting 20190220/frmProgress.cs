using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmProgress : Form
    {
        public frmProgress()
        {
            InitializeComponent();
        }

        private int _valueMin;

        public int valueMin
        {
            get { return _valueMin; }
            set { _valueMin = value; }
        }

        private int _valueMax;

        public int valueMax
        {
            get { return _valueMax; }
            set { _valueMax = value; }
        }

        private int _valueCount;

        public int valueCount
        {
            get { return _valueCount; }
            set { _valueCount = value; }
        }

        private void frmProgress_Load(object sender, EventArgs e)
        {
            progressBar1.Maximum = valueMax;
            progressBar1.Minimum = valueMin;

        }

        public void ShowProgress()
        {
            progressBar1.Value = valueCount;
        }
    }
}
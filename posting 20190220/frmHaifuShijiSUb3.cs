using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace posting
{
    public partial class frmHaifuShijiSUb3 : Form
    {
        Entity.���z�z��� cMaster = new Entity.���z�z���();

        const string MESSAGE_CAPTION = "���z�z���ҏW";
        int sMode;

        public frmHaifuShijiSUb3(int tempMode)
        {
            InitializeComponent();
            sMode = tempMode;
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            switch (sMode)
            {
                case 0:
                    lblID.Text = _FID.ToString();
                    lblAdd.Text = _Add;
                    txtBanchi.Focus();
                    break;

                case 1:
                    ShowData();
                    txtBanchi.Focus();
                    break;

                default:
                    break;
            }
        }
        
        private void ShowData()
        {
            lblID.Text = _ID.ToString();
            lblAdd.Text = _Add;

            //�f�[�^�擾
            OleDbDataReader dR;
            string sqlSTR;
            sqlSTR = "where ID = " + _ID.ToString();

            Control.���z�z��� cMihaifu = new Control.���z�z���();
            dR = cMihaifu.FillBy(sqlSTR);

            while (dR.Read())
            {
                cMaster.ID = int.Parse(dR["ID"].ToString());
                cMaster.�z�z�G���AID = int.Parse(dR["�z�z�G���AID"].ToString());
                cMaster.�Ԓn�� = dR["�Ԓn��"].ToString();
                cMaster.�}���V������ = dR["�}���V������"].ToString();
                cMaster.���R = int.Parse(dR["���R"].ToString());
                cMaster.���̑����e = dR["���̑����e"].ToString();
            }

            dR.Close();
            cMihaifu.Close();

            //�f�[�^��ʕ\��
            txtBanchi.Text = cMaster.�Ԓn��;
            txtManshon.Text = cMaster.�}���V������;
            txtRiyu.Text = cMaster.���R.ToString();
            txtSonota.Text = cMaster.���̑����e;

            //���R�E�v�\��
            //OleDbDataReader dR;
            Control.���z�z���R cRiyu = new Control.���z�z���R();

            dR = cRiyu.FillBy("where ID = " + txtRiyu.Text);

            //�E�v����\��
            while (dR.Read())
            {
                txtRiyuName.Text = dR["�E�v"].ToString().Trim();
            }

            dR.Close();
            cRiyu.Close();

        }

        //�\���f�[�^�̃v���p�e�B
        private int _ID;
        private string _Add;

        public int ID
        {
            set { this._ID = value; }
        }

        public string Add
        {
            set { _Add = value; }
        }

        private int _FID;

        public int FID
        {
            get { return _FID; }
            set { _FID = value; }
        }

        private string _Banchi;

        public string Banchi
        {
            get { return _Banchi; }
            set { _Banchi = value; }
        }

        private string _Manshon;

        public string Manshon
        {
            get { return _Manshon; }
            set { _Manshon = value; }
        }

        private int _Riyu;

        public int Riyu
        {
            get { return _Riyu; }
            set { _Riyu = value; }
        }

        private string _Sonota;

        public string Sonota
        {
            get { return _Sonota; }
            set { _Sonota = value; }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("���z�z����o�^���܂��B��낵���ł���","�m�F",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (fDataCheck() == false) return;
            
            Control.���z�z��� cMihaifu = new Control.���z�z���();

            switch (sMode)
            {
                case 0:
                    cMihaifu.DataInsert(cMaster);
                    break;

                case 1:
                    cMihaifu.DataUpdate(cMaster);
                    break;

            }

            this.Close();

        }

        private Boolean fDataCheck()
        {

            try
            {

                //���R�R�[�h�������H
                if (Utility.NumericCheck(txtRiyu.Text) == false)
                {
                    txtRiyu.Focus();
                    throw new Exception("���R����͂��Ă�������");
                }

                //�N���X�Ƀf�[�^�Z�b�g
                if (sMode == 0) cMaster.�z�z�G���AID = _FID;

                cMaster.�Ԓn�� = txtBanchi.Text;
                cMaster.�}���V������ = txtManshon.Text;
                cMaster.���R = int.Parse(txtRiyu.Text);
                cMaster.���̑����e = txtSonota.Text;
                cMaster.�o�^�N���� = DateTime.Today;
                cMaster.�ύX�N���� = DateTime.Today;

                if (sMode == 0) cMaster.�o�^�N���� = DateTime.Today;
                cMaster.�ύX�N���� = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void txtRiyu_Validating(object sender, CancelEventArgs e)
        {

            //���L����OK
            if (txtRiyu.Text == "")
            {
                txtRiyuName.Text = "";
                txtSonota.Enabled = false;
                return;
            }

            //�������H
            if (Utility.NumericCheck(txtRiyu.Text) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", "���z�z���R", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Cancel = true;
                return;
            }

            //�}�X�^�[�o�^����Ă��邩�H
            OleDbDataReader dR;
            Control.���z�z���R cRiyu = new Control.���z�z���R();

            dR = cRiyu.FillBy("where ID = " + txtRiyu.Text);

            if (dR.HasRows == false)
            {
                MessageBox.Show("�}�X�^�[���o�^�ł�", "���z�z���R", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Cancel = true;
                dR.Close();
                cRiyu.Close();
                return;
            }
            
            //�E�v����\��
            while(dR.Read())
            {
                txtRiyuName.Text = dR["�E�v"].ToString().Trim();
            }

            dR.Close();

            cRiyu.Close();

            //���̑��̂Ƃ�
            if (txtRiyuName.Text == "���̑�")
            {
                txtSonota.Enabled = true;
                txtSonota.Focus();
            }
            else
            {
                txtSonota.Enabled = false;
            }

        }

        private void txt_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtBanchi) { txtObj = txtBanchi; }
            if (sender == txtManshon) { txtObj = txtManshon; }

            if (sender == txtRiyu)
            {
                txtObj = txtRiyu;

                if (txtRiyu.Text == "0")
                {
                    txtRiyu.Text = "";
                }
            }

            if (sender == txtSonota) { txtObj = txtSonota; }

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;

        }

        private void txt_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtBanchi) { txtObj = txtBanchi; }
            if (sender == txtManshon) { txtObj = txtManshon; }
            if (sender == txtRiyu) { txtObj = txtRiyu; }
            if (sender == txtSonota) { txtObj = txtSonota; }

            txtObj.BackColor = Color.White;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmHaifuShijiSUb3_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

    }
}
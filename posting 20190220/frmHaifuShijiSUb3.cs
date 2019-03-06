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
        Entity.未配布情報 cMaster = new Entity.未配布情報();

        const string MESSAGE_CAPTION = "未配布情報編集";
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

            //データ取得
            OleDbDataReader dR;
            string sqlSTR;
            sqlSTR = "where ID = " + _ID.ToString();

            Control.未配布情報 cMihaifu = new Control.未配布情報();
            dR = cMihaifu.FillBy(sqlSTR);

            while (dR.Read())
            {
                cMaster.ID = int.Parse(dR["ID"].ToString());
                cMaster.配布エリアID = int.Parse(dR["配布エリアID"].ToString());
                cMaster.番地号 = dR["番地号"].ToString();
                cMaster.マンション名 = dR["マンション名"].ToString();
                cMaster.理由 = int.Parse(dR["理由"].ToString());
                cMaster.その他内容 = dR["その他内容"].ToString();
            }

            dR.Close();
            cMihaifu.Close();

            //データ画面表示
            txtBanchi.Text = cMaster.番地号;
            txtManshon.Text = cMaster.マンション名;
            txtRiyu.Text = cMaster.理由.ToString();
            txtSonota.Text = cMaster.その他内容;

            //理由摘要表示
            //OleDbDataReader dR;
            Control.未配布理由 cRiyu = new Control.未配布理由();

            dR = cRiyu.FillBy("where ID = " + txtRiyu.Text);

            //摘要名を表示
            while (dR.Read())
            {
                txtRiyuName.Text = dR["摘要"].ToString().Trim();
            }

            dR.Close();
            cRiyu.Close();

        }

        //表示データのプロパティ
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

            if (MessageBox.Show("未配布情報を登録します。よろしいですか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (fDataCheck() == false) return;
            
            Control.未配布情報 cMihaifu = new Control.未配布情報();

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

                //理由コード数字か？
                if (Utility.NumericCheck(txtRiyu.Text) == false)
                {
                    txtRiyu.Focus();
                    throw new Exception("理由を入力してください");
                }

                //クラスにデータセット
                if (sMode == 0) cMaster.配布エリアID = _FID;

                cMaster.番地号 = txtBanchi.Text;
                cMaster.マンション名 = txtManshon.Text;
                cMaster.理由 = int.Parse(txtRiyu.Text);
                cMaster.その他内容 = txtSonota.Text;
                cMaster.登録年月日 = DateTime.Today;
                cMaster.変更年月日 = DateTime.Today;

                if (sMode == 0) cMaster.登録年月日 = DateTime.Today;
                cMaster.変更年月日 = DateTime.Today;

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

            //無記入はOK
            if (txtRiyu.Text == "")
            {
                txtRiyuName.Text = "";
                txtSonota.Enabled = false;
                return;
            }

            //数字か？
            if (Utility.NumericCheck(txtRiyu.Text) == false)
            {
                MessageBox.Show("数字で入力してください", "未配布理由", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Cancel = true;
                return;
            }

            //マスター登録されているか？
            OleDbDataReader dR;
            Control.未配布理由 cRiyu = new Control.未配布理由();

            dR = cRiyu.FillBy("where ID = " + txtRiyu.Text);

            if (dR.HasRows == false)
            {
                MessageBox.Show("マスター未登録です", "未配布理由", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Cancel = true;
                dR.Close();
                cRiyu.Close();
                return;
            }
            
            //摘要名を表示
            while(dR.Read())
            {
                txtRiyuName.Text = dR["摘要"].ToString().Trim();
            }

            dR.Close();

            cRiyu.Close();

            //その他のとき
            if (txtRiyuName.Text == "その他")
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
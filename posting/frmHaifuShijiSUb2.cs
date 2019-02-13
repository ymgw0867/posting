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
    public partial class frmHaifuShijiSUb2 : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.配布エリア cMaster = new Entity.配布エリア();

        const string MESSAGE_CAPTION = "配布指示データ編集";
        const int MIHAIFU_ADD = 0;      //未配布情報新規登録
        const int MIHAIFU_UPDATE = 1;   //未配布情報更新

        public frmHaifuShijiSUb2()
        {
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            ShowData();

            GridviewSet.Setting(dataGridView1);
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));
        }

        // データグリッドビュークラス
        private class GridviewSet
        {

            //配布指示明細データグリッド
            public static void Setting(DataGridView tempDGV)
            {
                try
                {
                    //フォームサイズ定義

                    // 列スタイルを変更する

                    tempDGV.EnableHeadersVisualStyles = false;

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 181;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "PID");
                    tempDGV.Columns.Add("col2", "FID");
                    tempDGV.Columns.Add("col3", "番地号");
                    tempDGV.Columns.Add("col4", "マンション名");
                    tempDGV.Columns.Add("col5", "理由");
                    tempDGV.Columns.Add("col6", "その他内容");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 140;
                    tempDGV.Columns[3].Width = 175;
                    tempDGV.Columns[4].Width = 60;
                    tempDGV.Columns[5].Width = 265;

                    tempDGV.Columns[0].Visible = false;
                    tempDGV.Columns[1].Visible = false;

                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    
                    //tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

                    // 編集不可とする
                    tempDGV.ReadOnly = true;

                    // 追加行表示しない
                    tempDGV.AllowUserToAddRows = false;

                    // データグリッドビューから行削除を禁止する
                    tempDGV.AllowUserToDeleteRows = false;

                    // 手動による列移動の禁止
                    tempDGV.AllowUserToOrderColumns = false;

                    // 列サイズ変更禁止
                    tempDGV.AllowUserToResizeColumns = false;

                    // 行サイズ変更禁止
                    tempDGV.AllowUserToResizeRows = false;

                    // 行ヘッダーの自動調節
                    //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// データグリッドビューの指定行のデータを取得する
            /// </summary>
            /// <param name="dgv">対象とするデータグリッドビューオブジェクト</param>
            public static Boolean GetData(DataGridView dgv, ref Entity.未配布情報 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.未配布情報 mihaifu = new Control.未配布情報();
                OleDbDataReader dr;

                sqlStr = " where 未配布情報.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = mihaifu.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.配布エリアID = int.Parse(dr["配布エリアID"].ToString());
                        tempC.番地号 = dr["番地号"].ToString();
                        tempC.マンション名 = dr["マンション名"].ToString();
                        tempC.理由 = Int32.Parse(dr["理由"].ToString());
                        tempC.その他内容 = dr["その他内容"].ToString();
                    }
                }
                else
                {
                    dr.Close();
                    mihaifu.Close();
                    return false;
                }

                dr.Close();
                mihaifu.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV,int tempID)
            {
                string sqlSTRING = "";
                int iX;

                try
                {
                    tempDGV.RowCount = 0;

                    //未配布情報データのデータリーダーを取得する
                    OleDbDataReader dR;
                    Control.未配布情報 cMi = new Control.未配布情報();
                    sqlSTRING = "where 配布エリアID = " + tempID.ToString();
                    dR = cMi.FillBy(sqlSTRING);

                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = dR["ID"].ToString();
                        tempDGV[1, iX].Value = dR["配布エリアID"].ToString();
                        tempDGV[2, iX].Value = dR["番地号"].ToString();
                        tempDGV[3, iX].Value = dR["マンション名"].ToString();
                        tempDGV[4, iX].Value = Int32.Parse(dR["理由"].ToString());
                        tempDGV[5, iX].Value = dR["その他内容"].ToString();
                        iX++;
                    }

                    dR.Close();
                    cMi.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }

                tempDGV.CurrentCell = null;
            }


        }


        private void ShowData()
        {
            lblsID.Text = _SID.ToString();
            lblhDate.Text = _hDate;
            lblName.Text = _staffName;
            lblID.Text = _ID.ToString();
            lblcName.Text = _cName;
            lblfJyouken.Text = _fJyouken;
            lblfKeitai.Text = _fKeitai;
            lblAdd.Text = _Add;
            txtTanka.Text = _Tanka.ToString("#,##0.0");
            lblyMaisu.Text = _yMaisu.ToString("#,##0");
            txthMaisu.Text = _hMaisu.ToString("#,##0");
            
            //txtBanchi.Text = _Banchi;
            //txtManshon.Text = _Manshon;
            //txtRiyu.Text = _Riyu.ToString();
            //txtRiyuName.Text = "";

            ////理由摘要表示
            //OleDbDataReader dR;
            //Control.未配布理由 cRiyu = new Control.未配布理由();

            //dR = cRiyu.FillBy("where ID = " + txtRiyu.Text);

            ////摘要名を表示
            //while (dR.Read())
            //{
            //    txtRiyuName.Text = dR["摘要"].ToString().Trim();
            //}

            //dR.Close();
            //cRiyu.Close();

            //txtSonota.Text = _Sonota;

            if (_kanryo == 1)
            {
                checkBox1.Checked = true;
                label12.Visible = true;
            }
            else
            {
                checkBox1.Checked = false;
                label12.Visible = false;
            }

            txtEdaban.Text = _Edaban;

            //枝番有無
            if (_Edaban_Status == 0)
            {
                txtEdaban.Enabled = false;
            }
            else
            {
                txtEdaban.Enabled = true;
            }

            //未配布有無
            if (_Mihaifu_Status == 0)
            {
                //txtBanchi.Enabled = false;
                //txtManshon.Enabled = false;
                //txtRiyu.Enabled = false;
                //txtSonota.Enabled = false;

                dataGridView1.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
            }
            else
            {
                //txtBanchi.Enabled = true;
                //txtManshon.Enabled = true;
                //txtRiyu.Enabled = true;
                //txtSonota.Enabled = true;

                dataGridView1.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
            }
        }

        //表示データのプロパティ
        private int _SID;
        private string _hDate;
        private string _staffName;
        private int _ID;
        private string _cName;
        private string _fJyouken;
        private string _fKeitai;
        private string _Add;
        private double _Tanka;
        private int _yMaisu;
        private int _hMaisu;
        private int _kanryo;
        private string _Edaban;
        private int _Edaban_Status;

        public int SID
        {
            set { this._SID = value; }
        }

        public string hDate
        {
            set { this._hDate = value; }
        }

        public string staffName
        {
            set { this._staffName = value; }
        }

        public int ID
        {
            set
            { this._ID = value; }
        }

        public string cName
        {
            set { this._cName = value; }
        }

        public string fJyouken
        {
            set { this._fJyouken = value; }
        }

        public string fKeitai
        {
            set { _fKeitai = value; }
        }

        public string Add
        {
            set { _Add = value; }
        }

        public double Tanka
        {
            get { return _Tanka; }
            set { _Tanka = value; }
        }

        public int yMaisu
        {
            set { _yMaisu = value; }
        }

        public int hMaisu
        {
            get { return _hMaisu; }
            set { _hMaisu = value; }
        }

        public int kanryo
        {
            get { return _kanryo; }
            set { _kanryo = value; }
        }

        public string Edaban
        {
            get { return _Edaban; }
            set { _Edaban = value; }
        }

        public int Edaban_Status
        {
            get { return _Edaban_Status; }
            set { _Edaban_Status = value; }
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

        private int _Mihaifu_Status;

        public int Mihaifu_Status
        {
            get { return _Mihaifu_Status; }
            set { _Mihaifu_Status = value; }
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtTanka_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtEdaban) { txtObj = txtEdaban; }

            //if (sender == txtBanchi) { txtObj = txtBanchi; }
            //if (sender == txtManshon) { txtObj = txtManshon; }
            
            //if (sender == txtRiyu) 
            //{
            //    txtObj = txtRiyu;

            //    if (txtRiyu.Text == "0")
            //    {
            //        txtRiyu.Text = "";
            //    }
            //}

            //if (sender == txtSonota) { txtObj = txtSonota; }

            if (sender == txtTanka) { txtObj = txtTanka; }
            if (sender == txthMaisu) { txtObj = txthMaisu; }

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;

        }

        private void txtTanka_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtEdaban) { txtObj = txtEdaban; }

            //if (sender == txtBanchi) { txtObj = txtBanchi; }
            //if (sender == txtManshon) { txtObj = txtManshon; }
            //if (sender == txtRiyu) { txtObj = txtRiyu; }
            //if (sender == txtSonota) { txtObj = txtSonota; }

            if (sender == txtTanka) { txtObj = txtTanka; }
            if (sender == txthMaisu) { txtObj = txthMaisu; }

            txtObj.BackColor = Color.White;

        }

        private void txtTanka_Validating(object sender, CancelEventArgs e)
        {
            string str;
            double d;

            if (txtTanka.Text == null)
            {
                MessageBox.Show("数字で入力してください",MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtTanka.Text;

            if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        private void txthMaisu_Validating(object sender, CancelEventArgs e)
        {
            string str;
            int d;

            if (txthMaisu.Text == null)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txthMaisu.Text;

            if (Int32.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if ((Int32.Parse(txthMaisu.Text, System.Globalization.NumberStyles.Any) != 0) && 
                (Int32.Parse(lblyMaisu.Text, System.Globalization.NumberStyles.Any) != Int32.Parse(txthMaisu.Text, System.Globalization.NumberStyles.Any)))
            {
                MessageBox.Show("予定枚数と一致していません", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //枝番有無チェック：2018/01/11
            if (_Edaban_Status == 1 && txtEdaban.Text.Trim() == string.Empty)
            {
                if (MessageBox.Show("枝番記入欄が未登録ですがよろしいですか？", "枝番該当案件", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                {
                    txtEdaban.Focus();
                    return;
                }
            }

            //単価
            _Tanka = Double.Parse(txtTanka.Text,System.Globalization.NumberStyles.Any);

            //配布枚数
            _hMaisu = Int32.Parse(txthMaisu.Text, System.Globalization.NumberStyles.Any);
        
            //完了区分
            if (checkBox1.Checked == true)
            {
                _kanryo = 1;
            }
            else
            {
                _kanryo = 0;
            }

            //枝番記入
            _Edaban = txtEdaban.Text;

            ////番地・号
            //_Banchi = txtBanchi.Text;

            ////マンション名
            //_Manshon = txtManshon.Text;

            ////理由
            //if (txtRiyu.Text == "")
            //{
            //    _Riyu = 0;
            //}
            //else
            //{
            //    _Riyu = int.Parse(txtRiyu.Text);
            //}

            ////その他
            //_Sonota = txtSonota.Text;

            Close();
        }

        private void txthMaisu_KeyDown(object sender, KeyEventArgs e)
        {
            if (txthMaisu.SelectedText == "0")
            {
                if (e.KeyCode == Keys.Return)
                {
                    txthMaisu.Text = lblyMaisu.Text;
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmHaifuShijiSUb3 frm3 = new frmHaifuShijiSUb3(MIHAIFU_ADD);
            frm3.FID = int.Parse(lblID.Text);       //配布エリアID
            frm3.Add = lblAdd.Text;     //配布先住所

            frm3.ShowDialog();

            //再表示
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void GridEnter()
        {
            frmHaifuShijiSUb3 frm3 = new frmHaifuShijiSUb3(MIHAIFU_UPDATE);
            frm3.ID = int.Parse(dataGridView1[0, dataGridView1.SelectedRows[0].Index].Value.ToString());
            frm3.Add = lblAdd.Text;     //配布先住所

            //編集画面表示
            frm3.ShowDialog();

            //再表示
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            GridEnter();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("未配布情報データが選択されていません", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("選択された " + dataGridView1.SelectedRows.Count.ToString() + "件の未配布情報を削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                dataGridView1.CurrentCell = null;
                return;
            }

            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                int aID;
                aID = int.Parse(dataGridView1[0, r.Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                //レコード削除
                Control.未配布情報 dArea = new Control.未配布情報();

                if (dArea.DataDelete(aID) == false)
                {
                    MessageBox.Show("削除に失敗しました。ID：" + aID.ToString(), MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                dArea.Close();
            }

            //未配布情報再表示
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));
 
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;

namespace posting
{
    public partial class frmTesuuryouMeisai : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.支給控除 cMaster = new Entity.支給控除();
        int ShowStatus;

        const string MESSAGE_CAPTION = "配布手数料支給・控除明細";

        public frmTesuuryouMeisai(int tempStatus)
        {
            InitializeComponent();

            ShowStatus = tempStatus;
        }

        private void form_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView1);
            GridviewSet.ShowData(dataGridView1); 

            DispClear();

            //配布指示画面からの呼び出しのとき
            if (ShowStatus == 1)
            {
                dateTimePicker1.Value = F_hDate;
                txthCode.Text = F_配布員ID.ToString();
                lblName.Text = F_配布員名;
            }

        }

        // データグリッドビュークラス
        private class GridviewSet
        {

            /// <summary>
            /// データグリッドビューの定義を行います
            /// </summary>
            /// <param name="tempDGV">データグリッドビューオブジェクト</param>
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
                    tempDGV.Height = 180;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "ｺｰﾄﾞ");
                    tempDGV.Columns.Add("col2", "日付");
                    tempDGV.Columns.Add("col3", "配布員ID");
                    tempDGV.Columns.Add("col4", "氏名");
                    tempDGV.Columns.Add("col5", "摘要");
                    tempDGV.Columns.Add("col6", "単価");
                    tempDGV.Columns.Add("col7", "数量");
                    tempDGV.Columns.Add("col8", "金額");
                    tempDGV.Columns.Add("col9", "支給控除");
                    tempDGV.Columns.Add("col10", "登録年月日");
                    tempDGV.Columns.Add("col11", "変更年月日");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 200;
                    tempDGV.Columns[5].Width = 70;
                    tempDGV.Columns[6].Width = 70;
                    tempDGV.Columns[7].Width = 70;
                    tempDGV.Columns[8].Width = 80;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;

                    tempDGV.Columns[1].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = false;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.支給控除 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.支給控除 sCon = new Control.支給控除();
                OleDbDataReader dr;

                sqlStr = " where 支給控除.ID = " + int.Parse(dgv[0, dgv.SelectedRows[iX].Index].Value.ToString());
                dr = sCon.FillBy(sqlStr);

                if (dr.HasRows == false)
                {

                    dr.Close();
                    sCon.Close();
                    return false;
                }

                while (dr.Read())
                {
                    tempC.ID = int.Parse(dr["ID"].ToString());
                    tempC.日付 = DateTime.Parse(dr["日付"].ToString());
                    tempC.配布員ID = int.Parse(dr["配布員ID"].ToString());
                    tempC.配布員名 = dr["氏名"].ToString();
                    tempC.摘要 = dr["摘要"].ToString() + "";
                    tempC.単価 = double.Parse(dr["単価"].ToString(),System.Globalization.NumberStyles.Any);
                    tempC.数量 = int.Parse(dr["数量"].ToString(),System.Globalization.NumberStyles.Any);
                    tempC.金額 = double.Parse(dr["金額"].ToString(),System.Globalization.NumberStyles.Any);
                    tempC.支給控除区分 = int.Parse(dr["支給控除区分"].ToString());
                }

                dr.Close();
                sCon.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV)
            {
                int iX;

                try
                {
                    tempDGV.RowCount = 0;

                    //支給控除マスターのデータリーダーを取得する
                    OleDbDataReader dR;
                    Control.支給控除 sCon = new Control.支給控除();
                    dR = sCon.FillBy("order by ID desc");
                    iX = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = dR["ID"].ToString();
                        tempDGV[1, iX].Value = DateTime.Parse(dR["日付"].ToString());
                        tempDGV[2, iX].Value = dR["配布員ID"].ToString();
                        tempDGV[3, iX].Value = dR["氏名"].ToString();
                        tempDGV[4, iX].Value = dR["摘要"].ToString();
                        tempDGV[5, iX].Value = double.Parse(dR["単価"].ToString()).ToString("#,##0.0");
                        tempDGV[6, iX].Value = int.Parse(dR["数量"].ToString()).ToString("#,##0");
                        tempDGV[7, iX].Value = double.Parse(dR["金額"].ToString()).ToString("#,##0.0");

                        switch (dR["支給控除区分"].ToString())
                        {
                            case "0":
                                tempDGV[8, iX].Value = "支給";
                                break;

                            case "1":
                                tempDGV[8, iX].Value = "控除";
                                break;
                        }

                        tempDGV[9, iX].Value = DateTime.Parse(dR["登録年月日"].ToString());
                        tempDGV[10, iX].Value = DateTime.Parse(dR["変更年月日"].ToString());

                        iX++;
                    }

                    dR.Close();

                    sCon.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }

            }

        }

        //グリッドからデータを選択
        private void GridEnter()
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[1, dataGridView1.SelectedRows[iX].Index].Value + "が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "データ選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {

                    //データを取得する
                    if (GridviewSet.GetData(dataGridView1,ref cMaster) == false)
                    {
                        MessageBox.Show("該当するデータがマスターに登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'データ値を取得
                    dateTimePicker1.Value = cMaster.日付;
                    txthCode.Text = cMaster.配布員ID.ToString();
                    lblName.Text = cMaster.配布員名;
                    txtName2.Text = cMaster.摘要;
                    txtTanka.Text = cMaster.単価.ToString();
                    txtSuuryou.Text = cMaster.数量.ToString();
                    txtkingaku.Text = cMaster.金額.ToString();

                    comboBox1.SelectedIndex = cMaster.支給控除区分;

                    //ボタン状態
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //フォームモードステータス:変更削除

                    comboBox1.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "画面クリア", MessageBoxButtons.OK);
                }
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            // Enterkey以外は対象外
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            GridEnter();
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            GridEnter();
        }

        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {

            try
            {
                fMode.Mode = 0;

                comboBox1.SelectedIndex = -1;
                txthCode.Text = "";
                lblName.Text = "";
                txtName2.Text = "";
                txtTanka.Text = "";
                txtSuuryou.Text = "";
                txtkingaku.Text = "0";

                btnDel.Enabled = false;
                btnClr.Enabled = false;

                if (this.dataGridView1.RowCount > 0)
                {
                    btnCsv.Enabled = true;
                }
                else
                {
                    btnCsv.Enabled = false;
                }

                comboBox1.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("選択されているデータを変更しないで終了します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
                
            DispClear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {

                if (fDataCheck() == true)
                {
                    Control.支給控除 sCon = new Control.支給控除();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                sCon.Close();
                                return;
                            }

                            if (sCon.DataInsert(cMaster) == true)
                            {
                                MessageBox.Show("新規登録されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("新規登録に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                sCon.Close();
                                return;
                            }

                            if (sCon.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", "支給控除", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    sCon.Close();

                    DispClear();

                    GridviewSet.ShowData(dataGridView1);

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"更新処理",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        //登録データチェック
        private Boolean fDataCheck()
        {

            try
            {

                //登録モードのとき、コードをチェック
                if (fMode.Mode == 0)
                {

                    //// 数字か？
                    //if (txtCode.Text == null)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("コードは数字で入力してください");
                    //}

                    //str = this.txtCode.Text;

                    //if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    //{
                    //}
                    //else
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("コードは数字で入力してください");
                    //}

                    //// 未入力またはスペースのみは不可
                    //if ((this.txtCode.Text).Trim().Length < 1)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("コードを入力してください");
                    //}

                    ////ゼロは不可
                    //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("ゼロは登録できません");
                    //}

                    ////登録済みコードか調べる
                    //string sqlStr;
                    //Control.所属 Shozoku = new Control.所属();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Shozoku.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Shozoku.Close();
                    //    throw new Exception("既に登録済みのコードです");
                    //}

                    //dr.Close();
                    //Shozoku.Close();

                }

                //名称チェック
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("支給控除を選択してください");
                }

                //配布員チェック
                if (txthCode.Text == "")
                {
                    txthCode.Focus();
                    throw new Exception("配布員IDを選択してください");
                }

                //クラスにデータセット
                cMaster.日付 = DateTime.Parse(dateTimePicker1.Text);
                cMaster.配布員ID = int.Parse(txthCode.Text);
                cMaster.摘要 = txtName2.Text.ToString();
                cMaster.単価 = double.Parse(txtTanka.Text,System.Globalization.NumberStyles.Any);
                cMaster.数量 = int.Parse(txtSuuryou.Text,System.Globalization.NumberStyles.Any);
                cMaster.金額 = double.Parse(txtkingaku.Text, System.Globalization.NumberStyles.Any);
                cMaster.支給控除区分 = comboBox1.SelectedIndex;

                if (fMode.Mode == 0) cMaster.登録年月日 = DateTime.Today;
                cMaster.変更年月日 = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
                
            }

        }

        private void btnDel_Click(object sender, EventArgs e)
        {   
            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.支給控除 sCon = new Control.支給控除();
            if (sCon.DataDelete(cMaster.ID)==true)
                MessageBox.Show("削除されました", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            sCon.Close();

            DispClear();

            GridviewSet.ShowData(dataGridView1);
        }

        private void btnCsv_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void txthCode_Validating(object sender, CancelEventArgs e)
        {
            if (txthCode.Text == "") return;

            if (Utility.NumericCheck(txthCode.Text) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            OleDbDataReader dR;
            Control.配布員 cStaff = new Control.配布員();
            dR = cStaff.FillBy("where ID = " + txthCode.Text);

            if (dR.HasRows == false)
            {
                lblName.Text = "";

                MessageBox.Show("登録されていない配布員IDです", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            else
            {
                while (dR.Read())
                {
                    lblName.Text = dR["氏名"].ToString();                    
                }
            }

            dR.Close();
            cStaff.Close();

            return;
        }

        private void txthCode_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txthCode) txtObj = txthCode;
            if (sender == txtName2) txtObj = txtName2;
            if (sender == txtTanka) txtObj = txtTanka;
            if (sender == txtSuuryou) txtObj = txtSuuryou;
            if (sender == txtkingaku) txtObj = txtkingaku;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void txthCode_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txthCode) txtObj = txthCode;
            if (sender == txtName2) txtObj = txtName2;
            if (sender == txtTanka) txtObj = txtTanka;
            if (sender == txtSuuryou) txtObj = txtSuuryou;
            if (sender == txtkingaku) txtObj = txtkingaku;

            txtObj.BackColor = Color.White;
        }

        private void txtTanka_Validating(object sender, CancelEventArgs e)
        {
            double kin;

            if (txtTanka.Text == "")
            {
                txtTanka.Text = "0";
                return;
            }

            if (Utility.NumericCheck(txtTanka.Text) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (Utility.NumericCheck(txtSuuryou.Text) == true)
            {
                kin = double.Parse(txtTanka.Text) * int.Parse(txtSuuryou.Text);
                txtkingaku.Text = kin.ToString("#,##0");
                return;
            }
            else
            {
                txtkingaku.Text = "0";
            }

        }

        private void txtSuuryou_Validating(object sender, CancelEventArgs e)
        {
            try
            {

                double kin;

                if (txtSuuryou.Text == "")
                {
                    txtSuuryou.Text = "0";
                    return;
                }

                if (Utility.NumericCheck(txtSuuryou.Text) == false)
                {
                    MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }

                if (Utility.NumericCheck(txtTanka.Text) == true)
                {
                    kin = double.Parse(txtTanka.Text) * int.Parse(txtSuuryou.Text);
                    txtkingaku.Text = kin.ToString("#,##0");
                    return;
                }
                else
                {
                    txtkingaku.Text = "0";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "単価入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSuuryou.Focus();
                return;
            }

        }

        //プロパティ
        private DateTime F_hDate;

        public DateTime 配布日
        {
            set { F_hDate = value; }
        }

        private int F_配布員ID;

        public int 配布員ID
        {
            set { F_配布員ID = value; }
        }

        private string F_配布員名;

        public string 配布員名
        {
            get { return F_配布員名; }
            set { F_配布員名 = value; }
        }

    }
}
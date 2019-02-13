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
    public partial class frmTenkou : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.天候 cMaster = new Entity.天候();

        const string MESSAGE_CAPTION = "天候入力";

        public frmTenkou()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            GridviewSet.Setting(dataGridView1);
            Utility.ComboTenkou.load(comboBox1);
            DispClear();

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
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 310;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "日付");
                    tempDGV.Columns.Add("col2", "天候");
                    tempDGV.Columns.Add("col3", "日付");
                    tempDGV.Columns.Add("col4", "天候");

                    tempDGV.Columns[0].Width = 110;
                    tempDGV.Columns[1].Width = 110;
                    tempDGV.Columns[2].Width = 110;
                    tempDGV.Columns[3].Width = 110;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.天候 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.天候 Tenkou = new Control.天候();
                OleDbDataReader dR;

                sqlStr = " where 天候.日付 = '" + dgv[0, dgv.SelectedRows[iX].Index].Value.ToString() + "'";
                dR = Tenkou.FillBy(sqlStr);

                if (dR.HasRows == true)
                {
                    while (dR.Read() == true)
                    {
                        tempC.日付 = DateTime.Parse(dR["日付"].ToString());
                        tempC.天候名 = dR["天候"].ToString() + "";
                    }
                }
                else
                {
                    dR.Close();
                    Tenkou.Close();
                    return false;
                }

                dR.Close();
                Tenkou.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV,int tempYear,int tempMonth)
            {
                string sqlSTRING = "";
                string rDate;
                int iDay,iX,c1,c2;
                DateTime sDate;

                try
                {
                    //天候データのデータリーダーを取得する
                    OleDbDataReader dR;
                    Control.天候 dCon = new Control.天候();

                    tempDGV.RowCount = 0;

                    iDay = 0;
                    iX = 0;

                    while (true)
                    {
                        iDay++;

                        rDate = tempYear.ToString() + "/" + tempMonth.ToString() + "/" + iDay.ToString();

                        if (DateTime.TryParse(rDate, out sDate) == true)
                        {
                            if (iX < 16)
                            {
                                tempDGV.Rows.Add();
                                c1 = 0;
                                c2 = 1;
                            }
                            else
                            {
                                c1 = 2;
                                c2 = 3;
                            }

                            
                            tempDGV[c1, iX%16].Value = rDate + "(" + ("日月火水木金土").Substring(int.Parse(sDate.DayOfWeek.ToString("d")), 1) + ")";
                            tempDGV[c2, iX%16].Value = "";

                            sqlSTRING = "where 日付 = '" + rDate + "'";

                            dR = dCon.FillBy(sqlSTRING);

                            while (dR.Read())
                            {
                                tempDGV[c2, iX%16].Value = dR["天候"].ToString();
                            }

                            dR.Close();
                        }
                        else
                        {
                            break;
                        }

                        iX++;
                    }

                    dCon.Close();

                    tempDGV.CurrentCell = null;

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
            msgStr += dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value + "が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    //txtDATE.Text = cMaster.日付.ToShortDateString();

                    //ボタン状態
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

                txtYear.Text = "";
                txtMonth.Text = "";
                dateTimePicker1.Text = "";
                dateTimePicker1.Enabled = false;
                comboBox1.Text = "";
                comboBox1.Enabled = false;

                btnUpdate.Enabled = false;

                dateTimePicker1.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("選択されているデータを破棄します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question)== DialogResult.No )
                return;
                
            DispClear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {

                if (fDataCheck() == true)
                {
                    Control.天候 cTenkou = new Control.天候();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                cTenkou.Close();
                                return;
                            }

                            if (cTenkou.DataInsert(cMaster) == true)
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
                                cTenkou.Close();
                                return;
                            }

                            if (cTenkou.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    cTenkou.Close();

                    //日付別天候表示
                    GridviewSet.ShowData(dataGridView1, int.Parse(txtYear.Text), int.Parse(txtMonth.Text));

                    //天候コンボリロード
                    Utility.ComboTenkou.load(comboBox1);

                    comboBox1.Text = "";

                    //DispClear();

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
                    //if (txtDATE.Text == null)
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("コードは数字で入力してください");
                    //}

                    //str = this.txtDATE.Text;

                    //if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    //{
                    //}
                    //else
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("コードは数字で入力してください");
                    //}

                    //// 未入力またはスペースのみは不可
                    //if ((this.txtDATE.Text).Trim().Length < 1)
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("コードを入力してください");
                    //}

                    ////ゼロは不可
                    //if (Convert.ToInt32(this.txtDATE.Text.ToString()) == 0)
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("ゼロは登録できません");
                    //}

                    ////登録済みコードか調べる
                    //string sqlStr;
                    //Control.所属 Shozoku = new Control.所属();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtDATE.Text.ToString();
                    //dr = Shozoku.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtDATE.Focus();
                    //    dr.Close();
                    //    Shozoku.Close();
                    //    throw new Exception("既に登録済みのコードです");
                    //}

                    //dr.Close();
                    //Shozoku.Close();

                }

                //名称チェック
                if (comboBox1.Text.Trim().Length < 1)
                {
                    comboBox1.Focus();
                    throw new Exception("天候を入力してください");
                }

                //所属クラスにデータセット
                cMaster.日付 = dateTimePicker1.Value;
                cMaster.天候名  = comboBox1.Text.ToString();

                if (fMode.Mode == 0) cMaster.登録年月日 = DateTime.Today;
                cMaster.変更年月日 = DateTime.Today;

                //登録済みか？
                OleDbDataReader dR;
                Control.天候 cTenkou = new Control.天候();
                dR = cTenkou.FillBy("where 日付 = '" + dateTimePicker1.Text + "'");

                if (dR.HasRows == true)
                {
                    fMode.Mode = 1;
                }
                else
                {
                    fMode.Mode = 0;
                }

                dR.Close();
                cTenkou.Close();

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void txtEnter(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtYear)
            {
                objtxt = txtYear;
            }

            if (sender == txtMonth)
            {
                objtxt = txtMonth;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtYear)
            {
                objtxt = txtYear;
            }

            if (sender == txtMonth)
            {
                objtxt = txtMonth;
            }

            objtxt.BackColor = Color.White;

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

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (Utility.NumericCheck(txtYear.Text) == false)
                {
                    txtYear.Focus();
                    throw new Exception("年は数字で入力してください");
                }

                if (Utility.NumericCheck(txtMonth.Text) == false)
                {
                    txtMonth.Focus();
                    throw new Exception("月は数字で入力してください");
                }

                if ((int.Parse(txtMonth.Text) < 1) || (int.Parse(txtMonth.Text) > 12))
                {
                    txtMonth.Focus();
                    throw new Exception("月が正しくありません");
                }

                //カレンダー対象月表示
                dateTimePicker1.Enabled = true;
                dateTimePicker1.Value = DateTime.Parse(txtYear.Text + "/" + txtMonth.Text + "/01");

                //日付別天候表示
                GridviewSet.ShowData(dataGridView1, int.Parse(txtYear.Text), int.Parse(txtMonth.Text));

                //入力開放
                dateTimePicker1.Enabled = true;
                comboBox1.Enabled = true;
                btnUpdate.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void dateTimePicker1_Validating(object sender, CancelEventArgs e)
        {
            if ((dateTimePicker1.Value).Year.ToString() != txtYear.Text || (dateTimePicker1.Value).Month.ToString() != txtMonth.Text)
            {
                MessageBox.Show("対象年月と異なります", "年月", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Cancel = true;
                dateTimePicker1.Focus();
            }
        }

    }
}
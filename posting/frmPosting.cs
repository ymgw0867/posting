using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using MyLibrary;

namespace posting
{
    public partial class frmPosting : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Utility.areaMode aMode = new Utility.areaMode();
        Entity.受注 cMaster = new Entity.受注();
        Entity.配布エリア cArea = new Entity.配布エリア();

        const string MESSAGE_CAPTION = "ポスティングエリア";

        public frmPosting()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: このコード行はデータを 'darwinDataSet.受注' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);
            this.darwinDataSet.Clear();
            this.darwinDataSet.EnforceConstraints = false;
            this.受注TableAdapter.FillByPosting(darwinDataSet.受注);

            //if (dataGridView1.RowCount <= 8)
            //{
            //    dataGridView1.Columns[3].Width = 288;
            //}
            //else
            //{
            //    dataGridView1.Columns[3].Width = 271;
            //}

           

            //ポスティングエリア
            GridviewSet.AriaSetting(dataGridView2);

            //町名マスタ
            GridviewSet.TownSetting(dataGridView3);

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
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 163;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    //tempDGV.Columns.Add("col1", "ｺｰﾄﾞ");
                    //tempDGV.Columns.Add("col2", "名称");
                    //tempDGV.Columns.Add("col3", "備考");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 280;
                    tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static void AriaSetting(DataGridView tempDGV)
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
                    //tempDGV.Height = 230;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "ID");
                    tempDGV.Columns.Add("col2", "エリアID");
                    tempDGV.Columns.Add("col3", "配布エリア");
                    tempDGV.Columns.Add("col4", "予定枚数");
                    tempDGV.Columns.Add("col5", "指示№");

                    tempDGV.Columns[0].Visible = false;

                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 70;

                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = true;

                    // 編集制御
                    tempDGV.ReadOnly = true;
                    //tempDGV.Columns[0].ReadOnly = true;
                    //tempDGV.Columns[1].ReadOnly = true;
                    //tempDGV.Columns[2].ReadOnly = true;

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

            public static void TownSetting(DataGridView tempDGV)
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
                    //tempDGV.Height = 140;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "エリアID");
                    tempDGV.Columns.Add("col2", "町名");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static Boolean GetPostingData(int tempID, ref Entity.配布エリア tempC)
            {
                string sqlStr;

                Control.配布エリア dArea = new Control.配布エリア();
                OleDbDataReader dr;

                sqlStr = " where 配布エリア.ID = " + tempID.ToString();
                dr = dArea.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.町名ID = Int32.Parse(dr["町名ID"].ToString());
                        tempC.予定枚数 = Int32.Parse(dr["予定枚数"].ToString());
                        tempC.受注ID = long.Parse(dr["受注ID"].ToString());
                        tempC.配布指示ID = Int32.Parse(dr["配布指示ID"].ToString());
                        tempC.配布単価 =double.Parse(dr["配布単価"].ToString(), System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.配布日 = dr["配布日"].ToString();
                        tempC.実配布枚数 = Int32.Parse(dr["実配布枚数"].ToString());
                        tempC.実残数 = Int32.Parse(dr["実残数"].ToString());
                        tempC.報告枚数 = Int32.Parse(dr["報告枚数"].ToString());
                        tempC.報告残数 = Int32.Parse(dr["報告残数"].ToString());
                        tempC.併配区分 = Int32.Parse(dr["併配区分"].ToString());
                        tempC.枝番記入 = dr["枝番記入"].ToString();
                        tempC.完了区分 = Int32.Parse(dr["完了区分"].ToString());
                        tempC.ステータス = Int32.Parse(dr["ステータス"].ToString());
                    }
                }
                else
                {
                    dr.Close();
                    dArea.Close();
                    return false;
                }

                dr.Close();
                dArea.Close();
                return true;
            }


            public static void AreaShowData(DataGridView tempDGV,long tempID)
            {
                int iX;

                try
                {
                    tempDGV.RowCount = 0;

                    //配布エリアデータを取得する
                    OleDbDataReader dr;
                    Control.配布エリア dArea = new Control.配布エリア();
                    dr = dArea.FillBy("where 受注ID = " + tempID.ToString() + " order by 町名ID");

                    iX = 0;

                    while (dr.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = long.Parse(dr["ID"].ToString());
                        tempDGV[1, iX].Value = dr["町名ID"].ToString();
                        tempDGV[3, iX].Value = Int32.Parse(dr["予定枚数"].ToString());
                        tempDGV[4, iX].Value = Int32.Parse(dr["配布指示ID"].ToString());
                        iX++;
                    }

                    tempDGV.CurrentCell = null;

                    //if (tempDGV.RowCount <= 18)
                    //{
                    //    tempDGV.Columns[2].Width = 218;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[2].Width = 201;
                    //}

                    dr.Close();
                    dArea.Close();

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

            try
            {

                if (MessageBox.Show(dataGridView1[3, dataGridView1.SelectedRows[0].Index].Value.ToString() + " が選択されました。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                DispClear();

                //データを取得する
                OleDbDataReader dr;
                Control.受注 cOrder = new Control.受注();
                dr = cOrder.FillBy("where ID = " + dataGridView1[0,dataGridView1.SelectedRows[0].Index].Value.ToString());
                
                if (dr.HasRows == false)
                {
                    MessageBox.Show("該当するデータが登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                //'データ値を取得
                while (dr.Read())
                {
                    txtID.Text = dr["ID"].ToString();
                    txtCName.Text = "";
                    label8.Text = int.Parse(dr["枚数"].ToString()).ToString("#,##0");
                    label11.Text = double.Parse(dr["配布単価"].ToString(),System.Globalization.NumberStyles.Any).ToString("#,##0.00");

                    //得意先名
                    OleDbDataReader drt;
                    Control.得意先 Client = new Control.得意先();
                    drt = Client.FillBy("where ID = " + dr["得意先ID"].ToString());

                    while (drt.Read())
                    {
                        txtCName.Text = drt["略称"].ToString();
                    }

                    drt.Close();

                    txtChirashi.Text = dr["チラシ名"].ToString();   
                }

                dr.Close();

                cOrder.Close();

                //配布エリアデータ表示
                GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
                MaisuSum();

                //txtTotal.Text = GetMaisuTotal().ToString("#,##0");
                //int Zan;
                //Zan = Int32.Parse(textBox2.Text, System.Globalization.NumberStyles.Any) - Int32.Parse(txtTotal.Text, System.Globalization.NumberStyles.Any);
                //textBox3.Text = Zan.ToString("#,##0");
      
                //ボタン表示
                txtAdd.Enabled = true;
                txtAdel.Enabled = true;
                txtAclear.Enabled = true;

                button2.Enabled = true;
                button4.Enabled = true;

                tabPage3.Text = txtChirashi.Text + " ： ポスティングエリア表";

                txtAreaID.Enabled = true;
                txtAreaName.Enabled = true;
                txtHaihuMaisu.Enabled = true;
                textBox5.Enabled = true;

                txtAreaID.Focus();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "データ表示", MessageBoxButtons.OK);
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            if (dataGridView1.Rows.Count == 0) return;

            // Enterkey以外は対象外
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            GridEnter();
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //GridEnter();
        }

        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {

            try
            {
                fMode.Mode = 0;
                aMode.Mode = 0;

                txtID.Text = "";
                txtCName.Text = "";
                txtChirashi.Text = "";

                txtAreaID.Text = "";
                txtAreaName.Text = "";
                txtHaihuMaisu.Text = "";

                txtAreaID.Enabled = false;
                txtAreaName.Enabled = false;
                txtHaihuMaisu.Enabled = false;

                dataGridView2.Rows.Clear();
                dataGridView3.Rows.Clear();

                label11.Text = "";
                label8.Text = "";
                label7.Text = "";
                label10.Text = "";
                textBox5.Text = "";
                textBox5.Enabled = false;

                txtAdel.Enabled = false;

                button2.Enabled = false;
                button5.Enabled = false;
                button4.Enabled = false;

                txtAdd.Enabled = false;
                txtAdel.Enabled = false;
                txtAclear.Enabled = false;

                tabPage3.Text = "ポスティングエリア表";

                dataGridView1.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }
        
        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear2()
        {

            try
            {
                aMode.Mode = 0;

                txtAreaID.Text = "";
                txtAreaName.Text = "";
                txtHaihuMaisu.Text = "";

                textBox5.Text = "";
                txtAdel.Enabled = false;

                txtAreaID.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        //登録データチェック
        private Boolean fDataCheck()
        {
            string str;
            int d;

            try
            {

                //エリアIDチェック
                if (txtAreaID.Text == null)
                {
                    this.txtAreaID.Focus();
                    throw new Exception("エリアIDは数字で入力してください");
                }

                str = txtAreaID.Text;

                if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                {
                    this.txtAreaID.Focus();
                    throw new Exception("エリアIDが正しくありません");
                }

                if (Int32.Parse(txtAreaID.Text.ToString()) < 0)
                {
                    this.txtAreaID.Focus();
                    throw new Exception("エリアIDが正しくありません");
                }

                //配布枚数チェック
                if (txtHaihuMaisu.Text == null)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("配布枚数は数字で入力してください");
                }

                str = txtHaihuMaisu.Text;

                if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("配布枚数が正しくありません");
                }

                if (Int32.Parse(txtHaihuMaisu.Text.ToString(),System.Globalization.NumberStyles.AllowThousands) < 0)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("配布枚数が正しくありません");
                }

                //クラスにデータセット
                cArea.町名ID = Int32.Parse(txtAreaID.Text.ToString());
                cArea.予定枚数 = Int32.Parse(txtHaihuMaisu.Text.ToString(),System.Globalization.NumberStyles.AllowThousands);
                cArea.受注ID = long.Parse(txtID.Text.ToString());

                if (aMode.Mode == 0)
                {
                    cArea.配布指示ID = 0;

                    if (Utility.NumericCheck(label11.Text) == false)
                    {
                        cArea.配布単価 = 0;
                    }
                    else
                    {
                        cArea.配布単価 = double.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                    }
                    
                    cArea.配布日 = "";
                    cArea.実配布枚数 = 0;
                    cArea.実残数 = cArea.予定枚数;
                    cArea.報告枚数 = 0;
                    cArea.報告残数 = cArea.予定枚数;
                    cArea.併配区分 = 0;
                    cArea.完了区分 = 0;
                    cArea.ステータス = 0;
                    cArea.登録年月日 = DateTime.Today;
                }
                else
                {
                    cArea.実残数 = cArea.予定枚数 - cArea.実配布枚数;
                    cArea.報告残数 = cArea.予定枚数 - cArea.報告枚数;
                }

                cArea.変更年月日 = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void txtEnter(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == txtAreaID)
            {
                objtxt = txtAreaID;
            }

            if (sender == txtHaihuMaisu)
            {
                objtxt = txtHaihuMaisu;
            }

            if (sender == textBox5)
            {
                objtxt = textBox5;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {

           TextBox objtxt = new TextBox();

           string str;
           double d;

           try
           {
               if (sender == textBox1)
               {
                   objtxt = textBox1;
               }

               //配布エリアID
               if (sender == txtAreaID)
               {
                   objtxt = txtAreaID;

                   if (txtAreaID.Text == null) txtAreaID.Text = "0";

                   str = txtAreaID.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtAreaID.Text = "0";

                   txtAreaName.Text = GetTownName(txtAreaID.Text.ToString());
               }

               if (sender == txtHaihuMaisu)
               {
                   objtxt = txtHaihuMaisu;
               }

               if (sender == textBox5)
               {
                   objtxt = textBox5;
               }

               objtxt.BackColor = Color.White;
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message,"エラーメッセージ");
           }
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


        private void button1_Click(object sender, EventArgs e)
        {
            darwinDataSet ds = new darwinDataSet();
            ds.Clear();
            ds.EnforceConstraints = false;
            this.受注TableAdapter.FillByPostingName(ds.受注, "%" + textBox1.Text.ToString() + "%");
            dataGridView1.DataSource = ds.受注;

            if (dataGridView1.RowCount <= 8)
            {
                dataGridView1.Columns[3].Width = 288;
            }
            else
            {
                dataGridView1.Columns[3].Width = 271;
            }

            DispClear();
        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void txtAdd_Click(object sender, EventArgs e)
        {
            //配布エリア登録
            try
            {
                if (fDataCheck() == true)
                {

                    Control.配布エリア dArea = new Control.配布エリア();

                    switch (aMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                break;

                            if (dArea.DataInsert(cArea) == false)
                            {
                                MessageBox.Show("新規登録に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                break;

                            if (dArea.DataUpdate(cArea) == false)
                            {
                                MessageBox.Show("更新に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    dArea.Close();

                    DispClear2();
                    
                    txtAreaID.Focus();

                    //配布エリア再表示
                    GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
                    MaisuSum();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex != 1) return;

            dataGridView2[2, e.RowIndex].Value = GetTownName(dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString());
        }

        private string GetTownName(string tempID)
        {
            //配布エリア町名検索
            string strName = ""; 
            OleDbDataReader dr;

            Control.町名 cTown = new Control.町名();
            dr = cTown.FillBy("where ID = " + tempID);

            while (dr.Read())
            {
                strName = dr["名称"].ToString();
            }

            dr.Close();
            return strName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                //配布エリア町名検索
                OleDbDataReader dr;
                int iX = 0;

                Control.町名 cTown = new Control.町名();
                dr = cTown.FillBy("where 名称 like '%" + textBox5.Text.ToString() + "%' order by ID");
                dataGridView3.Rows.Clear();

                while (dr.Read())
                {
                    dataGridView3.Rows.Add();
                    dataGridView3[0, iX].Value = dr["ID"];
                    dataGridView3[1, iX].Value = dr["名称"];
                    iX++;
                }

                //if (dataGridView3.RowCount <= 11)
                //{
                //    dataGridView3.Columns[1].Width = 280;
                //}
                //else
                //{
                //    dataGridView3.Columns[1].Width = 263;
                //}

                dr.Close();
                cTown.Close();

                dataGridView3.Focus();
                dataGridView3.CurrentCell = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        //配布エリアの町名を一括登録する
        {
            if (MessageBox.Show("選択されている町名を一括追加登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView3.SelectedRows)
            {
                txtAreaID.Text = dataGridView3[0, r.Index].Value.ToString();
                txtHaihuMaisu.Text = "0";

                if (fDataCheck() == true)
                {
                    Control.配布エリア dArea = new Control.配布エリア();
                    if (dArea.DataInsert(cArea) == false)
                    {
                        MessageBox.Show(dataGridView3[1, r.Index].Value.ToString() + "の新規登録に失敗しました", "町名一括登録", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    dArea.Close();
                }

            }

            DispClear2();

            txtAreaID.Focus();

            //配布エリア再表示
            GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
            MaisuSum();
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            PostingGridEnter();
        }

        private void PostingGridEnter()
        {
            if (MessageBox.Show(dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString() + " が選択されました。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            int iX = 0;

            txtAreaID.Text = dataGridView2[1, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtAreaName.Text = dataGridView2[2, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtHaihuMaisu.Text = dataGridView2[3, dataGridView2.SelectedRows[iX].Index].Value.ToString();

            aMode.Mode = 1;
            aMode.RowIndex = dataGridView2.SelectedRows[iX].Index;
            txtAdel.Enabled = true;

            //配布エリアデータを取得
            GridviewSet.GetPostingData(Int32.Parse(dataGridView2[0, dataGridView2.SelectedRows[iX].Index].Value.ToString()), ref cArea);

            //エリアIDコード
            txtAreaID.Focus();
        }

        private void TownGridEnter()
        {
            if (MessageBox.Show(dataGridView3[1, dataGridView3.SelectedRows[0].Index].Value.ToString() + " が選択されました。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            txtAreaID.Text = dataGridView3[0, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtAreaName.Text = dataGridView3[1, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtHaihuMaisu.Focus();
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            TownGridEnter();
        }

        private int GetMaisuTotal()
        {
            int Total = 0;

            for (int i = 0; i <= dataGridView2.Rows.Count - 1; i++)
            {
                Total += Int32.Parse(dataGridView2[3, i].Value.ToString(), System.Globalization.NumberStyles.AllowThousands);
            }

            return Total;
        }

        private void txtAdel_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("ポスティングエリアデータが選択されていません", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("選択された " + dataGridView2.SelectedRows.Count.ToString() + "件のポスティングエリアデータを削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                dataGridView2.CurrentCell = null;
                return;
            }

            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                int aID;
                aID = int.Parse(dataGridView2[0, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                //レコード削除
                Control.配布エリア dArea = new Control.配布エリア();

                if (dArea.DataDelete(aID) == false)
                {
                    MessageBox.Show("削除に失敗しました。エリアID：" + aID.ToString(), MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                dArea.Close();
            }

            DispClear2();

            //配布エリア再表示
            GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
            MaisuSum();




            ////削除
            //if (MessageBox.Show(txtAreaName.Text + " を削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            //    return;

            ////データ削除
            //Control.配布エリア dArea = new Control.配布エリア();
            //if (dArea.DataDelete(cArea.ID) == true)
            //{
            //    MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //else
            //{
            //    MessageBox.Show("削除に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}

            //dArea.Close();

            //DispClear2();

            ////配布エリア再表示
            //GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
            //MaisuSum();

        }

        private void txtAclear_Click(object sender, EventArgs e)
        {
            DispClear();
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (dataGridView2.Rows.Count == 0) return;
            if (dataGridView2.SelectedRows.Count == 0) return;

            if (e.KeyCode.ToString() == "Return")
            {
                PostingGridEnter();
            }
        }

        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (dataGridView3.Rows.Count == 0) return;
            if (dataGridView3.SelectedRows.Count == 0) return;

            if (e.KeyCode.ToString() == "Return")
            {
                TownGridEnter();
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            button5.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GetExcelPos();

        }

        private void GetExcelPos()
        {

            DialogResult ret;

            //ダイアログボックスの初期設定
            openFileDialog1.Title = "ポスティングエリア表の選択";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel) return;

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "ポスティングエリアExcelシート取り込み", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int S_GYO = 2;    //エクセルファイル見出し行（明細は2行目から）

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            //string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            Excel.Range dRng;
            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            int iX = S_GYO;
            int Cnt = 0;
            int err;
            string cellID;
            string cellMaisu;
            int d;

            try
            {

                while (true) 
                {
                    err = 0;

                    //エリアID
                    dRng = (Excel.Range)oxlsSheet.Cells[iX,1];

                    //空白なら処理終了
                    if ((dRng.Text.ToString().Trim() + "") == "")
                        break;

                    cellID = dRng.Text.ToString().Trim();

                    //IDチェック
                    if (cellID == null)
                    {
                        err = 1;
                        MessageBox.Show("エリアIDが正しくありません。この行はスキップされます。　" + Environment.NewLine + iX.ToString() + "行目　：　" + cellID, "取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (int.TryParse(cellID, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                    {
                        err = 1;
                        MessageBox.Show("エリアIDが正しくありません。この行はスキップされます。　" + Environment.NewLine + iX.ToString() + "行目　：　" + cellID, "取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (Int32.Parse(cellID, System.Globalization.NumberStyles.Any) < 0)
                    {
                        err = 1;
                        MessageBox.Show("エリアIDが正しくありません。この行はスキップされます。　" + Environment.NewLine + iX.ToString() + "行目　：　" + cellID, "取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    //配布枚数
                    dRng = (Excel.Range)oxlsSheet.Cells[iX,3];
                    cellMaisu = dRng.Text.ToString().Trim();

                    //チェック
                    if (cellMaisu == null)
                    {
                        err = 1;
                        MessageBox.Show("配布枚数が正しくありません。この行はスキップされます。　" + Environment.NewLine + iX.ToString() + "行目　：　" + cellMaisu, "取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (int.TryParse(cellMaisu, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                    {
                        err = 1;
                        MessageBox.Show("配布枚数が正しくありません。この行はスキップされます。　" + Environment.NewLine + iX.ToString() + "行目　：　" + cellMaisu, "取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (Int32.Parse(cellMaisu,System.Globalization.NumberStyles.Any) < 0)
                    {
                        err = 1;
                        MessageBox.Show("配布枚数が正しくありません。この行はスキップされます。　" + Environment.NewLine + iX.ToString() + "行目　：　" + cellMaisu, "取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    //エラーのときは読み飛ばし
                    if (err == 0)
                    {

                        txtAreaID.Text = cellID.ToString();
                        txtHaihuMaisu.Text = cellMaisu.ToString();

                        if (fDataCheck() == true)
                        {
                            Control.配布エリア dArea = new Control.配布エリア();
                            if (dArea.DataInsert(cArea) == false)
                            {
                                MessageBox.Show("新規登録に失敗しました", "Excelシート取り込み", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                Cnt++;
                            }

                            dArea.Close();
                        }

                    }

                    iX++;
                }

                MessageBox.Show(Cnt.ToString() + " 件の配布エリアを取り込みました", "取り込み終了", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // 確認のためExcelのウィンドウを表示する
                //oXls.Visible = true;

                //印刷
                //oxlsSheet.PrintPreview(true);

                //保存処理
                oXls.DisplayAlerts = false;

                //Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                //Excelを終了
                oXls.Quit();
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            finally
            {
                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;
            }


            DispClear2();

            txtAreaID.Focus();

            //配布エリア再表示
            GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));

            MaisuSum();
        }

        private void MaisuSum()
        {
            label10.Text = GetMaisuTotal().ToString("#,##0");
            int Zan;
            Zan = Int32.Parse(label8.Text, System.Globalization.NumberStyles.Any) - Int32.Parse(label10.Text, System.Globalization.NumberStyles.Any);
            label7.Text = Zan.ToString("#,##0");

            if (Zan < 0)
            {
                label7.ForeColor = Color.Red;
                MessageBox.Show("ポスティングエリア枚数が配布枚数を超えています", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                label7.ForeColor = Color.Black;
            }

        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            txtAdel.Enabled = true;
        }

    }
}
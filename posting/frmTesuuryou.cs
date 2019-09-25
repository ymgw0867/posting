using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Linq;

namespace posting
{
    public partial class frmTesuuryou : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.会社情報 sMaster = new Entity.会社情報();
        Entity.全銀トレーラーレコード sZtr = new Entity.全銀トレーラーレコード();

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.配布員TableAdapter adp = new darwinDataSetTableAdapters.配布員TableAdapter();

        const string MESSAGE_CAPTION = "配布手数料明細";

        int[] fIDArray;

        public frmTesuuryou()
        {
            InitializeComponent();

            // 配布員データ読み込み
            adp.Fill(dts.配布員);
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView2);
            button1.Enabled = false;

            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            DispClear();

            // 配布員指定コンボボックス値セット 2016/02/01
            addCmbHaihuItems();
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     配布員指定コンボボックス値セット </summary>
        ///---------------------------------------------------------------
        private void addCmbHaihuItems()
        {
            comboBox2.Items.Add("全員");
            comboBox2.Items.Add("配布員指定");

            // 事業所マスターの事業所名を追加
            darwinDataSetTableAdapters.事業所TableAdapter jAdp = new darwinDataSetTableAdapters.事業所TableAdapter();
            jAdp.Fill(dts.事業所);

            int iX =0;

            foreach (var t in dts.事業所.OrderBy(a => a.ID))
            {
                comboBox2.Items.Add(t.名称);

                iX++;
                Array.Resize(ref fIDArray, iX);
                fIDArray[iX - 1] = t.ID;
            }

            comboBox2.SelectedIndex = 0;
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
                    tempDGV.Height = 505;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "事業所");
                    tempDGV.Columns.Add("col2", "配布員ID");
                    tempDGV.Columns.Add("col3", "配布員名");
                    tempDGV.Columns.Add("col4", "指示番号");
                    tempDGV.Columns.Add("col5", "日付");
                    tempDGV.Columns.Add("col6", "チラシ名");
                    tempDGV.Columns.Add("col7", "配布エリア");
                    tempDGV.Columns.Add("col8", "単価");
                    tempDGV.Columns.Add("col9", "配布数");
                    tempDGV.Columns.Add("col10", "合計");

                    tempDGV.Columns[0].Width = 70;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 90;
                    //tempDGV.Columns[5].Width = 200;
                    //tempDGV.Columns[6].Width = 200;
                    tempDGV.Columns[7].Width = 70;
                    tempDGV.Columns[8].Width = 80;
                    tempDGV.Columns[9].Width = 80;

                    tempDGV.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV, int tempYear, int tempMonth, string sDate, string eDate, string tempSKbn,int tempHIndex,int tempHID)
            {

                string sqlSTRING = "";
                int iX;

                try
                {
                    //カーソルを待機表示
                    Cursor.Current = Cursors.WaitCursor;

                    tempDGV.RowCount = 0;
                    
                    //データリーダーを取得する
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select * from ";
                    sqlSTRING += "(";

                    //配布実績明細
                    sqlSTRING += "select 配布指示.ID as 配布指示ID,配布指示.配布日,事業所.ID as 事業所ID,";
                    sqlSTRING += "事業所.名称 as 事業所名称, 配布員.ID as 配布員ID,";
                    sqlSTRING += "配布員.氏名, 得意先.ID as 得意先ID, 得意先.略称 as 得意先略称,";
                    sqlSTRING += "受注.チラシ名, 町名.名称, 配布エリア.配布単価 as 単価,";
                    sqlSTRING += "配布エリア.実配布枚数,";
                    sqlSTRING += "配布エリア.配布単価 * 配布エリア.実配布枚数 as 金額,受注.ID,'0' as sortID  ";
                    sqlSTRING += "from 配布指示 inner join 配布エリア ";
                    sqlSTRING += "on 配布指示.ID = 配布エリア.配布指示ID inner join 受注 ";
                    sqlSTRING += "on 配布エリア.受注ID = 受注.ID left join 配布員 ";
                    sqlSTRING += "on 配布指示.配布員ID = 配布員.ID left join 町名 ";
                    sqlSTRING += "on 配布エリア.町名ID = 町名.ID left join 事業所 ";
                    sqlSTRING += "on 配布員.事業所コード = 事業所.ID left join 得意先 ";
                    sqlSTRING += "on 受注.得意先ID = 得意先.ID ";

                    //週又は月
                    switch (tempSKbn)
                    {
                        case "月":
                            sqlSTRING += "where (year(配布指示.入力日) = ?) and (month(配布指示.入力日) = ?) ";
                            break;

                        case "週":
                            sqlSTRING += "where ((配布指示.入力日) >= ?) and ((配布指示.入力日) <= ?) ";
                            break;
                    }

                    sqlSTRING += "and (配布員.支払区分 = ?) ";

                    //配布員指定
                    if (tempHIndex == 1)
                    {
                            sqlSTRING += "and (配布員.ID = ?) ";
                    }
                    else if (tempHIndex >= 2)
                    {
                        // 2016/02/01
                        sqlSTRING += "and (事業所.ID = ?) ";
                    }


                    //交通費
                    sqlSTRING += "union ";

                    sqlSTRING += "select 配布指示.ID AS 配布指示ID,配布指示.配布日,事業所.ID AS 事業所ID,";
                    sqlSTRING += "事業所.名称 AS 事業所名称,配布員.ID AS 配布員ID,配布員.氏名,";
                    sqlSTRING += "'0' AS 得意先ID,'' AS 得意先略称,'交通費' AS チラシ名,'' AS 名称,";
                    sqlSTRING += "配布指示.交通費 AS 単価,'1' AS 実配布枚数,配布指示.交通費 AS 金額,";
                    sqlSTRING += "'999999999999' AS ID, '0' AS sortID ";

                    sqlSTRING += "FROM 配布指示 LEFT OUTER JOIN ";
                    sqlSTRING += "配布員 ON 配布指示.配布員ID = 配布員.ID LEFT OUTER JOIN ";
                    sqlSTRING += "事業所 ON 配布員.事業所コード = 事業所.ID ";
                    sqlSTRING += "WHERE (配布指示.交通費 > 0) ";

                    //週又は月
                    switch (tempSKbn)
                    {
                        case "月":
                            sqlSTRING += "and  (year(配布指示.入力日) = ?) and (month(配布指示.入力日) = ?) ";
                            break;

                        case "週":
                            sqlSTRING += "and  ((配布指示.入力日) >= ?) and ((配布指示.入力日) <= ?) ";
                            break;
                    }

                    sqlSTRING += "and (配布員.支払区分 = ?) ";

                    //配布員指定
                    if (tempHIndex == 1)
                    {
                        sqlSTRING += "and (配布員.ID = ?) ";
                    }
                    else if (tempHIndex >= 2)
                    {
                        // 2016/02/01
                        sqlSTRING += "and (事業所.ID = ?) ";
                    }

                    //支給・控除明細
                    sqlSTRING += "union ";

                    sqlSTRING += "select '0' as 配布指示ID,支給控除.日付 as 配布日,事業所.ID as 事業所ID,";
                    sqlSTRING += "事業所.名称 as 事業所名称,配布員.ID as 配布員ID,";
                    sqlSTRING += "配布員.氏名,0 as 得意先ID,'' as 得意先略称,";
                    sqlSTRING += "支給控除.摘要 as チラシ名,'' as 名称,";

                    //単価 : 控除項目のときは(-1)
                    sqlSTRING += "case 支給控除.支給控除区分 when 0 ";
                    sqlSTRING += "then 支給控除.単価 ";
                    sqlSTRING += "else 支給控除.単価 * (-1) ";
                    sqlSTRING += "end,";

                    sqlSTRING += "支給控除.数量 as 実配布枚数,";

                    //金額 : 控除項目のときは(-1)
                    sqlSTRING += "case 支給控除.支給控除区分 when 0 ";
                    sqlSTRING += "then 支給控除.金額 ";
                    sqlSTRING += "else 支給控除.金額 * (-1) ";
                    sqlSTRING += "end,";
                    
                    sqlSTRING += "999999999999 as ID,'2' as sortID ";

                    sqlSTRING += "from 支給控除 inner join 配布員 ";
                    sqlSTRING += "on 支給控除.配布員ID = 配布員.ID left join 事業所 ";
                    sqlSTRING += "on 配布員.事業所コード = 事業所.ID ";

                    //週又は月
                    switch (tempSKbn)
                    {
                        case "月":
                            sqlSTRING += "where (year(支給控除.日付) = ?) and (month(支給控除.日付) = ?) ";
                            break;

                        case "週":
                            sqlSTRING += "where ((支給控除.日付) >= ?) and ((支給控除.日付) <= ?) ";
                            break;
                    }

                    sqlSTRING += "and (配布員.支払区分 = ?) ";

                    //配布員指定
                    if (tempHIndex == 1)
                    {
                        sqlSTRING += "and (配布員.ID = ?) ";
                    }
                    else if (tempHIndex >= 2)
                    {
                        // 2016/02/01
                        sqlSTRING += "and (事業所.ID = ?) ";
                    }

                    sqlSTRING += ") as U ";
                    sqlSTRING += "order by 事業所ID,U.配布員ID,sortID,U.配布指示ID,U.ID";
                                        
                    //配布指示データのデータリーダーを取得する
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    switch (tempSKbn)
                    {
                        case "月":
                            SCom.Parameters.AddWithValue("@P1", tempYear);
                            SCom.Parameters.AddWithValue("@P2", tempMonth);
                            SCom.Parameters.AddWithValue("@P3", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 配布員または事業所指定
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH1", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P4", tempYear);
                            SCom.Parameters.AddWithValue("@P5", tempMonth);
                            SCom.Parameters.AddWithValue("@P6", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH2", tempHID);
                            //}

                            // 2016/02/01 配布員または事業所指定
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH2", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P7", tempYear);
                            SCom.Parameters.AddWithValue("@P8", tempMonth);
                            SCom.Parameters.AddWithValue("@P9", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH3", tempHID);
                            //}

                            // 2016/02/01 配布員または事業所指定
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH3", tempHID);
                            }

                            break;

                        case "週":
                            SCom.Parameters.AddWithValue("@P1", sDate);
                            SCom.Parameters.AddWithValue("@P2", eDate);
                            SCom.Parameters.AddWithValue("@P3", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 配布員または事業所指定
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH1", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P4", sDate);
                            SCom.Parameters.AddWithValue("@P5", eDate);
                            SCom.Parameters.AddWithValue("@P6", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 配布員または事業所指定
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH2", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P7", sDate);
                            SCom.Parameters.AddWithValue("@P8", eDate);
                            SCom.Parameters.AddWithValue("@P9", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 配布員または事業所指定
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH3", tempHID);
                            }

                            break;
                    }

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();
                    
                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = dR["事業所名称"].ToString();
                            tempDGV[1, iX].Value = int.Parse(dR["配布員ID"].ToString());
                            tempDGV[2, iX].Value = dR["氏名"].ToString() + "";
                            tempDGV[3, iX].Value = dR["配布指示ID"].ToString();
                            tempDGV[4, iX].Value = DateTime.Parse(dR["配布日"].ToString()).ToShortDateString();
                            tempDGV[5, iX].Value = dR["チラシ名"].ToString();
                            tempDGV[6, iX].Value = dR["名称"].ToString() + "";
                            tempDGV[7, iX].Value = double.Parse(dR["単価"].ToString()).ToString("#,##0.00");
                            tempDGV[8, iX].Value = int.Parse(dR["実配布枚数"].ToString());
                            tempDGV[9, iX].Value = System.Math.Floor(double.Parse(dR["金額"].ToString()) + 0.5);

                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();

                    cn.Close();

                    Con.Close();

                    //カーソル表示を戻す
                    Cursor.Current = Cursors.Default;
                }
                catch (Exception e)
                {
                    //カーソル表示を戻す
                    Cursor.Current = Cursors.Default;

                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);

                }

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("終了します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int tempHID = 0;

            if (DataCheck() == false) return;

            if (MessageBox.Show("配布手数料明細を表示します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            //明細表示
            //switch (comboBox2.SelectedIndex)
            //{
            //    case 0:
            //        tempHID = 0;
            //        break;

            //    case 1:
            //        tempHID = int.Parse(txtHID.Text);
            //        break;
            //}

            // 2016/02/01
            if (comboBox2.SelectedIndex == 0)
            {
                // 全員
                tempHID = 0;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                // 配布員指定
                tempHID = int.Parse(txtHID.Text);
            }
            else
            {
                // 事業所指定
                tempHID = fIDArray[comboBox2.SelectedIndex - 2];
            }
            
            GridviewSet.ShowData(dataGridView2, int.Parse(txtYear.Text), int.Parse(txtMonth.Text),dateTimePicker1.Value.ToShortDateString(),dateTimePicker2.Value.ToShortDateString(), comboBox1.Text,comboBox2.SelectedIndex,tempHID);
            
            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("該当者はありませんでした", MESSAGE_CAPTION);

                //印刷ボタン
                button1.Enabled = false;

                //振込データ作成関連オブジェクト
                label12.Enabled = false;
                label13.Enabled = false;
                label14.Enabled = false;
                label15.Enabled = false;
                txtfMonth.Enabled = false;
                txtfDay.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
            }
            else
            {
                //印刷ボタン
                button1.Enabled = true;

                //振込データ作成関連オブジェクト
                label12.Enabled = true;
                label13.Enabled = true;
                label14.Enabled = true;
                label15.Enabled = true;
                txtfMonth.Enabled = true;
                txtfDay.Enabled = true;
                button2.Enabled = true;

                //配布員別手数料データ出力ボタン
                button3.Enabled = true;
            }

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ShowPosting(dataGridView1, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("配布料請求書受領証を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //グリッド再表示
            int tempHID = 0;

            if (DataCheck() == false) return;

            // 2016/02/01
            if (comboBox2.SelectedIndex == 0)
            {
                // 全員
                tempHID = 0;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                // 配布員指定
                tempHID = int.Parse(txtHID.Text);
            }
            else
            {
                // 事業所指定
                tempHID = fIDArray[comboBox2.SelectedIndex - 2];
            }

            GridviewSet.ShowData(dataGridView2, int.Parse(txtYear.Text), int.Parse(txtMonth.Text), dateTimePicker1.Value.ToShortDateString(), dateTimePicker2.Value.ToShortDateString(), comboBox1.Text, comboBox2.SelectedIndex, tempHID);

            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("該当者はありませんでした", MESSAGE_CAPTION);
                return;
            }

            // 発行 2016/09/19
            newReport();
        }

        private void newReport()
        {
            // プリントダイアログ
            PrintDialog pd = new PrintDialog();
            if (pd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            //印刷処理
            int staffID = 0;
            int dCnt = 0;
            double Gur = 0;
            string staffName = "";
            string sKikan = "";

            string[] sCell01 = new string[1];
            string[] sCell02 = new string[1];
            string[] sCell03 = new string[1];
            string[] sCell04 = new string[1];
            string[] sCell05 = new string[1];
            string[] sCell06 = new string[1];
            string[] sCell07 = new string[1];

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布料請求書受領証, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[2];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= dataGridView2.RowCount - 1; iX++)
                    {
                        // 印刷
                        if ((staffID != 0) && (staffID != int.Parse(dataGridView2[1, iX].Value.ToString())))
                        {
                            // エクセルシートにデータをセット
                            xlsDataSet(ref oXlsBook, ref oxlsSheet, Gur, staffID, staffName, sKikan, sCell01, sCell02, sCell03, sCell04, sCell05, sCell06, sCell07);                            
                        }

                        if (staffID != int.Parse(dataGridView2[1, iX].Value.ToString()))
                        {
                            // 対象期間
                            switch (comboBox1.Text)
                            {
                                case "月":
                                    sKikan = "対象期間 : " + txtYear.Text + label1.Text + txtMonth.Text + label2.Text;
                                    break;

                                case "週":
                                    sKikan = "対象期間 : " + DateTime.Parse(dateTimePicker1.Value.ToString()).AddDays(-1).ToShortDateString() + label4.Text + DateTime.Parse(dateTimePicker2.Value.ToString()).AddDays(-1).ToShortDateString();
                                    break;
                            }

                            // 配布員名
                            staffName = dataGridView2[2, iX].Value.ToString();

                            // 手数料合計
                            Gur = 0;

                            // 配布員別の明細数
                            dCnt = 0;
                        }

                        //明細
                        Array.Resize(ref sCell01, dCnt + 1);
                        Array.Resize(ref sCell02, dCnt + 1);
                        Array.Resize(ref sCell03, dCnt + 1);
                        Array.Resize(ref sCell04, dCnt + 1);
                        Array.Resize(ref sCell05, dCnt + 1);
                        Array.Resize(ref sCell06, dCnt + 1);
                        Array.Resize(ref sCell07, dCnt + 1);

                        sCell01[dCnt] = dataGridView2[3, iX].Value.ToString();
                        sCell02[dCnt] = dataGridView2[4, iX].Value.ToString();
                        sCell03[dCnt] = dataGridView2[5, iX].Value.ToString();
                        sCell04[dCnt] = dataGridView2[6, iX].Value.ToString();
                        sCell05[dCnt] = dataGridView2[7, iX].Value.ToString();
                        sCell06[dCnt] = dataGridView2[8, iX].Value.ToString();
                        sCell07[dCnt] = dataGridView2[9, iX].Value.ToString();

                        // 手数料計算
                        Gur += double.Parse(dataGridView2[9, iX].Value.ToString());

                        // 配布員ID
                        staffID = int.Parse(dataGridView2[1, iX].Value.ToString());

                        dCnt++;
                    }

                    // エクセルシートにデータをセット
                    xlsDataSet(ref oXlsBook, ref oxlsSheet, Gur, staffID, staffName, sKikan, sCell01, sCell02, sCell03, sCell04, sCell05, sCell06, sCell07);

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // テンプレート用のシートを削除する
                    ((Excel.Worksheet)oXlsBook.Sheets["受領書"]).Delete();    // 配布料請求書受領書シート
                    ((Excel.Worksheet)oXlsBook.Sheets["請求書"]).Delete();    // 配布料請求書シート

                    // Excelのウィンドウを表示する
                    //oXls.Visible = true;

                    // BOOKで印刷
                    oXlsBook.PrintOut(1, Type.Missing, 1, false, pd.PrinterSettings.PrinterName, Type.Missing, Type.Missing, Type.Missing);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();
                }
                finally
                {
                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    MessageBox.Show("処理が終了しました", "印刷処理",MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        ///---------------------------------------------------------------------------------------
        /// <summary>
        ///     配布料請求書受領書シート、配布料請求書シートにデータをセットする </summary>
        /// <param name="oXlsBook">
        ///     ExcelBookオブジェクト </param>
        /// <param name="oxlsSheet">
        ///     ExcelSheetオブジェクト </param>
        /// <param name="Gur">
        ///     手数料合計 </param>
        /// <param name="staffID">
        ///     スタッフコード </param>
        /// <param name="staffName">
        ///     スタッフ名 </param>
        /// <param name="sKikan">
        ///     対象期間 </param>
        /// <param name="sCell01">
        ///     明細項目１ </param>
        /// <param name="sCell02">
        ///     明細項目２ </param>
        /// <param name="sCell03">
        ///     明細項目３ </param>
        /// <param name="sCell04">
        ///     明細項目４ </param>
        /// <param name="sCell05">
        ///     明細項目５ </param>
        /// <param name="sCell06">
        ///     明細項目６ </param>
        /// <param name="sCell07">
        ///     明細項目７ </param>
        ///---------------------------------------------------------------------------------------
        private void xlsDataSet(ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, double Gur, int staffID, string staffName, string sKikan,
                                  string[] sCell01, string[] sCell02, string[] sCell03, string[] sCell04,
                                  string[] sCell05, string[] sCell06, string[] sCell07)
        {
            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            Excel.Worksheet jyuryoSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet seikyuSheet = (Excel.Worksheet)oXlsBook.Sheets[2];

            const int S_GYO = 14;    //エクセルファイル明細開始行
            const int S_ROWSMAX = 7; //エクセルファイル列最大値

            // 配布料請求書受領書シートを追加する
            jyuryoSheet.Copy(Type.Missing, oxlsSheet);
            oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[oXlsBook.Sheets.Count];

            // 対象期間
            oxlsSheet.Cells[12, 3] = sKikan;

            // 配布員ID
            oxlsSheet.Cells[4, 1] = staffID.ToString();

            // 配布員名
            oxlsSheet.Cells[4, 2] = staffName + "　殿";

            // 金額
            oxlsSheet.Cells[10, 6] = Gur;

            for (int iz = 0; iz < sCell01.Length; iz++)
            {
                // 明細
                oxlsSheet.Cells[iz + S_GYO, 1] = sCell01[iz];
                oxlsSheet.Cells[iz + S_GYO, 2] = sCell02[iz];
                oxlsSheet.Cells[iz + S_GYO, 3] = sCell03[iz];
                oxlsSheet.Cells[iz + S_GYO, 4] = sCell04[iz];
                oxlsSheet.Cells[iz + S_GYO, 5] = sCell05[iz];
                oxlsSheet.Cells[iz + S_GYO, 6] = sCell06[iz];
                oxlsSheet.Cells[iz + S_GYO, 7] = sCell07[iz];

                // セル下部へ実線ヨコ罫線を引く
                rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }

            // 表全体に実線縦罫線を引く
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // 表全体の左端縦罫線
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // 表全体の右端縦罫線
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


            // 配布料請求書 ////////////////////////////////////////////////////////////////////////////////////
            
            // 配布料請求書シートを追加する
            seikyuSheet.Copy(Type.Missing, oxlsSheet);
            oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[oXlsBook.Sheets.Count];

            // 対象期間
            oxlsSheet.Cells[12, 3] = sKikan;

            // 配布員名
            oxlsSheet.Cells[4, 2] = staffID.ToString() + "  " + staffName;

            // 配布員住所
            OleDbDataReader dR;
            Control.配布員 sCon = new Control.配布員();
            dR = sCon.FillBy("where 配布員.ID = " + staffID.ToString());
            while (dR.Read())
            {
                oxlsSheet.Cells[9, 2] = dR["住所"];
            }

            dR.Close();
            sCon.Close();

            // 金額
            oxlsSheet.Cells[10, 6] = Gur;

            for (int iz = 0; iz < sCell01.Length; iz++)
            {
                // 明細
                oxlsSheet.Cells[iz + S_GYO, 1] = sCell01[iz];
                oxlsSheet.Cells[iz + S_GYO, 2] = sCell02[iz];
                oxlsSheet.Cells[iz + S_GYO, 3] = sCell03[iz];
                oxlsSheet.Cells[iz + S_GYO, 4] = sCell04[iz];
                oxlsSheet.Cells[iz + S_GYO, 5] = sCell05[iz];
                oxlsSheet.Cells[iz + S_GYO, 6] = sCell06[iz];
                oxlsSheet.Cells[iz + S_GYO, 7] = sCell07[iz];

                // セル下部へ実線ヨコ罫線を引く
                rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }

            // 表全体に実線縦罫線を引く
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // 表全体の左端縦罫線
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // 表全体の右端縦罫線
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        }

        private void KanryoReport(double Gur, string staffID, string staffName, string sKikan,
                                  string[] sCell01,string[] sCell02,string[] sCell03,string[] sCell04,
                                  string[] sCell05,string[] sCell06,string[] sCell07, string prnName)
        {
            const int S_GYO = 14;    //エクセルファイル明細開始行
            const int S_ROWSMAX = 7; //エクセルファイル列最大値

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布料請求書受領証, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {

                    //対象期間
                    oxlsSheet.Cells[12, 3] = sKikan;

                    //配布員ID
                    oxlsSheet.Cells[4, 1] = staffID;

                    //配布員名
                    oxlsSheet.Cells[4, 2] = staffName + "　殿";

                    //金額
                    oxlsSheet.Cells[10, 6] = Gur;

                    for (int iX = 0; iX < sCell01.Length; iX++)
                    {

                        //明細
                        oxlsSheet.Cells[iX + S_GYO, 1] = sCell01[iX];
                        oxlsSheet.Cells[iX + S_GYO, 2] = sCell02[iX];
                        oxlsSheet.Cells[iX + S_GYO, 3] = sCell03[iX];
                        oxlsSheet.Cells[iX + S_GYO, 4] = sCell04[iX];
                        oxlsSheet.Cells[iX + S_GYO, 5] = sCell05[iX];
                        oxlsSheet.Cells[iX + S_GYO, 6] = sCell06[iX];
                        oxlsSheet.Cells[iX + S_GYO, 7] = sCell07[iX];

                        ////セル上部へ実線ヨコ罫線を引く
                        //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //セル下部へ実線ヨコ罫線を引く
                        rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }

                    //表全体に実線縦罫線を引く
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の左端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の右端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    //oXls.Visible = true;

                    //印刷
                    //oxlsSheet.PrintPreview(false);
                    //oxlsSheet.PrintOut(1, Type.Missing, 1, false, oXls.ActivePrinter, Type.Missing, Type.Missing, Type.Missing);
                    oxlsSheet.PrintOut(1, Type.Missing, 1, false, prnName, Type.Missing, Type.Missing, Type.Missing);


                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();

                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
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
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void KanryoReportS(double Gur, int staffID, string staffName, string sKikan,
                                  string[] sCell01, string[] sCell02, string[] sCell03, string[] sCell04,
                                  string[] sCell05, string[] sCell06, string[] sCell07, string prnName)
        {

            const int S_GYO = 14;    //エクセルファイル明細開始行
            const int S_ROWSMAX = 7; //エクセルファイル列最大値

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布料請求書, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {

                    //対象期間
                    oxlsSheet.Cells[12, 3] = sKikan;

                    //配布員名
                    oxlsSheet.Cells[4, 2] = staffID.ToString() + "  " + staffName;

                    //配布員住所
                    OleDbDataReader dR;
                    Control.配布員 sCon = new Control.配布員();
                    dR = sCon.FillBy("where 配布員.ID = " + staffID.ToString());
                    while (dR.Read())
                    {
                        oxlsSheet.Cells[9, 2] = dR["住所"];
                    }

                    dR.Close();
                    sCon.Close();

                    //金額
                    oxlsSheet.Cells[10, 6] = Gur;

                    for (int iX = 0; iX < sCell01.Length; iX++)
                    {

                        //明細
                        oxlsSheet.Cells[iX + S_GYO, 1] = sCell01[iX];
                        oxlsSheet.Cells[iX + S_GYO, 2] = sCell02[iX];
                        oxlsSheet.Cells[iX + S_GYO, 3] = sCell03[iX];
                        oxlsSheet.Cells[iX + S_GYO, 4] = sCell04[iX];
                        oxlsSheet.Cells[iX + S_GYO, 5] = sCell05[iX];
                        oxlsSheet.Cells[iX + S_GYO, 6] = sCell06[iX];
                        oxlsSheet.Cells[iX + S_GYO, 7] = sCell07[iX];

                        ////セル上部へ実線ヨコ罫線を引く
                        //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //セル下部へ実線ヨコ罫線を引く
                        rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }

                    //表全体に実線縦罫線を引く
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の左端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の右端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    //oXls.Visible = true;

                    //印刷
                    //oxlsSheet.PrintPreview(false);
                    oxlsSheet.PrintOut(1, Type.Missing, 1, false, prnName, Type.Missing, Type.Missing, Type.Missing);


                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();

                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
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
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void txtYear_Validating(object sender, CancelEventArgs e)
        {

            string str;
            int d;

            if (txtYear.Text == null)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtYear.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

        }

        private void txtMonth_Validating(object sender, CancelEventArgs e)
        {

            string str;
            int d;

            if (txtMonth.Text == null)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtMonth.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("1～12で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtHID) txtObj = txtHID;
            if (sender == txtfMonth) txtObj = txtfMonth;
            if (sender == txtfDay) txtObj = txtfDay;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtHID) txtObj = txtHID;
            if (sender == txtfMonth) txtObj = txtfMonth;
            if (sender == txtfDay) txtObj = txtfDay;

            txtObj.BackColor = Color.White;
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)

            {
		        case 0:
                    
                    //週期間
                    label3.Enabled = false;
                    dateTimePicker1.Enabled = false;
                    label4.Enabled = false;
                    dateTimePicker2.Enabled = false;

                    //対象月
                    label5.Enabled = true;
                    txtYear.Enabled = true;
                    label1.Enabled = true;
                    txtMonth.Enabled = true;
                    label2.Enabled = true;

                    break;

                case 1:

                    //週期間
                    label3.Enabled = true;
                    dateTimePicker1.Enabled = true;
                    label4.Enabled = true;
                    dateTimePicker2.Enabled = true;

                    //対象月
                    label5.Enabled = false;
                    txtYear.Enabled = false;
                    label1.Enabled = false;
                    txtMonth.Enabled = false;
                    label2.Enabled = false;

                     break;
            }

        }

        private void DispClear()
        {

            //週期間
            label3.Enabled = false;
            dateTimePicker1.Enabled = false;
            label4.Enabled = false;
            dateTimePicker2.Enabled = false;

            //対象月
            label5.Enabled = false;
            txtYear.Enabled = false;
            label1.Enabled = false;
            txtMonth.Enabled = false;
            label2.Enabled = false;

            //配布員指定
            label6.Enabled = false;
            txtHID.Enabled = false;
            label7.Enabled = false;

            //振込データ作成関連オブジェクト
            label12.Enabled = false;
            label13.Enabled = false;
            label14.Enabled = false;
            label15.Enabled = false;
            txtfMonth.Enabled = false;
            txtfDay.Enabled = false;

            button2.Enabled = false;
            button3.Enabled = false;
        }

        private bool DataCheck()
        {
            try
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("月または週を選択してください");
                }

                if (comboBox2.SelectedIndex == -1)
                {
                    comboBox2.Focus();
                    throw new Exception("配布員指定を選択してください");
                }

                if (comboBox2.SelectedIndex == 1 && txtHID.Text == "")
                {
                    txtHID.Focus();
                    throw new Exception("配布員IDを入力してください");
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            switch (comboBox2.SelectedIndex)
            {
                case 0:
                    label6.Enabled = false;
                    txtHID.Enabled = false;
                    label7.Enabled = false;
                    break;

                case 1:
                    label6.Enabled = true;
                    txtHID.Enabled = true;
                    label7.Enabled = true;
                    break;
            }
        }

        private void txtHID_Validating(object sender, CancelEventArgs e)
        {
            int d;
            string str;

            // 未入力またはスペースのみは可
            if ((this.txtHID.Text).Trim().Length < 1)
            {
                label7.Text = "";
                return;
            }

            // 数字か？
            if (txtHID.Text == null)
            {
                MessageBox.Show("数字で入力してください", "配布員コード", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtHID.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", "配布員コード", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            //コード検証
            string sqlStr;
            Control.配布員 cStaff = new Control.配布員();
            OleDbDataReader dR;

            sqlStr = " where ID = " + txtHID.Text.ToString();
            dR = cStaff.FillBy(sqlStr);

            if (dR.HasRows == false)
            {
                MessageBox.Show("未登録コードです", "配布員コード", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                label7.Text = "";
                dR.Close();
                cStaff.Close();
                return;
            }
            else
            {
                while (dR.Read())
                {

                    ////配布員の支払単位と照合
                    //if (comboBox1.Text != dR["支払区分"].ToString())
                    //{
                    //    MessageBox.Show("指定した配布員の手数料支払単位は「" + comboBox1.Text + "」ではありません", "配布員", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    e.Cancel = true;
                    //    label7.Text = "";
                    //    dR.Close();
                    //    cStaff.Close();
                    //    return;
                    //}

                    label7.Text = dR["氏名"].ToString();
                }
            }

            dR.Close();
            cStaff.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (DataCheckZengin() == false) return;
                if (MessageBox.Show("全銀フォーマットの振込データを出力します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                Cursor.Current = Cursors.WaitCursor;

                DialogResult ret;

                //ダイアログボックスの初期設定
                saveFileDialog1.Title = "振込データ保存";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = "振込データ_" + DateTime.Now.ToString("yyyyMMddhhmmss");

                saveFileDialog1.Filter = "ﾃｷｽﾄﾌｧｲﾙ(*.txt)|*.txt";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret != System.Windows.Forms.DialogResult.OK) return;

                fileName = saveFileDialog1.FileName;

                //ヘッダクラス
                MakeZenginHead(fileName);

                //データクラス
                MakeZenginData(fileName);

                //トレーラクラス
                MakeZenginTr(fileName);

                //エンドクラス
                MakeZenginEnd(fileName);

                //終了メッセージ
                MessageBox.Show("処理終了", "振込データ作成");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private bool DataCheckZengin()
        {
            try
            {
                if (Utility.NumericCheck(txtfMonth.Text) == false)
                {
                    txtfMonth.Focus();
                    throw new Exception("振込月は数字で入力してください");
                }

                if (int.Parse(txtfMonth.Text, System.Globalization.NumberStyles.Any) < 1 || int.Parse(txtfMonth.Text, System.Globalization.NumberStyles.Any) > 12)
                {
                    txtfMonth.Focus();
                    throw new Exception("振込月が正しくありません");
                }

                if (Utility.NumericCheck(txtfDay.Text) == false)
                {
                    txtfDay.Focus();
                    throw new Exception("振込日は数字で入力してください");
                }

                if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) < 1)
                {
                    txtfDay.Focus();
                    throw new Exception("振込日が正しくありません");
                }

                switch (int.Parse(txtfMonth.Text,System.Globalization.NumberStyles.Any))
                {
                    case 1:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("振込日が正しくありません");break;
                    case 2:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 29) throw new Exception("振込日が正しくありません"); break;
                    case 3:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("振込日が正しくありません"); break;
                    case 4:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("振込日が正しくありません"); break;
                    case 5:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("振込日が正しくありません"); break;
                    case 6:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("振込日が正しくありません"); break;
                    case 7:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("振込日が正しくありません"); break;
                    case 8:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("振込日が正しくありません"); break;
                    case 9:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("振込日が正しくありません"); break;
                    case 10: if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("振込日が正しくありません"); break;
                    case 11: if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("振込日が正しくありません"); break;
                    case 12: if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("振込日が正しくありません"); break;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

        }
        private void MakeZenginHead(string tempPath)
        {
            const string DATAKBN = "1";
            const string SHUBETSU = "21";
            const string BANKCODE = "0005";
            const string BANKNAME = "ﾐﾂﾋﾞｼﾄｳｷﾖｳUFJ";

            string stTarget;
            string sBuff;
            
            Entity.全銀ヘッダ cHead = new Entity.全銀ヘッダ();

            //会社情報取得
            OleDbDataReader dR;
            Control.会社情報 cSys = new Control.会社情報();
            dR = cSys.Fill();

            while (dR.Read())
            {
                sMaster.ID = int.Parse(dR["ID"].ToString(),System.Globalization.NumberStyles.Any);
                sMaster.依頼人コード = dR["依頼人コード"].ToString();
                sMaster.依頼人名 = dR["依頼人名"].ToString();
                sMaster.金融機関コード = dR["金融機関コード"].ToString();
                sMaster.金融機関名 = dR["金融機関名"].ToString();
                sMaster.支店コード = dR["支店コード"].ToString();
                sMaster.支店名 = dR["支店名"].ToString();
                sMaster.口座種別 = int.Parse(dR["口座種別"].ToString(),System.Globalization.NumberStyles.Any);
                sMaster.口座番号 = dR["口座番号"].ToString();
            }

            dR.Close();
            cSys.Close();

            //ヘッダクラスにデータをセット
            cHead.データ区分 = DATAKBN;
            cHead.種別コード = SHUBETSU;
            cHead.コード区分 = " ";
            stTarget = "";
            cHead.振込依頼人コード = stTarget.PadLeft(10);
            cHead.振込依頼人名 = sMaster.依頼人名.PadRight(40);
            cHead.取組日 = txtfMonth.Text.PadLeft(2, '0') + txtfDay.Text.PadLeft(2, '0');
            cHead.仕向銀行番号 = BANKCODE;
            cHead.仕向銀行名 = BANKNAME.PadRight(15);
            cHead.仕向支店番号 = sMaster.支店コード.Trim().PadLeft(3, '0');
            cHead.仕向支店名 = sMaster.支店名.PadRight(15);
            cHead.預金種目 = sMaster.口座種別.ToString();
            cHead.口座番号 = sMaster.口座番号.Trim().PadLeft(7, '0');
            stTarget = "";
            cHead.ダミー = stTarget.PadLeft(17);
 
            //出力ライン
            sBuff = "";
            sBuff += cHead.データ区分;
            sBuff += cHead.種別コード;
            sBuff += cHead.コード区分;
            sBuff += cHead.振込依頼人コード;
            sBuff += cHead.振込依頼人名;
            sBuff += cHead.取組日;
            sBuff += cHead.仕向銀行番号;
            sBuff += cHead.仕向銀行名;
            sBuff += cHead.仕向支店番号;
            sBuff += cHead.仕向支店名;
            sBuff += cHead.預金種目;
            sBuff += cHead.口座番号;
            sBuff += cHead.ダミー;

            //テキストデータ出力
            System.IO.StreamWriter tZengin = new System.IO.StreamWriter(tempPath, false,System.Text.Encoding.GetEncoding(932));
            tZengin.WriteLine(sBuff);
            tZengin.Close();

        }

        private void MakeZenginData(string tempPath)
        {
            int staffID = 0;
            double Gur = 0;
            double GurTotal = 0;
            int dCnt = 0;

            //ファイルオープン
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true,System.Text.Encoding.GetEncoding(932));

            for (int iX = 0; iX <= dataGridView2.RowCount - 1; iX++)
            {

                //データ作成
                if ((staffID != 0) && (staffID != int.Parse(dataGridView2[1, iX].Value.ToString())))
                {
                    MakeZenginData_meisai(staffID, Gur, tZengin);
                    dCnt ++;
                }

                if (staffID != int.Parse(dataGridView2[1, iX].Value.ToString()))
                {
                    //手数料合計
                    Gur = 0;
                }

                //手数料計算
                Gur += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                GurTotal += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);

                //配布員ID
                staffID = int.Parse(dataGridView2[1, iX].Value.ToString());
            }

            MakeZenginData_meisai(staffID, Gur, tZengin);
            dCnt++;

            //ファイルクローズ
            tZengin.Close();

            //トレーラークラスに合計件数と合計金額をセット
            sZtr.合計件数 = dCnt.ToString().PadLeft(6, '0');
            sZtr.合計金額 = GurTotal.ToString().PadLeft(12, '0');
            
        }

        private void MakeZenginData_meisai(int tempID, double tempKin,System.IO.StreamWriter tempZengin)
        {
            const string DATAKUBUN = "2";
            const string SHINKICODE = "0";
            const string FURIKOMISHITEIKUBUN = "7";

            string stTarget;
            string sBuff;
            string zKouza;
            string hKouza;

            Entity.全銀データレコード cData = new Entity.全銀データレコード();

            //配布員情報を取得
            OleDbDataReader dR;
            Control.配布員 cStaff = new Control.配布員();
            dR = cStaff.FillBy("where ID = " + tempID.ToString());

            while (dR.Read())
            {
                cData.被仕向銀行番号 = dR["金融機関コード"].ToString().Trim().PadLeft(4, '0');
                cData.被仕向支店番号 = dR["支店コード"].ToString().Trim().PadLeft(3, '0');
                cData.預金種目 = dR["口座種別"].ToString();

                //口座番号全角→半角に変換
                zKouza = dR["口座番号"].ToString().Trim() + "";
                hKouza = Strings.StrConv(zKouza, VbStrConv.Narrow, 0);
                cData.口座番号 = hKouza.PadLeft(7, '0');

                cData.受取人名 = dR["口座名義カナ"].ToString().Trim().PadRight(30);
            }

            dR.Close();
            cStaff.Close();

            //クラスにデータをセット
            stTarget = "";
            cData.データ区分 = DATAKUBUN;
            cData.被仕向銀行名 = stTarget.PadLeft(15);
            cData.被仕向支店名 = stTarget.PadLeft(15);
            cData.手形交換所番号 = stTarget.PadRight(4);
            cData.振込金額 = tempKin.ToString().PadLeft(10, '0');
            cData.新規コード = SHINKICODE;
            cData.顧客コード1 = tempID.ToString().PadLeft(10, '0');
            cData.顧客コード2 = tempID.ToString().PadLeft(10, '0');
            cData.振込指定区分 = FURIKOMISHITEIKUBUN;
            cData.識別表示  = stTarget.PadLeft(1);
            cData.ダミー = stTarget.PadLeft(7);

            //出力ライン
            sBuff = "";
            sBuff += cData.データ区分;
            sBuff += cData.被仕向銀行番号;
            sBuff += cData.被仕向銀行名;
            sBuff += cData.被仕向支店番号;
            sBuff += cData.被仕向支店名;
            sBuff += cData.手形交換所番号;
            sBuff += cData.預金種目;
            sBuff += cData.口座番号;
            sBuff += cData.受取人名;
            sBuff += cData.振込金額;
            sBuff += cData.新規コード;
            sBuff += cData.顧客コード1;
            sBuff += cData.顧客コード2;
            sBuff += cData.振込指定区分;
            sBuff += cData.識別表示;
            sBuff += cData.ダミー;

            //テキストデータ出力
            tempZengin.WriteLine(sBuff);

        }

        private void MakeZenginTr(string tempPath)
        {
            
            const string DATAKUBUN = "8";
            string stTarget = "";
            string sbuff;

            //ファイルオープン
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true, System.Text.Encoding.GetEncoding(932));

            //クラスにデータセット
            sZtr.データ区分 = DATAKUBUN;
            sZtr.ダミー = stTarget.PadLeft(101);

            //ライン出力
            sbuff = "";
            sbuff += sZtr.データ区分;
            sbuff += sZtr.合計件数;
            sbuff += sZtr.合計金額;
            sbuff += sZtr.ダミー;

            //テキストデータ出力
            tZengin.WriteLine(sbuff);

            //ファイルクローズ
            tZengin.Close();

        }

        private void MakeZenginEnd(string tempPath)
        {

            const string DATAKUBUN = "9";
            string stTarget = "";
            string sbuff;

            Entity.全銀エンドレコード cEnd = new Entity.全銀エンドレコード();

            //ファイルオープン
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true, System.Text.Encoding.GetEncoding(932));

            //クラスにデータセット
            cEnd.データ区分 = DATAKUBUN;
            cEnd.ダミー = stTarget.PadLeft(119);

            //ライン出力
            sbuff = "";
            sbuff += cEnd.データ区分;
            sbuff += cEnd.ダミー;

            //テキストデータ出力
            tZengin.WriteLine(sbuff);

            //ファイルクローズ
            tZengin.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {            
            try
            {
                if (MessageBox.Show("配布員別手数料一覧を出力します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                Cursor.Current = Cursors.WaitCursor;

                DialogResult ret;
                string fNameHD;

                // ファイル名見出し
                switch (comboBox1.Text)
                {
                    case "週":
                        fNameHD = dateTimePicker1.Value.ToLongDateString() + "から" + dateTimePicker2.Value.ToLongDateString() + "まで_"; 
                        break;

                    case "月":
                        fNameHD = txtYear.Text + "年" + txtMonth.Text + "月_";
                        break;

                    default:
                        fNameHD = "";
                        break;
                }

                // ダイアログボックスの初期設定
                saveFileDialog1.Title = "配布員別手数料一覧";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = fNameHD + "配布員別手数料一覧_" + DateTime.Now.ToString("yyyyMMddhhmmss");

                saveFileDialog1.Filter = "csvﾌｧｲﾙ(*.csv)|*.csv";

                // ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret != System.Windows.Forms.DialogResult.OK) return;

                fileName = saveFileDialog1.FileName;

                // データクラス
                MakeIchiranData(fileName);

                // 終了メッセージ
                MessageBox.Show("処理終了", "配布員別手数料一覧");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private void MakeIchiranData(string tempPath)
        {
            int staffID = 0;
            string staffName = "";
            string officeName = "";
            double Gur = 0;
            double GurTotal = 0;
            int dCnt = 0;

            //ファイルオープン
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true,System.Text.Encoding.GetEncoding(932));

            for (int iX = 0; iX <= dataGridView2.RowCount - 1; iX++)
            {
                //見出し出力
                if (iX == 0)
                {
                    MakeIchiranData_Header(tZengin);
                }

                //データ作成
                if ((staffID != 0) && (staffID != int.Parse(dataGridView2[1, iX].Value.ToString())))
                {
                    MakeIchiranData_meisai(officeName, staffID, staffName, Gur, tZengin);
                    dCnt++;
                }

                if (staffID != int.Parse(dataGridView2[1, iX].Value.ToString()))
                {
                    //手数料合計
                    Gur = 0;
                }

                //手数料計算
                Gur += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                GurTotal += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);

                //配布員ID,氏名,事業所
                officeName = dataGridView2[0, iX].Value.ToString();
                staffID = int.Parse(dataGridView2[1, iX].Value.ToString());
                staffName = dataGridView2[2, iX].Value.ToString();
            }

            MakeIchiranData_meisai(officeName, staffID, staffName, Gur, tZengin);
            dCnt++;

            //ファイルクローズ
            tZengin.Close();

        }
        private void MakeIchiranData_Header(System.IO.StreamWriter tempZengin)
        {
            string sBuff;

            //出力ライン : 2015/07/16 マイナンバー追加
            sBuff = "事業所,配布員ID,配布員名,手数料,マイナンバー";

            //テキストデータ出力
            tempZengin.WriteLine(sBuff);

        }
        private void MakeIchiranData_meisai(string tempOffice,int tempID, string tempName,double tempKin, System.IO.StreamWriter tempZengin)
        {
            string sBuff;

            //出力ライン
            sBuff = "";
            sBuff += tempOffice  + ",";
            sBuff += tempID.ToString() + ",";
            sBuff += tempName + ",";
            sBuff += tempKin.ToString() + ",";

            // マイナンバー : 2015/07/16
            foreach (var t in dts.配布員.Where(a => a.ID == tempID))
            {
                sBuff += t.マイナンバー;
            }

            //テキストデータ出力
            tempZengin.WriteLine(sBuff);
        }
    }
}
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
    public partial class frmRuikeiRep : Form
    {
        const string MESSAGE_CAPTION = "累計表";
        const int SHINJUKU = 2;
        const int JYOUTOU = 3;

        public frmRuikeiRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            GridviewSet.Setting(dataGridView1);

            //事業所コンボ
            //Utility.ComboOffice.load(comboBox1);

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
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", (float)9.5, FontStyle.Regular);
                    
                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 595;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "日付");
                    tempDGV.Columns.Add("col2", "曜日");
                    tempDGV.Columns.Add("col3", "天候");
                    tempDGV.Columns.Add("col4", "新宿実績");
                    tempDGV.Columns.Add("col5", "城東実績");
                    tempDGV.Columns.Add("col6", "計");
                    tempDGV.Columns.Add("col7", "新宿収支");
                    tempDGV.Columns.Add("col8", "城東収支");
                    tempDGV.Columns.Add("col9", "計");
                    tempDGV.Columns.Add("col10", "新宿配布員");
                    tempDGV.Columns.Add("col11", "城東配布員");
                    tempDGV.Columns.Add("col12", "計");

                    tempDGV.Columns[0].Width = 55;
                    tempDGV.Columns[1].Width = 55;
                    tempDGV.Columns[2].Width = 60;
                    tempDGV.Columns[3].Width = 90;
                    tempDGV.Columns[4].Width = 90;
                    tempDGV.Columns[5].Width = 90;
                    tempDGV.Columns[6].Width = 90;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;
                    tempDGV.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[11].DefaultCellStyle.Format = "#,##0";

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

            public static void ShowData(DataGridView tempDGV,int temptYear,int temptMonth)
            {
                string sqlSTRING = "";

                const int GYOSU = 32;
                DateTime sDate;

                try
                {
                    Control.DataControl sdcon = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = sdcon.GetConnection();

                    //データリーダーを取得する
                    OleDbDataReader dR;

                    sqlSTRING += "select 配布日,事業所ID,count(distinct 配布員ID) AS 配布員数,SUM(売上) AS 売上, ";
                    sqlSTRING += "SUM(原価) AS 原価, SUM(売上) - SUM(原価) AS 収支 ";
                    sqlSTRING += "from ";
                    sqlSTRING += "(SELECT TOP (100) PERCENT 配布指示.配布日, 配布指示.配布員ID,";
                    sqlSTRING += "事業所.ID AS 事業所ID,受注.単価,配布エリア.配布単価,";
                    sqlSTRING += "配布エリア.実配布枚数,配布エリア.予定枚数,";
                    sqlSTRING += "受注.単価 * 配布エリア.予定枚数 AS 売上,";
                    sqlSTRING += "配布エリア.配布単価 * 配布エリア.実配布枚数 AS 原価 ";
                    sqlSTRING += "from 配布指示 INNER JOIN 配布エリア ";
                    sqlSTRING += "ON 配布指示.ID = 配布エリア.配布指示ID INNER JOIN 受注 ";
                    sqlSTRING += "ON 配布エリア.受注ID = 受注.ID INNER JOIN 配布員 ";
                    sqlSTRING += "ON 配布指示.配布員ID = 配布員.ID LEFT OUTER JOIN 事業所 ";
                    sqlSTRING += "ON 配布員.事業所コード = 事業所.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(YEAR(配布指示.配布日) = ?) AND (MONTH(配布指示.配布日) = ?) AND ";
                    sqlSTRING += "(事業所.ID = " + SHINJUKU.ToString() + ") ";
                    sqlSTRING += "order by 配布指示.配布日, 配布指示.配布員ID) AS sel_tbl ";
                    sqlSTRING += "group by 配布日, 事業所ID ";

                    sqlSTRING += "union ";

                    sqlSTRING += "select 配布日,事業所ID,count(distinct 配布員ID) AS 配布員数,SUM(売上) AS 売上, ";
                    sqlSTRING += "SUM(原価) AS 原価, SUM(売上) - SUM(原価) AS 収支 ";
                    sqlSTRING += "from ";
                    sqlSTRING += "(SELECT TOP (100) PERCENT 配布指示.配布日, 配布指示.配布員ID,";
                    sqlSTRING += "事業所.ID AS 事業所ID,受注.単価,配布エリア.配布単価,";
                    sqlSTRING += "配布エリア.実配布枚数,配布エリア.予定枚数,";
                    sqlSTRING += "受注.単価 * 配布エリア.予定枚数 AS 売上,";
                    sqlSTRING += "配布エリア.配布単価 * 配布エリア.実配布枚数 AS 原価 ";
                    sqlSTRING += "from 配布指示 INNER JOIN 配布エリア ";
                    sqlSTRING += "ON 配布指示.ID = 配布エリア.配布指示ID INNER JOIN 受注 ";
                    sqlSTRING += "ON 配布エリア.受注ID = 受注.ID INNER JOIN 配布員 ";
                    sqlSTRING += "ON 配布指示.配布員ID = 配布員.ID LEFT OUTER JOIN 事業所 ";
                    sqlSTRING += "ON 配布員.事業所コード = 事業所.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(YEAR(配布指示.配布日) = ?) AND (MONTH(配布指示.配布日) = ?) AND ";
                    sqlSTRING += "(事業所.ID = " + JYOUTOU.ToString() + ") ";
                    sqlSTRING += "order by 配布指示.配布日, 配布指示.配布員ID) AS sel_tbl ";
                    sqlSTRING += "group by 配布日, 事業所ID ";
                    sqlSTRING += "order by 配布日, 事業所ID ";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@year1", temptYear);
                    SCom.Parameters.AddWithValue("@month1", temptMonth);

                    SCom.Parameters.AddWithValue("@year2", temptYear);
                    SCom.Parameters.AddWithValue("@month2", temptMonth);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //グリッドビューに表示する
                    int iX = 0;
                    double gzUri = 0;
                    double gtUri = 0;
                    double gzShushi = 0;
                    double gtShushi = 0;
                    double gzNin = 0;
                    double gtNin = 0;

                    //グリッド作成
                    tempDGV.RowCount = GYOSU;

                    //初期化(ゼロセット）
                    foreach (DataGridViewRow iRow in tempDGV.Rows)
                    {
                        foreach (DataGridViewColumn iColumn in tempDGV.Columns)
                        {
                            if ((iColumn.Index == 1) || (iColumn.Index == 2))
                            {
                                tempDGV[iColumn.Index, iRow.Index].Value = "";
                            }
                            else
                            {
                                tempDGV[iColumn.Index, iRow.Index].Value = (double)(0);
                            }
                        }
                    }

                    //日付、曜日、天候をセット
                    for (int i = 0; i < GYOSU; i++)
                    {
                        int rDay;
                        string rDate;

                        rDay = i + 1;

                        //対象日付
                        rDate = temptYear.ToString() + "/" + temptMonth.ToString() + "/" + rDay.ToString();
                        
                        if (DateTime.TryParse(rDate, out sDate) == true)
                        {
                            //日付
                            tempDGV[0, i].Value = rDay.ToString("00");

                            //曜日
                            tempDGV[1, i].Value = ("日月火水木金土").Substring(int.Parse(sDate.DayOfWeek.ToString("d")), 1);

                            //天候
                            OleDbDataReader dRt;
                            Control.天候 sTenkou = new Control.天候();
                            dRt = sTenkou.FillBy("where 日付 = '" + sDate.ToShortDateString() + "'");

                            while (dRt.Read())
                            {
                                tempDGV[2, i].Value = dRt["天候"].ToString() + "";
                            }

                            dRt.Close();
                            sTenkou.Close();
                        }
                    }

                    //データ表示
                    int sOffice = 0;

                    while (dR.Read())
                    {
                        iX = DateTime.Parse(dR["配布日"].ToString()).Day - 1;

                        //新宿or城東の判断
                        if (dR["事業所ID"] == DBNull.Value)
                        {
                            sOffice = 0;
                        }
                        else
                        {
                            sOffice = int.Parse(dR["事業所ID"].ToString(), System.Globalization.NumberStyles.Any);
                        }

                        switch (sOffice)
                        {
                            case SHINJUKU: //新宿
                                tempDGV[3, iX].Value = double.Parse(dR["売上"].ToString(), System.Globalization.NumberStyles.Any); //売上
                                tempDGV[6, iX].Value = double.Parse(dR["収支"].ToString(), System.Globalization.NumberStyles.Any); //収支
                                tempDGV[9, iX].Value = double.Parse(dR["配布員数"].ToString(), System.Globalization.NumberStyles.Any); //配布員数

                                //合計
                                gtUri += double.Parse(dR["売上"].ToString(), System.Globalization.NumberStyles.Any); //売上
                                gtShushi += double.Parse(dR["収支"].ToString(), System.Globalization.NumberStyles.Any); //収支
                                gtNin += double.Parse(dR["配布員数"].ToString(), System.Globalization.NumberStyles.Any); //配布員数

                                break;

                            case JYOUTOU:
                                tempDGV[4, iX].Value = double.Parse(dR["売上"].ToString(), System.Globalization.NumberStyles.Any); //売上
                                tempDGV[7, iX].Value = double.Parse(dR["収支"].ToString(), System.Globalization.NumberStyles.Any); //収支
                                tempDGV[10, iX].Value = double.Parse(dR["配布員数"].ToString(), System.Globalization.NumberStyles.Any); //収支

                                //合計
                                gzUri += double.Parse(dR["売上"].ToString(), System.Globalization.NumberStyles.Any); //売上
                                gzShushi += double.Parse(dR["収支"].ToString(), System.Globalization.NumberStyles.Any); //収支
                                gzNin += double.Parse(dR["配布員数"].ToString(), System.Globalization.NumberStyles.Any); //配布員数
 
                                break;
                        }
                    }

                    //合計行
                    if (tempDGV.RowCount == 0)
                    {
                        MessageBox.Show("該当するデータがありませんでした", MESSAGE_CAPTION);
                    }
                    else
                    {
                        tempDGV[0, GYOSU -1].Value = "計";
                        tempDGV[3, GYOSU - 1].Value = gtUri;
                        tempDGV[4, GYOSU - 1].Value = gzUri;
                        tempDGV[6, GYOSU - 1].Value = gtShushi;
                        tempDGV[7, GYOSU - 1].Value = gzShushi;
                        tempDGV[9, GYOSU - 1].Value = gtNin;
                        tempDGV[10, GYOSU - 1].Value = gzNin;
                    }

                    //if (tempDGV.RowCount <= 25)
                    //{
                    //    tempDGV.Columns[2].Width = 318;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[2].Width = 301;
                    //}

                    dR.Close();
                    sdcon.Close();
                    cn.Close();

                    tempDGV.CurrentCell = null;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }

            }
        }

        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {

            try
            {

                txtYear.Text = "";
                txtMonth.Text = "";

                btnPrn.Enabled = false;
                button1.Enabled = false;
                txtYear.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            int tYear, tMonth;

            try
            {

                if (Utility.NumericCheck(txtYear.Text) == false)
                {
                    MessageBox.Show("対象年は数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtYear.Focus();
                    return;
                }

                if (Utility.NumericCheck(txtMonth.Text) == false)
                {
                    MessageBox.Show("対象月は数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMonth.Focus();
                    return;
                }

                if (int.Parse(txtMonth.Text, System.Globalization.NumberStyles.Any) < 1 || int.Parse(txtMonth.Text, System.Globalization.NumberStyles.Any) > 12)
                {
                    MessageBox.Show("対象月が正しくありません", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMonth.Focus();
                    return;
                }

                //当月
                tYear = int.Parse(txtYear.Text, System.Globalization.NumberStyles.Any);
                tMonth = int.Parse(txtMonth.Text, System.Globalization.NumberStyles.Any);

                //データ表示
                GridviewSet.ShowData(dataGridView1, tYear,tMonth);

                //見出し書き換え
                //dataGridView1.Columns[1].HeaderText = zMonth.ToString() + "月度実績";
                //dataGridView1.Columns[2].HeaderText = tMonth.ToString() + "月度実績";
                //dataGridView1.Columns[5].HeaderText = zMonth.ToString() + "月度収支";
                //dataGridView1.Columns[6].HeaderText = tMonth.ToString() + "月度収支";
                //dataGridView1.Columns[9].HeaderText = zMonth.ToString() + "月配布員";
                //dataGridView1.Columns[10].HeaderText = tMonth.ToString() + "月配布員";

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                    button1.Enabled = true;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"選択",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            EigyoReport(dataGridView1);
        }


        private void EigyoReport(DataGridView tempDGV)
        {
            const int S_GYO = 4;        //エクセルファイル見出し行
            //const int S_ROWSMAX = 13;   //エクセルファイル列最大値
            string exlHead = "";        //見出し

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル累計表, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX < tempDGV.RowCount; iX++)
                    {
                        //見出し
                        int tYear,tMonth;

                        tYear = int.Parse(txtYear.Text, System.Globalization.NumberStyles.Any);
                        tMonth = int.Parse(txtMonth.Text,System.Globalization.NumberStyles.Any);

                        exlHead = tYear.ToString() + "年 " + tMonth.ToString() + "月度 新宿・城東累計実績";
                        oxlsSheet.Cells[S_GYO - 3, 3] = exlHead;

                        //oxlsSheet.Cells[S_GYO - 1, 2] = tempDGV.Columns[1].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 3] = tempDGV.Columns[2].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 6] = tempDGV.Columns[5].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 7] = tempDGV.Columns[6].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 10] = tempDGV.Columns[9].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 11] = tempDGV.Columns[10].HeaderText;

                        //明細
                        if (tempDGV[0, iX].Value.ToString() == "0")
                        {
                            oxlsSheet.Cells[iX + S_GYO, 1] = "";
                            oxlsSheet.Cells[iX + S_GYO, 2] = "";
                            oxlsSheet.Cells[iX + S_GYO, 3] = "";
                            oxlsSheet.Cells[iX + S_GYO, 4] = "";
                            oxlsSheet.Cells[iX + S_GYO, 5] = "";
                            oxlsSheet.Cells[iX + S_GYO, 6] = "";
                            oxlsSheet.Cells[iX + S_GYO, 7] = "";
                            oxlsSheet.Cells[iX + S_GYO, 8] = "";
                            oxlsSheet.Cells[iX + S_GYO, 9] = "";
                            oxlsSheet.Cells[iX + S_GYO, 10] = "";
                            oxlsSheet.Cells[iX + S_GYO, 11] = "";
                            oxlsSheet.Cells[iX + S_GYO, 12] = "";
                        }
                        else
                        {
                            oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                            oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                            oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                            oxlsSheet.Cells[iX + S_GYO, 4] = double.Parse(tempDGV[3, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 5] = double.Parse(tempDGV[4, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 6] = double.Parse(tempDGV[5, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 7] = double.Parse(tempDGV[6, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 8] = double.Parse(tempDGV[7, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 9] = double.Parse(tempDGV[8, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 10] = double.Parse(tempDGV[9, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 11] = double.Parse(tempDGV[10, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 12] = double.Parse(tempDGV[11, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                        }
                    }

                    ////////セル上部へ実線ヨコ罫線を引く
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////セル下部へ実線ヨコ罫線を引く
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////表全体に実線縦罫線を引く
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////表全体の左端縦罫線
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////表全体の右端縦罫線
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    
                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    //印刷
                    oxlsSheet.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //ダイアログボックスの初期設定
                    saveFileDialog1.Title = MESSAGE_CAPTION;
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = exlHead;
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            switch (e.ColumnIndex)
	        {
		        case 3: //新宿実績  
                    dgvUri_Sai(3, 4, e.RowIndex);
                    break;

                case 4: //城東実績
                    dgvUri_Sai(3, 4, e.RowIndex);
                    break;

                case 6: //新宿収支
                    dgvUri_Sai(6, 7, e.RowIndex);
                    break;

                case 7: //城東収支
                    dgvUri_Sai(6, 7, e.RowIndex);
                    break;

                case 9: //新宿人数
                    dgvUri_Sai(9, 10, e.RowIndex);
                    break;

                case 10: //城東人数
                    dgvUri_Sai(9, 10, e.RowIndex);
                    break;
	        }

        }

        /// <summary>
        /// 合計計算
        /// </summary>
        /// <param name="tempCol1">前月カラム</param>
        /// <param name="tempCol2">当月カラム</param>
        /// <param name="tempRow">行</param>
        private void dgvUri_Sai(int tempCol1, int tempCol2, int tempRow)
        {
            double z,t,r;

            //合計計算
            if (dataGridView1[tempCol1, tempRow].Value == null) dataGridView1[tempCol1, tempRow].Value = (double)(0);
            if (dataGridView1[tempCol2, tempRow].Value == null) dataGridView1[tempCol2, tempRow].Value = (double)(0);

            z = double.Parse(dataGridView1[tempCol1,tempRow].Value.ToString(),System.Globalization.NumberStyles.Any);
            t = double.Parse(dataGridView1[tempCol2,tempRow].Value.ToString(),System.Globalization.NumberStyles.Any);
            r = t + z;

            dataGridView1[tempCol2 + 1, tempRow].Value = r;
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.BackColor = Color.LightGray;
            txtObj.SelectAll();
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.BackColor = Color.White;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmSchedule : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "配布スケジュール";

        public frmSchedule()
        {            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView2,DateTime.Parse(dateTimePicker1.Value.ToShortDateString()));

            //事業所コンボ
            Utility.ComboOffice.load(comboBox1);
            comboBox1.SelectedIndex = -1;
            comboBox1.Enabled = false;

            //チェックボックス
            checkBox1.Checked = false;

            button1.Enabled = false;

        }

        // データグリッドビュークラス
        private class GridviewSet
        {

            /// <summary>
            /// データグリッドビューの定義を行います
            /// </summary>
            /// <param name="tempDGV">データグリッドビューオブジェクト</param>
            public static void Setting(DataGridView tempDGV,DateTime tempDate)
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
                    tempDGV.Columns.Add("col1", "ID");
                    tempDGV.Columns.Add("col2", "チラシ名");
                    tempDGV.Columns.Add("col3", "事業所");
                    tempDGV.Columns.Add("col4", "担当者");
                    tempDGV.Columns.Add("col5", "開始日");
                    tempDGV.Columns.Add("col6", "終了日");
                    tempDGV.Columns.Add("col7", "配布形態");
                    tempDGV.Columns.Add("col8", "残枚数");
                    tempDGV.Columns.Add("col9", "人数");
                    tempDGV.Columns.Add("col10", "延日数");

                    tempDGV.Columns[1].Frozen = true;

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 217;
                    tempDGV.Columns[2].Width = 70;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 90;
                    tempDGV.Columns[5].Width = 90;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 70;
                    tempDGV.Columns[8].Width = 60;
                    tempDGV.Columns[9].Width = 66;

                    //tempDGV.Columns.Add("col10", "d1");
                    //tempDGV.Columns.Add("col11", "d2");
                    //tempDGV.Columns.Add("col12", "d3");
                    //tempDGV.Columns.Add("col13", "d4");
                    //tempDGV.Columns.Add("col14", "d5");
                    //tempDGV.Columns.Add("col15", "d6");
                    //tempDGV.Columns.Add("col16", "d7");
                    //tempDGV.Columns.Add("col17", "d8");
                    //tempDGV.Columns.Add("col18", "d9");
                    //tempDGV.Columns.Add("col19", "d10");
                    //tempDGV.Columns.Add("col20", "d11");
                    //tempDGV.Columns.Add("col21", "d12");
                    //tempDGV.Columns.Add("col22", "d13");
                    //tempDGV.Columns.Add("col23", "d14");


                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void Setting2(DataGridView tempDGV)
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
                    tempDGV.Columns.Add("col1", "配布場所");
                    tempDGV.Columns.Add("col2", "件数");

                    tempDGV.Columns[0].Width = 200;
                    tempDGV.Columns[1].Width = 57;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[1].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV,DateTime tempDate,int tempOfficeID,ComboBox tempCmb)
            {
                try
                {

                    int iX;
                    int colno;
                    int  Nin;
                    string htxt;
                    string sqlSTRING = "";
                    DateTime sDate;
                    DateTime eDate;

                    const int DKIKAN = 14; 
                    DateTime [] mDate = new DateTime [DKIKAN];
                    int[] mTotal = new int[DKIKAN];

                    eDate = tempDate.AddDays(DKIKAN - 1);

                    //カーソル待機
                    Cursor.Current = Cursors.WaitCursor;

                    //期間列削除
                    if (tempDGV.Columns.Count > 10)
                    {
                        for (int i = 0; i < DKIKAN; i++)
	                    {
	                        tempDGV.Columns.RemoveAt(10);
	                    }
                    }

                    //期間列追加
                    for (int i = 0; i < DKIKAN; i++)
                    {
                        colno = i + 10;
                        sDate = tempDate.AddDays(i);

                        htxt = sDate.Day.ToString() + Environment.NewLine + ("日月火水木金土").Substring(int.Parse(sDate.DayOfWeek.ToString("d")), 1);

                        tempDGV.Columns.Add("col" + colno.ToString(), htxt);
                        tempDGV.Columns[colno].Width = 40;
                        tempDGV.Columns[colno].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                        mDate[i] = sDate;

                        //土日なら背景色表示
                        if (sDate.DayOfWeek.ToString("d") == "0" || sDate.DayOfWeek.ToString("d") == "6")
                        {
                            tempDGV.Columns[colno].DefaultCellStyle.BackColor = Color.LightPink;
                        }
                    }

                    tempDGV.RowCount = 0;
                    
                    //データリーダーを取得する
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select 受注.ID,受注.チラシ名,社員.氏名,受注.配布開始日,受注.配布終了日,";
                    sqlSTRING += "配布形態.名称,配布形態.一人当たり枚数,受注.枚数,sum_tbl.配布枚数,事業所.名称 as 事業所名 ";
                    sqlSTRING += "from 受注 LEFT OUTER JOIN ";
                    sqlSTRING += "得意先 ON 受注.得意先ID = 得意先.ID LEFT OUTER JOIN ";
                    sqlSTRING += "社員 ON 得意先.担当社員コード = 社員.ID LEFT OUTER JOIN ";
                    sqlSTRING += "配布形態 ON 受注.配布形態 = 配布形態.ID LEFT OUTER JOIN ";

                    sqlSTRING += "(select 配布エリア.受注ID, SUM(配布エリア.予定枚数) AS 配布枚数 ";
                    sqlSTRING += "from 配布エリア INNER JOIN ";
                    sqlSTRING += "配布指示 ON 配布エリア.配布指示ID = 配布指示.ID ";
                    sqlSTRING += "where (配布指示.配布日 < ?) ";
                    sqlSTRING += "group by 配布エリア.受注ID) AS sum_tbl ON 受注.ID = sum_tbl.受注ID ";
                    sqlSTRING += "left join ";
                    sqlSTRING += "事業所 on 受注.事業所ID = 事業所.ID ";

                    sqlSTRING += "where ";

                    //事業所指定
                    if (tempCmb.SelectedIndex != -1)
                    {
                        sqlSTRING += "(受注.事業所ID = ?) and ";
                    }

                    sqlSTRING += "(((受注.配布開始日 >= ?) and (受注.配布開始日 <= ?)) or ";
                    sqlSTRING += "((受注.配布終了日 >= ?) and (受注.配布終了日 <= ?)) or ";
                    sqlSTRING += "((受注.配布開始日 <= ?) and (受注.配布終了日 >= ?))) ";
                    sqlSTRING += "order by 受注.ID";
                                        
                    //配布指示データのデータリーダーを取得する

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@sDate", tempDate);

                    //事業所指定
                    if (tempCmb.SelectedIndex != -1)
                    {
                        SCom.Parameters.AddWithValue("@office", tempOfficeID);
                    }

                    SCom.Parameters.AddWithValue("@d1", tempDate);
                    SCom.Parameters.AddWithValue("@d2", eDate);
                    SCom.Parameters.AddWithValue("@d3", tempDate);
                    SCom.Parameters.AddWithValue("@d4", eDate);
                    SCom.Parameters.AddWithValue("@d5", tempDate);
                    SCom.Parameters.AddWithValue("@d6", eDate);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();
                    
                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = long.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = dR["チラシ名"].ToString();
                            tempDGV[2, iX].Value = dR["事業所名"].ToString() + "";
                            tempDGV[3, iX].Value = dR["氏名"].ToString();
                            tempDGV[4, iX].Value = DateTime.Parse(dR["配布開始日"].ToString()).ToShortDateString();
                            tempDGV[5, iX].Value = DateTime.Parse(dR["配布終了日"].ToString()).ToShortDateString();
                            tempDGV[6, iX].Value = dR["名称"].ToString() + "";

                            if (dR["配布枚数"] == DBNull.Value)
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["枚数"].ToString());
                            }
                            else
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["枚数"].ToString()) - int.Parse(dR["配布枚数"].ToString());
                            }

                            if (dR["一人当たり枚数"] == DBNull.Value)
                            {
                                Nin = (int)0;
                            }
                            else
                            {
                                Nin = int.Parse(dR["一人当たり枚数"].ToString(), System.Globalization.NumberStyles.Any);
                            }

                            if (Nin == 0)
                            {
                                tempDGV[8, iX].Value = (double)(0);
                            }
                            else
                            {
                                tempDGV[8, iX].Value = System.Math.Floor(double.Parse(tempDGV[7,iX].Value.ToString(),System.Globalization.NumberStyles.Any) / Nin + 0.9);
                            }

                            TimeSpan tSpan;
                            tSpan = DateTime.Parse(dR["配布終了日"].ToString()) - DateTime.Parse(dR["配布開始日"].ToString());

                            tempDGV[9, iX].Value = tSpan.Days + 1;

                            //スケジュール欄に人数表示
                            //配布開始日が起点日以前のときは起点日に表示
                            if (DateTime.Parse(dR["配布開始日"].ToString()) < mDate[0])
                            {
                                tempDGV[10, iX].Value = tempDGV[8, iX].Value.ToString();
                                mTotal[0] += int.Parse(tempDGV[8, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                            }
                            else　//配布開始日に表示
                            {
                                for (int i = 0; i < mDate.Length; i++)
                                {
                                    if (DateTime.Parse(dR["配布開始日"].ToString()) == mDate[i])
                                    {
                                        tempDGV[i + 10, iX].Value = tempDGV[8, iX].Value.ToString();
                                        mTotal[i] += int.Parse(tempDGV[8, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                                    }
                                }
                            }

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

                    //if (tempDGV.RowCount <= 27)
                    //{
                    //    tempDGV.Columns[1].Width = 217;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 200;
                    //}

                    //スケジュールヘッダに合計人数表示
                    for (int i = 0; i < mTotal.Length; i++)
                    {
                        tempDGV.Columns[i + 10].HeaderText += Environment.NewLine + mTotal[i].ToString();
                    }

                    //カーソル表示戻す
                    Cursor.Current = Cursors.Default;

                    if (iX == 0)
                    {
                        MessageBox.Show("該当するデータがありません",MESSAGE_CAPTION);
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);

                    //カーソル表示戻す
                    Cursor.Current = Cursors.Default;
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

            int officeID = 0;

            if (checkBox1.Checked == true)
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    MessageBox.Show("事業所を選択してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    comboBox1.Focus();
                    return;
                }
                else
                {
                    //事業所コード
                    Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                    cmb1 = (Utility.ComboOffice)comboBox1.SelectedItem;
                    officeID = cmb1.ID;
                }
            }

            //画面表示
            GridviewSet.ShowData(dataGridView2, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()),officeID,comboBox1);
            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount > 0)
            {
                button1.Enabled = true;
            }
        }

        private void ShowPosting(DataGridView tempDGV,DateTime tempDate,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select 町名.名称,配布エリア.枝番記入,配布エリア.報告枚数 ";
            mySql += "from (配布エリア inner join 配布指示 ";
            mySql += "on 配布エリア.配布指示ID = 配布指示.ID) left join 町名 ";
            mySql += "on 配布エリア.町名ID = 町名.ID ";
            mySql += "where (配布エリア.受注ID = " + tempID.ToString() + ") and ";
            mySql += "(配布指示.配布日 = '" + tempDate.ToShortDateString() + "') ";
            mySql += "order by 町名ID";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["名称"].ToString() + " " + dR["枝番記入"].ToString() + "";
                tempDGV[1, iX].Value = int.Parse(dR["報告枚数"].ToString());

                iX++;
            }

            if (tempDGV.RowCount <= 27)
            {
                tempDGV.Columns[0].Width = 200;
            }
            else
            {
                tempDGV.Columns[0].Width = 183;
            }

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("配布スケジュール表を印刷します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            PrintReport();

        }

        private void PrintReport()
        {

            const int S_GYO = 5;        //エクセルファイル明細は5行目から印字
            const int S_ROWSMAX = 24;   //エクセルファイル列最大値
            const int DKIKAN = 14;      //期間列数

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布スケジュール, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];
                Excel.Range srng;

                try
                {
                    //見出し
                    int sPos,sIdx,sCnt;
                    string sHead;
                    oxlsSheet.Cells[1, 1] = "配布スケジュール  [" + dateTimePicker1.Value.ToShortDateString() + " ～ " + dateTimePicker1.Value.AddDays(13).ToShortDateString() + "]";

                    //日・曜日・人数
                    for (int i = 10; i < S_ROWSMAX; i++)
                    {
                        sHead = dataGridView2.Columns[i].HeaderText.Replace(Environment.NewLine, "*");
                        sPos = -1;
                        sIdx = 0;
                        sCnt = -3;

                        while (true)
	                    {
                            sPos =sHead.IndexOf("*", sPos + 1);

                            if (sPos == -1)
                            {
                                oxlsSheet.Cells[S_GYO + sCnt, i + 1] = sHead.Substring(sIdx, sHead.Length - sIdx);
                                break;
                            }

                            oxlsSheet.Cells[S_GYO + sCnt, i + 1] = sHead.Substring(sIdx, sPos -sIdx);
                            sIdx = sPos + 1;
                            sCnt++;
	                    }
                        
                    }

                    //配布エリア明細
                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        oxlsSheet.Cells[i + S_GYO, 1] = dataGridView2[0, i].Value.ToString();   //受注番号
                        oxlsSheet.Cells[i + S_GYO, 2] = dataGridView2[1, i].Value.ToString();   //チラシ名
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView2[2, i].Value.ToString();   //事業所名                        
                        oxlsSheet.Cells[i + S_GYO, 4] = dataGridView2[3, i].Value.ToString();   //担当者名
                        oxlsSheet.Cells[i + S_GYO, 5] = dataGridView2[4, i].Value.ToString();   //配布開始日
                        oxlsSheet.Cells[i + S_GYO, 6] = dataGridView2[5, i].Value.ToString();   //配布終了日
                        oxlsSheet.Cells[i + S_GYO, 7] = dataGridView2[6, i].Value.ToString();   //配布形態
                        oxlsSheet.Cells[i + S_GYO, 8] = int.Parse(dataGridView2[7, i].Value.ToString(), System.Globalization.NumberStyles.Any);   //残枚数
                        oxlsSheet.Cells[i + S_GYO, 9] = int.Parse(dataGridView2[8, i].Value.ToString(), System.Globalization.NumberStyles.Any);   //人数
                        oxlsSheet.Cells[i + S_GYO, 10] = int.Parse(dataGridView2[9, i].Value.ToString(), System.Globalization.NumberStyles.Any);   //延日数

                        for (int n = 1; n <= DKIKAN; n++)
                        {
                            if (dataGridView2[n + 9, i].Value != null)
                            {
                                oxlsSheet.Cells[i + S_GYO, n + 10] = int.Parse(dataGridView2[n + 9, i].Value.ToString() + "", System.Globalization.NumberStyles.Any);   //人数
                            }
                        }

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

                    //土日の列は背景色を変える
                    for (int i = S_ROWSMAX - DKIKAN + 1; i <= S_ROWSMAX; i++)
                    {
                        srng = (Excel.Range)oxlsSheet.Cells[3, i];
                        
                        if (srng.Text.ToString() == "土" || srng.Text.ToString() == "日")
                        {
                            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, i];
                            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, i];
                            oxlsSheet.get_Range(rng[0], rng[1]).Interior.ColorIndex = 15;
                        }
                    }

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
                    saveFileDialog1.FileName = MESSAGE_CAPTION;
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                comboBox1.Enabled = true;
            }
            else
            {
                comboBox1.Enabled = false;
                comboBox1.SelectedIndex = -1;
            }
        }


    }
}
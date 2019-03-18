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
    public partial class frmHaifuKanryoRep : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "配布完了報告書";
        const string FILE_TYPE_XLS = ".xls"; 

        public frmHaifuKanryoRep()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            GridviewSet.Setting(dataGridView2);
            GridviewSet.Setting2(dataGridView1);
            button1.Enabled = false;
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
                    tempDGV.Columns.Add("col1", "受注番号");
                    tempDGV.Columns.Add("col2", "チラシ名");
                    tempDGV.Columns.Add("col3", "納品数");
                    tempDGV.Columns.Add("col4", "前回まで");
                    tempDGV.Columns.Add("col5", "残部数");
                    tempDGV.Columns.Add("col6", "配布件数");
                    tempDGV.Columns.Add("col7", "当日残");

                    tempDGV.Columns[0].Width = 80;
                    //tempDGV.Columns[1].Width = 210;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;

                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
                    tempDGV.Columns.Add("col0", "配布日");
                    tempDGV.Columns.Add("col1", "配布場所");
                    tempDGV.Columns.Add("col2", "件数");

                    tempDGV.Columns[0].Width = 80;
                    //tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 57;

                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    tempDGV.Columns[0].DefaultCellStyle.Format = "yyyy/MM/dd";
                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV, DateTime tempDate, DateTime tempDateE, string chirashiName)
            {

                string sqlSTRING = "";
                int iX;
                
                try
                {
                    tempDGV.RowCount = 0;
                    
                    //データリーダーを取得する
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    //sqlSTRING = "";
                    //sqlSTRING += "select 受注.ID, 受注.チラシ名, sum(配布エリア.予定枚数) as 配布件数,";
                    //sqlSTRING += "受注.枚数 as 納品数 ";
                    //sqlSTRING += "from 配布エリア inner join 受注 ";
                    //sqlSTRING += "on 配布エリア.受注ID = 受注.ID inner join 配布指示 ";
                    //sqlSTRING += "on 配布エリア.配布指示ID = 配布指示.ID ";
                    //sqlSTRING += "where (配布指示.配布日 = ?) and (配布エリア.完了区分 = 1) ";
                    //sqlSTRING += "group by 受注.ID, 受注.チラシ名, 受注.枚数";
                       
                    sqlSTRING = "";
                    sqlSTRING += "select 受注.ID,受注.チラシ名,SUM(配布エリア.予定枚数) AS 配布件数, ";
                    sqlSTRING += "受注.枚数 AS 納品数, x.前回配布件数 ";
                    sqlSTRING += "from 配布エリア inner join 受注 ";
                    sqlSTRING += "on 配布エリア.受注ID = 受注.ID inner join 配布指示 ";
                    sqlSTRING += "on 配布エリア.配布指示ID = 配布指示.ID left join ";

                    sqlSTRING += "(";
                    sqlSTRING += "select 配布エリア.受注ID, sum(配布エリア.予定枚数) AS 前回配布件数 ";
                    sqlSTRING += "from 配布エリア inner join 配布指示 ";
                    sqlSTRING += "on 配布エリア.配布指示ID = 配布指示.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(配布エリア.完了区分 = 1) and ";
                    sqlSTRING += "(配布指示.配布日 < ?) ";
                    sqlSTRING += "group by 配布エリア.受注ID";
                    sqlSTRING += ") AS x ";

                    sqlSTRING += "on 受注.ID = x.受注ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(配布指示.配布日 >= ?) and ";
                    sqlSTRING += "(配布指示.配布日 <= ?) and "; 
                    sqlSTRING += "(配布エリア.完了区分 = 1) and ";
                    sqlSTRING += "(受注.チラシ名 like ?) ";
                    sqlSTRING += "group by 受注.ID, 受注.チラシ名, 受注.枚数, x.前回配布件数";
                 
                    //配布指示データのデータリーダーを取得する

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@sDate", tempDate);
                    SCom.Parameters.AddWithValue("@sDate2", tempDate);
                    SCom.Parameters.AddWithValue("@DateE", tempDateE);
                    SCom.Parameters.AddWithValue("@cName",  "" + "%" + "" + chirashiName + "" + "%" + "");

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
                            tempDGV[2, iX].Value = int.Parse(dR["納品数"].ToString());

                            //tempDGV[3, iX].Value = 0;
                            //tempDGV[4, iX].Value = int.Parse(dR["納品数"].ToString());
                            
                            tempDGV[5, iX].Value = int.Parse(dR["配布件数"].ToString());
                            
                            //tempDGV[6, iX].Value = int.Parse(dR["納品数"].ToString()) - 0 - int.Parse(dR["配布件数"].ToString());

                            if (dR["前回配布件数"] == DBNull.Value)
                            {
                                tempDGV[3, iX].Value = 0;
                                tempDGV[4, iX].Value = int.Parse(dR["納品数"].ToString());
                                tempDGV[6, iX].Value = int.Parse(dR["納品数"].ToString()) - 0 - int.Parse(dR["配布件数"].ToString());
                            }
                            else
                            {
                                tempDGV[3, iX].Value = int.Parse(dR["前回配布件数"].ToString()); 
                                tempDGV[4, iX].Value = int.Parse(dR["納品数"].ToString()) - int.Parse(dR["前回配布件数"].ToString());
                                tempDGV[6, iX].Value = int.Parse(dR["納品数"].ToString()) - int.Parse(dR["前回配布件数"].ToString()) - int.Parse(dR["配布件数"].ToString());
                            }

                            ////前回までの配布枚数
                            //string mySql = "";
                            //OleDbDataReader dR2;

                            //mySql += "select sum(配布エリア.予定枚数) as 前回配布件数 ";
                            //mySql += "from 配布エリア inner join 配布指示 ";
                            //mySql += "on 配布エリア.配布指示ID = 配布指示.ID ";
                            //mySql += "where ";
                            //mySql += "(配布エリア.受注ID = " + dR["ID"].ToString() + ") and ";
                            //mySql += "(配布エリア.完了区分 = 1) and ";
                            //mySql += "(配布指示.配布日 < '" + tempDate.ToShortDateString() + "') ";
                            //mySql += "group by 配布エリア.受注ID";

                            //Control.FreeSql fCon = new Control.FreeSql();
                            //dR2 = fCon.free_dsReader(mySql);

                            //while (dR2.Read())
                            //{
                            //    tempDGV[3, iX].Value = int.Parse(dR2["前回配布件数"].ToString());
                            //    tempDGV[4, iX].Value = int.Parse(dR["納品数"].ToString()) - int.Parse(dR2["前回配布件数"].ToString());
                            //    tempDGV[6, iX].Value = int.Parse(dR["納品数"].ToString()) - int.Parse(dR2["前回配布件数"].ToString()) - int.Parse(dR["配布件数"].ToString());
                            //}

                            //dR2.Close();
                            //fCon.Close();


                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    //該当なしのとき
                    if (tempDGV.RowCount == 0) MessageBox.Show("該当期間に配布実績はありませんでした","該当なし");

                    dR.Close();

                    cn.Close();

                    Con.Close();

                    //tempDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    //if (tempDGV.RowCount <= 27)
                    //{
                    //    tempDGV.Columns[1].Width = 217;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 200;
                    //}

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中のチラシデータ全てを選択状態にします。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.Rows[i].Selected = true;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("現在の選択状態を解除します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                dataGridView2.Rows[r.Index].Selected = false;
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
            this.Cursor = Cursors.WaitCursor;   // 2019/02/13

            GridviewSet.ShowData(dataGridView2, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), Convert.ToDateTime(dateTimePicker2.Value.ToShortDateString()), txtsName.Text);
            dataGridView2.CurrentCell = null;
            dataGridView1.RowCount = 0;

            this.Cursor = Cursors.Default;  // 2019/02/13

            button1.Enabled = false;
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ShowPosting(dataGridView1, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), Convert.ToDateTime(dateTimePicker2.Value.ToShortDateString()), long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
            
            button1.Enabled = true;
        }

        private void ShowPosting(DataGridView tempDGV, DateTime tempDate, DateTime tempDateE, long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select 配布指示.配布日,町名.名称,配布エリア.枝番記入,配布エリア.予定枚数 ";
            mySql += "from (配布エリア inner join 配布指示 ";
            mySql += "on 配布エリア.配布指示ID = 配布指示.ID) left join 町名 ";
            mySql += "on 配布エリア.町名ID = 町名.ID ";
            mySql += "where ";
            mySql += "(配布エリア.受注ID = " + tempID.ToString() + ") and ";
            mySql += "(配布エリア.完了区分 = 1) and ";
            mySql += "(配布指示.配布日 >= '" + tempDate.ToShortDateString() + "') and ";
            mySql += "(配布指示.配布日 <= '" + tempDateE.ToShortDateString() + "') ";
            mySql += "order by 配布指示.配布日,町名ID";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["配布日"];
                tempDGV[1, iX].Value = dR["名称"].ToString() + " " + dR["枝番記入"].ToString() + "";
                tempDGV[2, iX].Value = int.Parse(dR["予定枚数"].ToString());

                iX++;
            }

            //if (tempDGV.RowCount <= 27)
            //{
            //    tempDGV.Columns[1].Width = 200;
            //}
            //else
            //{
            //    tempDGV.Columns[1].Width = 183;
            //}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("配布完了報告書を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 32; //配布完了報告書の明細行数

            //const int G_COUNT = 10; //配布完了報告書の明細行数
            //int pCnt;

            //ページカウント
            //pCnt = dataGridView1.Rows.Count / G_COUNT + 1;

            //for (int i = 1; i <= pCnt; i++)
            //{
            //    KanryoReport(pCnt, i, G_COUNT);
            //}

            KanryoReport(G_COUNT);
        }

        private void KanryoReport(int tempFixRows)
        {

            const int S_GYO = 13;    //エクセルファイル明細は13行目から印字
            int i;

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布完了報告書, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range rng;

                try
                {
                    //得意先情報
                    long sID;
                    string sqlSTR;

                    sID = long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString());

                    sqlSTR = "";
                    sqlSTR += "select 得意先.名称,得意先.担当者名,得意先.FAX番号 ";
                    sqlSTR += "from 受注 inner join 得意先 ";
                    sqlSTR += "on 受注.得意先ID = 得意先.ID ";
                    sqlSTR += "where 受注.ID = " + sID.ToString();

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTR);

                    while (dR.Read())
                    {
                        oxlsSheet.Cells[1, 3] = dR["名称"].ToString() + " " + dR["担当者名"].ToString() + "様";
                        oxlsSheet.Cells[2, 3] = dR["FAX番号"].ToString();                        
                    }

                    dR.Close();
                    fCon.Close();

                    //納品数
                    oxlsSheet.Cells[8, 2] = int.Parse(dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                    
                    //前回まで
                    oxlsSheet.Cells[9, 2] = int.Parse(dataGridView2[3, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    //残部数
                    oxlsSheet.Cells[10, 2] = int.Parse(dataGridView2[4, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    //チラシ名
                    oxlsSheet.Cells[10, 3] = dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //配布エリア明細初期化：2019/03/18
                    for (int iX = S_GYO; iX < (tempFixRows + S_GYO); iX++)
                    {
                        oxlsSheet.Cells[iX, 2] = string.Empty;
                        oxlsSheet.Cells[iX, 3] = string.Empty;
                        oxlsSheet.Cells[iX, 4] = string.Empty;
                    }
                    
                    //配布エリア明細
                    i = 0;
                    while (true)
                    {

                        //印刷明細最大行超えのとき行挿入
                        if (i >= (tempFixRows - 1))
                        {
                            rng = (Excel.Range)oxlsSheet.Cells[i + S_GYO,1];
                            rng.EntireRow.Insert(Type.Missing, Microsoft.Office.Interop.Excel.XlDirection.xlDown);
                        }

                        oxlsSheet.Cells[i + S_GYO, 2] = dataGridView1[0, i].Value.ToString();
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView1[1, i].Value.ToString();
                        oxlsSheet.Cells[i + S_GYO, 4] = int.Parse(dataGridView1[2, i].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //グリッド最終行のとき終了
                        if (i == (dataGridView1.Rows.Count - 1)) break;

                        i++;
                    }

                    ////配布枚数合計
                    //oxlsSheet.Cells[45, 4] = int.Parse(dataGridView2[5, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    ////残部数
                    //oxlsSheet.Cells[46, 4] = int.Parse(dataGridView2[6, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

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
                    saveFileDialog1.Title = MESSAGE_CAPTION + "保存";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    //saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString() + "_" + DateTime.Today.ToLongDateString();
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString() + "_" + this.dateTimePicker1.Text + "-" + this.dateTimePicker2.Text + FILE_TYPE_XLS;


                    ////複数ページのとき、ページ数も付与
                    //if (tempPage > 1)
                    //{
                    //    saveFileDialog1.FileName += "_" + tempCurrentPage.ToString();
                    //}

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

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


    }
}
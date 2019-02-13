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
    public partial class frmHaifuShinchoku : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "配布進捗状況";

        public frmHaifuShinchoku()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShinchoku_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView2);
            GridviewSet.Setting2(dataGridView1);
            button1.Enabled = false;

            comboBox1.Items.Add("全て");
            comboBox1.Items.Add("未完了");
            comboBox1.Items.Add("完了");
            comboBox1.SelectedIndex = 0;
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
                    tempDGV.Height = 235;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "受注番号");
                    tempDGV.Columns.Add("col2", "得意先");
                    tempDGV.Columns.Add("col3", "チラシ名");
                    tempDGV.Columns.Add("col4", "担当者");
                    tempDGV.Columns.Add("col5", "税込売上");
                    tempDGV.Columns.Add("col6", "予定枚数");
                    tempDGV.Columns.Add("col7", "配布枚数");
                    tempDGV.Columns.Add("col8", "残枚数");
                    tempDGV.Columns.Add("col9", "完了日");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 78;
                    tempDGV.Columns[5].Width = 78;
                    tempDGV.Columns[6].Width = 78;
                    tempDGV.Columns[7].Width = 75;
                    tempDGV.Columns[8].Width = 80;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
                    tempDGV.Height = 253;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "受注番号");
                    tempDGV.Columns.Add("col2", "指示番号");
                    tempDGV.Columns.Add("col3", "配布日");
                    tempDGV.Columns.Add("col4", "ID");
                    tempDGV.Columns.Add("col5", "配布員");
                    tempDGV.Columns.Add("col6", "ID");
                    tempDGV.Columns.Add("col7", "町名");
                    tempDGV.Columns.Add("col8", "予定枚数");
                    tempDGV.Columns.Add("col9", "残枚数");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 90;
                    tempDGV.Columns[3].Width = 60;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 60;
                    tempDGV.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[7].Width = 78;
                    tempDGV.Columns[8].Width = 78;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV,int tempSel,string tempCName)
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

                    sqlSTRING = "";
                    sqlSTRING += "select 受注.ID,社員.氏名,得意先.略称,受注.チラシ名,受注.税込金額,";
                    sqlSTRING += "受注.枚数 as 納品数,h_tbl.配布件数,";
                    sqlSTRING += "受注.枚数 - h_tbl.配布件数 as 残枚数 ";
                    sqlSTRING += "from 受注 left join  ";

                    sqlSTRING += "(select 受注ID,sum(予定枚数) as 配布件数 ";
                    sqlSTRING += "from 配布エリア ";
                    sqlSTRING += "where 完了区分 = 1 ";
                    sqlSTRING += "group by 受注ID ) as h_tbl ";

                    sqlSTRING += "on 受注.ID = h_tbl.受注ID left join 得意先 ";
                    sqlSTRING += "on 受注.得意先ID = 得意先.ID left join 社員 ";
                    sqlSTRING += "on 得意先.担当社員コード = 社員.ID ";
                    sqlSTRING += "where (受注.受注種別ID = 1) ";

                    switch (tempSel)
                    {
                        case 1:
                            sqlSTRING += "and (((受注.枚数 - h_tbl.配布件数) > 0) or h_tbl.配布件数 is null) ";
                            break;

                        case 2:
                            sqlSTRING += "and (受注.枚数 - h_tbl.配布件数 = 0) ";
                            break;
                    }

                    if (tempCName.Trim().Length > 0)
                    {
                        sqlSTRING += "and (受注.チラシ名 like ?)";
                    }

                    sqlSTRING += "order by 受注.ID desc";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    if (tempCName.Trim().Length > 0)
                    {
                        SCom.Parameters.AddWithValue("@CName", "%" + tempCName + "%");
                    }

                    SCom.Connection = cn;
                
                    //配布進捗状況のデータリーダーを取得する
                    //Control.FreeSql fCon = new Control.FreeSql();
                    //dR = fCon.free_dsReader(sqlSTRING);

                    dR = SCom.ExecuteReader();
                    
                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = long.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = dR["略称"].ToString() + "";
                            tempDGV[2, iX].Value = dR["チラシ名"].ToString();
                            tempDGV[3, iX].Value = dR["氏名"].ToString() + "";
                            tempDGV[4, iX].Value = int.Parse(dR["税込金額"].ToString(),System.Globalization.NumberStyles.Any);
                            tempDGV[5, iX].Value = int.Parse(dR["納品数"].ToString(), System.Globalization.NumberStyles.Any);

                            if (dR["配布件数"] == DBNull.Value)
                            {
                                tempDGV[6, iX].Value = 0;
                            }
                            else
                            {
                                tempDGV[6, iX].Value = int.Parse(dR["配布件数"].ToString(), System.Globalization.NumberStyles.Any);
                            }

                            if (dR["残枚数"] == DBNull.Value)
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["納品数"].ToString(), System.Globalization.NumberStyles.Any);
                            }
                            else
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["残枚数"].ToString(), System.Globalization.NumberStyles.Any);
                            }

                            //配布完了日
                            if (int.Parse(tempDGV[7, iX].Value.ToString(), System.Globalization.NumberStyles.Any) == 0)
                            {
                                string sqlSTR;
                                OleDbDataReader r;
                                Control.FreeSql rCon = new Control.FreeSql();
                                sqlSTR = "";
                                sqlSTR += "select max(配布指示.配布日) as 完了日 ";
                                sqlSTR += "from 配布エリア inner join 配布指示 ";
                                sqlSTR += "on 配布エリア.配布指示ID = 配布指示.ID ";
                                sqlSTR += "where ";
                                sqlSTR += "(配布エリア.受注ID = " + long.Parse(dR["ID"].ToString()) + ") and ";
                                sqlSTR += "(配布エリア.完了区分 = 1)";

                                r = rCon.free_dsReader(sqlSTR);

                                while (r.Read())
                                {
                                    if (r["完了日"] != DBNull.Value)
                                    {
                                        tempDGV[8, iX].Value = DateTime.Parse(r["完了日"].ToString()).ToShortDateString();
                                    }
                                    else
                                    {
                                        tempDGV[8, iX].Value = "";
                                    }
                                }

                                r.Close();
                                rCon.Close();
                            }
                            else
                            {
                                tempDGV[8, iX].Value = "";
                            }


                            iX++;

                            //frmP.valueCount = iX;
                            //frmP.ShowProgress();
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();
                    Con.Close();
                    cn.Close();

                    //fCon.Close();

                    //frmP.Close();

                    //frmP.Dispose();

                    //if (tempDGV.RowCount <= 12)
                    //{
                    //    tempDGV.Columns[3].Width = 97;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[3].Width = 80;
                    //}

                }
                catch (Exception e)
                {
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
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("対象チラシを選択してください");
                comboBox1.Focus();
                return;
            }

            try
            {

                Cursor.Current = Cursors.WaitCursor;    //カーソルを待機表示

                GridviewSet.ShowData(dataGridView2, comboBox1.SelectedIndex,txtCName.Text);
                dataGridView2.CurrentCell = null;
                dataGridView1.RowCount = 0;

                button1.Enabled = false;
                label1.Text = "【ポスティング詳細】";
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ShowPosting(dataGridView1,long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
        }

        private void ShowPosting(DataGridView tempDGV,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select 受注.ID, 配布指示.ID as 指示番号,配布指示.配布日, 配布員.ID AS 配布員ID, 配布員.氏名 AS 配布員氏名,";
            mySql += "町名.ID AS 町名ID, 町名.名称 AS 町名,";
            mySql += "配布エリア.予定枚数,配布エリア.完了区分 ";
            mySql += "from 受注 inner join 配布エリア ";
            mySql += "on 受注.ID = 配布エリア.受注ID left join 配布指示 ";
            mySql += "on 配布エリア.配布指示ID = 配布指示.ID left join 町名 ";
            mySql += "on 配布エリア.町名ID = 町名.ID left join 配布員 ";
            mySql += "on 配布指示.配布員ID = 配布員.ID ";
            mySql += "where 受注.ID = " + tempID.ToString() + " ";
            mySql += "order by 配布指示.配布日 desc";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;
            button1.Enabled = false;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["ID"].ToString();

                if (dR["指示番号"] == DBNull.Value)
                {
                    tempDGV[1, iX].Value = "** 未設定 **";
                }
                else
                {
                    tempDGV[1, iX].Value = dR["指示番号"].ToString();
                }

                if (dR["配布日"] == DBNull.Value)
                {
                    tempDGV[2, iX].Value = "";
                }
                else
                {
                    tempDGV[2, iX].Value =  DateTime.Parse(dR["配布日"].ToString()).ToShortDateString();
                }

                if (dR["配布員ID"] == DBNull.Value)
                {
                    tempDGV[3, iX].Value = "";
                }
                else
                {
                    tempDGV[3, iX].Value = dR["配布員ID"].ToString();
                }

                if (dR["配布員氏名"] == DBNull.Value)
                {
                    tempDGV[4, iX].Value = "";
                }
                else
                {
                    tempDGV[4, iX].Value = dR["配布員氏名"].ToString();
                }

                if (dR["町名ID"] == DBNull.Value)
                {
                    tempDGV[5, iX].Value = "";
                }
                else
                {
                    tempDGV[5, iX].Value = dR["町名ID"].ToString();
                }

                if (dR["町名"] == DBNull.Value)
                {
                    tempDGV[6, iX].Value = "";
                }
                else
                {
                    tempDGV[6, iX].Value = dR["町名"].ToString();
                }

                if (dR["予定枚数"] == DBNull.Value)
                {
                    tempDGV[7, iX].Value = 0;
                }
                else
                {
                    tempDGV[7, iX].Value = int.Parse(dR["予定枚数"].ToString());
                }

                //残枚数
                if (dR["予定枚数"] == DBNull.Value)
                {
                    tempDGV[8, iX].Value = 0;
                }
                else
                {
                    if (dR["完了区分"].ToString() == "1")
                    {
                        tempDGV[8, iX].Value = 0;
                    }
                    else
                    {
                        tempDGV[8, iX].Value = int.Parse(dR["予定枚数"].ToString());
                    }
                }

                iX++;

                button1.Enabled = true;
            }

            //if (tempDGV.RowCount <= 13)
            //{
            //    tempDGV.Columns[6].Width = 330;
            //}
            //else
            //{
            //    tempDGV.Columns[6].Width = 313;
            //}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();

            label1.Text = "【" + dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString() +" 配布状況】";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("配布進捗状況を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 36; //配布完了報告書の明細行数

            int pCnt;

            //ページカウント
            pCnt = dataGridView1.Rows.Count / G_COUNT + 1;

            for (int i = 1; i <= pCnt; i++)
            {
                KanryoReport(pCnt, i, G_COUNT);
            }

        }

        private void KanryoReport(int tempPage, int tempCurrentPage, int tempFixRows)
        {

            const int S_GYO = 8;    //エクセルファイル明細は8行目から印字
            int dgvIndex;
            int i;
            int yMai, zMai;

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布進捗状況, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    //日付・ページ数
                    oxlsSheet.Cells[1, 44] = "DATE : " + DateTime.Today.ToShortDateString() + "  P." + tempCurrentPage.ToString() + "/" + tempPage.ToString();

                    //受注番号
                    oxlsSheet.Cells[3, 5] = long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);
                    
                    //得意先名
                    oxlsSheet.Cells[4, 5] = dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString();
                    
                    //チラシ名
                    oxlsSheet.Cells[5, 5] = dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //担当者名
                    oxlsSheet.Cells[3, 18] = dataGridView2[3, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //税込売上
                    oxlsSheet.Cells[4, 29] = int.Parse(dataGridView2[4, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                    //予定枚数
                    oxlsSheet.Cells[4, 34] = int.Parse(dataGridView2[5, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                    
                    //配布枚数
                    oxlsSheet.Cells[4, 39] = int.Parse(dataGridView2[6, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                    //残枚数
                    oxlsSheet.Cells[4, 44] = int.Parse(dataGridView2[7, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                    //完了日
                    oxlsSheet.Cells[4, 49] = dataGridView2[8, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //配布エリア明細
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //データグリッドビューの行インデックスを求める

                        //指示番号
                        oxlsSheet.Cells[i + S_GYO, 1] = dataGridView1[1, dgvIndex].Value.ToString();
                        
                        //配布日
                        oxlsSheet.Cells[i + S_GYO, 5] = dataGridView1[2, dgvIndex].Value.ToString();
                        
                        //配布員ID
                        oxlsSheet.Cells[i + S_GYO, 10] = dataGridView1[3, dgvIndex].Value.ToString();
                        
                        //配布員名
                        oxlsSheet.Cells[i + S_GYO, 13] = dataGridView1[4, dgvIndex].Value.ToString();
                        
                        //町名ID
                        oxlsSheet.Cells[i + S_GYO, 19] = dataGridView1[5, dgvIndex].Value.ToString();
                        
                        //町名
                        oxlsSheet.Cells[i + S_GYO, 22] = dataGridView1[6, dgvIndex].Value.ToString();
                        
                        //予定枚数
                        oxlsSheet.Cells[i + S_GYO, 39] = dataGridView1[7, dgvIndex].Value.ToString();
                        
                        //配布枚数
                        yMai = int.Parse(dataGridView1[7, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);
                        zMai = int.Parse(dataGridView1[8, dgvIndex].Value.ToString(), System.Globalization.NumberStyles.Any);
                        oxlsSheet.Cells[i + S_GYO, 44] = yMai - zMai;
                        
                        //残枚数
                        oxlsSheet.Cells[i + S_GYO, 49] = dataGridView1[8, dgvIndex].Value.ToString();

                        //グリッド最終行のとき終了
                        if (dgvIndex == (dataGridView1.Rows.Count - 1)) break;

                        //印刷明細最大行のとき終了
                        if (i == (tempFixRows - 1)) break;

                        i++;
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
                    saveFileDialog1.Title = MESSAGE_CAPTION + "保存";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //複数ページのとき、ページ数も付与
                    if (tempPage > 1)
                    {
                        saveFileDialog1.FileName += "_" + tempCurrentPage.ToString();
                    }

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

    }
}
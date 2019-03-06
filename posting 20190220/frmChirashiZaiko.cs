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
    public partial class frmChirashiZaiko : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "チラシ在庫管理表";

        public frmChirashiZaiko()
        {

            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView2);
            button1.Enabled = false;

            radioButton2.Checked = true;
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();
            txtChirashiName.Text = string.Empty;
            txtsOrderNum.Text = string.Empty;

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
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 508;

                    // 奇数行の色
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col9", "受注番号");
                    tempDGV.Columns.Add("col1", "商品名");
                    tempDGV.Columns.Add("col2", "搬入日");
                    tempDGV.Columns.Add("col3", "配布日");
                    tempDGV.Columns.Add("col4", "搬入数");
                    tempDGV.Columns.Add("col5", "配布指示");
                    tempDGV.Columns.Add("col6", "指示残数");
                    tempDGV.Columns.Add("col7", "配布完了");
                    tempDGV.Columns.Add("col8", "完了残数");

                    tempDGV.Columns[0].Width = 120;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 190;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 80;
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
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
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
                    
                    // 罫線
                    tempDGV.BorderStyle = BorderStyle.Fixed3D;
                    tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void ShowData(DataGridView tempDGV, int tempYear, int tempMonth, int temprb, string cName, Int64 orderNum)
            {
                string sqlSTRING = "";
                int iX;

                try
                {
                    //カーソル表示を待機状態
                    Cursor.Current = Cursors.WaitCursor;

                    tempDGV.RowCount = 0;
                    
                    //データリーダーを取得する
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select 受注.ID,受注.チラシ名,受注.納品予定日,受注.配布開始日,";
                    sqlSTRING += "受注.配布終了日,受注.枚数,t_tbl.配布数,";
                    sqlSTRING += "受注.枚数 - t_tbl.配布数 AS 残数,";
                    sqlSTRING += "k_tbl.完了配布数,";
                    sqlSTRING += "受注.枚数 - k_tbl.完了配布数 AS 完了残数,";
                    sqlSTRING += "受注.外注依頼日支払, 受注.外注委託枚数 ";
                    
                    sqlSTRING += "from 受注 inner join ";

                    sqlSTRING += "(SELECT 受注_1.ID,SUM(配布エリア.予定枚数) AS 配布数 ";
                    sqlSTRING += "from 受注 AS 受注_1 INNER JOIN 配布エリア ";
                    sqlSTRING += "ON 受注_1.ID = 配布エリア.受注ID ";
                    sqlSTRING += "where (配布エリア.配布指示ID <> 0) ";
                    sqlSTRING += "GROUP BY 受注_1.ID) AS t_tbl ";

                    sqlSTRING += "ON 受注.ID = t_tbl.ID ";
                    sqlSTRING += "left join ";

                    sqlSTRING += "(SELECT 受注_2.ID,SUM(配布エリア.予定枚数) AS 完了配布数 ";
                    sqlSTRING += "from 受注 AS 受注_2 INNER JOIN 配布エリア ";
                    sqlSTRING += "ON 受注_2.ID = 配布エリア.受注ID ";
                    sqlSTRING += "where (配布エリア.配布指示ID <> 0) and (配布エリア.完了区分 = 1)";
                    sqlSTRING += "GROUP BY 受注_2.ID) AS k_tbl ";

                    sqlSTRING += "ON 受注.ID = k_tbl.ID ";

                    sqlSTRING += "where (受注種別ID = 1) ";
                    sqlSTRING += "and (受注.外注依頼日支払 is Null) "; // 2015/11/19

                    if (temprb == 1)
                    {
                        sqlSTRING += "and (year(受注.納品予定日) = ?) and (month(受注.納品予定日) = ?) ";
                    }

                    //sqlSTRING += "order by 受注.納品予定日 DESC";

                    // 外注渡し分 2015/11/19
                    sqlSTRING += "union ";
                    sqlSTRING += "select 受注.ID,受注.チラシ名,受注.納品予定日,受注.配布開始日,";
                    sqlSTRING += "受注.配布終了日, 受注.枚数, 0 as 配布数,";
                    sqlSTRING += "受注.枚数 - 0 AS 残数,";
                    sqlSTRING += "0 as 完了配布数,";
                    sqlSTRING += "受注.枚数 - 受注.外注委託枚数 AS 完了残数,";
                    sqlSTRING += "受注.外注依頼日支払, 受注.外注委託枚数 ";
                    sqlSTRING += "from 受注 ";
                    sqlSTRING += "where (受注種別ID = 1) ";
                    sqlSTRING += "and (受注.外注依頼日支払 > '2000/01/01') ";

                    if (temprb == 1)
                    {
                        sqlSTRING += "and (year(受注.納品予定日) = ?) and (month(受注.納品予定日) = ?) ";
                    }

                    sqlSTRING += "order by 受注.納品予定日 DESC";
                    
                                        
                    //配布指示データのデータリーダーを取得する
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    if (temprb == 1)
                    {
                        SCom.Parameters.AddWithValue("@year", tempYear);
                        SCom.Parameters.AddWithValue("@month", tempMonth);
                        SCom.Parameters.AddWithValue("@year", tempYear);
                        SCom.Parameters.AddWithValue("@month", tempMonth);
                    }

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();
                    
                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            // チラシ名検索 2015/09/18
                            if (cName != string.Empty)
                            {
                                if (!dR["チラシ名"].ToString().Contains(cName))
                                {
                                    continue;
                                }
                            }

                            // 受注番号検索 2015/09/18
                            if (orderNum != 0)
                            {
                                if (!dR["ID"].ToString().Contains(orderNum.ToString()))
                                {
                                    continue;
                                }
                            }

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = dR["ID"].ToString();
                            tempDGV[1, iX].Value = dR["チラシ名"].ToString();

                            if (dR["納品予定日"] == DBNull.Value)
                            {
                                tempDGV[2, iX].Value = "";
                            }
                            else
                            {

                                tempDGV[2, iX].Value = DateTime.Parse(dR["納品予定日"].ToString()).ToShortDateString();
                            }

                            tempDGV[3, iX].Value = DateTime.Parse(dR["配布開始日"].ToString()).ToShortDateString() + "～" + DateTime.Parse(dR["配布終了日"].ToString()).ToShortDateString();
                            tempDGV[4, iX].Value = int.Parse(dR["枚数"].ToString());
                            tempDGV[5, iX].Value = int.Parse(dR["配布数"].ToString());
                            tempDGV[6, iX].Value = int.Parse(dR["残数"].ToString());

                            if (dR["完了配布数"] == DBNull.Value)
                            {
                                tempDGV[7, iX].Value = 0;
                                tempDGV[8, iX].Value = int.Parse(dR["枚数"].ToString());

                            }
                            else
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["完了配布数"].ToString());
                                tempDGV[8, iX].Value = int.Parse(dR["完了残数"].ToString());
                            }

                            // 外注先に渡したとき 2015/08/11
                            if (dR["外注依頼日支払"] != DBNull.Value)
                            {
                                // 委託枚数指定がないとき
                                if (Utility.strToInt(dR["外注委託枚数"].ToString()) == global.FLGOFF)
                                {
                                    tempDGV[7, iX].Value = int.Parse(dR["枚数"].ToString()); // 2015/11/19
                                    tempDGV[8, iX].Value = "0";
                                }
                                else
                                { 
                                    // 2015/11/19
                                    tempDGV[7, iX].Value = dR["外注委託枚数"].ToString();

                                    // 委託枚数指定のとき、枚数から委託枚数を差し引いた数を完了残数
                                    int m = Utility.strToInt(dR["枚数"].ToString()) - Utility.strToInt(dR["外注委託枚数"].ToString());
                                    tempDGV[8, iX].Value = m;
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

                    if (tempDGV.RowCount <= 27)
                    {
                        tempDGV.Columns[1].Width = 310;
                    }
                    else
                    {
                        tempDGV.Columns[1].Width = 293;
                    }

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
            int rb = 0;

            //if (MessageBox.Show("チラシ在庫管理表を表示します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
            //    return;

            if (radioButton1.Checked)
            {
                rb = 0;
            }
            else
            {
                rb = 1;
            }

            GridviewSet.ShowData(dataGridView2, int.Parse(txtYear.Text), int.Parse(txtMonth.Text), rb, txtChirashiName.Text, Utility.strToInt64(txtsOrderNum.Text));
            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("該当データがありませんでした", MESSAGE_CAPTION);
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ShowPosting(dataGridView1, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("チラシ在庫管理表を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            KanryoReport(dataGridView2);
        }

        private void KanryoReport(DataGridView tempDGV)
        {

            const int S_GYO = 4;        //エクセルファイル見出し行（明細は4行目から印字）
            const int S_ROWSMAX = 9;    //エクセルファイル列最大値

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル在庫管理表, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        //oxlsSheet.Cells[S_GYO - 2, 1] = txtYear.Text + "年" + txtMonth.Text + "月 稼働日数 " + "日以上" ;
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[3, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[4, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 6] = tempDGV[5, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[6, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[7, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 9] = tempDGV[8, iX].Value.ToString();

                        //セル下部へ実線ヨコ罫線を引く
                        rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    }

                    //////セル上部へ実線ヨコ罫線を引く
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////セル下部へ実線ヨコ罫線を引く
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

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

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.BackColor = Color.White;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = new RadioButton();

            if (sender == radioButton1)
            {
                txtYear.Enabled = false;
                txtMonth.Enabled = false;
                label1.Enabled = false;
                label2.Enabled = false;
            }
            else if (sender == radioButton2)
            {
                txtYear.Enabled = true;
                txtMonth.Enabled = true;
                label1.Enabled = true;
                label2.Enabled = true;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
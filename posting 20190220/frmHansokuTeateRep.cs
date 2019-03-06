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
    public partial class frmHansokuTeateRep : Form
    {
        const string MESSAGE_CAPTION = "販売促進手当";

        public frmHansokuTeateRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView1);

            //社員コンボ
            Utility.ComboShain.load(comboBox1);

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
                    tempDGV.Height = 469;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "入金日");
                    tempDGV.Columns.Add("col2", "得意先略称");
                    tempDGV.Columns.Add("col3", "入金額");
                    tempDGV.Columns.Add("col4", "手当");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    tempDGV.Columns[0].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";

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

            public static void ShowData(DataGridView tempDGV,DateTime tempDate1,DateTime tempDate2,int tempID,Label t1,Label t2)
            {
                string sqlSTRING = "";
                int nTotal = 0;
                int Total = 0;

                try
                {

                    Control.DataControl sdcon = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = sdcon.GetConnection();

                    //データリーダーを取得する
                    OleDbDataReader dR;

                    sqlSTRING += "select 入金.*,請求書.税率,得意先.略称 ";
                    sqlSTRING += "from 入金 inner join 請求書 ";
                    sqlSTRING += "on 入金.請求書ID = 請求書.ID left join 得意先 ";
                    sqlSTRING += "on 請求書.得意先ID = 得意先.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(入金.入金年月日 >= ?) and ";
                    sqlSTRING += "(入金.入金年月日 <= ?) and ";
                    sqlSTRING += "(得意先.担当社員コード = ?) ";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@Date1", tempDate1);
                    SCom.Parameters.AddWithValue("@Date2", tempDate2);
                    SCom.Parameters.AddWithValue("@SID", tempID);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //グリッドビューに表示する
                    int iX = 0;
                    int RT;
                    double sKin;

                    RT = Properties.Settings.Default.販促手当率;
                    tempDGV.RowCount = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = Convert.ToDateTime(dR["入金年月日"].ToString());
                        tempDGV[1, iX].Value = dR["略称"].ToString();
                        tempDGV[2, iX].Value = int.Parse(dR["金額"].ToString(), System.Globalization.NumberStyles.Any);
                        
                        sKin = double.Parse(dR["金額"].ToString(), System.Globalization.NumberStyles.Any) / (1 + double.Parse(dR["税率"].ToString(), System.Globalization.NumberStyles.Any) / 100);
                        sKin = Math.Floor(sKin * RT / 100 + 0.5);
                        tempDGV[3, iX].Value = int.Parse(sKin.ToString(), System.Globalization.NumberStyles.Any);

                        //入金額合計
                        nTotal += int.Parse(dR["金額"].ToString(), System.Globalization.NumberStyles.Any);

                        //手当合計
                        Total += int.Parse(sKin.ToString(),System.Globalization.NumberStyles.Any);

                        iX++;
                    }

                    if (tempDGV.RowCount == 0)
                    {
                        MessageBox.Show("該当する入金データがありませんでした",MESSAGE_CAPTION);
                    }

                    //if (tempDGV.RowCount <= 25)
                    //{
                    //    tempDGV.Columns[1].Width = 300;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 293;
                    //}

                    dR.Close();
                    sdcon.Close();
                    cn.Close();

                    //合計表示
                    t1.Text = nTotal.ToString("#,##0");
                    t2.Text = Total.ToString("#,##0");

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
                label5.Text = "";
                label6.Text = "";

                tDate.Checked = false;
                tDate2.Checked = false;
                comboBox1.SelectedIndex = -1;

                btnPrn.Enabled = false;
                tDate.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            DateTime StartDate;
            DateTime EndDate;
            int ShainID;

            try
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    MessageBox.Show("担当者を選択してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (tDate.Checked == false)
                {
                    StartDate = Convert.ToDateTime("1900/01/01");
                }
                else
                {
                    StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
                }
                
                if (tDate2.Checked == false)
                {
                    EndDate = Convert.ToDateTime("9999/12/31");
                }
                else
                {
                    EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
                }

                if (comboBox1.SelectedIndex == -1)
                {
                    ShainID = 0;
                }
                else
                {
                    Utility.ComboShain cmb1 = new Utility.ComboShain();
                    cmb1 = (Utility.ComboShain)comboBox1.SelectedItem;
                    ShainID = cmb1.ID;
                }

                GridviewSet.ShowData(dataGridView1, StartDate, EndDate, ShainID,label5,label6);

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
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
            const int S_GYO = 4;    //エクセルファイル見出し行（明細は3行目から印字）
            const int S_ROWSMAX = 4; //エクセルファイル列最大値
            string sMidashi;

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル販売促進手当, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        sMidashi = this.comboBox1.Text + "  ";

                        if (tDate.Checked == true)
                        {
                            sMidashi += this.tDate.Value.ToShortDateString() + " ～ ";
                        }

                        if (tDate2.Checked == true)
                        {
                            if (tDate.Checked == false) sMidashi += " ～ ";

                            sMidashi += this.tDate2.Value.ToShortDateString();
                        }

                        oxlsSheet.Cells[S_GYO - 2, 1] = sMidashi;
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[3, iX].Value.ToString();

                        //セル下部へ実線ヨコ罫線を引く
                        rng[0] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    }

                    ////セル上部へ実線ヨコ罫線を引く
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

                    //合計
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count + 1, 3] = this.label5.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 4] = this.label6.Text;

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

                    string msgHD = "";

                    if (tDate.Checked == true)
                    {
                        msgHD += tDate.Value.ToLongDateString() + "から";
                    }

                    if (tDate2.Checked == true)
                    {
                        msgHD += tDate2.Value.ToLongDateString() + "まで";
                    }

                    //ダイアログボックスの初期設定
                    saveFileDialog1.Title = MESSAGE_CAPTION;
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + comboBox1.Text + "_" + msgHD; 
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
                MessageBox.Show(e.Message, "販売促進手当", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
            label5.Text = "";
            label6.Text = "";
        }

    }
}
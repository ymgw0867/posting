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
using System.Linq;

namespace posting
{
    public partial class frmClientbyOrderRep : Form
    {
        const string MESSAGE_CAPTION = "クライアント別受注内容一覧";

        public frmClientbyOrderRep()
        {
            InitializeComponent();

            // データ読み込み
            rAdp.Fill(dts.新請求書);
            cAdp.Fill(dts.得意先);
            sAdp.Fill(dts.社員);
            jAdp.Fill(dts.受注1);
            kAdp.Fill(dts.受注種別);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.新請求書TableAdapter rAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.受注種別TableAdapter kAdp = new darwinDataSetTableAdapters.受注種別TableAdapter();

        bool mukouStatus = false;

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView1);

            //画面クリア
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
                    tempDGV.Height = 470;

                    // 奇数行の色
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "番号");
                    tempDGV.Columns.Add("col2", "クライアント名");
                    tempDGV.Columns.Add("col3", "受注内容");
                    tempDGV.Columns.Add("col4", "担当者");
                    tempDGV.Columns.Add("col5", "発行日");
                    tempDGV.Columns.Add("col6", "売上金額");
                    tempDGV.Columns.Add("col7", "粗利");
                    tempDGV.Columns.Add("col8", "入金予定日");
                    tempDGV.Columns.Add("col9", "配布開始日");
                    tempDGV.Columns.Add("col10", "配布終了日");
                    tempDGV.Columns.Add("col11", "区分");
                    tempDGV.Columns.Add("col12", "ID");


                    //tempDGV.Columns[1].Frozen = true;   // 2015/07/22

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 90;
                    tempDGV.Columns[4].Width = 90;
                    tempDGV.Columns[5].Width = 90;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 60;
                    tempDGV.Columns[11].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[3].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[5].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    //請求書IDを非表示とする
                    //tempDGV.Columns[10].Visible = false;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // 編集不可とする
                    tempDGV.ReadOnly = false;

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
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
        }
        
        private void ShowDataLinq(DataGridView g)
        {
            decimal sArari = 0;
            decimal sUriage = 0;

            try
            {
                // グリッドビューに表示する
                int iX = 0;

                g.RowCount = 0;

                //var s = dts.新請求書
                //        .Where(a => a.無効 == global.FLGOFF && a.Get受注1Rows().Count() > 0)
                //        .OrderBy(a => a.請求金額).ThenBy(a => a.請求書発行日);


                var s = dts.受注1.Where(a => a.新請求書Row != null).OrderBy(a => a.得意先ID).ThenBy(a => a.受注種別ID);

                // 請求書発行日

                s = s.Where(a => a.請求書発行日 >= hDate.Value.Date).OrderBy(a => a.得意先ID).ThenBy(a => a.受注種別ID);

                if (hDate2.Checked)
                {
                    s = s.Where(a => a.請求書発行日 <= hDate2.Value.Date).OrderBy(a => a.得意先ID).ThenBy(a => a.受注種別ID);
                }

                // 並び順
                if (rBtnOrder.Checked == true)
                {
                    s = s.OrderBy(a => a.受注種別ID).ThenBy(a => a.得意先ID);
                }

                if (s.Count() == 0)
                {
                    MessageBox.Show("該当する受注データはありません","検索結果",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                
                // グリッドへ表示
                foreach (var t in s)
                {
                    // クライアント名指定
                    if (txtsClientName.Text.Trim() != string.Empty)
                    {
                        // 得意先名
                        if (t.得意先Row == null)
                        {
                            continue;
                        }
                        else if (!t.得意先Row.略称.Contains(txtsClientName.Text.Trim()))
                        {
                            continue;
                        }
                    }

                    g.Rows.Add();

                    g[0, iX].Value = t.得意先ID.ToString();    // 得意先ID

                    // 得意先名・担当者
                    if (t.得意先Row == null)
                    {
                        g[1, iX].Value = string.Empty;
                        g[3, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[1, iX].Value = t.得意先Row.略称;

                        if (t.得意先Row.社員Row == null)
                        {
                            g[3, iX].Value = string.Empty;
                        }
                        else
                        {
                            g[3, iX].Value = t.得意先Row.社員Row.氏名;
                        }
                    }

                    // 受注内容
                    if (t.受注種別Row != null)
                    {
                        g[2, iX].Value = t.受注種別Row.名称;
                    }
                    else
                    {
                        g[2, iX].Value = "";
                    }

                    g[4, iX].Value =  t.請求書発行日.ToShortDateString();     // 請求書発行日
                    g[5, iX].Value = t.売上金額.ToString("#,##0");            // 売上金額
                    //g[6, iX].Value = (t.売上金額 - t.外注原価支払 - t.外注原価支払2 - t.外注原価支払3).ToString("#,##0"); // 粗利
                    g[6, iX].Value = (t.金額 - t.値引額 - t.外注原価支払 - t.外注原価支払2 - t.外注原価支払3).ToString("#,##0");   // 粗利 2017/04/17

                    // 入金予定日
                    if (t.新請求書Row.Is支払期日Null())
                    {
                        g[7, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[7, iX].Value = t.新請求書Row.支払期日.ToShortDateString();
                    }

                    // 配布開始日：2017/05/24
                    if (t.Is配布開始日Null())
                    {
                        g[8, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[8, iX].Value = t.配布開始日.ToShortDateString();
                    }

                    // 配布終了日：2017/05/24
                    if (t.Is配布終了日Null())
                    {
                        g[9, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[9, iX].Value = t.配布終了日.ToShortDateString();
                    }

                    // 区分：2017/05/24
                    g[10, iX].Value = "";
                    if (t.新請求書Row != null)
                    {
                        if (t.配布開始日 > t.入金予定日)
                        {
                            if (t.新請求書Row.入金完了 == global.FLGON)
                            {
                                g[10, iX].Value = global.FLGON.ToString();
                            }
                            else
                            {
                                g[10, iX].Value = "2";
                            }
                        }
                    }

                    if (!t.Is配布終了日Null())
                    {
                        if (t.配布終了日.Year == DateTime.Today.Year && t.配布終了日.Month == DateTime.Today.Month)
                        {
                            int syymm = t.請求書発行日.Year * 100 + t.請求書発行日.Month;
                            int yymm = DateTime.Today.Year * 100 + DateTime.Today.Month;

                            if (syymm > yymm)
                            {
                                g[10, iX].Value = "3";
                            }
                        }
                    }

                    // 受注確定書ID：2017/05/24
                    g[11, iX].Value = t.ID.ToString();


                    iX++;

                    sUriage += t.売上金額;
                    //sArari += (t.売上金額 - t.外注原価支払 - t.外注原価支払2 - t.外注原価支払3);
                    sArari += (t.金額 - t.値引額 - t.外注原価支払 - t.外注原価支払2 - t.外注原価支払3);    // 2017/04/17
                }
                
                g.CurrentCell = null;

                txtUriTotal.Text = sUriage.ToString("#,##0");
                txtArariTotal.Text = sArari.ToString("#,##0");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }
        
                
        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {
            try
            {
                rBtnOrder.Checked = true;
                btnPrn.Enabled = false;
                hDate.Checked = false;
                hDate2.Checked = false;
                txtsClientName.Text = "";
                rBtnOrder.Checked = false;
                rBtnClient.Checked = true;
                linkLabel1.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                //画面表示
                mukouStatus = false;
                ShowDataLinq(dataGridView1);
                mukouStatus = true;

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                    linkLabel1.Visible = true;
                }
                else
                {
                    btnPrn.Enabled = false;
                    linkLabel1.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "選択", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            finally
            {
                Cursor = Cursors.Default;
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
            object [,] rtnArray = new object[dataGridView1.RowCount, dataGridView1.ColumnCount];
            getPrnArray(dataGridView1, ref rtnArray);
            EigyoReport(dataGridView1, rtnArray);
        }

        private void getPrnArray(DataGridView tempDGV, ref object [,] rtnArray)
        {
            // グリッド情報を配列に取り込む
            for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
            {
                rtnArray[iX, 0] = tempDGV[0, iX].Value.ToString();
                rtnArray[iX, 1] = tempDGV[1, iX].Value.ToString();
                rtnArray[iX, 2] = tempDGV[2, iX].Value.ToString();
                rtnArray[iX, 3] = tempDGV[3, iX].Value.ToString();
                rtnArray[iX, 4] = tempDGV[4, iX].Value.ToString();
                rtnArray[iX, 5] = tempDGV[5, iX].Value.ToString();
                rtnArray[iX, 6] = tempDGV[6, iX].Value.ToString();
                rtnArray[iX, 7] = tempDGV[7, iX].Value.ToString();
                rtnArray[iX, 8] = tempDGV[8, iX].Value.ToString();
                rtnArray[iX, 9] = tempDGV[9, iX].Value.ToString();
                rtnArray[iX, 10] = tempDGV[10, iX].Value.ToString();
                rtnArray[iX, 11] = tempDGV[11, iX].Value.ToString();
            }
        }

        private void EigyoReport(DataGridView tempDGV, object [,] rtnArray)
        {
            const int S_GYO = 6;        // エクセルファイル明細印刷開始行
            //const int S_ROWSMAX = 8;    // エクセルファイル列最大値

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;
                
                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセルクライアント別受注一覧, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] aRng = new Microsoft.Office.Interop.Excel.Range[2];
                Excel.Range rng;

                try
                {
                    // 配列からシートセルに一括してデータをセットします
                    rng = oxlsSheet.Range[oxlsSheet.Cells[6, 1], oxlsSheet.Cells[6 + rtnArray.GetLength(0) - 1, oxlsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;

                    // 合計情報を印字
                    oxlsSheet.Cells[S_GYO - 3, 10] = int.Parse(txtUriTotal.Text, System.Globalization.NumberStyles.Any);
                    oxlsSheet.Cells[S_GYO - 3, 12] = int.Parse(txtArariTotal.Text, System.Globalization.NumberStyles.Any);

                    // 見出し
                    oxlsSheet.Cells[3, 1] = getFromToDate();

                    ////////セル上部へ実線ヨコ罫線を引く
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //セル下部へ実線ヨコ罫線を引く
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体に実線縦罫線を引く
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の左端縦罫線
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の右端縦罫線
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, oxlsSheet.UsedRange.Columns.Count];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    
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
                    saveFileDialog1.Title = "クライアント別受注一覧";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    DateTime dt = DateTime.Now;
                    saveFileDialog1.FileName = "クライアント別受注一覧 " + getFromToDate();
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

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
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {
                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();

                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    oXls = null;
                    oXlsBook = null;
                    oxlsSheet = null;

                    GC.Collect();

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

        private string getFromToDate()
        {
            string msg = hDate.Value.Year.ToString() + "年" + hDate.Value.Month.ToString() + "月" + hDate.Value.Day.ToString() + "日～";
            
            if (hDate2.Checked)
            {
                msg += hDate2.Value.Year.ToString() + "年" + hDate2.Value.Month.ToString() + "月" + hDate2.Value.Day.ToString() + "日";
            }
            
            return msg;
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.BackColor = Color.White;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (mukouStatus)
            {
                // 入金済み、備考
                if (e.ColumnIndex == 9 || e.ColumnIndex == 11)
                {
                    // ID取得
                    int sID = Utility.strToInt(dataGridView1["col11", e.RowIndex].Value.ToString());

                    // データ取得
                    var s = dts.新請求書.Single(a => a.ID == sID);

                    // 入金済チェック
                    if (e.ColumnIndex == 9)
                    {
                        if (dataGridView1["col10", e.RowIndex].Value.ToString() == "True")
                        {
                            s.入金完了 = global.FLGON;
                        }
                        else
                        {
                            s.入金完了 = global.FLGOFF;
                        }
                    }
                    
                    // 備考
                    if (e.ColumnIndex == 11)
                    {
                        if (dataGridView1["col12", e.RowIndex].Value == null)
                        {
                            s.備考 = string.Empty;
                        }
                        else
                        {
                            s.備考 = dataGridView1["col12", e.RowIndex].Value.ToString();
                        }
                    }
                    
                    s.変更年月日 = DateTime.Now;
                    s.ユーザーID = global.loginUserID;

                    // 新請求書データ更新
                    rAdp.Update(dts.新請求書);

                    // データ読み込み
                    rAdp.Fill(dts.新請求書);
                }
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCellAddress.X == 9)
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DateTime sDt;
            DateTime eDt;

            // 請求開始年月日
            sDt = hDate.Value.Date;

            // 請求終了年月日
            if (hDate2.Checked)
            {
                eDt = hDate2.Value.Date;
            }
            else
            {
                eDt = new DateTime(2900, 12, 31, 0, 0, 0);
            }

            // 受注内容別集計画面表示
            frmClientbyOrderGrpRep frm = new posting.frmClientbyOrderGrpRep(sDt, eDt, txtsClientName.Text, getFromToDate());
            frm.ShowDialog();
        }
    }
}
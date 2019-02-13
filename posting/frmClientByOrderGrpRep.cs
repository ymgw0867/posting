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
    public partial class frmClientbyOrderGrpRep : Form
    {
        const string MESSAGE_CAPTION = "クライアント別受注内容別集計";

        public frmClientbyOrderGrpRep(DateTime sDt, DateTime eDt, string cName, string msg)
        {
            InitializeComponent();

            // データ読み込み
            rAdp.Fill(dts.新請求書);
            cAdp.Fill(dts.得意先);
            sAdp.Fill(dts.社員);
            jAdp.Fill(dts.受注1);
            kAdp.Fill(dts.受注種別);

            _sDt = sDt;
            _eDt = eDt;
            _msg = msg;
            _cName = cName;

            this.Text = "クライアント別受注内容別集計  " + msg;
        }

        DateTime _sDt;
        DateTime _eDt;
        string _msg = string.Empty;
        string _cName = string.Empty;

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

            // 集計結果表示
            ShowDataLinq(dataGridView1);
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
                    tempDGV.Columns.Add("col6", "売上金額");
                    tempDGV.Columns.Add("col7", "粗利");
                    

                    //tempDGV.Columns[1].Frozen = true;   // 2015/07/22

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

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
            try
            {
                // グリッドビューに表示する
                int iX = 0;

                g.RowCount = 0;

                var j = dts.受注1
                    .Where(a => a.新請求書Row != null &&
                    a.請求書発行日 >= DateTime.Parse(_sDt.ToShortDateString()) &&
                    a.請求書発行日 <= DateTime.Parse(_eDt.ToShortDateString()))
                    .GroupBy(a => a.得意先ID)
                    .Select(b => new
                    {
                        sID = b.Key,
                        ss = b.GroupBy(a => a.受注種別ID)
                        .Select(c => new
                        {
                            orderID = c.Key,
                            sumUri = c.Sum(a => a.金額),
                            sumNebiki = c.Sum(a => a.値引額),
                            sumGenka1 = c.Sum(a => a.外注原価支払),
                            sumGenka2 = c.Sum(a => a.外注原価支払2),
                            sumGenka3 = c.Sum(a => a.外注原価支払3)
                        })
                        .OrderBy(a => a.orderID)
                    })
                    .OrderBy(a => a.sID);

                foreach (var t in j)
                {
                    string uID = t.sID.ToString();    // 得意先ID
                    string uName = "";

                    if (dts.得意先.Any(a => a.ID == t.sID))
                    {
                        foreach (var item in dts.得意先.Where(a => a.ID == t.sID))
                        {
                            uName = item.略称;
                        }
                    }

                    // クライアント名指定
                    if (_cName != string.Empty)
                    {
                        // 得意先名
                        if (!uName.Contains(_cName))
                        {
                            continue;
                        }
                    }

                    foreach (var d in t.ss)
                    {
                        g.Rows.Add();

                        g[0, iX].Value = uID;    // 得意先ID
                        g[1, iX].Value = uName;

                        if (dts.受注種別.Any(a => a.ID == d.orderID))
                        {
                            foreach (var item in dts.受注種別.Where(a => a.ID == d.orderID))
                            {
                                g[2, iX].Value = item.名称;
                            }
                        }

                        g[3, iX].Value = d.sumUri.ToString("#,##0");
                        g[4, iX].Value = (d.sumUri - d.sumNebiki - d.sumGenka1 - d.sumGenka2 - d.sumGenka3).ToString("#,##0");      // 2017/04/17
                        //g[4, iX].Value = (d.sumUri - d.sumGenka1 - d.sumGenka2 - d.sumGenka3).ToString("#,##0");

                        iX++;
                    }          
                }
                
                g.CurrentCell = null;

                if (g.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                }
                else
                {
                    btnPrn.Enabled = false;
                }
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
                btnPrn.Enabled = false;
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

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセルクライアント別受注別集計, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] aRng = new Microsoft.Office.Interop.Excel.Range[2];
                Excel.Range rng;

                try
                {
                    // 見出し
                    oxlsSheet.Cells[3, 1] = _msg;

                    // 配列からシートセルに一括してデータをセットします
                    rng = oxlsSheet.Range[oxlsSheet.Cells[6, 1], oxlsSheet.Cells[6 + rtnArray.GetLength(0) - 1, oxlsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;
                    
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
                    saveFileDialog1.Title = "クライアント別受注別集計";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    DateTime dt = DateTime.Now;
                    saveFileDialog1.FileName = this.Text;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void frmClientbyOrderGrpRep_Shown(object sender, EventArgs e)
        {
        }
    }
}
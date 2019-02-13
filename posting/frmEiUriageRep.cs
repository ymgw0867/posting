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
    public partial class frmEiUriageRep : Form
    {
        const string MESSAGE_CAPTION = "営業別売上表";

        public frmEiUriageRep()
        {
            InitializeComponent();

            jAdp.Fill(dts.受注1);
            sAdp.Fill(dts.新請求書);
            nAdp.Fill(dts.新入金);
            cAdp.Fill(dts.得意先);
            eAdp.Fill(dts.社員);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.新請求書TableAdapter sAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.新入金TableAdapter nAdp = new darwinDataSetTableAdapters.新入金TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter eAdp = new darwinDataSetTableAdapters.社員TableAdapter();

        private void form_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            Setting(dataGridView1);

            //社員コンボ
            Utility.ComboShain.load(comboBox1);

            DispClear();
        }

        #region グリッドビューカラム定義
        const string colNDt = "col0";           // 入金日
        const string colClient = "col1";        // 得意先
        const string colNyukin = "col2";        // 入金
        const string colEganka = "col3";        // 営業原価
        const string colGaichuhi = "col4";      // 外注費
        const string colArari1 = "col5";        // 粗利1
        const string colArari2 = "col6";        // 粗利2
        const string colArariSai = "col7";      // 粗利差異
        const string colGaichuhi2 = "col8";      // 外注費2
        const string colGaichuhi3 = "col9";      // 外注費3
        #endregion

        ///----------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///----------------------------------------------------------------
        private static void Setting(DataGridView tempDGV)
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
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);
                    
                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 542;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(colNDt, "入金日");
                tempDGV.Columns.Add(colClient, "得意先略称");
                tempDGV.Columns.Add(colNyukin, "入金額");
                tempDGV.Columns.Add(colEganka, "営業原価");
                tempDGV.Columns.Add(colArari1, "粗利１");
                tempDGV.Columns.Add(colGaichuhi, "外注費");
                tempDGV.Columns.Add(colGaichuhi2, "外注費２");
                tempDGV.Columns.Add(colGaichuhi3, "外注費３");
                tempDGV.Columns.Add(colArari2, "粗利２");
                tempDGV.Columns.Add(colArariSai, "粗利差異");

                tempDGV.Columns[colNDt].Width = 110;
                tempDGV.Columns[colClient].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colNyukin].Width = 100;
                tempDGV.Columns[colEganka].Width = 100;
                tempDGV.Columns[colArari1].Width = 100;
                tempDGV.Columns[colGaichuhi].Width = 100;
                tempDGV.Columns[colGaichuhi2].Width = 100;
                tempDGV.Columns[colGaichuhi3].Width = 100;
                tempDGV.Columns[colArari2].Width = 100;
                tempDGV.Columns[colArariSai].Width = 100;

                tempDGV.Columns[colNDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colNyukin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colEganka].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colArari1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colGaichuhi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colGaichuhi2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colGaichuhi3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colArari2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colArariSai].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                
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
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ShowData(DataGridView tempDGV,DateTime tempDate1,DateTime tempDate2,int tempID,Label t1,Label t2)
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

        ///---------------------------------------------------------------------------------------
        /// <summary>
        ///     グリッドにデータを表示します </summary>
        /// <param name="tempDGV">
        ///     dataGridViewオブジェクト</param>
        /// <param name="tempDate1">
        ///     開始年月日</param>
        /// <param name="tempDate2">
        ///     終了年月日</param>
        /// <param name="tempID">
        ///     社員番号</param>
        ///---------------------------------------------------------------------------------------
        private void ShowDataLinq(DataGridView tempDGV, DateTime tempDate1, DateTime tempDate2, int tempID)
        {
            int nTotal = 0;
            int Total = 0;
            int iX = 0;
            int arari1Tl = 0;
            int arari2Tl = 0;
            int gaichuhiTl = 0;
            int gaichuhiTl2 = 0;
            int gaichuhiTl3 = 0;
            int arariSaiTl = 0;

            try
            {
                tempDGV.RowCount = 0;

                foreach (var t in dts.新請求書.Where(a => a.入金完了 == global.FLGON && 
                                                          a.無効 == 0 && a.Get受注1Rows().Count() > 0 && 
                                                          a.Get新入金Rows().Count() > 0))
                {
                    bool hit = true;
                    int nyukin = 0;
                    DateTime nDate = DateTime.Today;

                    // 担当営業か検証
                    if (t.得意先Row == null || t.得意先Row.社員Row == null || t.得意先Row.社員Row.ID != tempID)
                    {
                        continue;
                    }

                    //if (t.得意先Row.社員Row == null)
                    //{
                    //    continue;
                    //}

                    //if (t.得意先Row.社員Row.ID != tempID)
                    //{
                    //    continue;
                    //}


                    // 全ての入金日が指定範囲内か検証
                    foreach (var m in t.Get新入金Rows())
                    {
                        if (m.入金年月日 < tempDate1 || m.入金年月日 > tempDate2)
                        {
                            hit = false;
                            break;
                        }

                        nyukin += m.金額;
                        nDate = m.入金年月日;
                    }

                    // 入金日範囲外あり：ネグる
                    if (!hit)
                    {
                        continue;
                    }

                    tempDGV.Rows.Add();

                    tempDGV[colNDt, iX].Value = nDate.ToShortDateString();

                    if (t.得意先Row != null)
                    {
                        tempDGV[colClient, iX].Value = t.得意先Row.略称;
                    }
                    else
                    {
                        tempDGV[colClient, iX].Value = string.Empty;
                    }

                    // 消費税を差し引く 2016/01/28
                    nyukin -= t.消費税;
                    tempDGV[colNyukin, iX].Value = nyukin.ToString("#,##0");

                    decimal genka = 0;
                    decimal gaichuhi = 0;
                    decimal gaichuhi2 = 0;
                    decimal gaichuhi3 = 0;

                    // 粗利算定
                    foreach (var j in t.Get受注1Rows())
                    {
                        //genka += ((decimal)j.枚数 * j.外注原価営業);
                        //gaichuhi += ((decimal)j.枚数 * j.外注原価支払);

                        // 2015/12/06 外注原価を単価から原価総額入力へ変更に伴う
                        genka += j.外注原価営業;
                        gaichuhi += j.外注原価支払; 
                        gaichuhi2 += j.外注原価支払2;     // 2016/10/24
                        gaichuhi3 += j.外注原価支払3;     // 2016/10/24
                    }

                    tempDGV[colEganka, iX].Value = genka.ToString("#,##0"); // 営業原価
                    tempDGV[colArari1, iX].Value = ((decimal)nyukin - genka).ToString("#,##0"); // 粗利１
                    tempDGV[colGaichuhi, iX].Value = gaichuhi.ToString("#,##0"); // 外注費
                    tempDGV[colGaichuhi2, iX].Value = gaichuhi2.ToString("#,##0"); // 外注費2
                    tempDGV[colGaichuhi3, iX].Value = gaichuhi3.ToString("#,##0"); // 外注費3
                    tempDGV[colArari2, iX].Value = ((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3).ToString("#,##0"); // 粗利２
                    tempDGV[colArariSai, iX].Value = (((decimal)nyukin - genka) - ((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3)).ToString("#,##0"); // 粗利差異

                    nTotal += nyukin;       // 入金額合計
                    Total += (int)genka;    // 営業原価
                    arari1Tl += (int)((decimal)nyukin - genka);     // 粗利１
                    gaichuhiTl += (int)gaichuhi;
                    gaichuhiTl2 += (int)gaichuhi2;
                    gaichuhiTl3 += (int)gaichuhi3;
                    arari2Tl += (int)((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3);  // 粗利２
                    arariSaiTl += (int)(((decimal)nyukin - genka) - ((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3));  // 粗利差異

                    iX++;
                }
                
                // 該当データなし
                if (tempDGV.RowCount == 0)
                {
                    MessageBox.Show("該当する入金済みデータ、および未収確定データがありませんでした", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                
                // 合計表示
                label5.Text = nTotal.ToString("#,##0");                 // 入金
                lblEgenka.Text = Total.ToString("#,##0");               // 営業原価
                lblArari1.Text = arari1Tl.ToString("#,##0");            // 粗利１
                lblGaichuhi.Text = gaichuhiTl.ToString("#,##0");        // 外注費
                lblGaichuhi2.Text = gaichuhiTl2.ToString("#,##0");      // 外注費2
                lblGaichuhi3.Text = gaichuhiTl3.ToString("#,##0");      // 外注費3
                lblArari2.Text = arari2Tl.ToString("#,##0");            // 粗利２
                lblArarisai.Text = arariSaiTl.ToString("#,##0");        // 粗利差異

                tempDGV.CurrentCell = null;
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
                label5.Text = "";
                lblGaichuhi.Text = "";

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

                // グリッド表示
                ShowDataLinq(dataGridView1, StartDate, EndDate, ShainID);

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
            const int S_GYO = 4;        //エクセルファイル見出し行（明細は3行目から印字）
            const int S_ROWSMAX = 8;    //エクセルファイル列最大値
            string sMidashi;
            double gTL = 0;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.営業売上表, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
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
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[colNDt, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[colClient, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[colNyukin, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[colEganka, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[colArari1, iX].Value.ToString();

                        double g = Utility.strToDouble(tempDGV[colGaichuhi, iX].Value.ToString()) +
                                   Utility.strToDouble(tempDGV[colGaichuhi2, iX].Value.ToString()) +
                                   Utility.strToDouble(tempDGV[colGaichuhi3, iX].Value.ToString());

                        oxlsSheet.Cells[iX + S_GYO, 6] = g.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[colArari2, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[colArariSai, iX].Value.ToString();

                        //セル下部へ実線ヨコ罫線を引く
                        rng[0] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        gTL += g;
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
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 4] = this.lblEgenka.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 5] = this.lblArari1.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 6] = gTL.ToString("#,##0");
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 7] = this.lblArari2.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 8] = this.lblArarisai.Text;

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
                MessageBox.Show(e.Message, "営業別売上表", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
            label5.Text = "";
            lblGaichuhi.Text = "";
        }

    }
}
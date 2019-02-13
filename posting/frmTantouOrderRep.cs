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
    public partial class frmTantouOrderRep : Form
    {
        const string MESSAGE_CAPTION = "営業担当者別受注一覧";
        const string COMBO_TOTAL = "合計";

        public frmTantouOrderRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: このコード行はデータを 'darwinDataSet.受注' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);

            //社員コンボ
            Utility.ComboShain.load(comboBox1);

            //////comboBox1.Items.Add(COMBO_TOTAL);   //社員合計項目を追加　2010/1/27

            DispClear();
        }

        // データグリッドビュークラス
        private class GridviewSet
        {
            ///-----------------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビューの定義を行います </summary>
            /// <param name="tempDGV">
            ///     データグリッドビューオブジェクト</param>
            ///-----------------------------------------------------------------------------
            public static void Setting(DataGridView tempDGV)
            {
                try
                {
                    //フォームサイズ定義

                    // 列スタイルを変更する
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    tempDGV.EnableHeadersVisualStyles = false;

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);
                    
                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 20;
                    tempDGV.RowTemplate.Height = 20;

                    // 全体の高さ
                    tempDGV.Height = 460;

                    // 奇数行の色
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "受注日");
                    tempDGV.Columns.Add("col2", "受注番号");
                    tempDGV.Columns.Add("col3", "チラシ名");
                    tempDGV.Columns.Add("col4", "担当");
                    tempDGV.Columns.Add("col5", "単価");
                    tempDGV.Columns.Add("col6", "枚数");
                    tempDGV.Columns.Add("col7", "受注合計");
                    tempDGV.Columns.Add("col8", "営業原価");
                    tempDGV.Columns.Add("col9", "粗利１");
                    tempDGV.Columns.Add("col10", "外注費");
                    tempDGV.Columns.Add("col11", "外注費２");
                    tempDGV.Columns.Add("col12", "外注費３");
                    tempDGV.Columns.Add("col13", "粗利２");
                    tempDGV.Columns.Add("col14", "粗利差異");

                    tempDGV.Columns[0].Width = 100;
                    tempDGV.Columns[1].Width = 110;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 100;
                    tempDGV.Columns[11].Width = 100;
                    tempDGV.Columns[12].Width = 100;
                    tempDGV.Columns[13].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    tempDGV.Columns[0].DefaultCellStyle.Format = "yyyy/MM/dd";
                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0.00";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[11].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[12].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[13].DefaultCellStyle.Format = "#,##0";

                    //チラシ名列はオートサイズモード
                    //tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static void ShowData(DataGridView tempDGV,DateTime tempDate1,DateTime tempDate2,int tempID)
            {
                string sqlSTRING = "";

                try
                {
                    if (tempID == 0)
                    {
                        tempDGV.Columns[3].Visible = true;
                    }
                    else
                    {
                        tempDGV.Columns[3].Visible = false;
                    }

                    Control.DataControl sdcon = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = sdcon.GetConnection();

                    //データリーダーを取得する
                    OleDbDataReader dR;

                    sqlSTRING += "select 受注.ID,受注.受注日,受注.チラシ名,受注.単価,受注.枚数,受注.金額,受注.値引額,社員.氏名, ";
                    sqlSTRING += "受注.外注原価支払, 受注.外注原価支払2, 受注.外注原価支払3, 受注.外注原価営業 ";
                    sqlSTRING += "from 受注 left join 得意先 on 受注.得意先ID = 得意先.ID ";
                    sqlSTRING += "left join 社員 on 得意先.担当社員コード = 社員.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(受注.受注日 >= ?) and ";
                    sqlSTRING += "(受注.受注日 <= ?) ";

                    if (tempID != 0) sqlSTRING += "and (得意先.担当社員コード = ?)";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@Date1", tempDate1);
                    SCom.Parameters.AddWithValue("@Date2", tempDate2);

                    if (tempID != 0) SCom.Parameters.AddWithValue("@SID", tempID);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //グリッドビューに表示する
                    int iX = 0;
                    int Total = 0;
                    decimal TotalGai = 0;
                    decimal TotalArari = 0;
                    decimal TotalGai2 = 0;
                    decimal TotalArari2 = 0;
                    decimal TotalGai3 = 0;
                    decimal TotalArari3 = 0;
                    decimal TotalGai4 = 0;
                    decimal TotalArari4 = 0;
                    decimal TotalArariSai = 0;

                    DateTime jDate = DateTime.Today; ;

                    tempDGV.RowCount = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = Convert.ToDateTime(dR["受注日"].ToString());
                        jDate = Convert.ToDateTime(dR["受注日"].ToString());
                        tempDGV[1, iX].Value = dR["ID"].ToString();
                        tempDGV[2, iX].Value = dR["チラシ名"].ToString();
                        tempDGV[3, iX].Value = dR["氏名"].ToString();                        
                        tempDGV[4, iX].Value = double.Parse(dR["単価"].ToString());
                        tempDGV[5, iX].Value = int.Parse(dR["枚数"].ToString(), System.Globalization.NumberStyles.Any);

                        // 2016/01/04 金額に値引を反映する
                        tempDGV[6, iX].Value = int.Parse(dR["金額"].ToString(), System.Globalization.NumberStyles.Any) - int.Parse(dR["値引額"].ToString(), System.Globalization.NumberStyles.Any);

                        // 営業原価　外注費 2015/09/18
                        //decimal gaichuhi = ((decimal)Utility.strToInt(dR["枚数"].ToString())) * Utility.strToDecimal(dR["外注原価営業"].ToString());
                        decimal gaichuhi = Utility.strToDecimal(dR["外注原価営業"].ToString()); // 2015/12/06
                        tempDGV[7, iX].Value = gaichuhi;

                        // 2016/01/04 金額に値引を反映する
                        decimal aRari = ((decimal)Utility.strToLong(dR["金額"].ToString())) - ((decimal)Utility.strToLong(dR["値引額"].ToString())) - gaichuhi;
                        tempDGV[8, iX].Value = aRari;

                        // 外注費 2015/09/18
                        decimal gaichuhi2 = Utility.strToDecimal(dR["外注原価支払"].ToString()); // 2015/12/06
                        tempDGV[9, iX].Value = gaichuhi2;

                        // 外注費2 2016/10/24
                        decimal gaichuhi3 = Utility.strToDecimal(dR["外注原価支払2"].ToString());
                        tempDGV[10, iX].Value = gaichuhi3;

                        // 外注費3 2016/10/24
                        decimal gaichuhi4 = Utility.strToDecimal(dR["外注原価支払3"].ToString());
                        tempDGV[11, iX].Value = gaichuhi4;

                        // 2016/01/04 金額に値引を反映する
                        decimal aRari2 = ((decimal)Utility.strToLong(dR["金額"].ToString())) - ((decimal)Utility.strToLong(dR["値引額"].ToString())) - gaichuhi2 - gaichuhi3 - gaichuhi4;
                        tempDGV[12, iX].Value = aRari2;

                        // 粗利差異 2015/09/18
                        tempDGV[13, iX].Value = aRari - aRari2;

                        Total += (int.Parse(dR["金額"].ToString(), System.Globalization.NumberStyles.Any) - int.Parse(dR["値引額"].ToString(), System.Globalization.NumberStyles.Any));
                        TotalGai += gaichuhi;   // 営業原価 2015/09/18
                        TotalArari += aRari;    // 粗利１合計 2015/09/18
                        TotalGai2 += gaichuhi2; // 外注費2合計 2015/09/18
                        TotalGai3 += gaichuhi3; // 外注費3合計 2016/11/08
                        TotalGai4 += gaichuhi4; // 外注費4合計 2016/11/08
                        TotalArari2 += aRari2;  // 粗利２合計 2015/09/18

                        TotalArariSai += (aRari -= aRari2);  // 差異合計 2016/10/24
                        iX++;
                    }

                    //合計行
                    if (tempDGV.RowCount == 0)
                    {
                        MessageBox.Show("該当するデータがありませんでした", MESSAGE_CAPTION);
                    }
                    else
                    {
                        tempDGV.Rows.Add();
                        tempDGV[0, iX].Value = "";
                        tempDGV[1, iX].Value = "";
                        tempDGV[2, iX].Value = "合計 : " + tempDGV.RowCount.ToString("#,##0") + " 件";
                        tempDGV[3, iX].Value = "";
                        tempDGV[4, iX].Value = "";
                        tempDGV[5, iX].Value = "";
                        tempDGV[6, iX].Value = Total;
                        tempDGV[7, iX].Value = TotalGai;        // 2015/09/18
                        tempDGV[8, iX].Value = TotalArari;      // 2015/09/18
                        tempDGV[9, iX].Value = TotalGai2;       // 2015/09/18
                        tempDGV[10, iX].Value = TotalGai3;    // 2015/09/18
                        tempDGV[11, iX].Value = TotalGai4;      // 2016/10/24
                        tempDGV[12, iX].Value = TotalArari2;      // 2016/10/24
                        tempDGV[13, iX].Value = TotalArariSai;  // 2015/09/18
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
                    if (comboBox1.Text != "")
                    {
                        MessageBox.Show("担当者の選択が正しくありません", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
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

                if (comboBox1.Text == "")
                {
                    ShainID = 0;
                }
                else
                {

                    //if (comboBox1.Text == COMBO_TOTAL)
                    //{
                    //    ShainID = 0;
                    //}
                    //else
                    //{

                    Utility.ComboShain cmb1 = new Utility.ComboShain();
                    cmb1 = (Utility.ComboShain)comboBox1.SelectedItem;
                    ShainID = cmb1.ID;

                    //}
                }

                GridviewSet.ShowData(dataGridView1, StartDate, EndDate, ShainID);

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
            int TempID;

            if (comboBox1.Text == "")
            {
                TempID = 1;
            }
            else
            {
                TempID = 0;
            }

            EigyoReport(dataGridView1,TempID);
        }


        private void EigyoReport(DataGridView tempDGV,int tempID)
        {
            const int S_GYO = 4;        // エクセルファイル見出し行（明細は3行目から印字）
            const int S_ROWSMAX = 11;   // エクセルファイル列最大値

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル営業担当者別受注一覧シート名, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        oxlsSheet.Cells[S_GYO - 2, 1] = this.comboBox1.Text;                // 担当
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();   // 受注日
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();   // 受注番号

                        //合計のとき
                        if (tempID == 1)
                        {
                            oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString() + "：" + tempDGV[3, iX].Value.ToString();
                        }
                        else
                        {
                            oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                        }

                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[4, iX].Value.ToString();   // 単価
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[5, iX].Value.ToString();   // 枚数
                        oxlsSheet.Cells[iX + S_GYO, 6] = tempDGV[6, iX].Value.ToString();   // 売上
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[7, iX].Value.ToString();   // 営業原価 2015/09/18
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[8, iX].Value.ToString();   // 粗利１   2015/09/18

                        // 外注費１，２，３合計 2016/10/24
                        double g = Utility.strToDouble(tempDGV[9, iX].Value.ToString()) +
                                Utility.strToDouble(tempDGV[10, iX].Value.ToString()) +
                                Utility.strToDouble(tempDGV[11, iX].Value.ToString());

                        oxlsSheet.Cells[iX + S_GYO, 9] = g.ToString();   // 外注費  2016/10/24
                        oxlsSheet.Cells[iX + S_GYO, 10] = tempDGV[12, iX].Value.ToString();   // 粗利２  2015/09/18
                        oxlsSheet.Cells[iX + S_GYO, 11] = tempDGV[13, iX].Value.ToString();   // 粗利差異  2015/09/18
                    }

                    //セル上部へ実線ヨコ罫線を引く
                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //セル下部へ実線ヨコ罫線を引く
                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

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

    }
}
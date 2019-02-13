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
    public partial class frmKadou2025 : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "配布員稼動日数一覧";

        public frmKadou2025()
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

            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();
            txtDays.Text = "20";
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
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 11, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 505;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "ID");
                    tempDGV.Columns.Add("col2", "配布員名");
                    tempDGV.Columns.Add("col3", "事業所");
                    tempDGV.Columns.Add("col4", "稼動数");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV,int tempYear,int tempMonth,int tempDays)
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
                    sqlSTRING += "select derivedtbl_1.ID, 配布員.氏名,事業所.名称, count(derivedtbl_1.ID) as 稼動数 ";
                    sqlSTRING += "from ";

                    sqlSTRING += "(select distinct 配布員ID as ID, 配布日 ";
                    sqlSTRING += "from 配布指示 ";
                    sqlSTRING += "where (year(配布日) = ?) AND (month(配布日) = ?) ";
                    sqlSTRING += "group by 配布員ID, 配布日) as derivedtbl_1 ";

                    sqlSTRING += "left join 配布員 ";
                    sqlSTRING += "on derivedtbl_1.ID = 配布員.ID ";
                    sqlSTRING += "left join 事業所 ";
                    sqlSTRING += "on 配布員.事業所コード = 事業所.ID ";
                    sqlSTRING += "group by derivedtbl_1.ID, 配布員.氏名,事業所.名称 ";
                    sqlSTRING += "having (count(derivedtbl_1.ID) >= ?) ";
                    sqlSTRING += "order by 稼動数 desc, derivedtbl_1.ID";
                                        
                    //配布指示データのデータリーダーを取得する
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@year", tempYear);
                    SCom.Parameters.AddWithValue("@month", tempMonth);
                    SCom.Parameters.AddWithValue("@days", tempDays);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();
                    
                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = int.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = dR["氏名"].ToString();
                            tempDGV[2, iX].Value = dR["名称"].ToString();
                            tempDGV[3, iX].Value = int.Parse(dR["稼動数"].ToString());

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
                    //    tempDGV.Columns[3].Width = 120;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[3].Width = 103;
                    //}

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
            if (MessageBox.Show("稼働日数一覧を表示します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            GridviewSet.ShowData(dataGridView2, int.Parse(txtYear.Text), int.Parse(txtMonth.Text),int.Parse(txtDays.Text));
            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("該当者はありませんでした", MESSAGE_CAPTION);
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

            if (MessageBox.Show("名簿を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 32; //配布完了報告書の明細行数
            int pCnt;

            //ページカウント
            pCnt = dataGridView2.Rows.Count / G_COUNT + 1;

            for (int i = 1; i <= pCnt; i++)
            {
                KanryoReport(dataGridView2);
            }

        }

        private void KanryoReport(DataGridView tempDGV)
        {

            const int S_GYO = 4;    //エクセルファイル見出し行（明細は3行目から印字）
            const int S_ROWSMAX = 5; //エクセルファイル列最大値

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布員稼動数一覧, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        oxlsSheet.Cells[S_GYO - 2, 1] = txtYear.Text + "年" + txtMonth.Text + "月 稼働日数 " + txtDays.Text + "日以上" ;
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[3, iX].Value.ToString();
                    }

                    ////セル上部へ実線ヨコ罫線を引く
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

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

                    //////DialogResult ret;

                    ////////ダイアログボックスの初期設定
                    //////saveFileDialog1.Title = "営業担当者別受注一覧";
                    //////saveFileDialog1.OverwritePrompt = true;
                    //////saveFileDialog1.RestoreDirectory = true;
                    //////saveFileDialog1.FileName = "営業担当者別受注一覧_" + comboBox1.Text;
                    //////saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

                    ////////ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    //////string fileName;
                    //////ret = saveFileDialog1.ShowDialog();

                    //////if (ret == System.Windows.Forms.DialogResult.OK)
                    //////{
                    //////    fileName = saveFileDialog1.FileName;
                    //////    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                    //////                    Type.Missing, Type.Missing, Type.Missing,
                    //////                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    //////                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //////}

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
            if (sender == txtDays) txtObj = txtDays;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtDays) txtObj = txtDays;

            txtObj.BackColor = Color.White;
        }

        private void txtDays_Validating(object sender, CancelEventArgs e)
        {

            string str;
            int d;

            if (txtDays.Text == null)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtDays.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (int.Parse(txtDays.Text) < 1)
            {
                MessageBox.Show("1以上で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }


    }
}
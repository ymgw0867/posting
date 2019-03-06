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
    public partial class frmMihaifuRep : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "未配布リスト";

        public frmMihaifuRep()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView2);
            GridviewSet.Setting2(dataGridView1);

            //////GridviewSet.ShowData(dataGridView2,"");
            //////dataGridView2.CurrentCell = null;
            //////dataGridView1.RowCount = 0;

            button2.Enabled = false;
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
                    tempDGV.Height = 595;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "受注番号");
                    tempDGV.Columns.Add("col2", "チラシ名");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 217;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
                    tempDGV.Height = 595;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "配布日");
                    tempDGV.Columns.Add("col2", "町名ID");
                    tempDGV.Columns.Add("col3", "町名");
                    tempDGV.Columns.Add("col4", "番地･号");
                    tempDGV.Columns.Add("col5", "マンション名");
                    tempDGV.Columns.Add("col6", "理由");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 66;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 200;
                    tempDGV.Columns[5].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    //tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    //tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    //tempDGV.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static void ShowData(DataGridView tempDGV,string tempCName)
            {

                string sqlSTRING = "";
                int iX;

                try
                {
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection Cn = new OleDbConnection();

                    Cn = Con.GetConnection();

                    tempDGV.RowCount = 0;
                    
                    //データリーダーを取得する
                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "SELECT 受注.ID,受注.チラシ名 ";
                    sqlSTRING += "from 受注 INNER JOIN 配布エリア ON 受注.ID = 配布エリア.受注ID ";
                    sqlSTRING += "INNER JOIN 未配布情報 ON 配布エリア.ID = 未配布情報.配布エリアID ";

                    sqlSTRING += "where 受注.チラシ名 like ? ";

                    sqlSTRING += "group by 受注.ID,受注.チラシ名 ";
                    sqlSTRING += "ORDER BY 受注.ID DESC";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@CName", "%" + tempCName + "%");

                    SCom.Connection = Cn;


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

                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();
                    Con.Close();
                    Cn.Close();

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
            dataGridView2.CurrentCell = null;
            dataGridView1.RowCount = 0;

            button2.Enabled = false;
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ShowPosting(dataGridView1, long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
            button2.Enabled = true;
        }

        private void ShowPosting(DataGridView tempDGV,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "SELECT 受注.ID,受注.チラシ名,配布エリア.町名ID,町名.名称,未配布情報.番地号,";
            mySql += "未配布情報.マンション名,未配布情報.理由,未配布理由.摘要,未配布情報.その他内容,";
            mySql += "配布指示.配布日 ";
            mySql += "from 受注 ";
            mySql += "inner join 配布エリア ON 受注.ID = 配布エリア.受注ID ";
            mySql += "inner join 未配布情報 on 配布エリア.ID = 未配布情報.配布エリアID ";
            mySql += "inner join 町名 ON 配布エリア.町名ID = 町名.ID ";
            mySql += "inner join 未配布理由 ON 未配布情報.理由 = 未配布理由.ID ";
            mySql += "inner join 配布指示 ON 配布エリア.配布指示ID = 配布指示.ID ";
            mySql += "where 受注.ID = " + tempID.ToString() + " ";
            mySql += " ORDER BY 配布日,配布エリア.町名ID,番地号";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;

            while (dR.Read())
            {
                tempDGV.Rows.Add();

                if (dR["配布日"] != DBNull.Value)
                {
                    tempDGV[0, iX].Value = DateTime.Parse(dR["配布日"].ToString()).ToShortDateString() + "";
                }
                else
                {
                    tempDGV[0, iX].Value = "";
                }

                tempDGV[1, iX].Value = dR["町名ID"].ToString() + "";
                tempDGV[2, iX].Value = dR["名称"].ToString() + "";
                tempDGV[3, iX].Value = dR["番地号"].ToString() + "";
                tempDGV[4, iX].Value = dR["マンション名"].ToString() + "";

                if (dR["摘要"].ToString() == "その他")
                {
                    tempDGV[5, iX].Value = dR["その他内容"].ToString() + "";
                }
                else
                {
                    tempDGV[5, iX].Value = dR["摘要"].ToString() + "";
                }

                iX++;
            }

            //////if (tempDGV.RowCount <= 27)
            //////{
            //////    tempDGV.Columns[3].Width = 200;
            //////}
            //////else
            //////{
            //////    tempDGV.Columns[3].Width = 183;
            //////}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("配布完了報告書を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 32; //配布完了報告書の明細行数
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

            const int S_GYO = 13;    //エクセルファイル明細は13行目から印字
            int dgvIndex;
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

                try
                {
                    //得意先情報
                    long sID;
                    string sqlSTR;

                    sID = long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString());

                    sqlSTR = "";
                    sqlSTR += "select 得意先.名称,得意先.担当者名,得意先.電話番号 ";
                    sqlSTR += "from 受注 inner join 得意先 ";
                    sqlSTR += "on 受注.得意先ID = 得意先.ID ";
                    sqlSTR += "where 受注.ID = " + sID.ToString();

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTR);

                    while (dR.Read())
                    {
                        oxlsSheet.Cells[1, 3] = dR["名称"].ToString() + " " + dR["担当者名"].ToString() + "様";
                        oxlsSheet.Cells[2, 3] = dR["電話番号"].ToString();                        
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
                    
                    //配布エリア明細
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //データグリッドビューの行インデックスを求める

                        //oxlsSheet.Cells[i + S_GYO, 2] = dateTimePicker1.Value.ToShortDateString();   //チラシ名
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView1[0, dgvIndex].Value.ToString();   //配布区分
                        oxlsSheet.Cells[i + S_GYO, 4] = int.Parse(dataGridView1[1, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //グリッド最終行のとき終了
                        if (dgvIndex == (dataGridView1.Rows.Count - 1)) break;

                        //印刷明細最大行のとき終了
                        if (i == (tempFixRows - 1)) break;

                        i++;
                    }

                    //配布枚数合計
                    oxlsSheet.Cells[45, 4] = int.Parse(dataGridView2[5, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    //残部数
                    oxlsSheet.Cells[46, 4] = int.Parse(dataGridView2[6, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

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
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString() + "_" + DateTime.Today.ToLongDateString();

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

        private void button2_Click_1(object sender, EventArgs e)
        {
            string csvTittle;

            csvTittle = dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString();
            MyLibrary.CsvOut.GridView(dataGridView1, csvTittle);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            GridviewSet.ShowData(dataGridView2, txtCName.Text);
            dataGridView2.CurrentCell = null;
            dataGridView1.RowCount = 0;
        }


    }
}
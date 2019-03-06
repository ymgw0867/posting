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
    public partial class frmSeikyuAdd : Form
    {
        Entity.請求書 cMaster = new Entity.請求書();
        Entity.入金 cNyukin = new Entity.入金();

        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "請求データ作成";

        private DateTime sD;
        private DateTime eD;

        public frmSeikyuAdd(int pMode)
        {
            
            InitializeComponent();

            //処理モード　0:登録、1:更新
            this.fMode.Mode  = pMode;
        }

        private void frmSeikyuAdd_Load(object sender, EventArgs e)
        {

            sD = Convert.ToDateTime("1900/01/01");
            eD = Convert.ToDateTime("9999/12/31");

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            //GridviewSet.Setting(dataGridView2);

            GridviewSet.Setting2(dataGridView1);
            GridviewSet.Setting3(dataGridView3,fMode.Mode);

            switch (fMode.Mode)
            {
                case 0: //新規登録
                    this.Text = "請求書データ登録";
                    label15.Text = "【未請求クライアント】" + Environment.NewLine + "配布終了日：クライアント名";

                    ListClient.load(listBox1,"",sD,eD);

                    //入金オブジェクト非表示
                    panel1.Hide();
                    label20.Hide();
                    label23.Hide();
                    label25.Hide();
                    txtZan.Hide();
                    checkBox1.Hide();
                    dataGridView4.Hide();

                    break;

                case 1: //編集
                    this.Text = "請求書データ編集";
                    label15.Text = "【請求書データ】" + Environment.NewLine + "配布終了日：クライアント名";
                    radioButton1.Checked = true;
                    radioButton2.Checked = false;

                    ListClient.loadData(listBox1, getRbtn(), "",sD,eD);

                    //入金オブジェクト表示
                    panel1.Show();
                    label20.Show();
                    label23.Show();
                    label25.Show();
                    txtZan.Show();
                    checkBox1.Show();
                    dataGridView4.Show();

                    GridviewSet.Setting4(dataGridView4);
                    
                    break;

            }

            DispClear();

            label1.Text = "【受注データ】";
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
                    tempDGV.Height = 145;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "ｺｰﾄﾞ");
                    tempDGV.Columns.Add("col2", "名称");
                    tempDGV.Columns.Add("col3", "担当者");
                    tempDGV.Columns.Add("col4", "〒");
                    tempDGV.Columns.Add("col5", "住所");
                    tempDGV.Columns.Add("col6", "ご担当者");
                    tempDGV.Columns.Add("col7", "TEL");
                    tempDGV.Columns.Add("col8", "FAX");

                    tempDGV.Columns[0].Width = 40;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 300;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 80;
                    //tempDGV.Columns[8].Width = 80;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    //tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";

                    //tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

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
                    tempDGV.Height = 145;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "受注番号");
                    tempDGV.Columns.Add("col2", "受注日");
                    tempDGV.Columns.Add("col3", "チラシ名");
                    tempDGV.Columns.Add("col4", "受注種別");
                    tempDGV.Columns.Add("col5", "単価");
                    tempDGV.Columns.Add("col6", "枚数");
                    tempDGV.Columns.Add("col7", "売上金額");
                    tempDGV.Columns.Add("col8", "消費税");
                    tempDGV.Columns.Add("col9", "税込金額");
                    tempDGV.Columns.Add("col10", "値引額");
                    tempDGV.Columns.Add("col11", "請求金額");
                    tempDGV.Columns.Add("col12", "サイズ");
                    tempDGV.Columns.Add("col13", "配布開始日");
                    tempDGV.Columns.Add("col14", "配布終了日");
                    tempDGV.Columns.Add("col15", "入金予定日");
                    tempDGV.Columns.Add("col16", "配布形態");
                    tempDGV.Columns.Add("col17", "税率");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;
                    tempDGV.Columns[11].Width = 90;
                    tempDGV.Columns[12].Width = 110;
                    tempDGV.Columns[13].Width = 110;
                    tempDGV.Columns[14].Width = 110;
                    tempDGV.Columns[15].Width = 120;
                    tempDGV.Columns[16].Width = 60;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void Setting3(DataGridView tempDGV,int tempMode)
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

                    switch (tempMode)
                    {
                        case 0:
                            tempDGV.Width = 970;
                            break;

                        case 1:
                            tempDGV.Width = 740;
                            break;
                    }

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "配布日");
                    tempDGV.Columns.Add("col2", "配布物");
                    tempDGV.Columns.Add("col3", "形態");
                    tempDGV.Columns.Add("col4", "チラシ名");
                    tempDGV.Columns.Add("col5", "単価");
                    tempDGV.Columns.Add("col6", "枚数");
                    tempDGV.Columns.Add("col7", "売上金額");
                    tempDGV.Columns.Add("col8", "消費税");
                    tempDGV.Columns.Add("col9", "税込金額");
                    tempDGV.Columns.Add("col10", "値引額");
                    tempDGV.Columns.Add("col11", "請求金額");
                    tempDGV.Columns.Add("col12", "受注ID");
                    tempDGV.Columns.Add("col13", "税率");

                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 100;
                    tempDGV.Columns[2].Width = 120;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;

                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void Setting4(DataGridView tempDGV)
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
                    tempDGV.Height = 73;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "ID");
                    tempDGV.Columns.Add("col2", "入金日");
                    tempDGV.Columns.Add("col3", "金額");

                    tempDGV.Columns[0].Visible = false;

                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 137;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// 請求対象得意先表示
            /// </summary>
            /// <param name="tempDGV"></param>
            public static void ShowData(DataGridView tempDGV)
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
                    sqlSTRING += "select DISTINCT 受注.得意先ID, 得意先.*,社員.氏名 ";
                    sqlSTRING += "from (受注 left join 得意先 ";
                    sqlSTRING += "on 受注.得意先ID = 得意先.ID) left join 社員 ";
                    sqlSTRING += "on 得意先.担当社員コード = 社員.ID ";
                    sqlSTRING += "where  (受注.請求書 = 1) AND (受注.請求書ID = 0) ";
                    sqlSTRING += "order by 受注.得意先ID";
                
                    //未請求得意先のデータリーダーを取得する
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);
                    
                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = long.Parse(dR["得意先ID"].ToString());
                            tempDGV[1, iX].Value = dR["名称"].ToString() + "";
                            tempDGV[2, iX].Value = dR["氏名"].ToString();
                            tempDGV[3, iX].Value = dR["請求先郵便番号"].ToString() + "";
                            tempDGV[4, iX].Value = dR["請求先住所1"].ToString() + dR["請求先住所2"].ToString() + "";
                            tempDGV[5, iX].Value = dR["担当者名"].ToString();
                            tempDGV[6, iX].Value = dR["電話番号"].ToString();
                            tempDGV[7, iX].Value = dR["FAX番号"].ToString();

                            iX++;

                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();

                    fCon.Close();

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

        private void ShowPosting(DataGridView tempDGV,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select 受注.*, 受注種別.名称 as 受注種別名称,判型.名称 as 判型名称,配布形態.名称 as 配布形態名称 ";
            mySql += "from ((受注 left join 受注種別 ";
            mySql += "on 受注.受注種別ID = 受注種別.ID) left join 判型 ";
            mySql += "on 受注.判型 = 判型.ID) left join 配布形態 ";
            mySql += "on 受注.配布形態 = 配布形態.ID ";
            mySql += "where ";
            mySql += "(受注.得意先ID = " + tempID.ToString() + ") and ";
            mySql += "(受注.請求書 = 1) and ";
            mySql += "(受注.請求書ID = 0) ";
            mySql += "order by 受注.ID desc";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;
            button1.Enabled = false;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["ID"].ToString();
                tempDGV[1, iX].Value = DateTime.Parse(dR["受注日"].ToString()).ToShortDateString();
                tempDGV[2, iX].Value = dR["チラシ名"].ToString();
                
                if (dR["受注種別名称"] == DBNull.Value)
                {
                    tempDGV[3, iX].Value = "";
                }
                else
                {
                    tempDGV[3, iX].Value = dR["受注種別名称"].ToString();
                }

                tempDGV[4, iX].Value = double.Parse(dR["単価"].ToString());
                tempDGV[5, iX].Value = int.Parse(dR["枚数"].ToString());
                tempDGV[6, iX].Value = double.Parse(dR["金額"].ToString());
                tempDGV[7, iX].Value = double.Parse(dR["消費税"].ToString());
                tempDGV[8, iX].Value = double.Parse(dR["税込金額"].ToString());
                tempDGV[9, iX].Value = double.Parse(dR["値引額"].ToString());
                tempDGV[10, iX].Value = double.Parse(dR["売上金額"].ToString());

                if (dR["判型名称"] == DBNull.Value)
                {
                    tempDGV[11, iX].Value = "";
                }
                else
                {
                    tempDGV[11, iX].Value = dR["判型名称"].ToString();
                }

                if (dR["配布開始日"] == DBNull.Value)
                {
                    tempDGV[12, iX].Value = "";
                }
                else
                {
                    tempDGV[12, iX].Value = DateTime.Parse(dR["配布開始日"].ToString()).ToShortDateString();
                }

                if (dR["配布終了日"] == DBNull.Value)
                {
                    tempDGV[13, iX].Value = "";
                }
                else
                {
                    tempDGV[13, iX].Value = DateTime.Parse(dR["配布終了日"].ToString()).ToShortDateString();
                }

                if (dR["入金予定日"] == DBNull.Value)
                {
                    tempDGV[14, iX].Value = "";
                }
                else
                {
                    tempDGV[14, iX].Value = DateTime.Parse(dR["入金予定日"].ToString()).ToShortDateString();
                }

                if (dR["配布形態名称"] == DBNull.Value)
                {
                    tempDGV[15, iX].Value = "";
                }
                else
                {
                    tempDGV[15, iX].Value = dR["配布形態名称"].ToString();
                }

                tempDGV[16, iX].Value = dR["税率"].ToString();

                iX++;
            }

            //if (tempDGV.RowCount <= 8)
            //{
            //    tempDGV.Columns[6].Width = 252;
            //}
            //else
            //{
            //    tempDGV.Columns[6].Width = 235;
            //}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();

            switch (fMode.Mode)
            {
                case 0:
                    label1.Text = "【" + listBox1.Text + " 受注データ 】";
                    break;

                case 1:
                    label1.Text = "【" + listBox1.Text.Substring(11,listBox1.Text.Length -11) + " 受注データ 】";
                    break;
            }

            dataGridView3.RowCount = 0;

            if (tempDGV.RowCount > 0)
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }

            button7.Enabled = true;
        }

        private void ShowClient(int tempID)
        {
            OleDbDataReader dR;
            Control.得意先 cCli = new Control.得意先();
            dR = cCli.FillBy("where ID = " + tempID.ToString());

            while(dR.Read())
            {
                cMaster.得意先ID = int.Parse(dR["ID"].ToString());

                txtClient.Text = dR["名称"].ToString() + "";                                            //得意先名
                txtKeisho.Text = dR["敬称"].ToString() + "";                                            //敬称
                txtZipCode.Text = dR["郵便番号"].ToString() + "";                                       //郵便番号
                txtAddress.Text = dR["請求先住所1"].ToString() + "　" + dR["請求先住所2"].ToString();   //住所
                txtTantou.Text = dR["担当者名"].ToString() + "";                                        //ご担当者
                txtTel.Text = dR["電話番号"].ToString() + "";                                           //TEL
                txtFax.Text = dR["FAX番号"].ToString() + "";                                            //FAX
            }

            dR.Close();
            cCli.Close();

        }

        /// <summary>
        /// 請求書データ読み込み
        /// </summary>
        /// <param name="tempID">請求書ID</param>
        private void ShowSeikyuData(int tempID)
        {
            int iX;
            OleDbDataReader dR,dRToku;
            Control.請求書 cData = new Control.請求書();
            dR = cData.FillBy("where ID = " + tempID.ToString());

            while (dR.Read())
            {
                GetData(dR);

                //得意先マスタ取得
                Control.得意先 cCli = new Control.得意先();
                dRToku = cCli.FillBy("where ID = " + cMaster.得意先ID.ToString());

                //得意先情報表示
                while (dRToku.Read())
	            {
                    txtClient.Text = dRToku["名称"].ToString() + "";    //得意先名
                    txtKeisho.Text = dRToku["敬称"].ToString() + "";    //敬称
                    txtZipCode.Text = dRToku["郵便番号"].ToString() + "";   //郵便番号
                    txtAddress.Text = dRToku["請求先住所1"].ToString() + "　" + dRToku["請求先住所2"].ToString();   //住所
                    txtTantou.Text = dRToku["担当者名"].ToString() + "";    //ご担当者
                    txtTel.Text = dRToku["電話番号"].ToString() + "";   //TEL
                    txtFax.Text = dRToku["FAX番号"].ToString() + "";    //FAX
	            }

                dRToku.Close();
                cCli.Close();
            }

            dR.Close();
            cData.Close();

            //請求書データ表示
            nDate.Checked = true;

            tabControl1.TabPages[0].Text = "請求書№" + cMaster.ID.ToString();

            nDate.Value = cMaster.入金予定日;
            label18.Text = cMaster.税率.ToString();
            label8.Text = cMaster.売上金額.ToString("#,##0");
            label9.Text = cMaster.消費税.ToString("#,##0");
            label10.Text = cMaster.値引額.ToString("#,##0");
            label11.Text = cMaster.請求金額.ToString("#,##0");
            txtMemo.Text = cMaster.備考;

            switch (cMaster.完了区分)
	        {
		        case 0:
                    checkBox1.Checked = false;
                    label24.Hide();
                    break;

                case 1:
                    checkBox1.Checked = true;
                    label24.Show();
                    break;
	        }

            txtZan.Text = cMaster.入金残.ToString("#,##0");

            //請求明細表示
            dataGridView3.RowCount = 0;

            Control.FreeSql fCon = new Control.FreeSql();

            string mySql = "";
            mySql += "select 受注.*, 受注種別.名称 as 受注種別名称,判型.名称 as 判型名称,配布形態.名称 as 配布形態名称 ";
            mySql += "from ((受注 left join 受注種別 ";
            mySql += "on 受注.受注種別ID = 受注種別.ID) left join 判型 ";
            mySql += "on 受注.判型 = 判型.ID) left join 配布形態 ";
            mySql += "on 受注.配布形態 = 配布形態.ID ";
            mySql += "where ";
            mySql += "(受注.請求書ID = " + cMaster.ID.ToString() + ")";
            mySql += "order by 受注.ID desc";

            dR = fCon.free_dsReader(mySql);

            while (dR.Read())
            {
                dataGridView3.Rows.Add();

                iX = dataGridView3.Rows.Count;

                if (dR["配布開始日"] == DBNull.Value)
                {
                    dataGridView3[0, iX - 1].Value = "";
                }
                else
                {
                    dataGridView3[0, iX - 1].Value = DateTime.Parse(dR["配布開始日"].ToString()).ToShortDateString();
                }

                if (dR["判型名称"] == DBNull.Value)
                {
                    dataGridView3[1, iX - 1].Value = "";
                }
                else
                {
                    dataGridView3[1, iX - 1].Value = dR["判型名称"].ToString();
                }

                if (dR["配布形態名称"] == DBNull.Value)
                {
                    dataGridView3[2, iX - 1].Value = "";
                }
                else
                {
                    dataGridView3[2, iX - 1].Value = dR["配布形態名称"].ToString();
                }

                dataGridView3[3, iX - 1].Value = dR["チラシ名"].ToString();
                dataGridView3[4, iX - 1].Value = double.Parse(dR["単価"].ToString());
                dataGridView3[5, iX - 1].Value = int.Parse(dR["枚数"].ToString());
                dataGridView3[6, iX - 1].Value = double.Parse(dR["金額"].ToString());
                dataGridView3[7, iX - 1].Value = double.Parse(dR["消費税"].ToString());
                dataGridView3[8, iX - 1].Value = double.Parse(dR["税込金額"].ToString());
                dataGridView3[9, iX - 1].Value = double.Parse(dR["値引額"].ToString());
                dataGridView3[10, iX - 1].Value = double.Parse(dR["売上金額"].ToString());
                dataGridView3[11, iX - 1].Value = dR["ID"].ToString();
                dataGridView3[12, iX - 1].Value = dR["税率"].ToString();

            }

            dR.Close();
            fCon.Close();

            dataGridView3.CurrentCell = null;

            //入金履歴表示
            ShowNyukin();

            //ボタン表示
            button1.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;

        }

        private void ShowNyukin()
        {
            //入金履歴表示
            int iX;
            OleDbDataReader dR;
            dataGridView4.RowCount = 0;
            Control.入金 cNyukin = new Control.入金();
            dR = cNyukin.FillBy("where 請求書ID = " + cMaster.ID.ToString());

            while (dR.Read())
            {
                dataGridView4.Rows.Add();

                iX = dataGridView4.Rows.Count;
                dataGridView4[0, iX - 1].Value = int.Parse(dR["ID"].ToString());
                dataGridView4[1, iX - 1].Value = DateTime.Parse(dR["入金年月日"].ToString()).ToShortDateString();
                dataGridView4[2, iX - 1].Value = int.Parse(dR["金額"].ToString(), System.Globalization.NumberStyles.Any);
            }

            if (dataGridView4.RowCount > 3)
            {
                dataGridView4.Columns[2].Width = 120;
            }
            else
            {
                dataGridView4.Columns[2].Width = 137;
            }

            dR.Close();
            cNyukin.Close();

            dataGridView4.CurrentCell = null;
        }

        private void GetData(OleDbDataReader dR)
        {
            cMaster.ID = int.Parse(dR["ID"].ToString());
            cMaster.得意先ID = int.Parse(dR["得意先ID"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.請求金額 = int.Parse(dR["請求金額"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.消費税 = int.Parse(dR["消費税"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.値引額 = int.Parse(dR["値引額"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.売上金額 = int.Parse(dR["売上金額"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.税率 = int.Parse(dR["税率"].ToString());
            cMaster.入金予定日 = DateTime.Parse(dR["入金予定日"].ToString());
            cMaster.発行日 = DateTime.Parse(dR["発行日"].ToString());
            cMaster.入金残 = int.Parse(dR["入金残"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.完了区分 = int.Parse(dR["完了区分"].ToString());
            cMaster.振込口座ID1 = int.Parse(dR["振込口座ID1"].ToString());
            cMaster.振込口座ID2 = int.Parse(dR["振込口座ID2"].ToString());
            cMaster.備考 = dR["備考"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("請求書を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 12; //請求書の明細行数

            int pCnt;

            //ページカウント
            pCnt = dataGridView3.Rows.Count / G_COUNT + 1;

            for (int i = 1; i <= pCnt; i++)
            {
                SeikyuReport(pCnt, i, G_COUNT);
            }

        }

        private void SeikyuReport(int tempPage, int tempCurrentPage, int tempFixRows)
        {

            const int S_GYO = 17;    //エクセルファイル明細は8行目から印字
            int dgvIndex;
            int i;
            DateTime sDate;

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル請求書, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    //ページ数
                    oxlsSheet.Cells[2, 11] = tempCurrentPage.ToString() + "/" + tempPage.ToString();

                    //得意先名
                    oxlsSheet.Cells[1, 1] = txtClient.Text;

                    //敬称
                    oxlsSheet.Cells[1, 5] = txtKeisho.Text;

                    //郵便番号
                    oxlsSheet.Cells[3, 1] = "〒 " + txtZipCode.Text;

                    //住所
                    oxlsSheet.Cells[3, 2] = txtAddress.Text;

                    //担当者名
                    oxlsSheet.Cells[4, 1] = txtTantou.Text;

                    //TEL
                    //oxlsSheet.Cells[5, 1] = "TEL： " + txtTel.Text;

                    //FAX
                    //oxlsSheet.Cells[5, 3] = "FAX： " + txtFax.Text;

                    //請求金額
                    if (tempCurrentPage == 1)
                    {
                        oxlsSheet.Cells[13, 3] = int.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                    }
                    else
                    {
                        oxlsSheet.Cells[13, 3] = "―";
                    }

                    //備考
                    oxlsSheet.Cells[30, 2] = txtMemo.Text;

                    //支払期日
                    oxlsSheet.Cells[35, 2] = nDate.Value.ToLongDateString();

                    //値引見出し
                    if (label10.Text == "0")
                    {
                        oxlsSheet.Cells[S_GYO - 1, 10] = ""; 
                    }
                    else
                    {
                        oxlsSheet.Cells[S_GYO - 1, 10] = "値引額"; 
                    }

                    //請求明細
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //データグリッドビューの行インデックスを求める

                        //配布日
                        sDate = DateTime.Parse(dataGridView3[0, dgvIndex].Value.ToString());
                        oxlsSheet.Cells[i + S_GYO, 1] = sDate.Month.ToString() + "/" + sDate.Day.ToString() + "～";

                        //配布物(サイズ)
                        oxlsSheet.Cells[i + S_GYO, 2] = dataGridView3[1, dgvIndex].Value.ToString();

                        //形態
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView3[2, dgvIndex].Value.ToString();

                        //チラシ名
                        oxlsSheet.Cells[i + S_GYO, 4] = dataGridView3[3, dgvIndex].Value.ToString();

                        //数量
                        oxlsSheet.Cells[i + S_GYO, 6] = int.Parse(dataGridView3[5, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //単価
                        oxlsSheet.Cells[i + S_GYO, 7] = double.Parse(dataGridView3[4, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //金額
                        oxlsSheet.Cells[i + S_GYO, 8] = int.Parse(dataGridView3[6, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //消費税
                        oxlsSheet.Cells[i + S_GYO, 9] = int.Parse(dataGridView3[7, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //値引額
                        oxlsSheet.Cells[i + S_GYO, 10] = int.Parse(dataGridView3[9, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //請求金額
                        oxlsSheet.Cells[i + S_GYO, 11] = int.Parse(dataGridView3[10, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //グリッド最終行のとき終了
                        if (dgvIndex == (dataGridView3.Rows.Count - 1)) break;

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
                    saveFileDialog1.FileName = "請求書" + "_" + txtClient.Text + DateTime.Today.Year.ToString() + "年" + DateTime.Today.Month.ToString() + "月";

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

        //得意先コンボボックスクラス
        private class ListClient
        {
            private int F_ID;
            private string F_Name;
            private int F_tID;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            public int tID
            {
                get
                {
                    return F_tID;
                }
                set
                {
                    F_tID = value;
                }
            }

            //未請求得意先マスターロード
            public static void load(ListBox tempObj, string tempClient, DateTime sDate, DateTime eDate)
            {
                try
                {
                    OleDbDataReader dR;
                    ListClient lst1;
                    string sqlSTRING;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    //データリーダーを取得する
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    sqlSTRING = "";
                    sqlSTRING += "select 受注.得意先ID,得意先.名称 as 得意先名称,受注.配布終了日 ";
                    sqlSTRING += "from 受注 inner join 得意先 ";
                    sqlSTRING += "on 受注.得意先ID = 得意先.ID ";
                    sqlSTRING += "where  (受注.請求書 = 1) AND (受注.請求書ID = 0) ";

                    if (tempClient != "")    //2009.09.09 クライアント検索
                    {
                        sqlSTRING += "and (得意先.略称 like '%" + tempClient + "%')";
                    }

                    sqlSTRING += " and (受注.配布終了日 >= '" + sDate + "') and (受注.配布終了日 <= '" + eDate + "') ";
                    sqlSTRING += "order by 受注.配布終了日 desc,受注.得意先ID";

                    //未請求得意先のデータリーダーを取得する
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);

                    while (dR.Read())
                    {
                        lst1 = new ListClient();
                        lst1.ID = Int32.Parse(dR["得意先ID"].ToString());
                        lst1.Name = DateTime.Parse(dR["配布終了日"].ToString()).ToShortDateString() + " " + dR["得意先名称"].ToString() + "";
                        tempObj.Items.Add(lst1);
                    }

                    dR.Close();
                    Con.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "未請求得意先リストボックスロード");
                }
            }

            //請求書データロード
            public static void loadData(ListBox tempObj, int tempS, string tempClient, DateTime sDate, DateTime eDate)
            {
                try
                {
                    OleDbDataReader dR;
                    ListClient lst1;
                    string sqlSTRING;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    //データリーダーを取得する
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    sqlSTRING = "";
                    sqlSTRING += "select 請求書.*,得意先.名称,受注.配布終了日 ";
                    sqlSTRING += "from 請求書 inner join 得意先 ";
                    sqlSTRING += "on 請求書.得意先ID = 得意先.ID ";
                    sqlSTRING += "inner join 受注 ";
                    sqlSTRING += "on 請求書.ID = 受注.請求書ID ";
                    sqlSTRING += "where (1 = 1) ";

                    if (tempS == 1)
                    {
                        sqlSTRING += "and (請求書.完了区分 = 0) ";
                    }

                    if (tempClient != "")    //2009.09.09 クライアント検索
                    {
                        sqlSTRING += "and (得意先.略称 like '%" + tempClient + "%')";
                    }

                    sqlSTRING += " and (受注.配布終了日 >= '" + sDate + "') and (受注.配布終了日 <= '" + eDate + "') ";

                    sqlSTRING += "order by 受注.配布終了日 desc";

                    //請求書のデータリーダーを取得する
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);

                    while (dR.Read())
                    {
                        lst1 = new ListClient();
                        lst1.ID = Int32.Parse(dR["ID"].ToString());
                        lst1.Name = DateTime.Parse(dR["配布終了日"].ToString()).ToShortDateString() + " " + dR["名称"].ToString() + "";

                        if (dR["完了区分"].ToString() == "1") lst1.Name += " 【完了】";

                        lst1.tID = Int32.Parse(dR["得意先ID"].ToString());

                        tempObj.Items.Add(lst1);
                    }

                    dR.Close();
                    Con.Close();
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "請求書リストボックスロード");
                }
            }


        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBox1.SelectedIndex == -1) return;

            if (MessageBox.Show(listBox1.Text  + "　が選択されました。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            ListClient lst1 = new ListClient();
            lst1 = (ListClient)listBox1.SelectedItem;

            switch (fMode.Mode)
            {
                case 0: //新規登録
                    ShowPosting(dataGridView1, lst1.ID);    //未請求受注データ表示
                    ShowClient(lst1.ID);                    //クライアント情報表示
                    break;

                case 1: //編集
                    ShowPosting(dataGridView1, lst1.tID);   //未請求受注データ表示
                    ShowSeikyuData(lst1.ID);                //請求データ表示
                    break;
            }

            panel2.Hide();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("受注データを選択してください", "請求明細登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("選択されている受注データを請求明細に追加します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {

                int iX;

                dataGridView3.Rows.Add();
                iX = dataGridView3.Rows.Count;

                dataGridView3[0, iX-1].Value = dataGridView1[12, r.Index].Value.ToString();
                dataGridView3[1, iX-1].Value = dataGridView1[11, r.Index].Value.ToString();
                dataGridView3[2, iX-1].Value = dataGridView1[15, r.Index].Value.ToString();
                dataGridView3[3, iX-1].Value = dataGridView1[2, r.Index].Value.ToString();
                dataGridView3[4, iX-1].Value = double.Parse(dataGridView1[4, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[5, iX-1].Value = int.Parse(dataGridView1[5, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[6, iX-1].Value = double.Parse(dataGridView1[6, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[7, iX-1].Value = double.Parse(dataGridView1[7, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[8, iX-1].Value = double.Parse(dataGridView1[8, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[9, iX-1].Value = double.Parse(dataGridView1[9, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[10, iX-1].Value = double.Parse(dataGridView1[10, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[11, iX - 1].Value = dataGridView1[0, r.Index].Value.ToString();
                dataGridView3[12, iX - 1].Value = int.Parse(dataGridView1[16, r.Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                //入金予定日
                if (dataGridView1[14, r.Index].Value.ToString() == "")
                {
                    nDate.Checked = false;
                }
                else
                {
                    nDate.Checked = true;
                    nDate.Value = DateTime.Parse(dataGridView1[14, r.Index].Value.ToString());
                }

                //税率
                if (label18.Text == "") label18.Text = dataGridView3[12, iX - 1].Value.ToString();

            }

            //請求合計金額計算
            SumKingaku();
            
            button1.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;

            //受注データから行削除する
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(r);
            }

        }

        private void SumKingaku()
        {

            double sTotal = 0;
            double sTax = 0;
            double sNebiki = 0;
            double sSeikyu = 0;

            //金額計算
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                sTotal += double.Parse(dataGridView3[6, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                sTax += double.Parse(dataGridView3[7, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                sNebiki += double.Parse(dataGridView3[9, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                sSeikyu += double.Parse(dataGridView3[10, i].Value.ToString(), System.Globalization.NumberStyles.Any);
            }

            label8.Text = sTotal.ToString("#,##0");
            label9.Text = sTax.ToString("#,##0");
            label10.Text = sNebiki.ToString("#,##0");
            label11.Text = sSeikyu.ToString("#,##0");

            dataGridView3.CurrentCell = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("受注データ全てを選択状態にします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Selected = true;
            } 

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try 
	        {	        
                if (dataGridView3.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請求明細を選択してください", "請求データ削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (MessageBox.Show("選択されている請求明細を削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                foreach (DataGridViewRow r in dataGridView3.SelectedRows)
                {
                    string mySql = "";
                    Control.FreeSql fCon = new Control.FreeSql();

                    //編集処理のとき受注データの請求書区分を初期化
                    if (fMode.Mode == 1)
                    {

                        mySql += "update 受注 set ";
                        mySql += "請求書ID = 0,";
                        mySql += "請求書発行日 = null,";
                        mySql += "変更年月日 = '" + DateTime.Today + "' ";
                        mySql += "where ID = " + dataGridView3[11, dataGridView3.SelectedRows[0].Index].Value.ToString();

                        if (fCon.Execute(mySql) == false)
                        {
                            fCon.Close();
                            throw new Exception("受注データの更新に失敗しました");
                        }
                    }

                    fCon.Close();

                    //データグリッド行削除
                    dataGridView3.Rows.Remove(r);
                }

                //請求金額再計算
                SumKingaku();

                //請求明細全てを削除した場合
                DateTime StartDate;
                DateTime EndDate;

                if (dataGridView3.RowCount == 0)
                {
                    switch (fMode.Mode)
                    {
                        case 0: //登録

                            button1.Enabled = false;
                            button5.Enabled = false;
                            button6.Enabled = false;

                            break;

                        case 1: //編集
                            MessageBox.Show("請求明細が全て削除されましたので請求書データは削除されます。", "請求書削除", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            Control.請求書 cSe = new Control.請求書();
                            cSe.DataDelete(cMaster.ID);
                            cSe.Close();

                            //請求書リストボックスリロード
                            if (tDate.Checked == false)
                            {
                                StartDate = sD;
                            }
                            else
                            {
                                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
                            }

                            if (tDate2.Checked == false)
                            {
                                EndDate = eD;
                            }
                            else
                            {
                                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
                            }
    
                            ListClient.loadData(listBox1, getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
                            DispClear();

                            break;
                    }

                }	
            }

	        catch (Exception ex)
	        {
                MessageBox.Show(ex.Message, "請求明細削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void DispClear()
        {
            switch (fMode.Mode)
            {
                case 0:
                    panel2.Hide();
                    break;

                case 1:
                    panel2.Show();
                    break; 
            }

            label24.Hide();

            textBox1.Text = "";
            tDate.Checked = false;
            tDate2.Checked = false;

            tabControl1.TabPages[0].Text = "請求書データ";

            txtClient.Text = "";
            txtKeisho.Text = "";
            txtZipCode.Text = "";
            txtAddress.Text = "";
            txtTantou.Text = "";
            txtTel.Text = "";
            txtFax.Text = "";
            nDate.Checked = false;

            label8.Text = "";
            label9.Text = "";
            label10.Text = "";
            label11.Text = "";
            label18.Text = "";

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;

            dataGridView1.RowCount = 0;
            dataGridView3.RowCount = 0;

            txtMemo.Text = "";

            label1.Text = "【受注データ】";

            //入金情報
            if (fMode.Mode == 1)
            {
                txtKingaku.Text = "";
                txtZan.Text = "";
                dataGridView4.RowCount = 0;
                button9.Enabled = false;
                button10.Enabled = false;
                nyukinMode = 0;
                checkBox1.Checked = false;
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("画面を取り消します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            
            DispClear();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {

                if (fDataCheck() == true)
                {
                    Control.DataControl Con;
                    OleDbConnection cn;
                    OleDbTransaction tran;
                    OleDbCommand SCom;

                    switch (fMode.Mode)
                    {
                        case 0:

                            //新規登録
                            if (MessageBox.Show("請求データを登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                return;

                            //IDを採番
                            string sqlStr = "";
                            int gID = (int)(1);

                            sqlStr = "select max(ID) as ID from 請求書 ";
                            OleDbDataReader dR;
                            Control.FreeSql fCon = new Control.FreeSql();
                            dR = fCon.free_dsReader(sqlStr);

                            while (dR.Read())
                            {
                                if (dR["ID"] == DBNull.Value)
                                {
                                    gID = (int)(1);
                                }
                                else
                                {
                                    gID = Int32.Parse(dR["ID"].ToString()) + 1;
                                }
                            }

                            dR.Close();
                            fCon.Close();

                            //IDを設定
                            cMaster.ID = gID;

                            //登録処理
                            Con = new Control.DataControl();
                            cn = new OleDbConnection();

                            cn = Con.GetConnection();

                            //トランザクション開始
                            tran = cn.BeginTransaction();

                            SCom = new OleDbCommand();
                            SCom.Connection = cn;
                            SCom.Transaction = tran;

                            try
                            {
                                //請求書データ登録処理
                                sqlStr = "";
                                sqlStr += "insert into 請求書 ";
                                sqlStr += "(ID,得意先ID,請求金額,消費税,値引額,売上金額,税率,入金予定日,発行日,";
                                sqlStr += "入金残,完了区分,振込口座ID1,振込口座ID2,備考,登録年月日,変更年月日) ";
                                sqlStr += "values (";
                                sqlStr += cMaster.ID + ",";
                                sqlStr += cMaster.得意先ID + ",";
                                sqlStr += cMaster.請求金額 + ",";
                                sqlStr += cMaster.消費税 + ",";
                                sqlStr += cMaster.値引額 + ",";
                                sqlStr += cMaster.売上金額 + ",";
                                sqlStr += cMaster.税率 + ",";
                                sqlStr += "'" + cMaster.入金予定日 + "',";
                                sqlStr += "'" + cMaster.発行日 + "',";
                                sqlStr += cMaster.入金残 + ",";
                                sqlStr += cMaster.完了区分 + ",";
                                sqlStr += cMaster.振込口座ID1 + ",";
                                sqlStr += cMaster.振込口座ID2 + ",";
                                sqlStr += "'" + cMaster.備考 + "',";
                                sqlStr += "'" + cMaster.登録年月日 + "',";
                                sqlStr += "'" + cMaster.変更年月日 + "')";

                                SCom.CommandText = sqlStr;

                                SCom.ExecuteNonQuery();

                                //受注データ更新
                                string sID;

                                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                                {
                                    sID = dataGridView3[11, i].Value.ToString();    //受注ID

                                    sqlStr = "";
                                    sqlStr += "update 受注 ";
                                    sqlStr += "set ";
                                    sqlStr += "請求書ID = " + gID.ToString() + ",";
                                    sqlStr += "請求書発行日 = '" + DateTime.Today + "',";
                                    sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
                                    sqlStr += "where (受注.ID = " + sID + ") ";

                                    SCom.CommandText = sqlStr;

                                    SCom.ExecuteNonQuery();
                                }

                                tran.Commit();

                                MessageBox.Show("新規登録されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();

                                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                MessageBox.Show("新規登録に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            }

                            cn.Close();

                            Con.Close();

                            break;

                        case 1: //更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                return;

                            //更新処理準備
                            Con = new Control.DataControl();
                            cn = new OleDbConnection();
                            cn = Con.GetConnection();

                            //トランザクション開始
                            tran = cn.BeginTransaction();

                            SCom = new OleDbCommand();
                            SCom.Connection = cn;
                            SCom.Transaction = tran;

                            try
                            {
                                //請求書データ更新
                                sqlStr = "";
                                sqlStr += "update 請求書 set ";
                                sqlStr += "得意先ID = " + cMaster.得意先ID + ",";
                                sqlStr += "請求金額 = " + cMaster.請求金額 + ",";
                                sqlStr += "消費税 = " + cMaster.消費税 + ",";
                                sqlStr += "値引額 = " + cMaster.値引額 + ",";
                                sqlStr += "売上金額 = " + cMaster.売上金額 + ",";
                                sqlStr += "税率 = " + cMaster.税率 + ",";
                                sqlStr += "入金予定日 = '" + cMaster.入金予定日 + "',";
                                sqlStr += "発行日 = '" + cMaster.発行日 + "',";
                                sqlStr += "入金残 = " + cMaster.入金残 + ",";
                                sqlStr += "完了区分 = " + cMaster.完了区分 + ",";
                                sqlStr += "振込口座ID1 = " + cMaster.振込口座ID1 + ",";
                                sqlStr += "振込口座ID2 = " + cMaster.振込口座ID2 + ",";
                                sqlStr += "備考 = '" + cMaster.備考 + "',";
                                sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
                                sqlStr += "where ID = " + cMaster.ID;

                                SCom.CommandText = sqlStr;

                                // SQLの実行
                                SCom.ExecuteNonQuery();

                                //受注データ更新
                                string sID;

                                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                                {
                                    sID = dataGridView3[11, i].Value.ToString();    //受注ID

                                    sqlStr = "";
                                    sqlStr += "update 受注 ";
                                    sqlStr += "set ";
                                    sqlStr += "請求書ID = " + cMaster.ID.ToString() + ",";
                                    sqlStr += "請求書発行日 = '" + DateTime.Today + "',";
                                    sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
                                    sqlStr += "where (受注.ID = " + sID + ") ";

                                    SCom.CommandText = sqlStr;

                                    SCom.ExecuteNonQuery();
                                }

                                tran.Commit();
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();

                                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                MessageBox.Show("更新に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }

                            cn.Close();
                            Con.Close();
                            break;
                    }

                    //得意先リストをリロード
                    DateTime StartDate;
                    DateTime EndDate;

                    if (tDate.Checked == false)
                    {
                        StartDate = sD;
                    }
                    else
                    {
                        StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
                    }

                    if (tDate2.Checked == false)
                    {
                        EndDate = eD;
                    }
                    else
                    {
                        EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
                    }
    
                    switch (fMode.Mode)
                    {

                        case 0:
                            ListClient.load(listBox1, textBox1.Text.Trim(), StartDate, EndDate);
                            break;

                        case 1:
                            ListClient.loadData(listBox1,getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
                            break;
                    }

                    //画面初期化
                    DispClear();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "登録処理", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }    
        }

        //登録データチェック
        private Boolean fDataCheck()
        {

            try
            {

                //編集モードのとき
                if (fMode.Mode == 1)
                {
                    if (int.Parse(txtZan.Text, System.Globalization.NumberStyles.Any) > 0 && checkBox1.Checked == true)
                    {
                        if (MessageBox.Show("請求残高がありますが入金完了にチェックが入っています。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            return false;
                    }
                }

                if (nDate.Checked == false)
                {
                    throw new Exception("入金予定日を登録してください");
                }

                //クラスにデータセット
                cMaster.請求金額 = int.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                cMaster.消費税 = int.Parse(label9.Text, System.Globalization.NumberStyles.Any);
                cMaster.売上金額  = int.Parse(label8.Text, System.Globalization.NumberStyles.Any);
                cMaster.値引額 = int.Parse(label10.Text, System.Globalization.NumberStyles.Any);
                cMaster.税率 = int.Parse(label18.Text, System.Globalization.NumberStyles.Any);
                cMaster.入金予定日 = DateTime.Parse(nDate.Value.ToShortDateString());

                switch (fMode.Mode)
	            {
                    case 0:
                        cMaster.発行日 = DateTime.Today;
                        cMaster.入金残 = int.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                        cMaster.完了区分 = 0;
                        break;

                    case 1:
                        cMaster.入金残 = int.Parse(txtZan.Text, System.Globalization.NumberStyles.Any);

                        if (checkBox1.Checked == true)
                        {
                            cMaster.完了区分 = 1;
                        }
                        else
                        {
                            cMaster.完了区分 = 0;
                        }
                        break;
	            } 

                cMaster.振込口座ID1 = 0;
                cMaster.振込口座ID2 = 0;
                cMaster.備考 = txtMemo.Text;
                
                if (fMode.Mode == 0) cMaster.登録年月日 = DateTime.Today;
                
                cMaster.変更年月日 = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の請求データを削除します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();
            cn = Con.GetConnection();

            OleDbTransaction tran;

            //トランザクション開始
            tran = cn.BeginTransaction();

            OleDbCommand SCom = new OleDbCommand();

            SCom.Connection = cn;
            SCom.Transaction = tran;

            string sqlSTR;

            try
            {
                //請求書データ削除
                sqlSTR = "";
                sqlSTR += "delete from 請求書 ";
                sqlSTR += "where ID = " + cMaster.ID.ToString();

                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                //受注データ更新・請求書ID初期化
                string sID;

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    sID = dataGridView3[11, i].Value.ToString();    //受注ID

                    sqlSTR = "";
                    sqlSTR += "update 受注 set ";
                    sqlSTR += "請求書ID = 0,";
                    sqlSTR += "請求書発行日 = null,";
                    sqlSTR += "変更年月日 = '" + DateTime.Today + "' ";
                    sqlSTR += "where ID = " + sID.ToString();

                    SCom.CommandText = sqlSTR;

                    SCom.ExecuteNonQuery();
                }

                //入金データ削除
                sqlSTR = "";
                sqlSTR += "delete from 入金 ";
                sqlSTR += "where 請求書ID = " + cMaster.ID.ToString();

                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                tran.Commit();

                MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                tran.Rollback();

                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MessageBox.Show("削除に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            cn.Close();
            Con.Close();

            //得意先リストをリロード
            DateTime StartDate;
            DateTime EndDate;

            if (tDate.Checked == false)
            {
                StartDate = sD;
            }
            else
            {
                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
            }

            if (tDate2.Checked == false)
            {
                EndDate = eD;
            }
            else
            {
                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
            }
            
            switch (fMode.Mode)
            {
                case 0:
                    ListClient.load(listBox1, textBox1.Text.Trim(), StartDate, EndDate);
                    break;

                case 1: 
                    ListClient.loadData(listBox1,getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
                    break;
            } 
            
            DispClear();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (Utility.NumericCheck(txtKingaku.Text) == false)
            {
                MessageBox.Show("入金額は数字で入力してください","入金額エラー",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                txtKingaku.Focus();
                return;
            }

            if (int.Parse(txtKingaku.Text,System.Globalization.NumberStyles.Any) > int.Parse(txtZan.Text,System.Globalization.NumberStyles.Any))
            {
                if (MessageBox.Show("請求残高を超えています。よろしいですか", "過入金", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    txtKingaku.Focus();
                    return;
                }
            }

            if (MessageBox.Show(dateTimePicker1.Value.ToShortDateString() + "の入金情報を登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            
            Control.入金 nCon = new Control.入金();

            switch (nyukinMode)
            {
                case 0: //新規登録

                    //クラスにデータをセットする
                    cNyukin.請求書ID = cMaster.ID;
                    cNyukin.入金年月日 = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                    cNyukin.金額 = int.Parse(txtKingaku.Text, System.Globalization.NumberStyles.Any);
                    cNyukin.備考 = "";
                    cNyukin.登録年月日 = DateTime.Today;
                    cNyukin.変更年月日 = DateTime.Today;

                    //入金データ登録
                    nCon.DataInsert(cNyukin);
                    break;

                case 1:

                    //クラスにデータをセットする
                    cNyukin.入金年月日 = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                    cNyukin.金額 = int.Parse(txtKingaku.Text, System.Globalization.NumberStyles.Any);
                    cNyukin.変更年月日 = DateTime.Today;

                    //入金データ登録
                    nCon.DataUpdate(cNyukin);
                    break;
            }

            nCon.Close();

            //入金履歴表示
            ShowNyukin();

            //残高表示
            ShowSeikyuZan();

            //入金処理モード初期化
            nyukinMode = 0;

            button10.Enabled = false;
            txtKingaku.Text = "";

        }

        //請求残高表示
        private void ShowSeikyuZan()
        {
            int sKin = 0;
            int sZan;

            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                sKin += int.Parse(dataGridView4[2, i].Value.ToString(), System.Globalization.NumberStyles.Any);
            }

            sZan = int.Parse(label11.Text, System.Globalization.NumberStyles.Any) - sKin;
            txtZan.Text = sZan.ToString("#,##0");

        }

        private void txtKingaku_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == textBox1) txtObj = textBox1;
            if (sender == txtClient) txtObj = txtClient;
            if (sender == txtKeisho) txtObj = txtKeisho;
            if (sender == txtZipCode) txtObj = txtZipCode;
            if (sender == txtAddress) txtObj = txtAddress;
            if (sender == txtTantou) txtObj = txtTantou;
            if (sender == txtTel) txtObj = txtTel;
            if (sender == txtFax) txtObj = txtFax;
            if (sender == txtKingaku) txtObj = txtKingaku;
            if (sender == txtMemo) txtObj = txtMemo;

            txtObj.BackColor = Color.LightGray;
            txtObj.SelectAll();
        }

        private void txtKingaku_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == textBox1) txtObj = textBox1;
            if (sender == txtClient) txtObj = txtClient;
            if (sender == txtKeisho) txtObj = txtKeisho;
            if (sender == txtZipCode) txtObj = txtZipCode;
            if (sender == txtAddress) txtObj = txtAddress;
            if (sender == txtTantou) txtObj = txtTantou;
            if (sender == txtTel) txtObj = txtTel;
            if (sender == txtFax) txtObj = txtFax;
            if (sender == txtKingaku) txtObj = txtKingaku;
            if (sender == txtMemo) txtObj = txtMemo;

            txtObj.BackColor = Color.White;
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView4.SelectedRows.Count == 0) return;

            //入金情報選択
            if (MessageBox.Show(dataGridView4[1, dataGridView4.SelectedRows[0].Index].Value.ToString() + "の入金データが選択されました。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ取得
            GetDataNyukin(int.Parse(dataGridView4[0, dataGridView4.SelectedRows[0].Index].Value.ToString()));

            //対象データの表示

            dateTimePicker1.Value = cNyukin.入金年月日;
            txtKingaku.Text = cNyukin.金額.ToString("#,##0");

            //削除ボタン表示
            button10.Enabled = true;

        }

        /// <summary>
        /// 入金データ取得
        /// </summary>
        /// <param name="tempID">ID</param>
        private void GetDataNyukin(int tempID)
        {
            OleDbDataReader dR;
            Control.入金 sNyukin = new Control.入金();
            dR = sNyukin.FillBy("where ID = " + tempID.ToString());

            while (dR.Read())
            {
                cNyukin.ID = int.Parse(dR["ID"].ToString());
                cNyukin.請求書ID = int.Parse(dR["請求書ID"].ToString());
                cNyukin.入金年月日 = DateTime.Parse(dR["入金年月日"].ToString());
                cNyukin.金額 = int.Parse(dR["金額"].ToString(),System.Globalization.NumberStyles.Any);
                cNyukin.備考 = dR["備考"].ToString();
            }

            dR.Close();
            sNyukin.Close();

            nyukinMode = 1;
        }

        //入金処理モード(0:登録,1:編集)
        private int nMode;

        private int nyukinMode
        {
            get { return nMode; }
            set { nMode = value; }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(dateTimePicker1.Value.ToShortDateString() + "の入金情報を削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            Control.入金 nCon = new Control.入金();
            nCon.DataDelete(cNyukin.ID);
            nCon.Close();

            //入金履歴表示
            ShowNyukin();

            //残高表示
            ShowSeikyuZan();

            //入金処理モード初期化
            nyukinMode = 0;

            button10.Enabled = false;
            txtKingaku.Text = "";
        }

        private int getRbtn()
        {
            if (radioButton1.Checked == true)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            DateTime StartDate;
            DateTime EndDate;

            if (tDate.Checked == false)
            {
                StartDate = sD;
            }
            else
            {
                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
            }

            if (tDate2.Checked == false)
            {
                EndDate = eD;
            }
            else
            {
                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
            }
            
            ListClient.loadData(listBox1, getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DateTime StartDate;
            DateTime EndDate;

            if (tDate.Checked == false)
                StartDate = sD;
            else
                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());

            if (tDate2.Checked == false)
                EndDate = eD;
            else
                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
            
            switch (fMode.Mode)
            {
                case 0:
                    ListClient.load(listBox1, textBox1.Text.Trim(), StartDate, EndDate);
                    break;

                case 1:
                    ListClient.loadData(listBox1, getRbtn(), textBox1.Text.Trim(),StartDate,EndDate);
                    break;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
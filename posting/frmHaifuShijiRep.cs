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
    public partial class frmHaifuShijiRep : Form
    {
        const string MESSAGE_CAPTION = "配布指示一覧";

        public frmHaifuShijiRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            viewSetting(dataGridView1);
            DispClear();
        }
        
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        private void viewSetting(DataGridView tempDGV)
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
                tempDGV.Height = 631;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "指示書№");
                tempDGV.Columns.Add("col2", "配布日");
                tempDGV.Columns.Add("col3", "配布員ID");
                tempDGV.Columns.Add("col4", "配布員名");
                tempDGV.Columns.Add("col5", "受注番号");
                tempDGV.Columns.Add("col6", "チラシ名");
                tempDGV.Columns.Add("col7", "コード");
                tempDGV.Columns.Add("col8", "エリア");
                tempDGV.Columns.Add("col9", "配布単価");
                tempDGV.Columns.Add("col10", "予定枚数");
                tempDGV.Columns.Add("col11", "配布枚数");
                tempDGV.Columns.Add("col12", "交通費");

                tempDGV.Columns[0].Width = 80;
                tempDGV.Columns[1].Width = 80;
                tempDGV.Columns[2].Width = 80;
                tempDGV.Columns[3].Width = 140;
                tempDGV.Columns[4].Width = 100;
                tempDGV.Columns[5].Width = 300;
                tempDGV.Columns[6].Width = 70;
                tempDGV.Columns[7].Width = 260;
                tempDGV.Columns[8].Width = 80;
                tempDGV.Columns[9].Width = 80;
                tempDGV.Columns[10].Width = 80;
                tempDGV.Columns[11].Width = 80;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                tempDGV.Columns[1].DefaultCellStyle.Format = "yyyy/MM/dd";
                tempDGV.Columns[6].DefaultCellStyle.Format = "0000";
                tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";
                tempDGV.Columns[11].DefaultCellStyle.Format = "#,##0";

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

        private void showData(DataGridView tempDGV)
        {
            string sqlSTRING = "";

            try
            {
                //データリーダーを取得する
                OleDbDataReader dR;

                sqlSTRING += "SELECT 配布指示.ID AS 配布指示ID, 配布指示.配布日 AS 配布指示配布日,";
                sqlSTRING += "配布指示.配布員ID,配布員.氏名, 配布エリア.受注ID,";
                sqlSTRING += "受注.チラシ名, 配布エリア.町名ID AS 町名コード, 町名.名称 AS 町名名称,";
                sqlSTRING += "配布エリア.配布単価, 配布エリア.予定枚数,"; 
                sqlSTRING += "配布エリア.実配布枚数, 配布指示.交通費 ";
                sqlSTRING += "FROM 配布指示 inner JOIN 配布エリア ";
                sqlSTRING += "ON 配布指示.ID = 配布エリア.配布指示ID inner JOIN 受注 ";
                sqlSTRING += "ON 配布エリア.受注ID = 受注.ID LEFT OUTER JOIN 配布員 ";
                sqlSTRING += "ON 配布指示.配布員ID = 配布員.ID LEFT OUTER JOIN 町名 ";
                sqlSTRING += "ON 配布エリア.町名ID = 町名.ID ";

                sqlSTRING += "where (配布指示.ID >= ? and 配布指示.ID <= ?) and ";
                sqlSTRING += "(配布指示.配布日 >= ? and 配布指示.配布日 <= ?) ";

                // 入力日設定 2016/09/19
                if (sInputDt.Checked)
                {
                    sqlSTRING += "and 配布指示.入力日 = ? ";
                }

                //// 町名設定
                //if (txtsChome.Text.Trim() != string.Empty)
                //{
                //    sqlSTRING += "and 町名.名称 like '*?*' ";
                //}

                sqlSTRING += "ORDER BY 配布指示ID DESC, 配布エリア.受注ID ";

                Control.DataControl Con = new Control.DataControl();
                OleDbConnection Cn = new OleDbConnection();
                Cn = Con.GetConnection();

                OleDbCommand sCom = new OleDbCommand();

                sCom.CommandText = sqlSTRING;

                // 指示№範囲設定 2016/09/19
                sCom.Parameters.AddWithValue("@sNoS", Utility.strToInt(txtsShijiS.Text));

                if (Utility.strToInt(txtsShijiE.Text) == 0)
                {
                    sCom.Parameters.AddWithValue("@sNoE", 2000000000);
                }
                else
                {
                    sCom.Parameters.AddWithValue("@sNoE", Utility.strToInt(txtsShijiE.Text));
                }

                // 配布開始日設定 2016/09/19
                if (sHaifuDtS.Checked)
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtS", sHaifuDtS.Value.ToShortDateString());
                }
                else
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtS", "1900/01/01");
                }

                // 配布終了日設定 2016/09/19
                if (sHaifuDtE.Checked)
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtE", sHaifuDtE.Value.ToShortDateString());
                }
                else
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtE", "2900/01/01");
                }

                // 入力日設定 2016/09/19
                if (sInputDt.Checked)
                {
                    sCom.Parameters.AddWithValue("@sInputDt", sInputDt.Value.ToShortDateString());
                }
                
                sCom.Connection = Cn;
                dR = sCom.ExecuteReader();

                //グリッドビューに表示する
                int iX = 0;

                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    // 丁目検索設定のとき 2016/09/19
                    if (txtsChome.Text.Trim() != string.Empty)
                    {
                        if (!dR["町名名称"].ToString().Contains(txtsChome.Text.Trim()))
                        {
                            continue;
                        }
                    }

                    // グリッドビュー表示
                    tempDGV.Rows.Add();

                    tempDGV[0, iX].Value = int.Parse(dR["配布指示ID"].ToString());
                    tempDGV[1, iX].Value = DateTime.Parse(dR["配布指示配布日"].ToString());
                    tempDGV[2, iX].Value = dR["配布員ID"].ToString();
                    tempDGV[3, iX].Value = dR["氏名"].ToString();
                    tempDGV[4, iX].Value = long.Parse(dR["受注ID"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[5, iX].Value = dR["チラシ名"].ToString();
                    tempDGV[6, iX].Value = int.Parse(dR["町名コード"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[7, iX].Value = dR["町名名称"].ToString();
                    tempDGV[8, iX].Value = double.Parse(dR["配布単価"].ToString(), System.Globalization.NumberStyles.Any).ToString("#,##0.00");
                    tempDGV[9, iX].Value = int.Parse(dR["予定枚数"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[10, iX].Value = int.Parse(dR["実配布枚数"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[11, iX].Value = int.Parse(dR["交通費"].ToString(), System.Globalization.NumberStyles.Any);
                        
                    iX++;
                }

                dR.Close();
                Con.Close();
                    
                //if (tempDGV.RowCount > 25)
                //{
                //    tempDGV.Columns[1].Width = 198;
                //}
                //else
                //{
                //    tempDGV.Columns[1].Width = 215;
                //}

                tempDGV.CurrentCell = null;

                // 表示件数表示 2016/09/19
                if (tempDGV.Rows.Count > 0)
                {
                    this.Text = MESSAGE_CAPTION + " 【" + tempDGV.Rows.Count.ToString("#,##0") + "件】";
                }
                else
                {
                    this.Text = MESSAGE_CAPTION;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
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
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("配布指示一覧を表示します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            Cursor.Current = Cursors.WaitCursor;
            showData(dataGridView1);
            dataGridView1.CurrentCell = null;
            Cursor.Current = Cursors.Default;

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("該当データがありませんでした", MESSAGE_CAPTION);
                btnPrn.Enabled = false;
            }
            else
            {
                btnPrn.Enabled = true;
            }
        }

        private void DispClear()
        {
            // 2016/09/19
            txtsShijiS.Text = string.Empty;
            txtsShijiE.Text = string.Empty;
            sHaifuDtS.Checked = false;
            sHaifuDtE.Checked = false;
            sInputDt.Checked = false;
            txtsChome.Text = string.Empty;
        }

        private void txtsShijiS_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
    }
}
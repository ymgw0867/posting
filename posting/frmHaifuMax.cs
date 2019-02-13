using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;

namespace posting
{
    public partial class frmHaifuMax : Form
    {
        const string MESSAGE_CAPTION = "市区町村別月間最多配布枚数";

        public frmHaifuMax()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            GridviewSet.Setting(dataGridView2);
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();
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
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", (float)11, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 541;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "コード");
                    tempDGV.Columns.Add("col2", "市区町村");
                    tempDGV.Columns.Add("col5", "配布枚数");

                    tempDGV.Columns[0].Width = 70;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 80;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

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

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void ShowData(DataGridView tempDGV,int tempYear,int tempMonth)
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
                    sqlSTRING += "SELECT m_tbl.ID,市区町村_1.都道府県,m_tbl.市区町村,SUM(m_tbl.枚数) AS 枚数 ";
                    sqlSTRING += "from ";
                    sqlSTRING += "(SELECT 市区町村.ID,市区町村.市区町村,配布エリア.町名ID,町名.名称,MAX(配布エリア.予定枚数) AS 枚数 ";
                    sqlSTRING += "FROM 配布エリア INNER JOIN ";
                    sqlSTRING += "配布指示 ON 配布エリア.配布指示ID = 配布指示.ID INNER JOIN ";
                    sqlSTRING += "町名 ON 配布エリア.町名ID = 町名.ID INNER JOIN ";
                    sqlSTRING += "市区町村 ON 町名.市区町村コード = 市区町村.ID ";
                    sqlSTRING += "WHERE (YEAR(配布指示.配布日) = ?) AND (MONTH(配布指示.配布日) = ?) ";
                    sqlSTRING += "GROUP BY 配布エリア.町名ID,町名.名称,市区町村.ID,市区町村.市区町村) ";
                    sqlSTRING += "AS m_tbl LEFT OUTER JOIN ";
                    sqlSTRING += "市区町村 AS 市区町村_1 ON m_tbl.ID = 市区町村_1.ID ";
                    sqlSTRING += "GROUP BY m_tbl.ID,m_tbl.市区町村,市区町村_1.都道府県 ";
                    sqlSTRING += "ORDER BY m_tbl.ID";

                    //配布指示データのデータリーダーを取得する
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@sYear", tempYear);
                    SCom.Parameters.AddWithValue("@sMonth",tempMonth);

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
                            tempDGV[1, iX].Value = dR["都道府県"].ToString() + " " + dR["市区町村"].ToString() + "";
                            tempDGV[2, iX].Value = int.Parse(dR["枚数"].ToString());

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

                    //if (tempDGV.Rows.Count > 29)
                    //{
                    //    tempDGV.Columns[1].Width = 333;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 350;
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
            if (MessageBox.Show("終了します" + Environment.NewLine + "よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, MESSAGE_CAPTION);
        }

        private void txtYear_Validating(object sender, CancelEventArgs e)
        {
            if (Utility.NumericCheck(txtYear.Text) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        private void txtMonth_Validating(object sender, CancelEventArgs e)
        {
            if (Utility.NumericCheck(txtMonth.Text) == false)
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

        private void button2_Click(object sender, EventArgs e)
        {
            GridviewSet.ShowData(dataGridView2,int.Parse(txtYear.Text.ToString()),int.Parse(txtMonth.Text.ToString()));
            dataGridView2.CurrentCell = null; //選択状態を回避する
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

    }
}
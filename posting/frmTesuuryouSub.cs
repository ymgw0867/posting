using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace posting
{
    public partial class frmTesuuryouSub : Form
    {
        const string MESSAGE_CAPTION = "配布員別手数料";

        public frmTesuuryouSub()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            GridviewSet.Setting(dataGridView2);
            GridviewSet.ShowData(dataGridView2);
            dataGridView2.CurrentCell = null; //選択状態を回避する
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
                    tempDGV.Height = 505;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "配布員ID");
                    tempDGV.Columns.Add("col2", "配布員名");
                    tempDGV.Columns.Add("col3", "手数料");

                    tempDGV.Columns[0].Width = 100;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV)
            {

                string sqlSTRING = "";
                int iX;

                try
                {
                    tempDGV.RowCount = 0;
                    
                    //データリーダーを取得する        
                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select * from 市区町村 ";
                    sqlSTRING += "order by ID";
                                        
                    //市区町村データのデータリーダーを取得する
                    Control.FreeSql cArea = new Control.FreeSql();
                    dR = cArea.free_dsReader(sqlSTRING);

                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = int.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = dR["都道府県"].ToString();
                            tempDGV[2, iX].Value = dR["市区町村"].ToString();

                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();
                    cArea.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //市区町村を登録する
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("市区町村を選択してください", "未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("選択中の市区町村を登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            int iX = 0;

            F_市区町村コード = int.Parse(dataGridView2[0, dataGridView2.SelectedRows[iX].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

        }

        //選択された市区町村コード
        private int F_市区町村コード;
        public int 市区町村コード
        {
            set
            {
                this.F_市区町村コード = value;
            }
            get
            {
                return this.F_市区町村コード;
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("終了します。現在、選択状態のデータは登録されません。" + Environment.NewLine + "よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            button4.DialogResult = DialogResult.No;
            this.Close();
        }

    }
}
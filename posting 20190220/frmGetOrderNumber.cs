using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmGetOrderNumber : Form
    {
        public frmGetOrderNumber()
        {
            InitializeComponent();

            // データ読み込み
            adp.Fill(dts.受注番号採番);
            cAdp.Fill(dts.得意先);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注番号採番TableAdapter adp = new darwinDataSetTableAdapters.受注番号採番TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        #region グリッドビューカラム定義
        string colID = "col1";
        string colNum = "col2";
        string colDt = "col6";
        string colMemo = "col3";
        string colAddDt = "col4";
        string colUpDt = "col5";
        string colUserID = "col7";
        string colClient = "col8";
        #endregion


        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビューの定義
            gridSetting(dataGridView1);

            // データグリッドビューデータ表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();
        }
        
        /// -------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// -------------------------------------------------------------------
        private void gridSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 361;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                tempDGV.Columns.Add(colID, "ｺｰﾄﾞ");
                tempDGV.Columns.Add(colNum, "受注番号");
                tempDGV.Columns.Add(colDt, "入庫日");
                tempDGV.Columns.Add(colClient, "クライアント");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colAddDt, "登録年月日");
                tempDGV.Columns.Add(colUpDt, "更新年月日");

                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[colNum].Width = 120;
                tempDGV.Columns[colDt].Width = 120;
                tempDGV.Columns[colAddDt].Width = 140;
                tempDGV.Columns[colMemo].Width = 200;
                tempDGV.Columns[colClient].Width = 200;

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

        /// ------------------------------------------------------------------
        /// <summary>
        ///     グリッドデータを表示する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------
        private void gridShow(DataGridView g)
        {
            g.Rows.Clear();
            int iX = 0;

            // グリッドに表示
            foreach (var t in dts.受注番号採番.Where(a => a.確定書入力 == 0).OrderByDescending(a => a.受注番号))
            {
                g.Rows.Add();
                g[colID, iX].Value = t.ID.ToString();
                g[colNum, iX].Value = t.受注番号;
                g[colDt, iX].Value = t.入庫日.ToShortDateString();

                if (t.得意先Row == null)
                {
                    g[colClient, iX].Value = string.Empty;
                }
                else
                {
                    g[colClient, iX].Value = t.得意先Row.略称;
                }

                g[colMemo, iX].Value = t.備考;
                g[colAddDt, iX].Value = t.登録年月日;
                iX++;
            }

            g.CurrentCell = null;
        }
        

        /// -------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// -------------------------------------------------------------
        private void dispClear()
        {
            button4.Enabled = false;
            _orderNum = string.Empty;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmLoginType_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 受注番号選択
            string num = getOrderNum();
            if (MessageBox.Show(num + "が選択されました。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            else
            {
                _orderNum = num;
                this.Close();
            }
        }
        
        private string getOrderNum()
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                return dataGridView1[colNum, dataGridView1.SelectedRows[0].Index].Value.ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // 受注番号選択
            string num = getOrderNum();
            if (MessageBox.Show(num + "が選択されました。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            else
            {
                _orderNum = num;
                this.Close();
            }
        }
        
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            button4.Enabled = true;
        }

        // 選択した受注番号
        public string _orderNum { get; set; }
    }
}

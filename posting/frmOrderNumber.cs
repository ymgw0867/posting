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
    public partial class frmOrderNumber : Form
    {
        public frmOrderNumber()
        {
            InitializeComponent();

            // データ読み込み
            adp.Fill(dts.受注番号採番);
            uAdp.Fill(dts.ログインユーザー);
            cAdp.Fill(dts.得意先);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注番号採番TableAdapter adp = new darwinDataSetTableAdapters.受注番号採番TableAdapter();
        darwinDataSetTableAdapters.ログインユーザーTableAdapter uAdp = new darwinDataSetTableAdapters.ログインユーザーTableAdapter();
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

        // 得意先コード
        int sClientCode = 0;

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビューの定義
            gridSetting(dataGridView1);
            clientGridSetting(dataGridView2);

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
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

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
                tempDGV.Height = 221;

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
                tempDGV.Columns.Add(colUserID, "ユーザーID");

                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[colNum].Width = 120;
                tempDGV.Columns[colDt].Width = 120;
                tempDGV.Columns[colAddDt].Width = 140;
                tempDGV.Columns[colUpDt].Width = 140;
                tempDGV.Columns[colUserID].Width = 140;
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

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.BorderStyle = BorderStyle.Fixed3D;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
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
            foreach (var t in dts.受注番号採番.OrderByDescending(a => a.受注番号))
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
                g[colUpDt, iX].Value = t.更新年月日;
                g[colUserID, iX].Value = t.ログインユーザーRowByログインユーザー_受注番号採番1.ログインユーザー;

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
            fMode.Mode = 0;
            fMode.ID = 0;
            sClientCode = 0;

            lblOrderNum.Text = string.Empty;
            txtMemo.Text = string.Empty;
            lblClientName.Text = string.Empty;

            txtsClient.Text = string.Empty;
            dataGridView2.Rows.Clear();

            button1.Enabled = true;
            button2.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 確認メッセージ
            if (fMode.Mode == 0)
            {
                if (MessageBox.Show("データを新規登録します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }
            else if (fMode.Mode == 1)
            {
                if (MessageBox.Show("データを更新します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            // エラーチェック
            if (!errCheck(fMode.Mode))
            {
                return;
            }

            // 登録・更新処理
            dataUpdate(fMode.Mode, fMode.ID);

            // グリッド表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     登録時エラーチェック </summary>
        /// <param name="sMode">
        ///     処理モード</param>
        /// <returns>
        ///     エラーなし:true、エラーあり:false</returns>
        /// -------------------------------------------------------------------------
        private bool errCheck(int sMode)
        {
            // 新規登録のとき
            if (sMode == 0)
            {
                // 受注番号チェック
                if (lblOrderNum.Text == string.Empty)
                {
                    MessageBox.Show("受注番号を取得してください", "受注番号エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    lblOrderNum.Focus();
                    return false;
                }
                
                // クライアントチェック
                if (sClientCode == global.FLGOFF)
                {
                    if (MessageBox.Show("クライアントが選択されていません。このまま登録してよろしいですか", "クライアント未入力", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No)
                    {
                        txtsClient.Focus();
                        return false;
                    }
                }
            }

            return true;
        }
        
        /// -------------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダ、タグデータ登録 </summary>
        /// <param name="sMode">
        ///     処理モード</param>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// -------------------------------------------------------------------------
        private void dataUpdate(int sMode, int sID)
        {
            // 新規登録
            if (sMode == 0)
            {
                darwinDataSet.受注番号採番Row r = dts.受注番号採番.New受注番号採番Row();
                r.受注番号 =  Int64.Parse(lblOrderNum.Text);
                r.入庫日 = DateTime.Parse(dtNyuko.Value.ToShortDateString());
                r.得意先ID = sClientCode;
                r.確定書入力 = 0;
                r.確定書入力日付 = DateTime.Parse("1900/01/01");
                r.確定書入力ユーザーID = 0;
                r.備考 = txtMemo.Text;
                r.登録年月日 = DateTime.Now;
                r.更新年月日 = DateTime.Now;
                r.ユーザーID = global.loginUserID;

                dts.受注番号採番.Add受注番号採番Row(r);
            }
            
            // データベース更新
            adp.Update(dts.受注番号採番);

            // データ読み込み
            adp.Fill(dts.受注番号採番);
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
            // 画面初期化
            dispClear();
        }
        
        /// ----------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダ、タグデータ表示 </summary>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// ----------------------------------------------------------------------
        private void getData(int sID)
        {
            // ログインタイプヘッダ
            darwinDataSet.受注番号採番Row r = dts.受注番号採番.Single(a => a.ID == sID);

            if (r.確定書入力 == 1)
            {
                MessageBox.Show("受注確定書と紐付られていますので編集は出来ません。","編集不可",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                return;
            }

            lblOrderNum.Text = r.受注番号.ToString();
            dtNyuko.Value = r.入庫日;

            if (r.得意先Row != null)
            {
                lblClientName.Text = r.得意先Row.略称;
            }
            else
            {
                lblClientName.Text = string.Empty;
            }
            
            txtMemo.Text = r.備考;
            
            // 処理モード
            fMode.Mode = 1;
            fMode.ID = r.ID;
            
            // 削除、取消ボタンの使用を可能とします
            button1.Enabled = false;
            button2.Enabled = true;
            button4.Enabled = true;

            // 受注番号取得ボタンは使用不可とします
            button5.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 確認メッセージ
            if (MessageBox.Show("表示中の受注番号を削除します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            { 
                return; 
            }

            // データ削除
            delData(fMode.ID);

            // グリッド表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();
        }
        
        /// ----------------------------------------------------------------------
        /// <summary>
        ///     データ削除 </summary>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// ----------------------------------------------------------------------
        private void delData(int sID)
        {
            // ヘッダデータ削除
            darwinDataSet.受注番号採番Row r = dts.受注番号採番.Single(a => a.ID == sID);
            r.Delete();
            
            // データベース更新
            adp.Update(dts.受注番号採番);

            // データ読み込み
            adp.Fill(dts.受注番号採番);
        }

        private void txtID_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;

            txtObj.BackColor = Color.LightSteelBlue;
            txtObj.SelectAll();
        }

        private void txtID_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;

            txtObj.BackColor = Color.White;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // IDを取得
            int sID = int.Parse(dataGridView1[colID, dataGridView1.SelectedRows[0].Index].Value.ToString());
            
            // データ表示
            getData(sID);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 受注番号を採番する
            //lblOrderNum.Text = getOrderNum(dtNyuko.Value).ToString();
            lblOrderNum.Text = Utility.getOrderNum(dtNyuko.Value).ToString();

            // 取消ボタンアクティブ
            button4.Enabled = true;
        }

        private void txtsClient_TextChanged(object sender, EventArgs e)
        {
            gridClientShow(dataGridView2, txtsClient.Text);
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     得意先名検索表示 </summary>
        /// <param name="d">
        ///     DataGridViewオブジェクト </param>
        /// <param name="sName">
        ///     得意先略称検索文字列</param>
        ///------------------------------------------------------------------
        private void gridClientShow(DataGridView d, string sName)
        {
            d.RowCount = 0;

            var s = dts.得意先.Where(a => a.略称.Contains(sName));

            foreach (var t in dts.得意先.Where(a => a.略称.Contains(sName)))
            {
                d.Rows.Add();
                d[colID, d.RowCount - 1].Value = t.ID;
                d[colClient, d.RowCount - 1].Value = t.略称; 
            }

            if (d.RowCount > 0)
            {
                d.CurrentCell = null;
            }
        }


        /// -------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// -------------------------------------------------------------------
        private void clientGridSetting(DataGridView tempDGV)
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
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 128;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列指定
                tempDGV.Columns.Add(colID, "No.");
                tempDGV.Columns.Add(colClient, "略称");

                // 各列幅指定
                tempDGV.Columns[colID].Width = 60;
                tempDGV.Columns[colClient].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列固定
                //tempDGV.Columns[colClient].Frozen = true;

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
                tempDGV.BorderStyle = BorderStyle.Fixed3D;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // IDを取得
            sClientCode = int.Parse(dataGridView2[colID, dataGridView2.SelectedRows[0].Index].Value.ToString());

            // クライアント名取得
            lblClientName.Text = dataGridView2[colClient, dataGridView2.SelectedRows[0].Index].Value.ToString();            
        }

    }
}

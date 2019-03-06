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
    public partial class frmLoginUser : Form
    {
        public frmLoginUser()
        {
            InitializeComponent();

            // データ読み込み
            hAdp.Fill(dts.ログインタイプヘッダ);
            tAdp.Fill(dts.ログインタイプタグ);
            uAdp.Fill(dts.ログインユーザー);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter hAdp = new darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter();
        darwinDataSetTableAdapters.ログインタイプタグTableAdapter tAdp = new darwinDataSetTableAdapters.ログインタイプタグTableAdapter();
        darwinDataSetTableAdapters.ログインユーザーTableAdapter uAdp = new darwinDataSetTableAdapters.ログインユーザーTableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        #region グリッドビューカラム定義
        string colID = "col1";
        string colName = "col2";
        string colType = "col6";
        string colMemo = "col3";
        string colAddDt = "col4";
        string colUpDt = "col5";
        #endregion

        // 管理者アカウント
        const string ADMINUSER = "administrator";

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // ログインタイプヘッダコンボボックスアイテムロード
            Utility.comboLogintype.itemLoad(cmbType);

            // データグリッドビューの定義
            gridSetting(dataGridView1);

            // データグリッドビューデータ表示
            gridShow(dataGridView1);

            // 受注データ保守コンボ
            cmbOrderSet();

            // 画面初期化
            dispClear();
        }
        
        private void cmbOrderSet()
        {
            comboBox1.Items.Add("自分が登録した受注データのみ対象とする");
            comboBox1.Items.Add("全ての受注データを対象とする");
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
                tempDGV.Height = 221;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                tempDGV.Columns.Add(colID, "ｺｰﾄﾞ");
                tempDGV.Columns.Add(colName, "ユーザーアカウント");
                tempDGV.Columns.Add(colType, "ログインタイプ");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colAddDt, "登録年月日");
                tempDGV.Columns.Add(colUpDt, "更新年月日");

                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[colName].Width = 200;
                tempDGV.Columns[colType].Width = 200;
                tempDGV.Columns[colAddDt].Width = 140;
                tempDGV.Columns[colUpDt].Width = 140;

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
            foreach (var t in dts.ログインユーザー)
            {
                g.Rows.Add();
                g[colID, iX].Value = t.ID.ToString();
                g[colName, iX].Value = t.ログインユーザー;
                
                if (t.ログインタイプヘッダRow != null)
                {
                    g[colType, iX].Value = t.ログインタイプヘッダRow.名称;
                }
                else
                {
                    g[colType, iX].Value = string.Empty;
                }

                g[colMemo, iX].Value = t.備考;
                g[colAddDt, iX].Value = t.登録年月日;
                g[colUpDt, iX].Value = t.更新年月日;

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
            txtName.Text = string.Empty;
            txtName.Enabled = true;
            txtPassword.Text = string.Empty;
            txtPassword.Enabled = true;
            txtPassword2.Text = string.Empty;
            txtPassword2.Enabled = true;
            txtMemo.Text = string.Empty;
            cmbType.SelectedIndex = -1;
            //checkBox1.Checked = false;
            checkBox1.Visible = false;
            comboBox1.SelectedIndex = -1;

            button2.Enabled = false;
            button4.Enabled = false;
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
                // IDチェック
                if (txtName.Text == string.Empty)
                {
                    MessageBox.Show("ユーザーアカウントが未入力です", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtName.Focus();
                    return false;
                }

                if (txtName.Text.Length < 6 || txtName.Text.Length > 14)
                {
                    MessageBox.Show("ユーザーアカウントは6～14文字で登録して下さい", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtName.Focus();
                    return false;
                }
                
                if (dts.ログインユーザー.Any(a => a.ログインユーザー.Equals(txtName.Text)))
                {
                    MessageBox.Show("既に登録済みのユーザーアカウントです", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtName.Focus();
                    return false;
                }
            }

            // パスワード
            if (txtPassword.Enabled)
            {
                if (txtPassword.Text.Trim() == string.Empty)
                {
                    MessageBox.Show("パスワードを入力して下さい", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPassword.Focus();
                    return false;
                }

                if (txtPassword.Text.Length < 6 || txtPassword.Text.Length > 14)
                {
                    MessageBox.Show("パスワードは6～14文字で登録して下さい", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPassword.Focus();
                    return false;
                }
                
                if (!txtPassword.Text.Equals(txtPassword2.Text))
                {
                    MessageBox.Show("パスワードと再入力パスワードが一致しません", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPassword.Focus();
                    return false;
                }
            }

            if (cmbType.SelectedIndex == -1)
            {
                MessageBox.Show("ログインタイプを選択してください", "ログインタイプ未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbType.Focus();
                return false;
            }

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("受注データ保守を選択してください", "受注データ保守未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                comboBox1.Focus();
                return false;
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
                // ヘッダ
                darwinDataSet.ログインユーザーRow r = dts.ログインユーザー.NewログインユーザーRow();
                r.ログインユーザー = txtName.Text;
                r.パスワード = txtPassword.Text;
                r.ログインタイプ = int.Parse(cmbType.SelectedValue.ToString());
                r.備考 = txtMemo.Text;
                r.受注案件保守 = comboBox1.SelectedIndex;
                r.登録ユーザー = global.loginUserID;
                r.登録年月日 = DateTime.Now;
                r.更新年月日 = DateTime.Now;
                dts.ログインユーザー.AddログインユーザーRow(r);                
            }
            else if (sMode == 1)    // 更新
            {
                // ヘッダ
                darwinDataSet.ログインユーザーRow r = dts.ログインユーザー.Single(a => a.ID == sID);

                if (txtPassword.Visible)
                {
                    r.パスワード = txtPassword.Text;
                }

                r.ログインタイプ = int.Parse(cmbType.SelectedValue.ToString());
                r.備考 = txtMemo.Text;
                r.受注案件保守 = comboBox1.SelectedIndex;
                r.登録ユーザー = global.loginUserID;
                r.更新年月日 = DateTime.Now;
            }
            
            // データベース更新
            uAdp.Update(dts.ログインユーザー);

            // データ読み込み
            uAdp.Fill(dts.ログインユーザー);
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプタグデータ削除 </summary>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// ----------------------------------------------------------------------
        private void delTagRow(int sHid)
        {
            // ログインタイプタグ
            foreach (darwinDataSet.ログインタイプタグRow  t in dts.ログインタイプタグ.Where(a => a.ヘッダID == sHid))
            {
                t.Delete();
            }
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
            darwinDataSet.ログインユーザーRow r = dts.ログインユーザー.Single(a => a.ID == sID);

            txtName.Text = r.ログインユーザー;
            txtPassword.Text = r.パスワード;
            cmbType.SelectedValue = r.ログインタイプ;
            comboBox1.SelectedIndex = r.受注案件保守;
            txtMemo.Text = r.備考;
            
            // 処理モード
            fMode.Mode = 1;
            fMode.ID = r.ID;

            // ユーザーアカウント、パスワードの編集は不可とします
            txtName.Enabled = false;

            // パスワード編集チェック
            checkBox1.Visible = true;
            txtPassword.Enabled = false;
            txtPassword2.Enabled = false;

            // 削除、取消ボタンの使用を可能とします
            button2.Enabled = true;
            button4.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 確認メッセージ
            if (MessageBox.Show("表示中のユーザーアカウントを削除します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
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
            darwinDataSet.ログインユーザーRow hr = dts.ログインユーザー.Single(a => a.ID == sID);
            hr.Delete();
            
            // データベース更新
            uAdp.Update(dts.ログインユーザー);

            // データ読み込み
            uAdp.Fill(dts.ログインユーザー);
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
            // administratorは編集不可
            if (dataGridView1[colName, dataGridView1.SelectedRows[0].Index].Value.ToString() == ADMINUSER)
            {
                MessageBox.Show("管理者アカウントの変更、削除は出来ません。", "管理者アカウント", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // IDを取得
            int sID = int.Parse(dataGridView1[colID, dataGridView1.SelectedRows[0].Index].Value.ToString());
            
            // データ表示
            getData(sID);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                txtPassword.Enabled = true;
                txtPassword2.Enabled = true;
            }
            else
            {
                txtPassword.Enabled = false;
                txtPassword2.Enabled = false;
            }
        }
    }
}

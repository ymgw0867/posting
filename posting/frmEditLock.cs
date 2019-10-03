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
    public partial class frmEditLock : Form
    {
        public frmEditLock()
        {
            InitializeComponent();

            // データ読み込み
            hAdp.Fill(dts.ログインタイプヘッダ);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter hAdp = new darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter();
      
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

            checkedListBox1.CheckOnClick = true;
            checkedListBox2.CheckOnClick = true;
            checkedListBox3.CheckOnClick = true;

            // ログインタイプヘッダコンボボックスアイテムロード
            Utility.comboLogintype.itemLoad(checkedListBox1);
            Utility.comboLogintype.itemLoad(checkedListBox2);
            Utility.comboLogintype.itemLoad(checkedListBox3);

            checkedListBox1.SelectedIndex = -1;

            //// データグリッドビューの定義
            //gridSetting(dataGridView1);

            //// データグリッドビューデータ表示
            //gridShow(dataGridView1);

            // 受注データ保守コンボ
            //cmbOrderSet();

            // 画面初期化
            dispClear();
        }
        
        
        /// -------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// -------------------------------------------------------------
        private void dispClear()
        {
            fMode.Mode = 0;
            fMode.ID = 0;
            //cmbType.SelectedIndex = -1;
            ////checkBox1.Checked = false;
            //checkBox1.Visible = false;
            //comboBox1.SelectedIndex = -1;

            //button2.Enabled = false;
            //button4.Enabled = false;
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

            //// グリッド表示
            //gridShow(dataGridView1);

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
            //if (sMode == 0)
            //{
            //    // IDチェック
            //    if (txtName.Text == string.Empty)
            //    {
            //        MessageBox.Show("ユーザーアカウントが未入力です", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        txtName.Focus();
            //        return false;
            //    }

            //    if (txtName.Text.Length < 6 || txtName.Text.Length > 14)
            //    {
            //        MessageBox.Show("ユーザーアカウントは6～14文字で登録して下さい", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        txtName.Focus();
            //        return false;
            //    }
                
            //    if (dts.ログインユーザー.Any(a => a.ログインユーザー.Equals(txtName.Text)))
            //    {
            //        MessageBox.Show("既に登録済みのユーザーアカウントです", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        txtName.Focus();
            //        return false;
            //    }
            //}

            //// パスワード
            //if (txtPassword.Enabled)
            //{
            //    if (txtPassword.Text.Trim() == string.Empty)
            //    {
            //        MessageBox.Show("パスワードを入力して下さい", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        txtPassword.Focus();
            //        return false;
            //    }

            //    if (txtPassword.Text.Length < 6 || txtPassword.Text.Length > 14)
            //    {
            //        MessageBox.Show("パスワードは6～14文字で登録して下さい", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        txtPassword.Focus();
            //        return false;
            //    }
                
            //    if (!txtPassword.Text.Equals(txtPassword2.Text))
            //    {
            //        MessageBox.Show("パスワードと再入力パスワードが一致しません", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        txtPassword.Focus();
            //        return false;
            //    }
            //}

            //if (cmbType.SelectedIndex == -1)
            //{
            //    MessageBox.Show("ログインタイプを選択してください", "ログインタイプ未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    cmbType.Focus();
            //    return false;
            //}

            //if (comboBox1.SelectedIndex == -1)
            //{
            //    MessageBox.Show("受注データ保守を選択してください", "受注データ保守未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    comboBox1.Focus();
            //    return false;
            //}

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
            //// 新規登録
            //if (sMode == 0)
            //{
            //    // ヘッダ
            //    darwinDataSet.ログインユーザーRow r = dts.ログインユーザー.NewログインユーザーRow();
            //    r.ログインユーザー = txtName.Text;
            //    r.パスワード = txtPassword.Text;
            //    r.ログインタイプ = int.Parse(cmbType.SelectedValue.ToString());
            //    r.備考 = txtMemo.Text;
            //    r.受注案件保守 = comboBox1.SelectedIndex;
            //    r.登録ユーザー = global.loginUserID;
            //    r.登録年月日 = DateTime.Now;
            //    r.更新年月日 = DateTime.Now;
            //    dts.ログインユーザー.AddログインユーザーRow(r);                
            //}
            //else if (sMode == 1)    // 更新
            //{
            //    // ヘッダ
            //    darwinDataSet.ログインユーザーRow r = dts.ログインユーザー.Single(a => a.ID == sID);

            //    if (txtPassword.Visible)
            //    {
            //        r.パスワード = txtPassword.Text;
            //    }

            //    r.ログインタイプ = int.Parse(cmbType.SelectedValue.ToString());
            //    r.備考 = txtMemo.Text;
            //    r.受注案件保守 = comboBox1.SelectedIndex;
            //    r.登録ユーザー = global.loginUserID;
            //    r.更新年月日 = DateTime.Now;
            //}
            
            //// データベース更新
            //uAdp.Update(dts.ログインユーザー);

            //// データ読み込み
            //uAdp.Fill(dts.ログインユーザー);
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
        
    }
}

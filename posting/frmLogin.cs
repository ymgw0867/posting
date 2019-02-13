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
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();

            // ログインステータス
            global.loginStatus = false;

            // ログインユーザーデータ読み込み
            adp.Fill(dts.ログインユーザー);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.ログインユーザーTableAdapter adp = new darwinDataSetTableAdapters.ログインユーザーTableAdapter();

        // ログイントライ回数
        int lTry = 0;

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmLogin_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 認証結果検証
            if (!SerachloginUser())
            {
                // 3回ログインエラーでシステムは終了します
                if (lTry > 2)
                {
                    MessageBox.Show("3回続けてログインに失敗しました。システムを終了します。","中止",MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Environment.Exit(0);
                }
            }
            else
            {
                global.loginStatus = true;
                this.Close();
            }
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     ユーザー認証 </summary>
        /// <returns>
        ///     認証成功：true、認証失敗：false</returns>
        /// ---------------------------------------------------------
        private bool SerachloginUser()
        {
            // ログインユーザー
            if (txtName.Text == string.Empty)
            {
                MessageBox.Show("ログインユーザー名を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                lTry++;
                txtName.Focus();
                return false;
            }

            // パスワード
            if (txtPassword.Text == string.Empty)
            {
                MessageBox.Show("パスワードを入力してください", "エラー", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                lTry++;
                txtPassword.Focus();
                return false;
            }

            // ユーザー認証
            if (!dts.ログインユーザー.Any(a => a.ログインユーザー == txtName.Text && a.パスワード == txtPassword.Text))
            {
                lTry++;
                MessageBox.Show("ユーザー認証に失敗しました。" + Environment.NewLine + "ログインユーザー名とパスワードを確認してください。", "認証エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtName.Focus();
                return false;
            }
            else
            {
                var s = dts.ログインユーザー.Single(a => a.ログインユーザー == txtName.Text && a.パスワード == txtPassword.Text);
                global.loginUserID = s.ID;
                global.loginType = s.ログインタイプ;
                global.loginUser = s.ログインユーザー;
                global.loginOrderMntType = s.受注案件保守;
            }

            // 認証成功
            return true;
        }

        private void txtName_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;
            txtObj.SelectAll();
            //txtObj.BackColor = SystemColors.ControlLight;
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;
            //txtObj.BackColor = Color.White;
        }
    }
}

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
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                Utility.comboLogintype ss = (Utility.comboLogintype)checkedListBox1.Items[i];
                if (ss.Lock == global.FLGON || ss.ID == global.adminID)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
                else
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }

            // ログインタイプヘッダコンボボックスアイテムロード
            Utility.comboLogintype.itemLoad(checkedListBox2);
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                Utility.comboLogintype ss = (Utility.comboLogintype)checkedListBox2.Items[i];
                if (ss.seigen == global.FLGON)
                {
                    checkedListBox2.SetItemChecked(i, true);
                }
                else
                {
                    checkedListBox2.SetItemChecked(i, false);
                }
            }

            // ログインタイプヘッダコンボボックスアイテムロード
            Utility.comboLogintype.itemLoad(checkedListBox3);
            for (int i = 0; i < checkedListBox3.Items.Count; i++)
            {
                Utility.comboLogintype ss = (Utility.comboLogintype)checkedListBox3.Items[i];
                if (ss.Jyuryo == global.FLGON || ss.ID == global.adminID)
                {
                    checkedListBox3.SetItemCheckState(i, CheckState.Checked);
                }
                else
                {
                    checkedListBox3.SetItemCheckState(i, CheckState.Unchecked);
                }
            }

            checkedListBox1.SelectedIndex = -1;
            checkedListBox2.SelectedIndex = -1;
            checkedListBox3.SelectedIndex = -1;
            
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
            if (MessageBox.Show("データを登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            
            // エラーチェック
            if (!errCheck(fMode.Mode))
            {
                return;
            }

            // 登録・更新処理
            dataUpdate(fMode.Mode, fMode.ID);

            // 閉じる
            Close();
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
            bool val = true;

            foreach (int item in checkedListBox1.CheckedIndices)
            {
                Utility.comboLogintype cc = (Utility.comboLogintype)checkedListBox1.Items[item];
                //c/*c.ID;*/

                foreach (int t in checkedListBox2.CheckedIndices)
                {
                    Utility.comboLogintype ss = (Utility.comboLogintype)checkedListBox2.Items[t];
                    if (ss.ID == cc.ID)
                    {
                        string msg = "ロック権限を有する「" + cc.Name + "」を編集制限を受ける設定にはできません";  
                        MessageBox.Show(msg, "チェック確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        val = false;
                        break;
                    }
                }

                if (!val)
                {
                    break;
                }
            }


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

            return val;
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
            Cursor = Cursors.WaitCursor;

            try
            {
                // 受注個別制限ロック権限更新
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    Utility.comboLogintype ls = (Utility.comboLogintype)checkedListBox1.Items[i];
                    int pID = Utility.strToInt(ls.ID.ToString());
                    int pLock;

                    if (checkedListBox1.GetItemChecked(i))
                    {
                        // 権限有り
                        pLock = global.FLGON;
                    }
                    else
                    {
                        // 権限なし
                        pLock = global.FLGOFF;
                    }

                    // ログインタイプヘッダデータ更新
                    hAdp.UpdateLock(DateTime.Now, pLock, pID);
                }

                // 受注個別制限更新
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    Utility.comboLogintype ls = (Utility.comboLogintype)checkedListBox2.Items[i];
                    int pID = Utility.strToInt(ls.ID.ToString());
                    int pLock;

                    if (checkedListBox2.GetItemChecked(i))
                    {
                        // 制限有り
                        pLock = global.FLGON;
                    }
                    else
                    {
                        // 制限なし
                        pLock = global.FLGOFF;
                    }

                    // ログインタイプヘッダデータ更新
                    hAdp.UpdateSeigen(DateTime.Now, pLock, pID);
                }

                // 注文書受領済み権限更新
                for (int i = 0; i < checkedListBox3.Items.Count; i++)
                {
                    Utility.comboLogintype ls = (Utility.comboLogintype)checkedListBox3.Items[i];
                    int pID = Utility.strToInt(ls.ID.ToString());
                    int pLock;

                    if (checkedListBox3.GetItemChecked(i))
                    {
                        // 権限有り
                        pLock = global.FLGON;
                    }
                    else
                    {
                        // 権限なし
                        pLock = global.FLGOFF;
                    }

                    // ログインタイプヘッダデータ更新
                    hAdp.UpdateJyuryo(DateTime.Now, pLock, pID);
                }

                MessageBox.Show("終了しました", "編集制限設定", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
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
        
    }
}

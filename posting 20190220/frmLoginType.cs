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
    public partial class frmLoginType : Form
    {
        public frmLoginType()
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

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // ログインタイプヘッダ・リストビューアイテム定義
            listView2Initial();

            // ログインタイプタグ・リストビューアイテム定義 
            listView1Initial();

            // 画面初期化
            dispClear();
        }

        /// -----------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダ・リストビューアイテム定義 </summary>
        /// -----------------------------------------------------------------------
        private void listView2Initial()
        {
            listView2.FullRowSelect = true;
            listView2.Columns.Add("Tag", 60);
            listView2.Columns.Add("ログインタイプ", 140);

            // ログインタイプヘッダ・リストビューアイテム表示
            hdView();
        }

        /// -----------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプタグ・リストビューアイテム定義 </summary>
        /// -----------------------------------------------------------------------
        private void listView1Initial()
        {
            // ログインタイプタグ・リストビュー定義
            listView1.CheckBoxes = true;
            listView1.Columns.Add("Tag", 60);
            listView1.Columns.Add("メニュータイトル", 200);
            listView1.Columns.Add("カテゴリ", 200);

            // メニュータイトルクラス 2015/07/07
            string[] m;
            clsMenu _cm = new clsMenu();

            // メニュータイトルCSVの読込 2015/07/07
            _cm.loadMenu();

            // メニュータイトルリストビューのデータ追加
            foreach (var t in _cm.menuCsv)
            {
                m = t.Split(',');

                // 既定のCSV構成でないときネグる
                if (m.Length < 3)
                {
                    continue;
                }

                // リストビューアイテム作成
                ListViewItem item = new ListViewItem(m[0]);

                // サブアイテム追加
                item.SubItems.Add(m[1]);
                item.SubItems.Add(m[2]);

                // アイテムをリストビューに追加
                this.listView1.Items.Add(item);
            }
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダ・リストビューアイテム表示 </summary>
        /// --------------------------------------------------------------------
        private void hdView()
        {
            // アイテムクリア
            listView2.Items.Clear();

            // アイテム表示
            foreach (var t in dts.ログインタイプヘッダ)
            {
                // リストビューアイテム作成
                ListViewItem item = new ListViewItem(t.Id.ToString());

                // サブアイテム追加
                item.SubItems.Add(t.名称);

                // アイテムをリストビューに追加
                this.listView2.Items.Add(item);
            }
        }

        /// -------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// -------------------------------------------------------------
        private void dispClear()
        {
            fMode.Mode = 0;
            fMode.ID = 0;
            txtID.Text = string.Empty;
            txtID.Enabled = true;
            txtName.Text = string.Empty;
            txtMemo.Text = string.Empty;

            // ログインタイプタグ
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                listView1.Items[i].Checked = false;
            }

            button2.Enabled = false;
            button4.Enabled = false;
            linkLabel2.Enabled = false;
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

            // 画面初期化
            dispClear();

            // ログオンタイプヘッダ再表示
            hdView();
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
                if (Utility.strToInt(txtID.Text) == 0)
                {
                    MessageBox.Show("コードが未入力です", "コードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtID.Focus();
                    return false;
                }
                
                if (dts.ログインタイプヘッダ.Any(a => a.Id == (Utility.strToInt(txtID.Text))))
                {
                    MessageBox.Show("既に登録済みのコードです", "コードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtID.Focus();
                    return false;
                }

                // 名称
                if (txtName.Text.Trim() == string.Empty)
                {
                    MessageBox.Show("名称を登録して下さい", "名称エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtName.Focus();
                    return false;
                }
            }

            // メニュータイトルチェック
            bool itemsCheck = false;

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].Checked)
                {
                    itemsCheck = true;
                    break;
                }
            }

            if (!itemsCheck)
            {
                MessageBox.Show("メニュータイトルは一つ以上チェックしてください", "メニュータイトル未チェックエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtName.Focus();
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
                darwinDataSet.ログインタイプヘッダRow hr = dts.ログインタイプヘッダ.NewログインタイプヘッダRow();
                hr.Id = int.Parse(txtID.Text);
                hr.名称 = txtName.Text;
                hr.備考 = txtMemo.Text;
                hr.登録年月日 = DateTime.Now;
                hr.更新年月日 = DateTime.Now;
                dts.ログインタイプヘッダ.AddログインタイプヘッダRow(hr);
                
                // タグデータ登録
                addTagRow(int.Parse(txtID.Text));
            }
            else if (sMode == 1)    // 更新
            {
                // ヘッダ
                darwinDataSet.ログインタイプヘッダRow hr = dts.ログインタイプヘッダ.Single(a => a.Id == sID);
                hr.名称 = txtName.Text;
                hr.備考 = txtMemo.Text;
                hr.更新年月日 = DateTime.Now;

                // タグデータ削除
                delTagRow(sID);

                // タグデータ登録
                addTagRow(sID);
            }
            
            // データベース更新
            hAdp.Update(dts.ログインタイプヘッダ);
            tAdp.Update(dts.ログインタイプタグ);

            // データ読み込み
            hAdp.Fill(dts.ログインタイプヘッダ);
            tAdp.Fill(dts.ログインタイプタグ);
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプタグデータ新規登録 </summary>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// ----------------------------------------------------------------------
        private void addTagRow(int sID)
        {
            // タグ
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked)
                {
                    darwinDataSet.ログインタイプタグRow tr = dts.ログインタイプタグ.NewログインタイプタグRow();
                    tr.ヘッダID = sID;
                    tr.tag = Utility.strToInt(item.Text);
                    tr.登録年月日 = DateTime.Now;
                    tr.更新年月日 = DateTime.Now;
                    dts.ログインタイプタグ.AddログインタイプタグRow(tr);
                }
            }
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

        private void listView2_Click(object sender, EventArgs e)
        {
            // 画面初期化
            dispClear();

            // 選択したタグIDを取得
            int r = listView2.SelectedItems[0].Index;

            // データを表示
            getData(Utility.strToInt(listView2.SelectedItems[0].Text));
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
            darwinDataSet.ログインタイプヘッダRow hr = dts.ログインタイプヘッダ.Single(a => a.Id == sID);

            txtID.Text = hr.Id.ToString();
            txtName.Text = hr.名称;
            txtMemo.Text = hr.備考;

            // ログインタイプタグ
            foreach (var t in dts.ログインタイプタグ.Where(a => a.ヘッダID == hr.Id))
            {
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    if (listView1.Items[i].Text == t.tag.ToString())
                    {
                        listView1.Items[i].Checked = true;
                        break;
                    }
                }
            }

            // 処理モード
            fMode.Mode = 1;
            fMode.ID = hr.Id;

            // コード編集は不可とします
            txtID.Enabled = false;

            // 削除、取消ボタンの使用を可能とします
            button2.Enabled = true;
            button4.Enabled = true;
            linkLabel2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // ログインユーザーで使用されているか調べる
            if (!delErrCheck(fMode.ID))
            {
                return;
            }

            // 確認メッセージ
            if (MessageBox.Show("表示中のデータを削除します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            { 
                return; 
            }

            // データ削除
            delData(fMode.ID);

            // 画面初期化
            dispClear();

            // ログオンタイプヘッダ再表示
            hdView();
        }
        
        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     ログインユーザーで使用されているか調べる </summary>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// <returns>
        ///     ログインユーザー登録なし;true, ログインユーザー登録あり:false</returns>
        /// --------------------------------------------------------------------------------
        private bool delErrCheck(int sID)
        {
            // ログインユーザーで使用されている件数を調べる
            int cnt = dts.ログインユーザー.Count(a => a.ログインタイプ == sID);

            if (cnt > 0)
            {
                MessageBox.Show("ログインユーザー " + cnt.ToString()  + "名のログインタイプに使用されています。" + Environment.NewLine + "削除はできません。" , "コードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
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
            darwinDataSet.ログインタイプヘッダRow hr = dts.ログインタイプヘッダ.Single(a => a.Id == sID);
            hr.Delete();

            // タグデータ削除
            delTagRow(sID);
            
            // データベース更新
            hAdp.Update(dts.ログインタイプヘッダ);
            tAdp.Update(dts.ログインタイプタグ);

            // データ読み込み
            hAdp.Fill(dts.ログインタイプヘッダ);
            tAdp.Fill(dts.ログインタイプタグ);
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

        private void txtID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     全てのチェックをオン・オフする </summary>
        /// <param name="status">
        ///     true, false</param>
        /// ---------------------------------------------------------
        private void listViewCheckAll(bool status)
        {
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                listView1.Items[i].Checked = status;                
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 確認メッセージ
            if (MessageBox.Show("全てのメニュータイトルをチェックします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 全てのチェックをオンにする
            listViewCheckAll(true);
            linkLabel2.Enabled = true;
        }

        private void linkLabel2_Click(object sender, EventArgs e)
        {
            // 確認メッセージ
            if (MessageBox.Show("全てのメニュータイトルのチェックをはずします。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 全てのチェックをオフにする
            listViewCheckAll(false);
        }
    }
}

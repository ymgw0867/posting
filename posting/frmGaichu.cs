using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;
using System.Linq;

namespace posting
{
    public partial class frmGaichu : Form
    {
        public frmGaichu()
        {
            InitializeComponent();

            // データを読み込む
            adp.Fill(dts.外注先);
            lAdp.Fill(dts.ログインユーザー);
        }

        string[] zipArray = null;

        Utility.formMode fMode = new Utility.formMode();
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.外注先TableAdapter adp = new darwinDataSetTableAdapters.外注先TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.ログインユーザーTableAdapter lAdp = new darwinDataSetTableAdapters.ログインユーザーTableAdapter();

        const string MESSAGE_CAPTION = "外注先マスター";

        // データグリッドカラム定義
        string colID = "col1";
        string colName = "col2";
        string colZip = "col3";
        string colAdd1 = "col4";
        string colAdd2 = "col5";
        string colAddDate = "col6";
        string colUpDate = "col7";
        string colUserID = "col8";
        string colSday = "col9";    // 支払日 2018/01/03

        bool jStatus = false;

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // グリッドビュー定義
            gridViewSetting(dataGridView1);

            // グリッドビューデータ表示
            gridViewDataShow(dataGridView1, textBox1.Text);
            
            //担当社員コンボ
            Utility.ComboShain.load(cmbShain);

            // 画面初期化
            DispClear();
            textBox1.Text = string.Empty;

            // 郵便番号CSV読み込み
            Utility.zipCsvLoad(ref zipArray);
        }

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void gridViewSetting(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 217;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(colID, "コード");
                tempDGV.Columns.Add(colName, "名称");
                tempDGV.Columns.Add(colZip, "郵便番号");
                tempDGV.Columns.Add(colAdd1, "住所１");
                tempDGV.Columns.Add(colAdd2, "住所２");
                tempDGV.Columns.Add(colAddDate, "登録年月日");
                tempDGV.Columns.Add(colUpDate, "更新年月日");
                tempDGV.Columns.Add(colUserID, "ユーザーID");
                tempDGV.Columns.Add(colSday, "支払日");

                tempDGV.Columns[colID].Width = 90;
                tempDGV.Columns[colName].Width = 200;
                tempDGV.Columns[colZip].Width = 110;
                tempDGV.Columns[colAdd1].Width = 200;
                tempDGV.Columns[colAdd2].Width = 200;
                tempDGV.Columns[colAddDate].Width = 160;
                tempDGV.Columns[colUpDate].Width = 160;
                tempDGV.Columns[colUserID].Width = 110;
                tempDGV.Columns[colSday].Width = 90;

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

        private void gridViewDataShow(DataGridView dg, string sName)
        {
            int r = 0;
            dg.Rows.Clear();

            foreach (var t in dts.外注先.OrderBy(a => a.ID))
            {
                if (sName ==string.Empty || (sName != string.Empty && t.名称.Contains(sName)))
                {
                    dg.Rows.Add();
                    dg[colID, r].Value = t.ID.ToString();
                    dg[colName, r].Value = t.名称;
                    dg[colZip, r].Value = t.郵便番号;
                    dg[colAdd1, r].Value = t.住所1;
                    dg[colAdd2, r].Value = t.住所2;
                    dg[colAddDate, r].Value = t.登録年月日.ToString();
                    dg[colUpDate, r].Value = t.更新年月日.ToString();

                    if (t.ログインユーザーRow == null)
                    {
                        dg[colUserID, r].Value = string.Empty;
                    }
                    else
                    {
                        dg[colUserID, r].Value = t.ログインユーザーRow.ログインユーザー;
                    }

                    dg[colSday, r].Value = t.支払日.ToString();

                    r++;
                }
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     グリッドからデータを選択 </summary>
        ///-----------------------------------------------------------
        private void GridEnter()
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[colName, dataGridView1.SelectedRows[iX].Index].Value + "が選択されました" + Environment.NewLine;
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "外注先選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    //データを取得する
                    int sID = int.Parse(dataGridView1[colID, dataGridView1.SelectedRows[0].Index].Value.ToString());

                    darwinDataSet.外注先Row r = (darwinDataSet.外注先Row)dts.外注先.Single(a => a.ID == sID);

                    // データ値を取得
                    txtFuri.Text = r.フリガナ;
                    txtName2.Text = r.名称;
                    txtTantou.Text = r.担当者名;
                    txtBusho.Text = r.担当部署;
                    mtxtZipCode.Text = r.郵便番号;
                    txtAddress1.Text = r.住所1;
                    txtAddress2.Text = r.住所2;
                    txtTel.Text = r.電話番号;
                    txtFax.Text = r.FAX番号;
                    txtEmail.Text = r.eMail;
                    cmbShain.SelectedValue = r.担当営業;
                    txtMemo.Text = r.備考;
                    txtShiharaibi.Text = r.支払日.ToString();

                    //IDテキストボックスは編集不可とする
                    //txtCode.Enabled = false;

                    //ボタン状態
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     // フォームモードステータス:変更削除
                    fMode.ID = sID;

                    txtName2.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "画面クリア", MessageBoxButtons.OK);
                }
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            // Enterkey以外は対象外
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            GridEnter();
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //GridEnter();
        }

        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {
            try
            {
                fMode.Mode = 0;

                txtFuri.Text = "";
                txtName2.Text = "";
                txtTantou.Text = "";
                txtBusho.Text = "";
                mtxtZipCode.Text = "";
                txtAddress1.Text = "";
                txtAddress2.Text = "";
                txtTel.Text = "";
                txtFax.Text = "";
                txtEmail.Text = "";
                cmbShain.SelectedIndex = -1;
                txtMemo.Text = "";
                txtShiharaibi.Text = "";

                btnDel.Enabled = false;
                btnClr.Enabled = false;

                if (this.dataGridView1.RowCount > 0)
                {
                    btnCsv.Enabled = true;
                }
                else
                {
                    btnCsv.Enabled = false;
                }

                txtName2.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("選択されているデータを変更しないで終了します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
                
            DispClear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                if (fDataCheck())
                {
                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                return;
                            }

                            if (dataAdd())
                            {
                                MessageBox.Show("新規登録されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("新規登録に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                return;
                            }

                            dataUpDate();

                            break;
                    }

                    // データベース更新
                    adp.Update(dts.外注先);

                    //データを 'darwinDataSet.外注先' テーブルに読み込みます。
                    adp.Fill(dts.外注先);

                    // 画面初期化
                    DispClear();

                    // グリッド再表示
                    gridViewDataShow(dataGridView1, textBox1.Text);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"更新処理",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }


        /// -----------------------------------------------------------------
        /// <summary>
        ///     外注先データセット更新 </summary>
        /// <returns>
        ///     true:行更新成功、false:行更新失敗</returns>
        /// -----------------------------------------------------------------
        private bool dataUpDate()
        {
            try
            {
                darwinDataSet.外注先Row s = dts.外注先.Single(a => a.ID == fMode.ID);
                dataRowSet(s);
                
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
        
        /// -----------------------------------------------------------------
        /// <summary>
        ///     外注先新規登録 </summary>
        /// <returns>
        ///     true:行追加成功、false:行追加失敗</returns>
        /// -----------------------------------------------------------------
        private bool dataAdd()
        {
            try
            {
                darwinDataSet.外注先Row r = dts.外注先.New外注先Row();
                dataRowSet(r);

                // データセットに行追加
                dts.外注先.Add外注先Row(r);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     外注先データRowにデータセット </summary>
        /// <param name="r">
        ///     darwinDataSet.外注先Row </param>
        /// --------------------------------------------------------------------
        private darwinDataSet.外注先Row dataRowSet(darwinDataSet.外注先Row r)
        {
            r.名称 = txtName2.Text;
            r.フリガナ = txtFuri.Text;
            r.担当者名 = txtTantou.Text;
            r.担当部署 = txtBusho.Text;

            if (((mtxtZipCode.Text).Replace("-", "")).Trim() == "")
            {
                r.郵便番号 = "";
            }
            else
            {
                r.郵便番号 = mtxtZipCode.Text;
            }

            r.住所1 = txtAddress1.Text;
            r.住所2 = txtAddress2.Text;
            r.電話番号 = txtTel.Text;
            r.FAX番号 = txtFax.Text;
            r.eMail = txtEmail.Text;

            Utility.ComboShain cmb1 = new Utility.ComboShain();

            if (cmbShain.SelectedIndex == -1)
            {
                r.担当営業 = 0;
            }
            else
            {
                cmb1 = (Utility.ComboShain)cmbShain.SelectedItem;
                r.担当営業 = cmb1.ID;
            }

            r.備考 = txtMemo.Text;

            if (fMode.Mode == 0)
            {
                r.登録年月日 = DateTime.Now;
            }

            r.更新年月日 = DateTime.Now;
            r.ユーザーID = global.loginUserID;
            r.支払日 = Utility.strToInt(txtShiharaibi.Text);

            return r;
        }


        ///--------------------------------------------------------------
        /// <summary>
        ///     登録データチェック </summary>
        /// <returns>
        ///     true:エラーなし, false:エラー有り</returns>
        ///--------------------------------------------------------------
        private Boolean fDataCheck()
        {
            try
            {
                // 名称チェック
                if (txtName2.Text.Trim().Length < 1)
                {
                    txtName2.Focus();
                    throw new Exception("名称を入力してください");
                }

                // 支払日チェック
                if (Utility.strToInt(txtShiharaibi.Text) > 31)
                {
                    txtShiharaibi.Focus();
                    throw new Exception("支払日が正しくありません");
                }

                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        private void txtEnter(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();
            MaskedTextBox objMtxt = new MaskedTextBox();

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}
            
            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;

                // 2019/03/16
                MyTextKana.TextKana.textEnter(txtName2);
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }
            
            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == txtShiharaibi)
            {
                objtxt = txtShiharaibi;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

            objMtxt.SelectAll();
            objMtxt.BackColor = Color.LightGray;
        }

        private void txtLeave(object sender, EventArgs e)
       {
           TextBox objtxt = new TextBox();
           MaskedTextBox objMtxt = new MaskedTextBox();

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}
            
            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;
            }
            
            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == txtShiharaibi)
            {
                objtxt = txtShiharaibi;
            }

            objtxt.BackColor = Color.White;
            objMtxt.BackColor = Color.White;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            dataDel();
        }

        private void dataDel()
        {
            // 削除確認：2018/01/04
            if (MessageBox.Show("削除処理が選択されました。続行してよろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // データ読み込み：2018/01/04
            if (!jStatus)
            {
                Cursor = Cursors.WaitCursor;
                jAdp.Fill(dts.受注1);
                jStatus = true;
                Cursor = Cursors.Default;
            }

            // 受注データに登録されているときは削除不可とする
            if (dts.受注1.Any(a => a.外注先ID営業 == fMode.ID || a.外注先ID支払 == fMode.ID))
            {
                MessageBox.Show(txtName2.Text.ToString() + " は受注データに登録されています", txtName2.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // 削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // データ削除
            if (dts.外注先.Any(a => a.ID == fMode.ID))
            {
                darwinDataSet.外注先Row r = dts.外注先.Single(a => a.ID == fMode.ID);
                r.Delete();

                // データベース更新
                adp.Update(dts.外注先);

                // メッセージ
                MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // メッセージ
                MessageBox.Show("該当する外注先データがありません", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            // 画面初期化
            DispClear();

            //データを 'darwinDataSet.外注先' テーブルに読み込みます。
            adp.Fill(dts.外注先);

            // グリッド再表示
            gridViewDataShow(dataGridView1, textBox1.Text);
        }
        
        private void btnCsv_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            gridViewDataShow(dataGridView1, textBox1.Text);
        }
        
        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmShain();

            frm.ShowDialog();
            Utility.ComboShain.load(cmbShain);
        }
        
        private void txtName2_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/16
            string furi = MyTextKana.TextKana.textBox_KeyDown(txtName2, sender, e);

            // 2019/03/16
            txtFuri.Text += furi;
        }

        private void txtName2_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtName2.Text) == false)
            {
                txtFuri.Text = "";
            }
        }

        private void mtxtZipCode_TextChanged(object sender, EventArgs e)
        {
            if (zipArray == null)
            {
                return;
            }

            MaskedTextBox mTxt = null;
            TextBox txtAdd = null;
            bool mc = false;

            mTxt = mtxtZipCode;
            txtAdd = txtAddress1;

            string zipText = mTxt.Text.Replace("-", "");

            if (zipText.Length == 7)
            {
                foreach (var t in zipArray)
                {
                    string[] r = t.Split(',');

                    if (zipText == r[2].Replace("\"", ""))
                    {
                        // 住所
                        string ad = r[6].Replace("\"", "") + r[7].Replace("\"", "") + r[8].Replace("\"", "");

                        // 同じ住所ならば再表示しない
                        if (!txtAdd.Text.Contains(ad))
                        {
                            txtAdd.Text = ad;
                        }
                        txtAdd.Focus();
                        mc = true;
                        break;
                    }
                }
            }
        }

        private void txtShiharaibi_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
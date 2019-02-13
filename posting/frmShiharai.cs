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
    public partial class frmShiharai : Form
    {
        public frmShiharai(int sID)
        {
            InitializeComponent();

            // データ読み込み
            adp.Fill(dts.外注支払);
            uAdp.Fill(dts.ログインユーザー);
            gAdp.Fill(dts.外注先);
            jAdp.Fill(dts.受注1);

            _dataID = sID;
        }

        // ID
        int _dataID = 0;

        // cmbStatus
        bool cmbStatus = true;

        // 外注先
        //int gaichuNum = 0;
        const int G1 = 1;
        const int G2 = 2;
        const int G3 = 3;
        int gID = 0;

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.外注支払TableAdapter adp = new darwinDataSetTableAdapters.外注支払TableAdapter();
        darwinDataSetTableAdapters.ログインユーザーTableAdapter uAdp = new darwinDataSetTableAdapters.ログインユーザーTableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        // 支払コード
        string shiCode = string.Empty;

        Utility.消費税率 cTax = new Utility.消費税率();

        #region グリッドビューカラム定義
        string colID = "col1";
        string colGcd = "col2";
        string colDt = "col6";
        string colMemo = "col3";
        string colAddDt = "col4";
        string colUpDt = "col5";
        string colUserID = "col7";
        string colClient = "col8";
        string colKingaku = "col9";
        string colChouseiGaku = "col10";
        string colChouseiMemo = "col11";
        #endregion

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            //Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 外注先コンボボックスアイテムロード
            Utility.comboGaichu.itemLoad(cmbGaichu);
            //Utility.comboGaichu cg = new Utility.comboGaichu();
            //cg.itemLoad(cmbGaichu); 

            // データグリッドビューの定義
            gridSetting(dataGridView1);

            // データグリッドビューデータ表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();

            // データ表示
            if (_dataID != 0)
            {
                getData(_dataID);
            }
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
                tempDGV.Columns.Add(colID, "");
                tempDGV.Columns.Add(colDt, "日付");
                tempDGV.Columns.Add(colClient, "外注先名称");
                tempDGV.Columns.Add(colKingaku, "支払金額");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colChouseiGaku, "調整金額");
                tempDGV.Columns.Add(colChouseiMemo, "調整備考");
                tempDGV.Columns.Add(colAddDt, "登録年月日");
                tempDGV.Columns.Add(colUpDt, "更新年月日");
                tempDGV.Columns.Add(colUserID, "ユーザーID");

                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[colDt].Width = 110;
                tempDGV.Columns[colClient].Width = 200;
                tempDGV.Columns[colKingaku].Width = 110;
                tempDGV.Columns[colMemo].Width = 200;
                tempDGV.Columns[colChouseiGaku].Width = 110;
                tempDGV.Columns[colChouseiMemo].Width = 200;
                tempDGV.Columns[colAddDt].Width = 140;
                tempDGV.Columns[colUpDt].Width = 140;
                tempDGV.Columns[colUserID].Width = 140;

                tempDGV.Columns[colDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colChouseiGaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
            foreach (var t in dts.外注支払.OrderByDescending(a => a.日付))
            {
                g.Rows.Add();
                g[colID, iX].Value = t.ID.ToString();

                if (t.外注先Row == null)
                {
                    g[colClient, iX].Value = string.Empty;
                }
                else
                {
                    g[colClient, iX].Value = t.外注先Row.名称;
                }

                g[colDt, iX].Value = t.日付.ToShortDateString();
                g[colKingaku, iX].Value = t.支払金額.ToString("#,##0");
                g[colMemo, iX].Value = t.備考;
                g[colChouseiGaku, iX].Value = t.調整額.ToString("n0");
                g[colChouseiMemo, iX].Value = t.調整備考;
                g[colAddDt, iX].Value = t.登録年月日;
                g[colUpDt, iX].Value = t.変更年月日;

                if (t.ログインユーザーRow == null)
                {
                    g[colUserID, iX].Value = string.Empty;
                }
                else
                {
                    g[colUserID, iX].Value = t.ログインユーザーRow.ログインユーザー;
                }
                
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
            txtKingaku.Text = string.Empty;
            cmbGaichu.SelectedIndex = -1;
            txtMemo.Text = string.Empty;
            checkedListBox1.Items.Clear();
            checkedListBox1.Enabled = true;
            linkLabel1.Visible = false;

            button1.Enabled = true;
            button2.Enabled = false;
            button4.Enabled = false;

            txtSai.Text = string.Empty;
            txtSaiMemo.Text = string.Empty;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // エラーチェック
            if (!errCheck(fMode.Mode))
            {
                return;
            }
            
            string _msg = kingakuKakunin(); 

            // 確認メッセージ
            if (fMode.Mode == 0)
            {
                if (MessageBox.Show(_msg,"確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }
            else if (fMode.Mode == 1)
            {
                if (MessageBox.Show(_msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }
            
            // 登録・更新処理
            dataUpdate(fMode.Mode, fMode.ID);

            // グリッド表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();
        }


        private string kingakuKakunin()
        {
            decimal k = 0;
            decimal g = 0;
            decimal z = 0;

            foreach (var t in checkedListBox1.CheckedItems)
            {
                Int64 fID = Int64.Parse(t.ToString().Substring(0, 12));

                // 2016/12/19
                int sNum = Utility.strToInt(t.ToString().Substring(13, 1));

                var r = dts.受注1.Single(a => a.ID == fID);                

                //k = r.外注原価支払 * (decimal)r.枚数;

                // 2015/12/06 外注原価を単価から原価総額入力へ変更に伴う
                if (r.外注先ID支払 == gID && sNum == 1)    // 2016/10/17, 2016/11/17, 2016/12/19
                {
                    k = r.外注原価支払;
                    g += r.外注原価支払;
                }
                else if (r.外注先ID支払2 == gID && sNum == 2)    // 2016/10/17, 2016/11/17, 2016/12/19
                {
                    k = r.外注原価支払2;
                    g += r.外注原価支払2;
                }
                else if (r.外注先ID支払3 ==gID && sNum == 3)    // 2016/10/17, 2016/11/17, 2016/12/19
                {
                    k = r.外注原価支払3;
                    g += r.外注原価支払3;
                }

                // 2016/02/01 消費税表示撤廃
                //// 税率取得
                //cTax.Ritsu = Utility.GetTaxRT(r.受注日);

                    // 2016/02/01 消費税表示撤廃
                    //// 消費税額計算 
                    //z += Utility.GetTax(k, cTax.Ritsu);
            }

            string msg = "入力した支払金額を登録してよろしいですか？";
            msg += Environment.NewLine + Environment.NewLine;
            msg += "入力金額：" + Utility.strToInt(txtKingaku.Text).ToString("#,##0") + Environment.NewLine;
            msg += "精算金額：" + Utility.strToInt(txtSai.Text).ToString("#,##0") + Environment.NewLine;
            msg += "支払金額：" + g.ToString("#,##0") + Environment.NewLine;

            // 2016/02/01 消費税表示撤廃
            //msg += "税込金額：" + (g + z).ToString("#,##0") + Environment.NewLine;

            return msg;
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
            // 支払金額と精算額がともに０のとき
            if (Utility.strToInt(txtKingaku.Text) == 0 && Utility.strToInt(txtSai.Text) == 0)
            {
                MessageBox.Show("支払金額または差異精算額を入力してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtKingaku.Focus();
                return false;
            }

            if (checkedListBox1.SelectedItems.Count == 0)
            {
                MessageBox.Show("対象案件を選択してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtKingaku.Focus();
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
                darwinDataSet.外注支払Row r = dts.外注支払.New外注支払Row();
                dts.外注支払.Add外注支払Row(dataSetup(sMode, r));
            }
            else
            {
                // 更新
                darwinDataSet.外注支払Row r = dts.外注支払.Single(a => a.ID == sID);
                dataSetup(sMode, r);
            }
            
            // 外注先支払完了フラグ更新
            shiharaiFlgUpdate();

            // データベース更新
            adp.Update(dts.外注支払);
            jAdp.Update(dts.受注1);

            // データ読み込み
            adp.Fill(dts.外注支払);
            jAdp.Fill(dts.受注1);
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     darwinDataSet.外注支払Rowにデータ値をセットする </summary>
        /// <param name="r">
        ///     darwinDataSet.外注支払Row</param>
        /// <returns>
        ///     darwinDataSet.外注支払Row</returns>
        /// ------------------------------------------------------------------------
        private darwinDataSet.外注支払Row dataSetup(int sMode, darwinDataSet.外注支払Row r)
        {
            if (cmbGaichu.SelectedIndex == -1)
            {
                r.外注先コード = 0;
            }
            else
            {
                Utility.comboGaichu cmb = (Utility.comboGaichu)cmbGaichu.SelectedItem;
                r.外注先コード = cmb.ID;
            }

            r.日付 = DateTime.Parse(dtNyuko.Value.ToShortDateString());
            r.支払金額 = Utility.strToInt(txtKingaku.Text);
            r.備考 = txtMemo.Text;

            r.調整額 = Utility.strToInt(txtSai.Text);

            if (Utility.strToInt(txtSai.Text) != 0)
            {
                r.調整日付 = dtNyuko.Value.ToShortDateString();
            }
            else
            {
                r.調整日付 = string.Empty;
            }

            r.調整備考 = txtSaiMemo.Text;

            if (sMode == 0)
            {
                DateTime dt = DateTime.Now;
                shiCode = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                r.支払コード = shiCode;

                r.登録年月日 = DateTime.Now;
            }

            r.変更年月日 = DateTime.Now;
            r.ユーザーID = global.loginUserID;

            return r;
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     外注先支払完了更新 </summary>
        /// ---------------------------------------------------------
        private void shiharaiFlgUpdate()
        {
            if (checkedListBox1.Enabled)
            {
                foreach (var t in checkedListBox1.CheckedItems)
                {
                    Int64 fID = Int64.Parse(t.ToString().Substring(0, 12));

                    // 2016/12/19
                    int sNum = Utility.strToInt(t.ToString().Substring(13, 1));

                    darwinDataSet.受注1Row r = dts.受注1.Single(a => a.ID == fID);

                    // 2016/10/17, 2016/11/17, 2016/12/19
                    if (r.外注先ID支払 == gID && sNum == 1)
                    {
                        r.外注支払ID = shiCode;
                    }
                    else if (r.外注先ID支払2 == gID && sNum == 2)
                    {
                        r.外注支払ID2 = shiCode;
                    }
                    else if(r.外注先ID支払3 == gID && sNum == 3)
                    {
                        r.外注支払ID3 = shiCode;
                    }

                    r.変更年月日 = DateTime.Now;
                    r.ユーザーID = global.loginUserID;
                }
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
            darwinDataSet.外注支払Row r = dts.外注支払.Single(a => a.ID == sID);

            dtNyuko.Value = r.日付;
            cmbGaichu.SelectedValue = r.外注先コード;
            gID = r.外注先コード;     // 2016/11/17
            txtKingaku.Text = r.支払金額.ToString();
            txtMemo.Text = r.備考;
            txtSai.Text = r.調整額.ToString("n0");
            txtSaiMemo.Text = r.調整備考;

            checkedListBox1.Items.Clear();
            foreach (var t in dts.受注1.Where(a => (a.外注支払ID == r.支払コード) ||
                                                   (a.外注支払ID2 == r.支払コード) ||
                                                   (a.外注支払ID3 == r.支払コード)))
            {
                //decimal gaku = t.外注原価支払 * (decimal)t.枚数;

                decimal gaku = 0;

                // 2015/12/06 外注原価を「単価から原価総額入力へ変更」に伴う
                // 2016/10/17
                if (t.外注支払ID == r.支払コード)
                {
                    //gaichuNum = G1;
                    gaku = t.外注原価支払;
                    checkedListBox1.Items.Add(t.ID + "_1 " + t.外注支払日支払.ToShortDateString() + "  " + t.チラシ名 + "  " + gaku.ToString("C"));
                }

                if (t.外注支払ID2 == r.支払コード)
                {
                    //gaichuNum = G2;
                    gaku = t.外注原価支払2;
                    checkedListBox1.Items.Add(t.ID + "_2 " + t.外注支払日支払2.ToShortDateString() + "  " + t.チラシ名 + "  " + gaku.ToString("C"));
                }

                if(t.外注支払ID3 == r.支払コード)
                {
                    //gaichuNum = G3;
                    gaku = t.外注原価支払3;
                    checkedListBox1.Items.Add(t.ID + "_3 " + t.外注支払日支払3.ToShortDateString() + "  " + t.チラシ名 + "  " + gaku.ToString("C"));
                }

                // 2016/02/01 消費税表示撤廃
                //// 税率取得
                //cTax.Ritsu = Utility.GetTaxRT(t.受注日);

                // 2016/02/01 消費税表示撤廃
                //// 消費税額計算 
                //decimal KingakuTax = Utility.GetTax(gaku, cTax.Ritsu);

                // 2016/02/01 消費税表示撤廃
                //checkedListBox1.Items.Add(t.ID + "  " + t.外注支払日支払.ToShortDateString() + "  " + t.チラシ名 + "  " + gaku.ToString("C") + "  税込 " + (gaku + KingakuTax).ToString("C"));

            }

            checkedListBox1.Enabled = false;
            linkLabel1.Visible = true;


            // 処理モード
            fMode.Mode = 1;
            fMode.ID = r.ID;
            shiCode = r.支払コード;
            
            // 削除、取消ボタンの使用を可能とします
            button1.Enabled = true;
            button2.Enabled = true;
            button4.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 確認メッセージ
            if (MessageBox.Show("表示中の支払データを削除します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            { 
                return; 
            }

            // データ削除
            delData(fMode.ID, shiCode);

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
        /// <param name="shiID">
        ///     支払コード</param>
        /// ----------------------------------------------------------------------
        private void delData(int sID, string shiID)
        {
            // 完了案件支払コードクリア
            if (dts.受注1.Any(a => a.外注支払ID == shiID))
            {
                foreach (var t in dts.受注1
                    .Where(a => (a.外注支払ID == shiID) || (a.外注支払ID2 == shiID) || (a.外注支払ID3 == shiID)))
                {
                    if (t.外注支払ID == shiID)
                    {
                        t.外注支払ID = string.Empty;
                    }
                    else if (t.外注支払ID2 == shiID)
                    {
                        t.外注支払ID2 = string.Empty;
                    }
                    else if (t.外注支払ID3 == shiID)
                    {
                        t.外注支払ID3 = string.Empty;
                    }
                }

                //foreach (var t in dts.受注1.Where(a => a.外注支払ID == shiID))
                //{
                //    t.外注支払ID = string.Empty;
                //}

                jAdp.Update(dts.受注1);
                jAdp.Fill(dts.受注1);
            }

            // 外注支払データ削除
            darwinDataSet.外注支払Row r = dts.外注支払.Single(a => a.ID == sID);
            r.Delete();
            
            // データベース更新
            adp.Update(dts.外注支払);

            // データ読み込み
            adp.Fill(dts.外注支払);
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

        private void btnClient_Click(object sender, EventArgs e)
        {

        }

        private void txtKingaku_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void cmbGaichu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbStatus)
            {
                if (fMode.Mode == 0)
                {
                    if (cmbGaichu.SelectedIndex != -1)
                    {
                        showListBoxAdd();
                    }
                }
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     リストボックスへ選択した外注先支払案件表示 </summary>
        ///---------------------------------------------------------------------
        private void showListBoxAdd()
        {
            Utility.comboGaichu cmb = (Utility.comboGaichu)cmbGaichu.SelectedItem;

            checkedListBox1.Items.Clear();
            foreach (var t in dts.受注1
                        .Where(a => ((a.外注先ID支払 == cmb.ID && a.外注支払ID.Trim() == string.Empty) || 
                               (a.外注先ID支払2 == cmb.ID && a.外注支払ID2.Trim() == string.Empty) ||
                               (a.外注先ID支払3 == cmb.ID && a.外注支払ID3.Trim() == string.Empty)) && 
                               !a.Is外注支払日支払Null())
                               .OrderBy(a => a.外注支払日支払))
            {
                //decimal gaku = t.外注原価支払 * (decimal)t.枚数;

                decimal gaku = 0;
                gID = cmb.ID;

                // 2015/12/06 外注原価を「単価から原価総額入力へ変更」に伴う
                if (t.外注先ID支払 == gID && t.外注支払ID.Trim() == string.Empty)
                {
                    //gaichuNum = G1;
                    gaku = t.外注原価支払;

                    // 2018/03/07
                    if (t.Is外注支払日支払Null())
                    {
                        checkedListBox1.Items.Add(t.ID + "_1 " + " " + " " + t.チラシ名 + " " + gaku.ToString("C"));
                    }
                    else
                    {

                        checkedListBox1.Items.Add(t.ID + "_1 " + t.外注支払日支払.ToShortDateString() + " " + t.チラシ名 + " " + gaku.ToString("C"));
                    }
                }

                if (t.外注先ID支払2 == gID && t.外注支払ID2.Trim() == string.Empty)
                {
                    //gaichuNum = G2;
                    gaku = t.外注原価支払2;
                    checkedListBox1.Items.Add(t.ID + "_2 " + t.外注支払日支払2.ToShortDateString() + " " + t.チラシ名 + " " + gaku.ToString("C"));
                }

                if (t.外注先ID支払3 == gID && t.外注支払ID3.Trim() == string.Empty)
                {
                    //gaichuNum = G3;
                    gaku = t.外注原価支払3;
                    checkedListBox1.Items.Add(t.ID + "_3 " + t.外注支払日支払3.ToShortDateString() + " " + t.チラシ名 + " " + gaku.ToString("C"));
                }

                // 2016/02/01 消費税表示撤廃
                //// 税率取得
                //cTax.Ritsu = Utility.GetTaxRT(t.受注日);

                // 2016/02/01 消費税表示撤廃
                //// 消費税額計算 
                //decimal KingakuTax = Utility.GetTax(gaku, cTax.Ritsu);

                // 2016/02/01 消費税表示撤廃
                //checkedListBox1.Items.Add(t.ID + "  " + t.外注支払日支払.ToShortDateString() + "  " + t.チラシ名 + "  " + gaku.ToString("C") + "  税込 " + (gaku + KingakuTax).ToString("C"));
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 完了案件再表示
            checkRetry(shiCode);
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     完了案件解除、再入力 </summary>
        /// <param name="sID">
        ///     外注支払コード </param>
        ///----------------------------------------------------------------------
        private void checkRetry(string sID)
        {
            if (MessageBox.Show("実行すると現在選択されている完了案件は解除され、完了案件を再入力する必要があります。" + Environment.NewLine + Environment.NewLine + "よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            foreach (var t in dts.受注1
                .Where(a => (a.外注支払ID == sID) || (a.外注支払ID2 == sID) || (a.外注支払ID3 == sID)))
            {
                if (t.外注支払ID == sID)
                {
                    t.外注支払ID = string.Empty;
                }

                if (t.外注支払ID2 == sID)
                {
                    t.外注支払ID2 = string.Empty;
                }

                if (t.外注支払ID3 == sID)
                {
                    t.外注支払ID3 = string.Empty;
                }
            }

            jAdp.Update(dts.受注1);
            jAdp.Fill(dts.受注1);

            showListBoxAdd();

            checkedListBox1.Enabled = true;
        }

        private void gaichuMnt()
        {
            cmbStatus = false;

            // 外注先マスター保守
            frmGaichu frm = new frmGaichu();
            frm.ShowDialog();

            // 外注先コンボボックスアイテムロード
            int idx = cmbGaichu.SelectedIndex;
            Utility.comboGaichu.itemLoad(cmbGaichu);
            cmbGaichu.SelectedIndex = idx;

            // 外注先データ読み込み
            gAdp.Fill(dts.外注先);
            
            cmbStatus = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 外注先マスター保守
            gaichuMnt();
        }

        private void txtSai_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

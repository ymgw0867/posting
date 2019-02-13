using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmNyukinRep2015 : Form
    {
        public frmNyukinRep2015()
        {
            InitializeComponent();

            // データ読み込み
            jAdp.Fill(dts.新請求書);
            cAdp.Fill(dts.得意先);
            sAdp.Fill(dts.社員);
            nAdp.Fill(dts.新入金);
            rAdp.Fill(dts.受注1);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.新請求書TableAdapter jAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.新入金TableAdapter nAdp = new darwinDataSetTableAdapters.新入金TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter rAdp = new darwinDataSetTableAdapters.受注1TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        bool mukouStatus = false;

        #region グリッドビューカラム定義
        string colSel = "col0";         // チェック列
        string colID = "col1";          // 請求書番号
        string colClient = "col2";      // 得意先
        string colKbn = "col3";         // 依頼内容
        string colKingaku = "col4";     // 小計
        string colTax = "col5";         // 消費税
        string colDt = "col6";          // 請求日
        string colSDt = "col7";         // 支払予定日
        string colSel2 = "col8";        // チェック列
        string colEtan = "col9";        // 営業担当者
        string colBtn = "col10";        // 詳細ボタン
        string colCnt = "col12";        // 案件数
        string colMemo = "col13";       // メモ
        string colNyukin = "col14";     // 入金
        string colNDt = "col15";        // 入金日
        string colZan = "col16";        // 残金
        string colMukou = "col17";      // 無効
        string colUserID = "col11";
        string colClientID = "col18";   // クライアントID
        string colSeisan = "col19";     // 精算金額 2016/06/16
        string colSeisanTl = "col20";   // 精算合計 2016/06/16
        string colTantou = "col21";     // 営業担当者 2016/09/16
        string colMaeuke = "col22";     // 前受金 2017/05/24
        string colFuri = "col23";       // フリガナ 2017/08/08
        string colKouza = "col24";      // 口座 2017/08/08
        #endregion

        const string CHKON = "1";
        const string CHKOFF = "0";

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            //Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            //// 得意先コンボボックスアイテムロード
            //Utility.ComboClient.itemsLoad(cmbClient);

            // データグリッドビューの定義
            gridSetting(dataGridView1);

            //// データグリッドビューデータ表示
            //gridShow(dataGridView1, DateTime.Parse(dtNyuko.Value.ToShortDateString()));

            // 画面初期化
            dtNyuko.Checked = false;
            dtNyuko2.Checked = false;       // 2017/05/24
            button2.Enabled = false;        // 2016/05/11
            button1.Enabled = false;        // 2017/08/14
            txtClient.Text = string.Empty;
            txtTantou.Text = string.Empty;
            dtNyukin.Checked = false;       // 2017/08/14
            dtNyukin2.Checked = false;      // 2017/08/14

            //dispClear();
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
                tempDGV.Height = 578;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列指定
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "詳細";
                tempDGV.Columns.Add(btn);
                tempDGV.Columns[0].HeaderText = "";
                tempDGV.Columns[0].Name = colBtn;

                tempDGV.Columns.Add(colID, "No.");
                tempDGV.Columns.Add(colClient, "社名");
                tempDGV.Columns.Add(colFuri, "フリガナ");   // 2017/08/10
                tempDGV.Columns.Add(colKingaku, "小計");
                tempDGV.Columns.Add(colTax, "消費税");
                tempDGV.Columns.Add(colSDt, "支払期日");
                tempDGV.Columns.Add(colNDt, "入金日");
                
                //Utility.DataGridViewMaskedTextBoxColumn nd2 = new Utility.DataGridViewMaskedTextBoxColumn();
                //nd2.Mask = "0000/00/00";
                //tempDGV.Columns.Add(nd2);
                //tempDGV.Columns[6].Name = colNDt;
                //tempDGV.Columns[6].HeaderText = "入金日";

                DataGridViewCheckBoxColumn cbc2 = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc2);
                tempDGV.Columns[8].HeaderText = "入金済";
                tempDGV.Columns[8].Name = colSel2;

                tempDGV.Columns.Add(colNyukin, "入金額");
                tempDGV.Columns.Add(colZan, "残金");
                tempDGV.Columns.Add(colSeisan, "精算額");    // 2016/06/16
                tempDGV.Columns.Add(colSeisanTl, "請求残");  // 2016/06/16
                tempDGV.Columns.Add(colKouza, "口座");       // 2017/08/10

                DataGridViewCheckBoxColumn cbc = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc);
                tempDGV.Columns[14].HeaderText = "無効";
                tempDGV.Columns[14].Name = colMukou;

                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colClientID, "CID");
                tempDGV.Columns.Add(colTantou, "営業担当者");
                tempDGV.Columns.Add(colMaeuke, "前受金含");

                // 各列幅指定
                tempDGV.Columns[colBtn].Width = 40;
                tempDGV.Columns[colID].Width = 50;
                tempDGV.Columns[colClient].Width = 220;
                tempDGV.Columns[colFuri].Width = 160;
                tempDGV.Columns[colKingaku].Width = 90;
                tempDGV.Columns[colTax].Width = 70;
                tempDGV.Columns[colSDt].Width = 100;
                tempDGV.Columns[colNDt].Width = 100;
                tempDGV.Columns[colSel2].Width = 50;
                tempDGV.Columns[colNyukin].Width = 90;
                tempDGV.Columns[colSeisan].Width = 90;
                tempDGV.Columns[colZan].Width = 90;
                tempDGV.Columns[colKouza].Width = 80;
                tempDGV.Columns[colMukou].Width = 50;
                tempDGV.Columns[colClientID].Width = 60;
                tempDGV.Columns[colTantou].Width = 100;
                tempDGV.Columns[colMaeuke].Width = 80;
                tempDGV.Columns[colMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colNyukin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colSeisan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colSeisanTl].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colTax].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colZan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colNDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colSDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colClientID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colMaeuke].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colKouza].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列固定
                //tempDGV.Columns[colClient].Frozen = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                //tempDGV.ReadOnly = true;
                foreach (DataGridViewColumn d in tempDGV.Columns)
                {
                    if (d.Name == colSel2 || d.Name == colMukou || d.Name == colMemo)
                    {
                        d.ReadOnly = false;
                    }
                    else
                    {
                        d.ReadOnly = true;
                    }
                }

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
                //tempDGV.BorderStyle = BorderStyle.None;
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
        private void gridShow(DataGridView g, DateTime dt, DateTime edt, DateTime ndt, DateTime nedt)
        {
            g.Rows.Clear();
            int iX = 0;

            var s = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0).OrderBy(a => a.支払期日).ThenBy(a => a.ID);

            // 支払期日
            //if (dt.ToShortDateString() != "1900/01/01")
            //{
            //    s = s.Where(a => a.支払期日 == dt).OrderBy(a => a.支払期日).ThenBy(a => a.ID);
            //}
            
            s = s.Where(a => a.支払期日 >= dt && a.支払期日 <= edt).OrderBy(a => a.支払期日).ThenBy(a => a.ID);

            foreach (var t in s)
            {
                // 未入金のみ指定のとき
                if (chkMinyukin.Checked)
                {
                    if (t.Get新入金Rows().Count() > 0)
                    {
                        continue;
                    }
                }

                // 前受金のみ指定のとき：2017/05/24
                if (chkMaeuke.Checked)
                {
                    if (!isMaeuke(t.ID))
                    {
                        continue;
                    }
                }

                // クライアント名指定のとき
                if (txtClient.Text != string.Empty)
                {
                    if (t.得意先Row == null)
                    {
                        continue;
                    }

                    // 検索は略称とフリガナを対象とする
                    if (!t.得意先Row.略称.Contains(txtClient.Text.Trim()) && !t.得意先Row.フリガナ.Contains(txtClient.Text.Trim()))
                    {
                        continue;
                    }
                }

                // 営業担当者指定の時
                if (txtTantou.Text != string.Empty)
                {
                    if (t.得意先Row == null)
                    {
                        continue;
                    }

                    if (t.得意先Row.社員Row == null)
                    {
                        continue;
                    }

                    if (!t.得意先Row.社員Row.氏名.Contains(txtTantou.Text.Trim()))
                    {
                        continue;
                    }
                }


                // 入金日範囲 : 2017/08/14

                //if (gDt == string.Empty)
                //{
                //    continue;
                //}

                if (dtNyukin.Checked || dtNyukin2.Checked)
                {
                    string gDt = getNyukinDate(t.ID);
                    DateTime outdt;

                    if (DateTime.TryParse(gDt, out outdt))
                    {
                        if (outdt < ndt)
                        {
                            continue;
                        }

                        if (outdt > nedt)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        continue;
                    }
                }

                g.Rows.Add();

                g[colID, iX].Value = t.ID.ToString();

                if (dts.得意先.Any(a => a.ID == t.得意先ID))
                {
                    var cn = dts.得意先.Single(a => a.ID == t.得意先ID);
                    g[colClient, iX].Value = cn.略称;
                    g[colFuri, iX].Value = cn.フリガナ;
                }
                else
                {
                    g[colClient, iX].Value = string.Empty;
                    g[colFuri, iX].Value = string.Empty;
                }
                
                g[colKingaku, iX].Value = t.売上金額.ToString("#,##0");
                g[colTax, iX].Value = t.消費税.ToString("#,##0");
                g[colSDt, iX].Value = t.支払期日.ToShortDateString();
                g[colNDt, iX].Value = getNyukinDate(t.ID);

                if (t.入金完了 == global.FLGON)
                {
                    g[colSel2, iX].Value = true;
                }
                else
                {
                    g[colSel2, iX].Value = false;
                }

                g[colNyukin, iX].Value = getNyukingaku(t.ID).ToString("#,##0");
                g[colZan, iX].Value = t.残金.ToString("#,##0");
                g[colSeisan, iX].Value = t.精算額.ToString("n0");   // 2016/06/16
                g[colSeisanTl, iX].Value = (t.残金 + t.精算額).ToString("n0");   // 2016/06/16

                // 口座 2017/08/10
                if (t.口座 == null)
                {
                    g[colKouza, iX].Value = string.Empty;
                }
                else
                {
                    g[colKouza, iX].Value = t.口座;
                }

                if (t.無効 == global.FLGON)
                {
                    g[colMukou, iX].Value = true;
                }
                else
                {
                    g[colMukou, iX].Value = false;
                }

                if (t.Is備考Null())
                {
                    g[colMemo, iX].Value = string.Empty;
                }
                else
                {
                    g[colMemo, iX].Value = t.備考;
                }

                // 2016/05/12
                if (t.得意先Row == null)
                {
                    g[colClientID, iX].Value = string.Empty;
                }
                else
                {
                    g[colClientID, iX].Value = t.得意先Row.ID.ToString();
                }

                // 2016/09/16
                if (t.得意先Row == null)
                {
                    g[colTantou, iX].Value = string.Empty;
                }
                else if (t.得意先Row.社員Row == null)
                {
                    g[colTantou, iX].Value = string.Empty;
                }
                else
                {
                    g[colTantou, iX].Value = t.得意先Row.社員Row.氏名;
                }

                // 前受金 2017/05/24
                if (isMaeuke(t.ID))
                {
                    g[colMaeuke, iX].Value = global.FLGON.ToString();
                }
                else
                {
                    g[colMaeuke, iX].Value = global.FLGOFF.ToString();
                }

                iX++;
            }
            
            g.CurrentCell = null;

            if (iX > 0)
            {
                this.Text = "入金管理  " + iX.ToString("#,##0") + "件";
            }
            else
            {
                this.Text = "入金管理";
                MessageBox.Show("対象となるデータはありませんでした","検索結果",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            }
        }

        private bool isMaeuke(int sID)
        {
            bool rtn = false;

            foreach (var t in dts.受注1.Where(a => a.請求書ID == sID))
            {
                // 前受金判定
                if (t.入金予定日 < t.配布開始日)
                {
                    rtn = true;
                    break;
                }

            }

            return rtn;
        }



        private void button1_Click(object sender, EventArgs e)
        {

        }

        /// ----------------------------------------------------------------------------------
        /// <summary>
        ///     任意の新請求書データの入金日を求める </summary>
        /// <param name="sID">
        ///     新請求書ID</param>
        /// <returns>
        ///     入金日</returns>
        /// ----------------------------------------------------------------------------------
        private string getNyukinDate(int sID)
        {
            string dt = string.Empty;

            foreach (var t in dts.新入金.Where(a => a.請求書ID == sID).OrderByDescending(a => a.入金年月日))
            {
                dt = t.入金年月日.ToShortDateString();
                break;
            }

            return dt;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     任意の新請求書データの入金額を求める </summary>
        /// <param name="sID">
        ///     新請求書ID</param>
        /// <returns>
        ///     入金額</returns>
        ///------------------------------------------------------------------------------------
        private int getNyukingaku(int sID)
        {
            int dt = 0;

            var s = dts.新入金
               .Where(a => a.請求書ID == sID)
                   .GroupBy(a => a.請求書ID)
                   .Select(cg => new
                   {
                       cCode = cg.Key,
                       kingaku = cg.Sum(a => a.金額)
                   });

            foreach (var t in s)
            {
                dt = t.kingaku;
                break;
            }

            return dt;
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
            //    darwinDataSet.受注番号採番Row r = dts.受注番号採番.New受注番号採番Row();
            //    r.受注番号 =  Int64.Parse(lblOrderNum.Text);
            //    r.入庫日 = DateTime.Parse(dtNyuko.Value.ToShortDateString());

            //    if (cmbClient.SelectedIndex == -1)
            //    {
            //        r.得意先ID = 0;
            //    }
            //    else
            //    {
            //        Utility.ComboClient cmb = (Utility.ComboClient)cmbClient.SelectedItem;
            //        r.得意先ID = cmb.ID;
            //    }

            //    r.確定書入力 = 0;
            //    r.確定書入力日付 = DateTime.Parse("1900/01/01");
            //    r.確定書入力ユーザーID = 0;
            //    r.備考 = txtMemo.Text;
            //    r.登録年月日 = DateTime.Now;
            //    r.更新年月日 = DateTime.Now;
            //    r.ユーザーID = global.loginUserID;

            //    dts.受注番号採番.Add受注番号採番Row(r);
            //}
            
            //// データベース更新
            //adp.Update(dts.受注番号採番);

            //// データ読み込み
            //adp.Fill(dts.受注番号採番);
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
            //dispClear();
        }
        
        
        private void button5_Click(object sender, EventArgs e)
        {
            // 請求書データ表示
            mukouStatus = false;
            dataShow(0, 0);
            mukouStatus = true;
        }

        /// -------------------------------------------------------
        /// <summary>
        ///     請求書データ表示 </summary>
        /// -------------------------------------------------------
        private void dataShow(int col, int row)
        {
            // 支払期日範囲
            string dt = "1900/01/01";
            string edt = "2900/12/31";

            // 入金日範囲 : 2017/08/14
            string ndt = "1900/01/01";
            string nedt = "2900/12/31";

            this.Cursor = Cursors.WaitCursor;

            // 支払期日範囲
            if (dtNyuko.Checked)
            {
                dt = dtNyuko.Value.ToShortDateString();
            }

            if (dtNyuko2.Checked)
            {
                edt = dtNyuko2.Value.ToShortDateString();
            }

            // 入金日範囲 : 2017/08/14
            if (dtNyukin.Checked)
            {
                ndt = dtNyukin.Value.ToShortDateString();
            }

            if (dtNyukin2.Checked)
            {
                nedt = dtNyukin2.Value.ToShortDateString();
            }

            // データ表示
            gridShow(dataGridView1, DateTime.Parse(dt), DateTime.Parse(edt), DateTime.Parse(ndt), DateTime.Parse(nedt));

            if (dataGridView1.RowCount > 0)
            {
                dataGridView1.CurrentCell = dataGridView1[col, row];
                dataGridView1.CurrentCell = null;
                button2.Enabled = true;
                button1.Enabled = true; // 2017/08/14
            }
            else
            {
                button2.Enabled = false;
                button1.Enabled = false; // 2017/08/14
            }

            this.Cursor = Cursors.Default;
        }
        
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // ID取得
            int sID = Utility.strToInt(dataGridView1[colID, e.RowIndex].Value.ToString());

            // 詳細ボタン
            if (e.ColumnIndex == 0)
            {
                frmNyukinItem2015 frm = new frmNyukinItem2015(sID);
                frm.ShowDialog();
                bool status = frm.dataProperty;     // 2017/08/15
                frm.Dispose(); 

                // 2017/08/15
                if (status)
                {
                    // 新入金データ更新
                    nAdp.Fill(dts.新入金);

                    // 請求書データ再表示
                    jAdp.Fill(dts.新請求書);
                    dataShow(0, e.RowIndex);
                }
            }

            //// 無効チェック
            //if (e.ColumnIndex == 9)
            //{
            //    // データ取得
            //    var s = dts.新請求書.Single(a => a.ID == sID);

            //    if (dataGridView1[colMukou, e.RowIndex].Value.ToString() == "True")
            //    {
            //        s.無効 = global.FLGON;
            //    }
            //    else
            //    {
            //        s.無効 = global.FLGOFF;
            //    }

            //    s.変更年月日 = DateTime.Now;
            //    s.ユーザーID = global.loginUserID;
            //}

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (mukouStatus)
            {
                // 入金済み、無効チェック、備考
                if (e.ColumnIndex == 8 || e.ColumnIndex == 14 || e.ColumnIndex == 15)
                {
                    // ID取得
                    int sID = Utility.strToInt(dataGridView1[colID, e.RowIndex].Value.ToString());

                    // データ取得
                    var s = dts.新請求書.Single(a => a.ID == sID);

                    // 入金済チェック
                    if (e.ColumnIndex == 8)
                    {
                        if (dataGridView1[colSel2, e.RowIndex].Value.ToString() == "True")
                        {
                            s.入金完了 = global.FLGON;
                        }
                        else
                        {
                            s.入金完了 = global.FLGOFF;
                        }
                    }

                    // 無効チェック
                    if (e.ColumnIndex == 14)
                    {
                        if (dataGridView1[colMukou, e.RowIndex].Value.ToString() == "True")
                        {
                            s.無効 = global.FLGON;
                        }
                        else
                        {
                            s.無効 = global.FLGOFF;
                        }
                    }

                    // 備考
                    if (e.ColumnIndex == 15)
                    {
                        if (dataGridView1[colMemo, e.RowIndex].Value == null)
                        {
                            s.備考 = string.Empty;
                        }
                        else
                        {
                            s.備考 = dataGridView1[colMemo, e.RowIndex].Value.ToString();
                        }
                    }


                    s.変更年月日 = DateTime.Now;
                    s.ユーザーID = global.loginUserID;

                    // 請求書データ再表示
                    jAdp.Update(dts.新請求書);
                    //jAdp.Fill(dts.新請求書);
                }

                //dataShow();
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCellAddress.X == 8 || dataGridView1.CurrentCellAddress.X == 14)
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "入金管理");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            putKaikeiCsv();
        }

        private void putKaikeiCsv()
        {
            //ダイアログボックスの初期設定
            SaveFileDialog sf = new SaveFileDialog();
            sf.Title = "会計CSV出力";
            sf.OverwritePrompt = true;
            sf.RestoreDirectory = true;
            sf.FileName = "勘定奉行汎用データ_入金";
            sf.Filter = "Microsoft Office Excelファイル(*.csv)|*.csv";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName = string.Empty;
            DialogResult ret = sf.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                SaveData(sf.FileName, dataGridView1);
            }

        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     会計データ出力処理 </summary>
        /// <param name="fName">
        ///     出力ファイル名</param>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        ///---------------------------------------------------------------------
        private void SaveData(string fName, DataGridView g)
        {
            string wrkOutputData = string.Empty;
            Boolean iniFlg = true;
            Boolean pblFirstGyouFlg = true;

            //出力ファイルインスタンス作成
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(fName, false, System.Text.Encoding.GetEncoding(932));

            try
            {
                // 表示中データを読み出す
                for (int i = 0; i < g.RowCount; i++)
                {
                    //// 未入金データはネグる
                    //if (Utility.nullToStr(g[colNDt, i].Value) == string.Empty)
                    //{
                    //    continue;
                    //}

                    //// 無効は対象外
                    //if (g[colMukou, i].Value.ToString() == "True")
                    //{
                    //    continue;
                    //}

                    //ヘッダファイル出力
                    if (pblFirstGyouFlg)
                    {
                        wrkOutputData = string.Empty;
                        wrkOutputData += Entity.OutPutHeader.dn01 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd00 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd01 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd02 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd03 + ",";

                        wrkOutputData += Entity.OutPutHeader.kr01 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr02 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr03 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr04 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr05 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr06 + ",";

                        wrkOutputData += Entity.OutPutHeader.ks01 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks02 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks03 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks04 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks05 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks06 + ",";

                        wrkOutputData += Entity.OutPutHeader.tk01;

                        outFile.WriteLine(wrkOutputData);
                    }
                    
                    //出力データ作成
                    //wrkOutputData = SetDataNew(iX, i, st, pblFirstGyouFlg, OutData);
                    if (pblFirstGyouFlg)
                    {
                        wrkOutputData = "*,";
                    }
                    else
                    {
                        wrkOutputData = ",";
                    }

                    wrkOutputData += "1,";      // 部門指定方法
                    wrkOutputData += "0,";      // 伝票部門コード
                    wrkOutputData += g[colNDt, i].Value.ToString() + ",";  // 入金日
                    wrkOutputData += ",";       // 伝票番号
                    wrkOutputData += "20,";     // 借方部門コード
                    wrkOutputData += "111,";    // 借方勘定科目コード
                    wrkOutputData += g[colKouza, i].Value.ToString() + ",";     // 借方補助科目コード(金融機関）
                    wrkOutputData += "4,";      // 借方事業区分コード
                    wrkOutputData += "2,";      // 借方端数処理
                    wrkOutputData += Utility.strToInt(g[colNyukin, i].Value.ToString()) + ",";    // 借方本体金額
                    wrkOutputData += "20,";     // 貸方部門コード
                    wrkOutputData += "135,";    // 貸方勘定科目コード
                    wrkOutputData += "4,";      // 貸方事業区分コード
                    wrkOutputData += "2,";      // 貸方端数処理
                    wrkOutputData += (Utility.strToInt(g[colClientID, i].Value.ToString()) + 990000000) + ",";      // 取引先コード
                    wrkOutputData += Utility.strToInt(g[colNyukin, i].Value.ToString()) + ",";    // 貸方本体金額
                    wrkOutputData += g[colFuri, i].Value.ToString().Replace(",","");              // 摘要

                    // ファイルへ出力            
                    outFile.WriteLine(wrkOutputData);
                    pblFirstGyouFlg = false;
                }
                
                //ファイルクローズ
                outFile.Close();

                // 終了メッセージ
                MessageBox.Show("作成終了", "勘定奉行汎用データ", MessageBoxButtons.OK);

                ////出力ファイル削除
                //utility.FileDelete(global.WorkDir + global.DIR_OK, global.OUTFILE);

                ////一時ファイルを出力ファイルにコピー
                //File.Copy(global.WorkDir + global.DIR_OK + global.tmpFile, global.WorkDir + global.DIR_OK + global.OUTFILE);

                ////一時ファイル削除
                //utility.FileDelete(global.WorkDir + global.DIR_OK, global.tmpFile);
            }
            catch (Exception e)
            {
                MessageBox.Show("データ変換中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }
    }
}

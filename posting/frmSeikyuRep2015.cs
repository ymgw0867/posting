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
    public partial class frmSeikyuRep2015 : Form
    {
        public frmSeikyuRep2015()
        {
            InitializeComponent();

            // データ読み込み
            jAdp.Fill(dts.新請求書);
            cAdp.Fill(dts.得意先);
            sAdp.Fill(dts.社員);
            aAdp.Fill(dts.受注1);
            nAdp.Fill(dts.新入金);
            pAdp.Fill(dts.受注種別);
            zAdp.Fill(dts.判型);
            kAdp.Fill(dts.配布形態);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.新請求書TableAdapter jAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter aAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.新入金TableAdapter nAdp = new darwinDataSetTableAdapters.新入金TableAdapter();
        darwinDataSetTableAdapters.受注種別TableAdapter pAdp = new darwinDataSetTableAdapters.受注種別TableAdapter();
        darwinDataSetTableAdapters.判型TableAdapter zAdp = new darwinDataSetTableAdapters.判型TableAdapter();
        darwinDataSetTableAdapters.配布形態TableAdapter kAdp = new darwinDataSetTableAdapters.配布形態TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

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
        string colNebiki = "col15";     // 値引
        string colZan = "col16";        // 請求残
        string colUserID = "col11";
        string colClientID = "col17";   // クライアントID
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
            //dispClear();
            linkLabel1.Tag = CHKOFF;
            linkLabel1.Text = "全てチェック";
            linkLabel1.Visible = false;
            dtNyuko.Checked = false;

            button1.Enabled = false;
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
                DataGridViewCheckBoxColumn cbc = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc);
                tempDGV.Columns[0].HeaderText = "";
                tempDGV.Columns[0].Name = colSel;

                tempDGV.Columns.Add(colID, "No.");
                tempDGV.Columns.Add(colClient, "社名");
                tempDGV.Columns.Add(colKingaku, "小計");
                tempDGV.Columns.Add(colNebiki, "値引");
                tempDGV.Columns.Add(colTax, "消費税");
                tempDGV.Columns.Add(colZan, "請求残");
                tempDGV.Columns.Add(colDt, "請求書発行日");
                tempDGV.Columns.Add(colSDt, "支払予定日");

                DataGridViewCheckBoxColumn cbc2 = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc2);
                tempDGV.Columns[9].HeaderText = "発行済";
                tempDGV.Columns[9].Name = colSel2;

                tempDGV.Columns.Add(colEtan, "営業担当者");
                tempDGV.Columns.Add(colCnt, "案件数");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colClientID, "CID");

                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "詳細";
                tempDGV.Columns.Add(btn);
                tempDGV.Columns[14].HeaderText = "";
                tempDGV.Columns[14].Name = colBtn;

                // 各列幅指定
                tempDGV.Columns[colSel].Width = 50;
                tempDGV.Columns[colID].Width = 60;
                tempDGV.Columns[colClient].Width = 200;
                tempDGV.Columns[colKingaku].Width = 100;
                tempDGV.Columns[colNebiki].Width = 80;
                tempDGV.Columns[colTax].Width = 80;
                tempDGV.Columns[colZan].Width = 100;
                tempDGV.Columns[colDt].Width = 110;
                tempDGV.Columns[colSDt].Width = 110;
                tempDGV.Columns[colSel2].Width = 60;
                tempDGV.Columns[colCnt].Width = 70;
                tempDGV.Columns[colBtn].Width = 60;
                tempDGV.Columns[colMemo].Width = 300;
                tempDGV.Columns[colClientID].Width = 60;

                tempDGV.Columns[colID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colNebiki].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colTax].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colZan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colSDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colCnt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colClientID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列固定
                tempDGV.Columns[colClient].Frozen = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                //tempDGV.ReadOnly = true;
                foreach (DataGridViewColumn d in tempDGV.Columns)
                {
                    if (d.Name == colSel)
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
        private void gridShow(DataGridView g, DateTime dt)
        {
            g.Rows.Clear();
            int iX = 0;

            Cursor = Cursors.WaitCursor;

            var s = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0).OrderBy(a => a.請求書発行日).ThenBy(a => a.ID);

            // 請求残あり
            if (checkBox2.Checked)
            {
                s = s.Where(a => a.入金完了 == 0).OrderBy(a => a.請求書発行日).ThenBy(a => a.ID);
            }

            if (dt.ToShortDateString() != "1900/01/01")
            {
                s = s.Where(a => a.請求書発行日 == dt).OrderBy(a => a.請求書発行日).ThenBy(a => a.ID);
            }

            // 請求残あり
            if (checkBox1.Checked)
            {
                s = s.Where(a => a.残金 != 0).OrderBy(a => a.請求書発行日).ThenBy(a => a.ID);
            }

            // クライアントID指定の時
            if (txtCID.Text != string.Empty)
            {
                s = s.Where(a => a.得意先ID == Utility.strToInt(txtCID.Text)).OrderBy(a => a.請求書発行日).ThenBy(a => a.ID);
            }

            foreach (var t in s)
            {
                // クライアント名指定のとき
                if (txtClient.Text != string.Empty)
                {
                    if (t.得意先Row == null)
                    {
                        continue;
                    }

                    if (!t.得意先Row.略称.Contains(txtClient.Text.Trim()))
                    {
                        continue;
                    }
                }

                // 営業担当者指定の時
                if (txtEtantou.Text != string.Empty)
                {
                    if (t.得意先Row == null)
                    {
                        continue;
                    }

                    if (t.得意先Row.社員Row == null)
                    {
                        continue;
                    }

                    if (!t.得意先Row.社員Row.氏名.Contains(txtEtantou.Text.Trim()))
                    {
                        continue;
                    }
                }

                g.Rows.Add();

                g[colSel, iX].Value = false;
                g[colID, iX].Value = t.ID.ToString();

                if (dts.得意先.Any(a => a.ID == t.得意先ID))
                {
                    var cn = dts.得意先.Single(a => a.ID == t.得意先ID);
                    g[colClient, iX].Value = cn.略称;

                    // 営業担当者
                    if (cn.社員Row != null)
                    {
                        g[colEtan, iX].Value = cn.社員Row.氏名;
                    }
                    else
                    {
                        g[colEtan, iX].Value = string.Empty;
                    }
                }
                else
                {
                    g[colClient, iX].Value = string.Empty;
                    g[colEtan, iX].Value = string.Empty;
                }

                g[colKingaku, iX].Value = t.売上金額.ToString("#,##0");
                g[colNebiki, iX].Value = t.値引額.ToString("#,##0");
                g[colTax, iX].Value = t.消費税.ToString("#,##0");
                g[colZan, iX].Value = t.残金.ToString("#,##0");
                g[colDt, iX].Value = t.請求書発行日.ToShortDateString();
                g[colSDt, iX].Value = t.支払期日.ToShortDateString();
                g[colSel2, iX].Value = t.請求書発行済;
                g[colCnt, iX].Value = t.明細数.ToString();

                if (!t.Is備考Null())
                {
                    g[colMemo, iX].Value = t.備考;
                }
                else
                {
                    g[colMemo, iX].Value = string.Empty;
                }

                g[colClientID, iX].Value = t.得意先ID.ToString();

                iX++;
            }
            
            g.CurrentCell = null;

            if (iX > 0)
            {
                this.Text = "請求一覧  " + iX.ToString("#,##0") + "件";
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
                MessageBox.Show("該当するデータは存在しませんでした","データなし",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }

            Cursor = Cursors.Default;
        }

        private void button1_Click(object sender, EventArgs e)
        {

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
            dataShow(0, 0);
        }

        /// -------------------------------------------------------
        /// <summary>
        ///     請求書データ表示 </summary>
        /// -------------------------------------------------------
        private void dataShow(int col, int row)
        {
            string dt = "1900/01/01";

            if (dtNyuko.Checked)
            {
                dt = dtNyuko.Value.ToShortDateString();
            }

            // データ表示
            gridShow(dataGridView1, DateTime.Parse(dt));

            if (dataGridView1.RowCount > 0)
            {
                dataGridView1.CurrentCell = dataGridView1[col, row];
                dataGridView1.CurrentCell = null;
                linkLabel1.Visible = true;
            }
            else
            {
                linkLabel1.Visible = false;
            }
        }
        

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (linkLabel1.Tag.ToString() == CHKOFF)
            {
                // チェックオン
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1[colSel, i].Value = true;
                }

                // 表示変更
                linkLabel1.Text = "全てチェックオフ";
                linkLabel1.Tag = CHKON;
            }
            else
            {
                // チェックオフ
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1[colSel, i].Value = false;
                }

                // 表示変更
                linkLabel1.Text = "全てチェック";
                linkLabel1.Tag = CHKOFF;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int pCnt = 0;

            // チェック件数カウント
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[colSel, i].Value.ToString() == "True")
                {
                    pCnt++;
                }
            }

            if (pCnt == 0)
            {
                MessageBox.Show("請求書を発行するデータがチェックされていません","対象データなし",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                if (MessageBox.Show(pCnt.ToString() + "件の請求書を発行します。よろしいですか？", "発行確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            // 請求書発行
            seikyuRep();
        }
        
        /// ---------------------------------------------------------
        /// <summary>
        ///     請求書発行 </summary>
        /// ---------------------------------------------------------
        private void seikyuRep()
        {
            int _Pages = 0;
            int _Rows = 0;
            int sIr = 26;
            int _RowsPrn = 21;

            decimal _nebikigo = 0;      // 売上金額 - 値引額
            decimal _tax = 0;           // 消費税
            decimal _seikyu = 0;        // 請求金額

            bool openStatus = true;

            string seikyuNum = string.Empty;    // 請求書番号
            string corp = string.Empty;         // 請求先

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                // 請求書テンプレートBOOK
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル請求書, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                // 印刷用BOOK
                Excel.Workbook oXlsBookPrn = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル受注確定書印刷, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                // 請求書テンプレートシート
                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                // 請求書印刷シート
                Excel.Worksheet oxlsSheetPrn = null;

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        // チェック行
                        if (dataGridView1[colSel, i].Value.ToString() == "True")
                        {
                            int iX = 0;

                            // 受注データ取得
                            foreach (var it in dts.新請求書.Where(a => a.ID == Utility.strToInt(dataGridView1[colID, i].Value.ToString())))
                            {
                                // 明細数
                                foreach (var t in it.Get受注1Rows().OrderBy(a => a.配布開始日))
                                {
                                    if (iX == 0 || _Rows >= _RowsPrn)
                                    {
                                        // ページ計印字
                                        if (!openStatus)
                                        {
                                            oxlsSheetPrn.Cells[47, 16] = _nebikigo.ToString();  // 小計
                                            oxlsSheetPrn.Cells[48, 16] = _tax.ToString();       // 消費税
                                            oxlsSheetPrn.Cells[49, 16] = _seikyu.ToString();    // 合計

                                            // 備考
                                            if (it.Is備考Null())
                                            {
                                                oxlsSheetPrn.Cells[47, 4] = string.Empty;
                                            }
                                            else
                                            {
                                                oxlsSheetPrn.Cells[47, 4] = it.備考;
                                            }

                                            // ページ計初期化
                                            _nebikigo = 0;
                                            _tax = 0;
                                            _seikyu = 0;
                                        }

                                        // ページ数加算
                                        iX++;
                                        _Pages = iX;

                                        // ページシートを追加する
                                        oxlsSheet.Copy(Type.Missing, oXlsBookPrn.Sheets[_Pages]);

                                        // 印刷用カレントシート
                                        oxlsSheetPrn = (Excel.Worksheet)oXlsBookPrn.Sheets[_Pages + 1];

                                        // 対象行をゼロにしページを初期化
                                        _Rows = 0;

                                        // 請求書№
                                        seikyuNum = t.請求書発行日.Year.ToString() + t.請求書発行日.Month.ToString().PadLeft(2, '0') + t.請求書ID.ToString().PadLeft(5, '0');
                                        oxlsSheetPrn.Cells[1, 16] = "№ " + seikyuNum;

                                        // 発行日
                                        oxlsSheetPrn.Cells[2, 15] = t.請求書発行日.ToShortDateString();

                                        //// ページ数
                                        //oxlsSheetPrn.Cells[3, 16] = "(" + iX.ToString() + "/" + it.明細数.ToString() + ")";

                                        corp = string.Empty;

                                        // 請求先住所
                                        if (it.得意先Row != null)
                                        {
                                            oxlsSheetPrn.Cells[2, 3] = "〒 " + it.得意先Row.請求先郵便番号;
                                            oxlsSheetPrn.Cells[3, 3] = it.得意先Row.請求先都道府県 + " " + it.得意先Row.請求先住所1;
                                            oxlsSheetPrn.Cells[4, 3] = it.得意先Row.請求先住所2;

                                            // 請求先名
                                            //if (it.得意先Row.Is請求先名称Null())
                                            //{
                                            //    corp = it.得意先Row.略称;
                                            //}
                                            //else
                                            //{
                                            //    corp = it.得意先Row.請求先名称;
                                            //}

                                            // 2019/02/21
                                            if (it.得意先Row.請求先名称 == null)
                                            {
                                                corp = it.得意先Row.略称;
                                            }
                                            else
                                            {
                                                corp = it.得意先Row.請求先名称;
                                            }

                                            // 2019/02/21 コメント化
                                            //// 部署・担当者名 
                                            //string tn = (it.得意先Row.部署名 + " " + it.得意先Row.請求先担当者名).Trim();

                                            //if (tn != string.Empty)
                                            //{
                                            //    oxlsSheetPrn.Cells[8, 3] = tn + " 様";
                                            //    oxlsSheetPrn.Cells[6, 3] = corp;
                                            //}
                                            //else
                                            //{
                                            //    oxlsSheetPrn.Cells[8, 3] = string.Empty;
                                            //    oxlsSheetPrn.Cells[6, 3] = corp + " 御中";
                                            //}

                                            // 部署・担当者名（得意先@請求先部署名、得意先@請求先敬称を使用）: 2019/02/21
                                            string tn = (it.得意先Row.請求先部署名 + " " + it.得意先Row.請求先担当者名).Trim();

                                            if (tn != string.Empty)
                                            {
                                                oxlsSheetPrn.Cells[8, 3] = tn + " " + it.得意先Row.請求先敬称;
                                                oxlsSheetPrn.Cells[6, 3] = corp;
                                            }
                                            else
                                            {
                                                oxlsSheetPrn.Cells[8, 3] = string.Empty;
                                                oxlsSheetPrn.Cells[6, 3] = corp + " " + it.得意先Row.請求先敬称;
                                            }
                                        }

                                        // 1ページ目
                                        if (_Pages == 1)
                                        {
                                            // 合計請求金額
                                            //oxlsSheetPrn.Cells[22, 5] = dataGridView1[colKingaku, i].Value.ToString().Trim().Replace(",", "");
                                            //oxlsSheetPrn.Cells[22, 5] = it.請求金額.ToString();
                                            oxlsSheetPrn.Cells[22, 5] = it.残金.ToString();   // 残金を請求金額とする 2015/10/20
                                        }
                                        else
                                        {
                                            // 合計請求金額
                                            oxlsSheetPrn.Cells[22, 5] = "********************";
                                        }

                                        // 支払期日
                                        oxlsSheetPrn.Cells[55, 4] = dataGridView1[colSDt, i].Value.ToString();

                                        // 開始ステータス
                                        openStatus = false;
                                    }

                                    // 明細内容
                                    //oxlsSheetPrn.Cells[sIr + _Rows, 1] = t.受注日.Month.ToString() + "." + t.受注日.Day.ToString().PadLeft(2, '0');    // 日付
                                    oxlsSheetPrn.Cells[sIr + _Rows, 1] = t.配布開始日.Month.ToString() + "." + (t.配布開始日.Day.ToString().PadLeft(2, '0'));    // 日付

                                    // 受注内容
                                    if (t.受注種別Row != null)
                                    {
                                        oxlsSheetPrn.Cells[sIr + _Rows, 3] = t.受注種別Row.名称;
                                    }
                                    else
                                    {
                                        oxlsSheetPrn.Cells[sIr + _Rows, 3] = string.Empty;
                                    }
                                           
                                    // サイズ
                                    if (t.判型Row != null)
                                    {
                                        oxlsSheetPrn.Cells[sIr + _Rows, 5] = t.判型Row.名称;
                                    }
                                    else
                                    {
                                        oxlsSheetPrn.Cells[sIr + _Rows, 5] = string.Empty;
                                    }

                                    // チラシ名
                                    oxlsSheetPrn.Cells[sIr + _Rows, 6] = t.チラシ名;

                                    // 配布形態
                                    if (t.配布形態Row != null)
                                    {
                                        oxlsSheetPrn.Cells[sIr + _Rows, 12] = t.配布形態Row.名称;
                                    }
                                    else
                                    {
                                        oxlsSheetPrn.Cells[sIr + _Rows, 12] = string.Empty;
                                    }
   
                                    oxlsSheetPrn.Cells[sIr + _Rows, 14] = t.単価.ToString("n2");
                                    oxlsSheetPrn.Cells[sIr + _Rows, 15] = t.枚数.ToString();
                                    oxlsSheetPrn.Cells[sIr + _Rows, 16] = t.金額.ToString();

                                    // 行加算
                                    _Rows++;

                                    // 値引額があるとき値引行を印字します
                                    if (t.値引額 > 0)
                                    {
                                        // 明細内容
                                        oxlsSheetPrn.Cells[sIr + _Rows, 6] = "値引";
                                        oxlsSheetPrn.Cells[sIr + _Rows, 14] = (t.値引額 * (-1)).ToString("#,0");
                                        oxlsSheetPrn.Cells[sIr + _Rows, 15] = "1";
                                        oxlsSheetPrn.Cells[sIr + _Rows, 16] = (t.値引額 * (-1)).ToString();

                                        // 行加算
                                        _Rows++;
                                    }

                                    // ページ計加算
                                    _nebikigo += (t.金額 - t.値引額);    // 小計
                                    _tax += t.消費税;                   // 消費税
                                    _seikyu += t.売上金額;              // 合計
                                }

                                // 入金額 2015/10/20
                                foreach (var nt in it.Get新入金Rows().OrderBy(a => a.入金年月日))
                                {
                                    // 明細内容
                                    oxlsSheetPrn.Cells[sIr + _Rows, 1] = nt.入金年月日.Month.ToString() + "." + (nt.入金年月日.Day.ToString().PadLeft(2, '0'));    // 日付
                                    oxlsSheetPrn.Cells[sIr + _Rows, 6] = "入金 ";
                                    oxlsSheetPrn.Cells[sIr + _Rows, 14] = (nt.金額 * (-1)).ToString("#,0");
                                    oxlsSheetPrn.Cells[sIr + _Rows, 15] = "1";
                                    oxlsSheetPrn.Cells[sIr + _Rows, 16] = (nt.金額 * (-1)).ToString();

                                    // 行加算
                                    _Rows++;

                                    // ページ計加算
                                    //_nebikigo -= nt.金額;     // 小計
                                }


                                // ページ計印字
                                oxlsSheetPrn.Cells[47, 16] = _nebikigo.ToString();  // 小計
                                oxlsSheetPrn.Cells[48, 16] = _tax.ToString();       // 消費税
                                oxlsSheetPrn.Cells[49, 16] = _seikyu.ToString();    // 合計

                                // 備考
                                if (it.Is備考Null())
                                {
                                    oxlsSheetPrn.Cells[47, 4] = string.Empty;
                                }
                                else
                                {
                                    oxlsSheetPrn.Cells[47, 4] = it.備考;
                                }

                                // 営業担当 2015/11/18
                                if (it.得意先Row != null && it.得意先Row.社員Row != null)
                                {
                                    oxlsSheetPrn.Cells[57, 15] = it.得意先Row.社員Row.氏名;
                                }
                                else
                                {
                                    oxlsSheetPrn.Cells[57, 15] = string.Empty;
                                }
                            }
                        }
                    }
                    
                    // 印刷用BOOKの1番目のシートは削除する
                    ((Excel.Worksheet)oXlsBookPrn.Sheets[1]).Delete();

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    //oXls.Visible = true;

                    // 印刷
                    oXlsBookPrn.PrintOutEx(1, Type.Missing, 1, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    // ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "請求書発行";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    //saveFileDialog1.FileName = "請求書_" + corp + seikyuNum;
                    saveFileDialog1.FileName = "請求書";
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    // ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBookPrn.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    // 処理終了メッセージ
                    MessageBox.Show("終了しました","請求書発行",MessageBoxButtons.OK,MessageBoxIcon.Information);

                    // 請求書発行フラグ更新
                    seikyuRepFlgUpdate();

                    // 請求書データ再表示
                    dataShow(0, 0);
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "請求書発行", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                finally
                {
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);
                    oXlsBookPrn.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheetPrn);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBookPrn);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "請求書作成", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     新請求書データ：請求書発行済みフラグ書き込み </summary>
        /// ------------------------------------------------------------------------
        private void seikyuRepFlgUpdate()
        {
            int cnt = 0;

            // 待機カーソル
            this.Cursor = Cursors.WaitCursor;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                // チェック行
                if (dataGridView1[colSel, i].Value.ToString() == "True")
                {
                    // 請求書ID
                    int sID = Utility.strToInt(dataGridView1[colID, i].Value.ToString());

                    if (dts.新請求書.Any(a => a.ID == sID))
                    {
                        var s = dts.新請求書.Single(a => a.ID == sID);
                        s.請求書発行済 = global.FLGON;
                        cnt++;
                    }
                }
            }

            if (cnt > 0)
            {
                // データベース更新
                jAdp.Update(dts.新請求書);
            }

            // カーソル戻す
            this.Cursor = Cursors.Default;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 詳細ボタン
            if (e.ColumnIndex == 14)
            {
                int sID = Utility.strToInt(dataGridView1[colID, e.RowIndex].Value.ToString());

                frmSeikyuItem2015 frm = new frmSeikyuItem2015(sID);
                frm.ShowDialog();

                // 請求書データ再表示
                jAdp.Fill(dts.新請求書);
                dataShow(0, e.RowIndex);
            }
        }

        private void txtCID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
    }
}

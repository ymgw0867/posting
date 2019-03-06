using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MyLibrary;

namespace posting
{
    public partial class frmOrderRecord : Form
    {
        public frmOrderRecord()
        {
            InitializeComponent();

            // データ読み込み
            jAdp.Fill(dts.受注確定書発行記録);
            uAdp.Fill(dts.ログインユーザー);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注確定書発行記録TableAdapter jAdp = new darwinDataSetTableAdapters.受注確定書発行記録TableAdapter();
        darwinDataSetTableAdapters.ログインユーザーTableAdapter uAdp = new darwinDataSetTableAdapters.ログインユーザーTableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        bool kCheck = false;

        #region グリッドビューカラム定義
        string colSel = "col0";         // チェック列
        string colID = "col1";          // ID
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
        string colUserID = "col11";
        string colName = "col16";       // 商品名
        string colUserID1 = "col17";    // 確認者１
        string colUserID2 = "col18";    // 確認者２
        string colUserID3 = "col19";    // 確認者３
        string colUserID4 = "col20";    // 確認者４
        string colUserID5 = "col21";    // 確認者５
        string colChkDt1 = "col22";     // 確認日１
        string colChkDt2 = "col23";     // 確認日２
        string colChkDt3 = "col24";     // 確認日３
        string colChkDt4 = "col25";     // 確認日４
        string colChkDt5 = "col26";     // 確認日５
        string colChk1 = "col27";       // 確認１
        string colChk2 = "col28";       // 確認２
        string colChk3 = "col29";       // 確認３
        string colChk4 = "col30";       // 確認４
        string colChk5 = "col31";       // 確認５
        #endregion

        const string CHKON = "1";
        const string CHKOFF = "0";

        DateTime nullDate = DateTime.Parse("1900/01/01");

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            //Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // ログインユーザーコンボボックスアイテムロード
            Utility.comboLoginUser.itemLoad(comboBox1);

            // データグリッドビューの定義
            gridSetting(dataGridView1);

            //// データグリッドビューデータ表示
            //gridShow(dataGridView1, DateTime.Parse(dtNyuko.Value.ToShortDateString()));

            // 画面初期化
            dtNyuko.Checked = false;
            comboBox1.SelectedIndex = -1;
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
                btn.Text = "表示";
                tempDGV.Columns.Add(btn);
                tempDGV.Columns[0].HeaderText = "";
                tempDGV.Columns[0].Name = colBtn;

                tempDGV.Columns.Add(colDt, "発行日時");
                tempDGV.Columns.Add(colClient, "クライアント名");
                tempDGV.Columns.Add(colName, "商品名");
                tempDGV.Columns.Add(colUserID, "発行者");

                DataGridViewCheckBoxColumn cbc1 = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc1);
                tempDGV.Columns[5].HeaderText = "確認１";
                tempDGV.Columns[5].Name = colChk1;

                tempDGV.Columns.Add(colChkDt1, "");
                tempDGV.Columns.Add(colUserID1, "");

                DataGridViewCheckBoxColumn cbc2 = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc2);
                tempDGV.Columns[8].HeaderText = "確認２";
                tempDGV.Columns[8].Name = colChk2;

                tempDGV.Columns.Add(colChkDt2, "");
                tempDGV.Columns.Add(colUserID2, "");

                DataGridViewCheckBoxColumn cbc3 = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc3);
                tempDGV.Columns[11].HeaderText = "確認３";
                tempDGV.Columns[11].Name = colChk3;

                tempDGV.Columns.Add(colChkDt3, "");
                tempDGV.Columns.Add(colUserID3, "");

                DataGridViewCheckBoxColumn cbc4 = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc4);
                tempDGV.Columns[14].HeaderText = "確認４";
                tempDGV.Columns[14].Name = colChk4;

                tempDGV.Columns.Add(colChkDt4, "");
                tempDGV.Columns.Add(colUserID4, "");

                DataGridViewCheckBoxColumn cbc5 = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(cbc5);
                tempDGV.Columns[17].HeaderText = "確認５";
                tempDGV.Columns[17].Name = colChk5;

                tempDGV.Columns.Add(colChkDt5, "");
                tempDGV.Columns.Add(colUserID5, "");

                tempDGV.Columns.Add(colMemo, "パス");
                tempDGV.Columns.Add(colID, "SID");


                // 各列幅指定
                tempDGV.Columns[colDt].Width = 140;
                tempDGV.Columns[colClient].Width = 200;
                tempDGV.Columns[colName].Width = 200;
                tempDGV.Columns[colUserID].Width = 100;

                tempDGV.Columns[colChk1].Width = 50;
                tempDGV.Columns[colChkDt1].Width = 140;
                tempDGV.Columns[colUserID1].Width = 100;

                tempDGV.Columns[colChk2].Width = 50;
                tempDGV.Columns[colChkDt2].Width = 140;
                tempDGV.Columns[colUserID2].Width = 100;

                tempDGV.Columns[colChk3].Width = 50;
                tempDGV.Columns[colChkDt3].Width = 140;
                tempDGV.Columns[colUserID3].Width = 100;

                tempDGV.Columns[colChk4].Width = 50;
                tempDGV.Columns[colChkDt4].Width = 140;
                tempDGV.Columns[colUserID4].Width = 100;

                tempDGV.Columns[colChk5].Width = 50;
                tempDGV.Columns[colChkDt5].Width = 140;
                tempDGV.Columns[colUserID5].Width = 100;

                tempDGV.Columns[colBtn].Width = 50;
                tempDGV.Columns[colMemo].Visible = false;
                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[colDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChk1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChkDt1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChk2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChkDt2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChk3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChkDt3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChk4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChkDt4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChk5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colChkDt5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
                    if (d.Name == colChk1 || d.Name == colChk2 || d.Name == colChk3 || d.Name == colChk4 || d.Name == colChk5)
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
        private void gridShow(DataGridView g)
        {
            g.Rows.Clear();
            int iX = 0;

            var s = dts.受注確定書発行記録.Where(a => a.ID > 0).OrderByDescending(a => a.発行日);

            // 発行日
            if (dtNyuko.Checked)
            {
                s = s.Where(a => a.発行日.Date == DateTime.Parse(dtNyuko.Value.ToShortDateString())).OrderByDescending(a => a.発行日);
            }

            // 発行者
            if (comboBox1.Text != string.Empty)
            {
                int sss = int.Parse(comboBox1.SelectedValue.ToString());
                s = s.Where(a => a.発行者 == sss).OrderByDescending(a => a.発行日);
            }

            foreach (var t in s)
            {
                // クライアント名指定のとき
                if (txtClient.Text != string.Empty)
                {
                    if (!t.クライアント名.Contains(txtClient.Text.Trim()))
                    {
                        continue;
                    }
                }

                g.Rows.Add();

                g[colDt, iX].Value = t.発行日;
                g[colClient, iX].Value = t.クライアント名;
                g[colName, iX].Value = t.商品名;

                if (t.ログインユーザーRowByログインユーザー_受注確定書発行記録 != null)
                {
                    g[colUserID, iX].Value = t.ログインユーザーRowByログインユーザー_受注確定書発行記録.ログインユーザー;
                }
                else
                {
                    g[colUserID, iX].Value = string.Empty;
                }

                // 確認１
                g[colChk1, iX].Value = t.確認1;

                if (t.確認1 == global.FLGON)
                {
                    g[colChkDt1, iX].Value = t.確認日1;

                    if (t.ログインユーザーRowByログインユーザー_受注確定書発行記録1 != null)
                    {
                        g[colUserID1, iX].Value = t.ログインユーザーRowByログインユーザー_受注確定書発行記録1.ログインユーザー;
                    }
                    else
                    {
                        g[colUserID1, iX].Value = string.Empty;
                    }
                }
                else
                {
                    g[colChkDt1, iX].Value = string.Empty;
                    g[colUserID1, iX].Value = string.Empty;
                }

                // 確認２
                g[colChk2, iX].Value = t.確認2;

                if (t.確認2 == global.FLGON)
                {
                    g[colChkDt2, iX].Value = t.確認日2;

                    if (t.ログインユーザーRowByログインユーザー_受注確定書発行記録2 != null)
                    {
                        g[colUserID2, iX].Value = t.ログインユーザーRowByログインユーザー_受注確定書発行記録2.ログインユーザー;
                    }
                    else
                    {
                        g[colUserID2, iX].Value = string.Empty;
                    }
                }
                else
                {
                    g[colChkDt2, iX].Value = string.Empty;
                    g[colUserID2, iX].Value = string.Empty;
                }

                // 確認３
                g[colChk3, iX].Value = t.確認3;

                if (t.確認3 == global.FLGON)
                {
                    g[colChkDt3, iX].Value = t.確認日3;

                    if (t.ログインユーザーRowByログインユーザー_受注確定書発行記録3 != null)
                    {
                        g[colUserID3, iX].Value = t.ログインユーザーRowByログインユーザー_受注確定書発行記録3.ログインユーザー;
                    }
                    else
                    {
                        g[colUserID3, iX].Value = string.Empty;
                    }
                }
                else
                {
                    g[colChkDt3, iX].Value = string.Empty;
                    g[colUserID3, iX].Value = string.Empty;
                }

                // 確認４
                g[colChk4, iX].Value = t.確認4;

                if (t.確認4 == global.FLGON)
                {
                    g[colChkDt4, iX].Value = t.確認日4;

                    if (t.ログインユーザーRowByログインユーザー_受注確定書発行記録4 != null)
                    {
                        g[colUserID4, iX].Value = t.ログインユーザーRowByログインユーザー_受注確定書発行記録4.ログインユーザー;
                    }
                    else
                    {
                        g[colUserID4, iX].Value = string.Empty;
                    }
                }
                else
                {
                    g[colChkDt4, iX].Value = string.Empty;
                    g[colUserID4, iX].Value = string.Empty;
                }

                // 確認５
                g[colChk5, iX].Value = t.確認5;

                if (t.確認5 == global.FLGON)
                {
                    g[colChkDt5, iX].Value = t.確認日5;

                    if (t.ログインユーザーRowByログインユーザー_受注確定書発行記録5 != null)
                    {
                        g[colUserID5, iX].Value = t.ログインユーザーRowByログインユーザー_受注確定書発行記録5.ログインユーザー;
                    }
                    else
                    {
                        g[colUserID5, iX].Value = string.Empty;
                    }
                }
                else
                {
                    g[colChkDt5, iX].Value = string.Empty;
                    g[colUserID5, iX].Value = string.Empty;
                }

                g[colMemo, iX].Value = t.受注確定書パス;     // 受注確定書パス
                g[colID, iX].Value = t.ID;                   // ID

                iX++;
            }

            if (iX > 0)
            {
                this.Text = "受注確定書発行記録  " + iX.ToString("#,##0") + "件";
                g.CurrentCell = null;
                kCheck = true;
            }
            else
            {
                MessageBox.Show("該当するデータはありませんでした", "受注確定書発行記録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                kCheck = false;
            }
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
            // 受注確定書発行記録データ表示
            dataShow(0, 0);
        }

        /// -------------------------------------------------------
        /// <summary>
        ///     受注確定書発行記録データ表示 </summary>
        /// -------------------------------------------------------
        private void dataShow(int col, int row)
        {
            // データ表示
            gridShow(dataGridView1);

            if (dataGridView1.RowCount > 0)
            {
                dataGridView1.CurrentCell = dataGridView1[col, row];
                dataGridView1.CurrentCell = null;
            }
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     請求書発行 </summary>
        /// ---------------------------------------------------------
        //private void seikyuRep()
        //{
        //    int _Pages = 0;
        //    int _Rows = 0;
        //    int sIr = 26;
        //    int _RowsPrn = 21;

        //    decimal _nebikigo = 0;      // 売上金額 - 値引額
        //    decimal _tax = 0;           // 消費税
        //    decimal _seikyu = 0;        // 請求金額

        //    bool openStatus = true;

        //    try
        //    {
        //        //マウスポインタを待機にする
        //        this.Cursor = Cursors.WaitCursor;

        //        string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

        //        Excel.Application oXls = new Excel.Application();

        //        // 請求書テンプレートBOOK
        //        Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル請求書, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing));

        //        // 印刷用BOOK
        //        Excel.Workbook oXlsBookPrn = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル受注確定書印刷, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing));

        //        // 請求書テンプレートシート
        //        Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

        //        // 請求書印刷シート
        //        Excel.Worksheet oxlsSheetPrn = null;

        //        Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

        //        try
        //        {
        //            for (int i = 0; i < dataGridView1.RowCount; i++)
        //            {
        //                // チェック行
        //                if (dataGridView1[colSel, i].Value.ToString() == "True")
        //                {
        //                    int iX = 0;

        //                    // 受注データ取得
        //                    foreach (var it in dts.新請求書.Where(a => a.ID == Utility.strToInt(dataGridView1[colID, i].Value.ToString())))
        //                    {
        //                        // 明細数
        //                        foreach (var t in it.Get受注1Rows().OrderBy(a => a.ID))
        //                        {
        //                            if (iX == 0 || _Rows >= _RowsPrn)
        //                            {
        //                                // ページ計印字
        //                                if (!openStatus)
        //                                {
        //                                    oxlsSheetPrn.Cells[47, 16] = _nebikigo.ToString();  // 小計
        //                                    oxlsSheetPrn.Cells[48, 16] = _tax.ToString();       // 消費税
        //                                    oxlsSheetPrn.Cells[49, 16] = _seikyu.ToString();    // 合計

        //                                    // 備考
        //                                    if (it.Is備考Null())
        //                                    {
        //                                        oxlsSheetPrn.Cells[47, 4] = string.Empty;
        //                                    }
        //                                    else
        //                                    {
        //                                        oxlsSheetPrn.Cells[47, 4] = it.備考;
        //                                    }

        //                                    // ページ計初期化
        //                                    _nebikigo = 0;
        //                                    _tax = 0;
        //                                    _seikyu = 0;
        //                                }

        //                                // ページ数加算
        //                                iX++;
        //                                _Pages = iX;

        //                                // ページシートを追加する
        //                                oxlsSheet.Copy(Type.Missing, oXlsBookPrn.Sheets[_Pages]);

        //                                // 印刷用カレントシート
        //                                oxlsSheetPrn = (Excel.Worksheet)oXlsBookPrn.Sheets[_Pages + 1];

        //                                // 対象行をゼロにしページを初期化
        //                                _Rows = 0;

        //                                // 請求書№
        //                                oxlsSheetPrn.Cells[1, 16] = "№ " + t.請求書発行日.Year.ToString() + t.請求書発行日.Month.ToString().PadLeft(2, '0') + t.請求書ID.ToString().PadLeft(5, '0');

        //                                // 発行日
        //                                oxlsSheetPrn.Cells[2, 15] = t.請求書発行日.ToShortDateString();

        //                                //// ページ数
        //                                //oxlsSheetPrn.Cells[3, 16] = "(" + iX.ToString() + "/" + it.明細数.ToString() + ")";

        //                                // 請求先住所
        //                                oxlsSheetPrn.Cells[2, 3] = "〒 " + t.得意先Row.請求先郵便番号;
        //                                oxlsSheetPrn.Cells[3, 3] = t.得意先Row.請求先都道府県 + " " + t.得意先Row.請求先住所1;
        //                                oxlsSheetPrn.Cells[4, 3] = t.得意先Row.請求先住所2;

        //                                string corp = string.Empty;

        //                                // 請求先名
        //                                if (t.得意先Row.Is請求先名称Null())
        //                                {
        //                                    corp = t.得意先Row.略称;
        //                                }
        //                                else
        //                                {
        //                                    corp = t.得意先Row.請求先名称;
        //                                }

        //                                // 部署・担当者名
        //                                string tn = (t.得意先Row.部署名 + " " + t.得意先Row.担当者名).Trim();

        //                                if (tn != string.Empty)
        //                                {
        //                                    oxlsSheetPrn.Cells[8, 3] = tn + " 様";
        //                                    oxlsSheetPrn.Cells[6, 3] = corp;
        //                                }
        //                                else
        //                                {
        //                                    oxlsSheetPrn.Cells[8, 3] = string.Empty;
        //                                    oxlsSheetPrn.Cells[6, 3] = corp + " 御中";
        //                                }

        //                                // 1ページ目
        //                                if (_Pages == 1)
        //                                {
        //                                    // 合計請求金額
        //                                    //oxlsSheetPrn.Cells[22, 5] = dataGridView1[colKingaku, i].Value.ToString().Trim().Replace(",", "");
        //                                    oxlsSheetPrn.Cells[22, 5] = it.請求金額.ToString();
        //                                }
        //                                else
        //                                {
        //                                    // 合計請求金額
        //                                    oxlsSheetPrn.Cells[22, 5] = "********************";
        //                                }

        //                                // 支払期日
        //                                oxlsSheetPrn.Cells[55, 4] = dataGridView1[colSDt, i].Value.ToString();

        //                                // 開始ステータス
        //                                openStatus = false;
        //                            }

        //                            // 明細内容
        //                            oxlsSheetPrn.Cells[sIr + _Rows, 1] = t.チラシ名;
        //                            oxlsSheetPrn.Cells[sIr + _Rows, 14] = t.単価.ToString();
        //                            oxlsSheetPrn.Cells[sIr + _Rows, 15] = t.枚数.ToString();
        //                            oxlsSheetPrn.Cells[sIr + _Rows, 16] = t.金額.ToString();

        //                            // 行加算
        //                            _Rows++;

        //                            // 値引額があるとき値引行を印字します
        //                            if (t.値引額 > 0)
        //                            {
        //                                // 明細内容
        //                                oxlsSheetPrn.Cells[sIr + _Rows, 1] = "値引";
        //                                oxlsSheetPrn.Cells[sIr + _Rows, 14] = (t.値引額 * (-1)).ToString();
        //                                oxlsSheetPrn.Cells[sIr + _Rows, 15] = "1";
        //                                oxlsSheetPrn.Cells[sIr + _Rows, 16] = (t.値引額 * (-1)).ToString();

        //                                // 行加算
        //                                _Rows++;
        //                            }

        //                            // ページ計加算
        //                            _nebikigo += (t.金額 - t.値引額);    // 小計
        //                            _tax += t.消費税;                   // 消費税
        //                            _seikyu += t.売上金額;              // 合計
        //                        }

        //                        // ページ計印字
        //                        oxlsSheetPrn.Cells[47, 16] = _nebikigo.ToString();  // 小計
        //                        oxlsSheetPrn.Cells[48, 16] = _tax.ToString();       // 消費税
        //                        oxlsSheetPrn.Cells[49, 16] = _seikyu.ToString();    // 合計

        //                        // 備考
        //                        if (it.Is備考Null())
        //                        {
        //                            oxlsSheetPrn.Cells[47, 4] = string.Empty;
        //                        }
        //                        else
        //                        {
        //                            oxlsSheetPrn.Cells[47, 4] = it.備考;
        //                        }
        //                    }
        //                }
        //            }

        //            // 印刷用BOOKの1番目のシートは削除する
        //            ((Excel.Worksheet)oXlsBookPrn.Sheets[1]).Delete();

        //            //マウスポインタを元に戻す
        //            this.Cursor = Cursors.Default;

        //            // 確認のためExcelのウィンドウを表示する
        //            //oXls.Visible = true;

        //            // 印刷
        //            oXlsBookPrn.PrintOutEx(1, Type.Missing, 1, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        //            // ウィンドウを非表示にする
        //            oXls.Visible = false;

        //            // 保存処理
        //            oXls.DisplayAlerts = false;

        //            DialogResult ret;

        //            // ダイアログボックスの初期設定
        //            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
        //            saveFileDialog1.Title = "請求書発行";
        //            saveFileDialog1.OverwritePrompt = true;
        //            saveFileDialog1.RestoreDirectory = true;
        //            saveFileDialog1.FileName = "請求書";
        //            saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

        //            // ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
        //            string fileName;
        //            ret = saveFileDialog1.ShowDialog();

        //            if (ret == System.Windows.Forms.DialogResult.OK)
        //            {
        //                fileName = saveFileDialog1.FileName;
        //                oXlsBookPrn.SaveAs(fileName, Type.Missing, Type.Missing,
        //                                Type.Missing, Type.Missing, Type.Missing,
        //                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
        //                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //            }

        //            // 処理終了メッセージ
        //            MessageBox.Show("終了しました","請求書発行",MessageBoxButtons.OK,MessageBoxIcon.Information);

        //            // 請求書発行フラグ更新
        //            seikyuRepFlgUpdate();

        //            // 請求書データ再表示
        //            dataShow(0, 0);
        //        }

        //        catch (Exception e)
        //        {
        //            MessageBox.Show(e.Message, "請求書発行", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        //            //Bookをクローズ
        //            oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

        //            //Excelを終了
        //            oXls.Quit();
        //        }

        //        finally
        //        {
        //            // ウィンドウを非表示にする
        //            oXls.Visible = false;

        //            // 保存処理
        //            oXls.DisplayAlerts = false;

        //            // Bookをクローズ
        //            oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);
        //            oXlsBookPrn.Close(Type.Missing, Type.Missing, Type.Missing);

        //            // Excelを終了
        //            oXls.Quit();

        //            // COM オブジェクトの参照カウントを解放する 
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheetPrn);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBookPrn);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

        //            //マウスポインタを元に戻す
        //            this.Cursor = Cursors.Default;
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message, "請求書作成", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //    }

        //    //マウスポインタを元に戻す
        //    this.Cursor = Cursors.Default;
        //}

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 表示ボタン
            if (e.ColumnIndex == 0)
            {
                string sPath = dataGridView1[colMemo, e.RowIndex].Value.ToString();
                frmOrderRecordXls frm = new frmOrderRecordXls(sPath);
                frm.ShowDialog();
            }

            // 確認
            if (e.ColumnIndex == 5 || e.ColumnIndex == 8 || e.ColumnIndex == 11 ||
                e.ColumnIndex == 14 || e.ColumnIndex == 17)
            {
                if (dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString() == "True")
                {
                    dataGridView1[e.ColumnIndex + 1, e.RowIndex].Value = DateTime.Now;
                    dataGridView1[e.ColumnIndex + 2, e.RowIndex].Value = global.loginUser;
                }
                else
                {
                    dataGridView1[e.ColumnIndex + 1, e.RowIndex].Value = string.Empty;
                    dataGridView1[e.ColumnIndex + 2, e.RowIndex].Value = string.Empty;
                }
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCellAddress.X == 5 || dataGridView1.CurrentCellAddress.X == 8 ||
                dataGridView1.CurrentCellAddress.X == 11 || dataGridView1.CurrentCellAddress.X == 14 ||
                dataGridView1.CurrentCellAddress.X == 17)
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!kCheck)
            {
                return;
            }

            // 確認１～５チェックのとき
            if (e.ColumnIndex == 7 || e.ColumnIndex == 10 || e.ColumnIndex == 13 || e.ColumnIndex == 16 ||
                e.ColumnIndex == 19)
            {
                int sID = int.Parse(dataGridView1[colID, e.RowIndex].Value.ToString());
                var r = dts.受注確定書発行記録.Single(a => a.ID == sID);

                if (e.ColumnIndex == 7)
                {
                    if ((bool)dataGridView1[colChk1, e.RowIndex].Value)
                    {
                        r.確認1 = global.FLGON;
                        r.確認日1 = DateTime.Parse(dataGridView1[colChkDt1, e.RowIndex].Value.ToString());
                        r.確認者1 = global.loginUserID;
                    }
                    else
                    {
                        r.確認1 = global.FLGOFF;
                        r.確認日1 = nullDate;
                        r.確認者1 = global.FLGOFF;
                    }
                }
                else if (e.ColumnIndex == 10)
                {
                    if ((bool)dataGridView1[colChk2, e.RowIndex].Value)
                    {
                        r.確認2 = global.FLGON;
                        r.確認日2 = DateTime.Parse(dataGridView1[colChkDt2, e.RowIndex].Value.ToString());
                        r.確認者2 = global.loginUserID;
                    }
                    else
                    {
                        r.確認2 = global.FLGOFF;
                        r.確認日2 = nullDate;
                        r.確認者2 = global.FLGOFF;
                    }
                }
                else if (e.ColumnIndex == 13)
                {
                    if ((bool)dataGridView1[colChk3, e.RowIndex].Value)
                    {
                        r.確認3 = global.FLGON;
                        r.確認日3 = DateTime.Parse(dataGridView1[colChkDt3, e.RowIndex].Value.ToString());
                        r.確認者3 = global.loginUserID;
                    }
                    else
                    {
                        r.確認3 = global.FLGOFF;
                        r.確認日3 = nullDate;
                        r.確認者3 = global.FLGOFF;
                    }
                }
                else if (e.ColumnIndex == 16)
                {
                    if ((bool)dataGridView1[colChk4, e.RowIndex].Value)
                    {
                        r.確認4 = global.FLGON;
                        r.確認日4 = DateTime.Parse(dataGridView1[colChkDt4, e.RowIndex].Value.ToString());
                        r.確認者4 = global.loginUserID;
                    }
                    else
                    {
                        r.確認4 = global.FLGOFF;
                        r.確認日4 = nullDate;
                        r.確認者4 = global.FLGOFF;
                    }
                }
                else if (e.ColumnIndex == 19)
                {
                    if ((bool)dataGridView1[colChk5, e.RowIndex].Value)
                    {
                        r.確認5 = global.FLGON;
                        r.確認日5 = DateTime.Parse(dataGridView1[colChkDt5, e.RowIndex].Value.ToString());
                        r.確認者5 = global.loginUserID;
                    }
                    else
                    {
                        r.確認5 = global.FLGOFF;
                        r.確認日5 = nullDate;
                        r.確認者5 = global.FLGOFF;
                    }
                }
                
                // データベース更新
                jAdp.Update(dts.受注確定書発行記録);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, this.Text);
        }
    }
}

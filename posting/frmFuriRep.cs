using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmFuriRep : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "振り表";

        public frmFuriRep()
        {
            InitializeComponent();

            // データ読込
            adp.Fill(dts.受注1);
            cAdp.Fill(dts.得意先);
            gAdp.Fill(dts.外注先);
            sAdp.Fill(dts.社員);
            hAdp.Fill(dts.判型);
            fAdp.Fill(dts.振り表);
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            gridSetting(dataGridView2);

            button1.Enabled = false;
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter adp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.判型TableAdapter hAdp = new darwinDataSetTableAdapters.判型TableAdapter();
        darwinDataSetTableAdapters.振り表TableAdapter fAdp = new darwinDataSetTableAdapters.振り表TableAdapter();
        
        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void gridSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 500;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "依頼日");
                tempDGV.Columns.Add("col12", "受注番号");
                tempDGV.Columns.Add("col2", "クライアント名");
                tempDGV.Columns.Add("col3", "チラシ名");
                tempDGV.Columns.Add("col4", "サイズ");
                tempDGV.Columns.Add("col5", "受注単価");
                tempDGV.Columns.Add("col6", "枚数");
                tempDGV.Columns.Add("col7", "予定表の配布期間");
                tempDGV.Columns.Add("col12", "渡した日");
                tempDGV.Columns.Add("col13", "渡した担当者");
                tempDGV.Columns.Add("col8", "");
                tempDGV.Columns.Add("col9", "振り先");
                tempDGV.Columns.Add("col10", "フリ原価");
                tempDGV.Columns.Add("col11", "営業担当");

                tempDGV.Columns[0].Width = 90;
                tempDGV.Columns[1].Width = 130;
                tempDGV.Columns[2].Width = 200;
                tempDGV.Columns[3].Width = 300;
                tempDGV.Columns[4].Width = 100;
                tempDGV.Columns[5].Width = 100;
                tempDGV.Columns[6].Width = 100;
                tempDGV.Columns[7].Width = 160;
                tempDGV.Columns[8].Width = 90;
                tempDGV.Columns[9].Width = 110;
                tempDGV.Columns[10].Width = 60;
                tempDGV.Columns[11].Width = 200;
                tempDGV.Columns[12].Width = 100;
                tempDGV.Columns[13].Width = 100;

                //tempDGV.Columns[0].Visible = false;
                //tempDGV.Columns[9].Visible = false;
                //tempDGV.Columns[10].Visible = false;
                //tempDGV.Columns[11].Visible = false;
                //tempDGV.Columns[12].Visible = false;

                tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0.00";
                tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                tempDGV.Columns[12].DefaultCellStyle.Format = "#,##0";

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                tempDGV.Columns[2].Frozen = true;

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

                // 罫線
                tempDGV.BorderStyle = BorderStyle.Fixed3D;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     グリッドにデータを表示 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト </param>
        /// -----------------------------------------------------------------
        private void gridShowData_org(DataGridView g)
        {
            int iX;

            try
            {
                // カーソル表示を待機状態
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // 振り表データ絞り込み検索
                var s = dts.受注1
                    .Where(a => !a.Is外注依頼日支払Null() && a.受注種別ID == 1 &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布開始日 && a.配布開始日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布終了日 && a.配布終了日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.配布開始日 <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.配布終了日))
                    .OrderBy(a => a.外注依頼日支払);

                // グリッドビューに表示する
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.外注依頼日支払.ToShortDateString();
                    g[1, iX].Value = t.ID;

                    if (t.得意先Row == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.得意先Row.略称;
                    }

                    g[3, iX].Value = t.チラシ名;

                    if (t.判型Row == null)
                    {
                        g[4, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[4, iX].Value = t.判型Row.名称;
                    }

                    g[5, iX].Value = t.単価;
                    g[6, iX].Value = t.枚数;

                    string sDay = string.Empty;
                    string eDay = string.Empty;

                    if (!t.Is配布開始日Null())
                    {
                        sDay = t.配布開始日.Month + "/" + t.配布開始日.Day;
                    }

                    if (!t.Is配布終了日Null())
                    {
                        eDay = t.配布終了日.Month + "/" + t.配布終了日.Day;
                    }

                    g[7, iX].Value = sDay + " ～ " + eDay;

                    if (t.Is外注渡し日Null())
                    {
                        g[8, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[8, iX].Value = t.外注渡し日.ToShortDateString();
                    }

                    if (t.外注受け渡し担当者 == null)
                    {
                        g[9, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[9, iX].Value = t.外注受け渡し担当者;
                    }

                    g[10, iX].Value = "→";

                    if (t.外注先RowBy外注先_受注1 == null)
                    {
                        g[11, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[11, iX].Value = t.外注先RowBy外注先_受注1.名称 + " 様";
                    }

                    g[12, iX].Value = t.外注原価支払;

                    if (t.得意先Row != null && t.得意先Row.社員Row != null)
                    {
                        g[13, iX].Value = t.得意先Row.社員Row.氏名;
                    }
                    else
                    {
                        g[13, iX].Value = string.Empty;
                    }

                    iX++;
                }

                // 外注先別枚数計
                var ss = dts.受注1
                    .Where(a => !a.Is外注依頼日支払Null() && a.受注種別ID == 1 &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布開始日 && a.配布開始日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布終了日 && a.配布終了日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.配布開始日 <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.配布終了日))
                    .GroupBy(a => new { a.外注先ID支払, a.外注先RowBy外注先_受注1.名称 })
                    .Select(gg => new
                    {
                        gg.Key.外注先ID支払,
                        gg.Key.名称,
                        合計 = gg.Sum(a => a.枚数)
                    });

                decimal ttl = 0;
                foreach (var tt in ss)
                {
                    g.Rows.Add();
                    g[11, iX].Value = tt.名称 + " 様";
                    g[12, iX].Value = tt.合計.ToString("#,##0");
                    ttl += tt.合計;
                    iX++;
                }

                // 総合計
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[11, iX].Value = "合計";
                    g[12, iX].Value = ttl.ToString("#,##0");

                    // カレントカーソル
                    dataGridView2.CurrentCell = null;
                }

                //カーソル表示を戻す
                Cursor.Current = Cursors.Default;
            }
            catch (Exception e)
            {
                //カーソル表示を戻す
                Cursor.Current = Cursors.Default;

                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }


        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     グリッドにデータを表示 2017/01/30 「振り表」テーブルセット </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト </param>
        /// ---------------------------------------------------------------------------
        private void gridShowData(DataGridView g)
        {
            int iX;

            try
            {
                // カーソル表示を待機状態
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // 振り表データ絞り込み検索
                var s = dts.振り表
                    .Where(a => !a.Is外注依頼日支払Null() &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布開始日 && a.配布開始日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布終了日 && a.配布終了日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.配布開始日 <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.配布終了日))
                    .OrderBy(a => a.外注依頼日支払);

                // グリッドビューに表示する
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.外注依頼日支払.ToShortDateString();
                    g[1, iX].Value = t.ID;

                    if (t.得意先Row == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.得意先Row.略称;
                    }

                    g[3, iX].Value = t.チラシ名;

                    if (t.判型Row == null)
                    {
                        g[4, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[4, iX].Value = t.判型Row.名称;
                    }

                    g[5, iX].Value = t.単価;
                    g[6, iX].Value = t.枚数;

                    string sDay = string.Empty;
                    string eDay = string.Empty;

                    if (!t.Is配布開始日Null())
                    {
                        sDay = t.配布開始日.Month + "/" + t.配布開始日.Day;
                    }

                    if (!t.Is配布終了日Null())
                    {
                        eDay = t.配布終了日.Month + "/" + t.配布終了日.Day;
                    }

                    g[7, iX].Value = sDay + " ～ " + eDay;

                    if (t.Is外注渡し日Null())
                    {
                        g[8, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[8, iX].Value = t.外注渡し日.ToShortDateString();
                    }

                    if (t.外注受け渡し担当者 == null)
                    {
                        g[9, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[9, iX].Value = t.外注受け渡し担当者;
                    }

                    g[10, iX].Value = "→";

                    if (t.外注先Row == null)
                    {
                        g[11, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[11, iX].Value = t.外注先Row.名称 + " 様";
                    }

                    g[12, iX].Value = t.外注原価支払;

                    if (t.得意先Row != null && t.得意先Row.社員Row != null)
                    {
                        g[13, iX].Value = t.得意先Row.社員Row.氏名;
                    }
                    else
                    {
                        g[13, iX].Value = string.Empty;
                    }

                    iX++;
                }

                // 外注先別枚数計
                var ss = dts.振り表
                    .Where(a => !a.Is外注依頼日支払Null() &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布開始日 && a.配布開始日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.配布終了日 && a.配布終了日 <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.配布開始日 <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.配布終了日))
                    .GroupBy(a => new { a.外注先ID支払, a.外注先Row.名称 })
                    .Select(gg => new
                    {
                        gg.Key.外注先ID支払,
                        gg.Key.名称,
                        合計 = gg.Sum(a => a.枚数)
                    });

                decimal ttl = 0;
                foreach (var tt in ss)
                {
                    g.Rows.Add();
                    g[11, iX].Value = tt.名称 + " 様";
                    g[12, iX].Value = tt.合計.ToString("#,##0");
                    ttl += tt.合計;
                    iX++;
                }

                // 総合計
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[11, iX].Value = "合計";
                    g[12, iX].Value = ttl.ToString("#,##0");

                    // カレントカーソル
                    dataGridView2.CurrentCell = null;
                }

                //カーソル表示を戻す
                Cursor.Current = Cursors.Default;
            }
            catch (Exception e)
            {
                //カーソル表示を戻す
                Cursor.Current = Cursors.Default;

                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("終了します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 配布日範囲チェック
            if (DateTime.Parse(iraiDtS.Value.ToShortDateString()) > DateTime.Parse(iraiDtE.Value.ToShortDateString()))
            {
                MessageBox.Show("配布日範囲が正しくありません", "配布日範囲エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // データ表示
            gridShowData(dataGridView2);
            
            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("該当データがありませんでした", MESSAGE_CAPTION);
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ShowPosting(dataGridView1, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("振り表を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            KanryoReport(dataGridView2);
        }

        private void KanryoReport(DataGridView g)
        {
            const int S_GYO = 4;        // エクセルファイル見出し行（明細は4行目から印字）
            const int S_ROWSMAX = 14;   // エクセルファイル列最大値
            int r = S_GYO;              // 印刷する行
            bool tl = true;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル振り表, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    // 対象配布日範囲
                    string hs = string.Empty;
                    string he = string.Empty;

                    if (iraiDtS.Checked)
                    {
                        hs = iraiDtS.Value.ToShortDateString();
                    }

                    if (iraiDtE.Checked)
                    {
                        he = iraiDtE.Value.ToShortDateString();
                    }

                    oxlsSheet.Cells[1, 3] = "配布日： " + hs + " ～ " + he;

                    // 明細
                    //for (int i = 0; i < 10; i++)
                    //{
                        for (int iX = 0; iX <= g.RowCount - 1; iX++)
                        {
                            if (g[0, iX].Value != null)
                            {
                                // 1行空き
                                oxlsSheet.Cells[r, 1] = " ";

                                //セル下部へ実線ヨコ罫線を引く
                                rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                                rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                r++;

                                // 明細
                                oxlsSheet.Cells[r, 1] = g[0, iX].Value.ToString();     // 依頼日
                                oxlsSheet.Cells[r, 2] = g[1, iX].Value.ToString();     // 受注番号
                                oxlsSheet.Cells[r, 3] = g[2, iX].Value.ToString();     // クライアント名
                                oxlsSheet.Cells[r, 4] = g[3, iX].Value.ToString();     // チラシ名
                                oxlsSheet.Cells[r, 5] = g[4, iX].Value.ToString();     // サイズ
                                oxlsSheet.Cells[r, 6] = g[5, iX].Value.ToString();     // 受注単価
                                oxlsSheet.Cells[r, 7] = g[6, iX].Value.ToString();     // 枚数
                                oxlsSheet.Cells[r, 8] = g[7, iX].Value.ToString();     // 予定表の配布期間
                                oxlsSheet.Cells[r, 9] = g[8, iX].Value.ToString();     // 渡した日
                                oxlsSheet.Cells[r, 10] = g[9, iX].Value.ToString();     // 渡した担当者
                                oxlsSheet.Cells[r, 11] = g[10, iX].Value.ToString();     // 
                                oxlsSheet.Cells[r, 12] = g[11, iX].Value.ToString();    // 振り先
                                oxlsSheet.Cells[r, 13] = g[12, iX].Value.ToString();   // 振り単価
                                oxlsSheet.Cells[r, 14] = g[13, iX].Value.ToString();   // 営業担当
                            }
                            else
                            {
                                if (tl)
                                {
                                    // 1行空き
                                    oxlsSheet.Cells[r, 1] = " ";

                                    //セル下部へ実線ヨコ罫線を引く
                                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    r++;
                                }

                                // 振り先別合計
                                oxlsSheet.Cells[r, 12] = g[11, iX].Value.ToString();    // 振り先
                                oxlsSheet.Cells[r, 13] = g[12, iX].Value.ToString();   // 振り合計
                                tl = false;
                            }

                            //セル下部へ実線ヨコ罫線を引く
                            rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                            r++;
                        }
                    //}
                    
                    //////セル上部へ実線ヨコ罫線を引く
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////セル下部へ実線ヨコ罫線を引く
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    // 表全体に実線縦罫線を引く
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    // 表全体の左端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    // 表全体の右端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //// 外注先別：支払額
                    //var s = dts.受注1
                    //    .Where(a => !a.Is外注依頼日支払Null())
                    //    .GroupBy(a => new { a.外注先ID支払, a.外注先Row.名称 })
                    //    .Select(gg => new
                    //    {
                    //        gg.Key.外注先ID支払,
                    //        gg.Key.名称,
                    //        合計 = gg.Sum(a => (decimal)a.枚数 * a.外注原価支払)
                    //    });

                    //r++;
                    //foreach (var t in s)
                    //{
                    //    oxlsSheet.Cells[r, 10] = t.名称;    // 振り先
                    //    oxlsSheet.Cells[r, 11] = t.合計;   // 振り合計
                    //    r++;
                    //}

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    // 印刷
                    oxlsSheet.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    // ダイアログボックスの初期設定
                    saveFileDialog1.Title = MESSAGE_CAPTION;
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = MESSAGE_CAPTION;
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

                    // ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();
                }

                finally
                {
                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }
    }
}
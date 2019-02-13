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
using MyLibrary;

namespace posting
{
    public partial class frmShiharaiYotei : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "外注支払予定表";

        public frmShiharaiYotei()
        {
            InitializeComponent();

            // データ読込
            adp.Fill(dts.受注1);
            sAdp.Fill(dts.外注支払);
            gAdp.Fill(dts.外注先);
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            gridSettingShiharai(dataGridView2);
            gridSetting(dataGridView1);

            // 年月
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            //button1.Enabled = false;
            linkLabel1.Visible = false;
            linkLabel2.Visible = false;

            // 支払予定表示
            showShiharaiYotei();
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter adp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.外注支払TableAdapter sAdp = new darwinDataSetTableAdapters.外注支払TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();

        Utility.消費税率 cTax = new Utility.消費税率();

        ///--------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///--------------------------------------------------------------------
        private void gridSettingShiharai(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 222;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "予定日");
                tempDGV.Columns.Add("col2", "コード");
                tempDGV.Columns.Add("col3", "支払先名");
                tempDGV.Columns.Add("col5", "支払額");
                //tempDGV.Columns.Add("col6", "消費税額");
                //tempDGV.Columns.Add("col7", "税込金額");

                tempDGV.Columns[0].Width = 110;
                tempDGV.Columns[1].Width = 60;
                tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[3].Width = 120;
                //tempDGV.Columns[4].Width = 120;
                //tempDGV.Columns[5].Width = 120;

                tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                //tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                //tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                //tempDGV.Columns[2].Frozen = true;

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

        ///--------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///--------------------------------------------------------------------
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
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 222;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "予定日");
                tempDGV.Columns.Add("col2", "コード");
                tempDGV.Columns.Add("col3", "支払先名");
                tempDGV.Columns.Add("col4", "内容");
                tempDGV.Columns.Add("col5", "支払額");
                //tempDGV.Columns.Add("col7", "消費税額");  // 2016/02/01
                //tempDGV.Columns.Add("col8", "税込金額");  // 2016/02/01
                tempDGV.Columns.Add("col6", "支払実施日");
                
                tempDGV.Columns[0].Width = 110;
                tempDGV.Columns[1].Width = 60;
                tempDGV.Columns[2].Width = 180;
                tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[4].Width = 120;

                //tempDGV.Columns[5].Width = 90;    // 2016/02/01
                //tempDGV.Columns[6].Width = 120;   // 2016/02/01
                
                tempDGV.Columns[5].Width = 110;
                
                tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                //tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";  // 2016/02/01
                //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";  // 2016/02/01

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                //tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;     // 2016/02/01
                //tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;     // 2016/02/01
                
                tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
        private void gridShowShiharai(DataGridView g)
        {
            int iX;
            decimal tl = 0;
            decimal tlTax = 0;
            decimal tlZeikomi = 0;

            try
            {
                //カーソル表示を待機状態
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // 支払予定表データ検索：2015/11/19 a.外注先Row != null を条件に追加
                var s = dts.受注1
                    .Where(a => a.外注先RowBy外注先_受注1 != null &&
                           !a.Is外注支払日支払Null() &&
                                 a.外注支払日支払.Year == int.Parse(txtYear.Text) && a.外注支払日支払.Month == int.Parse(txtMonth.Text))
                    .GroupBy(a => new { a.外注支払日支払, a.外注先ID支払, a.外注先RowBy外注先_受注1.名称 })
                    .Select(gg => new
                    {
                        gg.Key.外注支払日支払,
                        gg.Key.外注先ID支払,
                        gg.Key.名称,
                        //支払額 = gg.Sum(a => a.外注原価支払 * a.枚数)

                        // 2015/12/06 外注原価を「単価から原価総額入力へ変更」に伴う
                        支払額 = gg.Sum(a => a.外注原価支払)
                    })
                    .OrderBy(a => a.外注支払日支払).ThenBy(a => a.外注先ID支払);

                // グリッドビューに表示する
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.外注支払日支払.ToShortDateString();
                    g[1, iX].Value = t.外注先ID支払;

                    if (t.名称 == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.名称;
                    }

                    g[3, iX].Value = t.支払額.ToString("#,##0");

                    // 消費税額表示廃止　2016/02/01
                    //// 消費税額取得
                    //decimal KingakuTax = 0;
                    //foreach (var it in dts.受注1.Where(a => !a.Is外注支払日支払Null() && a.外注支払日支払 == t.外注支払日支払 && a.外注先ID支払 == t.外注先ID支払))
                    //{
                    //    //税率再取得
                    //    cTax.Ritsu = Utility.GetTaxRT(it.受注日);

                    //    //消費税額計算 
                    //    //KingakuTax += Utility.GetTax(((decimal)it.枚数 * it.外注原価支払), cTax.Ritsu);
                    //    // 2015/12/06 外注原価を「単価から原価総額入力へ変更」に伴う
                    //    KingakuTax += Utility.GetTax(it.外注原価支払, cTax.Ritsu);
                    //}

                    // 消費税額表示廃止　2016/02/01
                    //// 消費税
                    //g[4, iX].Value = KingakuTax.ToString("#,##0");

                    // 消費税額表示廃止　2016/02/01
                    //// 税込金額
                    //g[5, iX].Value = (t.支払額 + KingakuTax).ToString("#,##0");

                    iX++;

                    tl += t.支払額;

                    // 消費税額表示廃止　2016/02/01
                    //tlTax += KingakuTax;
                    //tlZeikomi += (t.支払額 + KingakuTax);
                }

                // 総合計
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[2, iX].Value = "合計";
                    g[3, iX].Value = tl.ToString("#,##0");

                    //g[4, iX].Value = tlTax.ToString("#,##0");     // 2016/02/01
                    //g[5, iX].Value = tlZeikomi.ToString("#,##0"); // 2016/02/01

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

        /// -----------------------------------------------------------------
        /// <summary>
        ///     グリッドにデータを表示 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト </param>
        /// -----------------------------------------------------------------
        private void gridShowData(DataGridView g, DateTime sDt, int sID)
        {
            int iX;
            decimal tl = 0;
            decimal tlTax = 0;
            decimal tlZeikomi = 0;

            try
            {
                //カーソル表示を待機状態
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // 指定の外注先の支払予定データ検索
                var s = dts.受注1
                    .Where(a => !a.Is外注支払日支払Null() && a.外注支払日支払 == sDt && a.外注先ID支払 == sID)
                    .OrderBy(a => a.外注支払日支払);

                // グリッドビューに表示する
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.外注支払日支払.ToShortDateString();
                    g[1, iX].Value = t.外注先ID支払;

                    if (t.外注先RowBy外注先_受注1 == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.外注先RowBy外注先_受注1.名称;
                    }

                    g[3, iX].Value = t.チラシ名;

                    //decimal kingaku = (decimal)t.枚数 * t.外注原価支払;

                    // 2015/12/06 外注原価を「単価から原価総額入力へ変更」に伴う
                    decimal kingaku = t.外注原価支払;
                    g[4, iX].Value = kingaku.ToString("#,##0");
                    
                    //税率再取得
                    cTax.Ritsu = Utility.GetTaxRT(t.受注日);

                    // 消費税額表示廃止　2016/02/01
                    ////消費税額計算 
                    //decimal KingakuTax = Utility.GetTax(kingaku, cTax.Ritsu);

                    // 消費税額表示廃止　2016/02/01
                    //// 消費税
                    //g[5, iX].Value = KingakuTax.ToString("#,##0");

                    // 消費税額表示廃止　2016/02/01
                    //// 税込金額
                    //g[6, iX].Value = (kingaku + KingakuTax).ToString("#,##0");

                    // 外注支払日
                    if (t.外注支払Row == null)
                    {
                        g[5, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[5, iX].Value = t.外注支払Row.日付.ToShortDateString();
                    }

                    iX++;

                    tl += kingaku;

                    // 消費税額表示廃止　2016/02/01
                    //tlTax += KingakuTax;
                    //tlZeikomi += (kingaku + KingakuTax);
                }

                // 総合計
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[2, iX].Value = "合計";
                    g[4, iX].Value = tl.ToString("#,##0");

                    // 2016/02/01
                    //g[5, iX].Value = tlTax.ToString("#,##0");
                    //g[6, iX].Value = tlZeikomi.ToString("#,##0");

                    // カレントカーソル
                    dataGridView1.CurrentCell = null;

                    // CSV
                    linkLabel2.Visible = true;
                }
                else
                {
                    linkLabel2.Visible = false;
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
            // 年月チェック
            if (Utility.strToInt(txtMonth.Text) < 1 || Utility.strToInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("指定月が正しくありません", "指定月エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            // データ表示
            gridShowShiharai(dataGridView2);

            // 明細グリッドクリア
            dataGridView1.Rows.Clear();
            
            // 該当データなし
            if (dataGridView2.RowCount == 0)
            {
                label5.Text = "該当データがありませんでした";
                label5.Visible = true;
                //button1.Enabled = false;
            }
            else
            {
                label5.Text = "";
                label5.Visible = false;
                //button1.Enabled = true;
            }
        }

        private void showShiharaiYotei()
        {
            // データ表示
            gridShowShiharai(dataGridView2);

            // 明細グリッドクリア
            dataGridView1.Rows.Clear();
            linkLabel2.Visible = false;

            // 該当データなし
            if (dataGridView2.RowCount == 0)
            {
                label5.Text = "該当データがありませんでした";
                label5.Visible = true;
                //button1.Enabled = false;
                linkLabel1.Visible = false;
            }
            else
            {
                label5.Text = "";
                label5.Visible = false;
                //button1.Enabled = true;
                linkLabel1.Visible = true;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int sID = 0;
            DateTime sDt = DateTime.Parse("1900/01/01");

            if (dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value == null)
            {
                return;
            }
            else
            {
                // 外注先コード取得
                sID = Utility.strToInt(dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString());
            }

            if (dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value == null)
            {
                return;
            }
            else
            {
                // 支払日取得
                if (!DateTime.TryParse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString(), out sDt))
                {
                    return;
                }
            }

            // 指定外注先明細データ表示
            gridShowData(dataGridView1, sDt, sID);
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void txtMonth_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 前月
            yearMonthBefore();
            showShiharaiYotei();
        }

        private void yearMonthBefore()
        {
            int m = Utility.strToInt(txtMonth.Text);
            m--;

            if (m == 0)
            {
                txtYear.Text = (Utility.strToInt(txtYear.Text) - 1).ToString();
                txtMonth.Text = "12";
            }
            else
            {
                txtMonth.Text = m.ToString();
            } 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 次月
            yearMonthAfter();
            showShiharaiYotei();
        }

        private void yearMonthAfter()
        {
            int m = Utility.strToInt(txtMonth.Text);
            m++;

            if (m > 12)
            {
                txtYear.Text = (Utility.strToInt(txtYear.Text) + 1).ToString();
                txtMonth.Text = (m - 12).ToString();
            }
            else
            {
                txtMonth.Text = m.ToString();
            }
        }

        private void txtMonth_TextChanged(object sender, EventArgs e)
        {
            yearMonthChange();
        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            yearMonthChange();
        }

        private void yearMonthChange()
        {
            // 年月チェック
            if (Utility.strToInt(txtYear.Text) < 2014)
            {
                txtYear.Focus();
                return;
            }

            if (Utility.strToInt(txtMonth.Text) < 1 || Utility.strToInt(txtMonth.Text) > 12)
            {
                txtMonth.Focus();
                return;
            }

            // 支払予定表示
            showShiharaiYotei();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, MESSAGE_CAPTION);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION + "明細");
        }
    }
}
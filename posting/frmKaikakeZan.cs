using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MyLibrary;

namespace posting
{
    public partial class frmKaikakeZan : Form
    {
        public frmKaikakeZan()
        {
            InitializeComponent();

            // データ読み込み
            uAdp.Fill(dts.外注先);
            rAdp.Fill(dts.受注1);
            nAdp.Fill(dts.外注支払);
            adp.Fill(dts.買掛元帳);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.買掛元帳TableAdapter adp = new darwinDataSetTableAdapters.買掛元帳TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter uAdp = new darwinDataSetTableAdapters.外注先TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter rAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.外注支払TableAdapter nAdp = new darwinDataSetTableAdapters.外注支払TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        #region グリッドビューカラム定義
        string colZengetsu = "col6";
        string colID = "col1";
        string colGcd = "col2";
        string colKousei = "col3";
        string colAddDt = "col4";
        string colUpDt = "col5";
        string colUserID = "col7";
        string colClient = "col8";
        string colKaikake = "col9";
        string colKingaku = "col10";
        string colName = "col11";
        string colSeikyuNum = "col12";
        string colZandaka = "col13";
        string colKbn = "col14";
        #endregion

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビューの定義
            gridSetting(dataGridView1);
                        
            // 画面初期化
            dispClear();
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
                tempDGV.Height = 522;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                //tempDGV.Columns.Add(colDt, "日付");

                tempDGV.Columns.Add(colID, "ID");
                tempDGV.Columns.Add(colClient, "外注先");
                tempDGV.Columns.Add(colZengetsu, "前月繰越");
                tempDGV.Columns.Add(colKaikake, "買掛金額");
                tempDGV.Columns.Add(colKingaku, "支払金額");
                tempDGV.Columns.Add(colZandaka, "当月残高");
                tempDGV.Columns.Add(colKousei, "構成比");

                //tempDGV.Columns.Add(colID, "入金ID");
                //tempDGV.Columns.Add(colSeikyuNum, "請求書ID");
                //tempDGV.Columns.Add(colKbn, "区分");
                ////tempDGV.Columns.Add(colAddDt, "登録年月日");
                ////tempDGV.Columns.Add(colUpDt, "更新年月日");
                ////tempDGV.Columns.Add(colUserID, "ユーザーID");

                //tempDGV.Columns[colID].Visible = false;
                //tempDGV.Columns[colSeikyuNum].Visible = false;
                //tempDGV.Columns[colKbn].Visible = false;

                //tempDGV.Columns[colDt].Width = 100;
                tempDGV.Columns[colID].Width = 60;
                tempDGV.Columns[colClient].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colZengetsu].Width = 110;
                tempDGV.Columns[colKaikake].Width = 110;
                tempDGV.Columns[colKingaku].Width = 110;
                tempDGV.Columns[colZandaka].Width = 110;
                tempDGV.Columns[colKousei].Width = 100;
                
                tempDGV.Columns[colID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colZengetsu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colKaikake].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colZandaka].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colKousei].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                //tempDGV.Columns[colClient].Frozen = true;

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
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical;
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
        private void gridShow(DataGridView g, int col, int row)
        {
            // カーソルを待機にする
            Cursor = Cursors.WaitCursor;

            // グリッドビュー
            g.Rows.Clear();

            int iX = 0;
            DateTime sDt = DateTime.Parse("1900/01/01");
            DateTime sDate = DateTime.Parse(txtYear.Text + "/" + txtMonth.Text + "/01");
            decimal zenKurikoshi = 0;
            decimal zenKurikoshiTotal = 0;
            decimal karikataTotal = 0;
            decimal kashikataTotal = 0;
            decimal zandakaTotal = 0;

            // グリッドに指定期間の明細表示
            //var s = dts.売掛元帳.OrderBy(a => a.日付); 

            // 2015/12/09
            //var s = dts.売掛元帳
            //    .Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0 && 
            //           a.日付.Year == Utility.strToInt(txtYear.Text) && a.日付.Month == Utility.strToInt(txtMonth.Text))
            //    .OrderBy(a => a.日付);

            // 借方合計 : 2016/06/24
            decimal kaikakeTl = dts.買掛元帳
                .Where(a => a.日付 < sDate.AddMonths(1)).Sum(a => a.外注原価支払);

            // 貸方合計 : 2016/06/24
            int shiharaiTl = dts.買掛元帳
                .Where(a => a.日付 < sDate.AddMonths(1)).Sum(a => a.支払金額);

            // 残高合計
            decimal zanTL = kaikakeTl - shiharaiTl;

            foreach (var it in dts.外注先.OrderBy(a => a.ID))
            {
                // 指定以外の外注先はネグる
                if (txtsClient.Text != string.Empty)
                {
                    if (!it.名称.Contains(txtsClient.Text))
                    {
                        continue;
                    }
                }

                // 前月繰越
                zenKurikoshi = setKurikoshi(sDate, it.ID);

                // 当月取引
                // 条件式：「a.外注先Row != null」を追加 2018/05/17
                var s = dts.買掛元帳.Where(a => a.外注先Row != null && a.外注先ID支払 == it.ID && a.日付.Year == Utility.strToInt(txtYear.Text) && a.日付.Month == Utility.strToInt(txtMonth.Text))
                            .GroupBy(a => new { a.外注先Row.名称 })
                            .Select(cg => new
                            {
                                cName = cg.Key,
                                kaikake = cg.Sum(a => a.外注原価支払),
                                shiharai = cg.Sum(a => a.支払金額)
                            });

                bool tougetsu = false;

                foreach (var t in s)
                {
                    tougetsu = true;

                    // 明細
                    g.Rows.Add();
                    iX = g.RowCount - 1;

                    // 外注先コード
                    g[colID, iX].Value = it.ID;

                    // 買掛先名称
                    g[colClient, iX].Value = t.cName.名称;

                    // 前月繰越
                    g[colZengetsu, iX].Value = zenKurikoshi.ToString("#,0");

                    // 買掛金
                    g[colKaikake, iX].Value = t.kaikake.ToString("#,0");

                    // 支払金額
                    g[colKingaku, iX].Value = t.shiharai.ToString("#,0");

                    // 当月残高
                    g[colZandaka, iX].Value = (zenKurikoshi + t.kaikake - t.shiharai).ToString("#,0");

                    // 構成比
                    g[colKousei, iX].Value = ((zenKurikoshi + t.kaikake - t.shiharai) / zanTL * 100).ToString("n2");

                    // 合計値加算
                    karikataTotal += t.kaikake;
                    kashikataTotal += t.shiharai;
                    zandakaTotal += (zenKurikoshi + t.kaikake - t.shiharai);
                }

                // 当月取引なし
                if (!tougetsu)
                {
                    // 前月繰越があるとき表示する
                    if (zenKurikoshi != 0)
                    {
                        // 明細
                        g.Rows.Add();
                        iX = g.RowCount - 1;

                        // 得意先コード
                        g[colID, iX].Value = it.ID;

                        // 買掛先名称
                        g[colClient, iX].Value = it.名称;

                        // 前月繰越
                        g[colZengetsu, iX].Value = zenKurikoshi.ToString("#,0");

                        // 買掛金
                        g[colKaikake, iX].Value = 0;

                        // 支払金額
                        g[colKingaku, iX].Value = 0;

                        // 当月残高
                        g[colZandaka, iX].Value = zenKurikoshi.ToString("#,0");

                        // 構成比
                        g[colKousei, iX].Value = (zenKurikoshi / zanTL * 100).ToString("n2");

                        // 残高合計加算
                        zandakaTotal += zenKurikoshi;
                    }
                }

                // 前月繰越合計加算
                zenKurikoshiTotal += zenKurikoshi;
            }

            // 総計
            if (g.RowCount > 0)
            {
                g.Rows.Add();
                iX = g.RowCount - 1;
                g[colID, iX].Value = "";
                g[colClient, iX].Value = "総合計";
                g[colZengetsu, iX].Value = zenKurikoshiTotal.ToString("#,0");
                g[colKaikake, iX].Value = karikataTotal.ToString("#,0");
                g[colKingaku, iX].Value = kashikataTotal.ToString("#,0");
                g[colZandaka, iX].Value = zandakaTotal.ToString("#,0");
                g[colKousei, iX].Value = "100.00";

                g.CurrentCell = null;
                button2.Enabled = true;

                // カレント行
                g.CurrentCell = g[col, row];
                g.CurrentCell = null;
            }
            else
            {
                MessageBox.Show("対象となるデータがありませんでした", "買掛残高表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
            }

            // カーソルを戻す
            Cursor = Cursors.Default;
        }
        
        /// -----------------------------------------------------------------
        /// <summary>
        ///     明細開始時の前月繰越金額を取得する </summary>
        /// <param name="dt">
        ///     データ日付 </param>
        /// <param name="sClient">
        ///     指定外注先</param>
        /// -----------------------------------------------------------------
        private decimal setKurikoshi(DateTime dt, int sClient)
        {
            decimal sKaikake = 0;
            int sShiharai = 0;
            decimal sKurikoshi = 0;

            // 指定期間以前の繰越金額を求める
            // 条件式：「a.外注先Row != null」を追加 2018/05/17
            var s = dts.買掛元帳.Where(a => a.日付 < dt && a.外注先Row != null && a.外注先Row.ID == sClient)
                    .GroupBy(a => new { a.外注先Row.名称 })
                    .Select(cg => new
                    {
                        cName = cg.Key,
                        kaikake = cg.Sum(a => a.外注原価支払),
                        shiharai = cg.Sum(a => a.支払金額)
                    });
                
            foreach (var t in s)
            {
                sKaikake += t.kaikake;
                sShiharai += t.shiharai;
            }

            sKurikoshi = sKaikake - sShiharai;

            return sKurikoshi;
        }
        
        /// -------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// -------------------------------------------------------------
        private void dispClear()
        {
            fMode.Mode = 0;
            fMode.ID = 0;
            txtYear.Text = DateTime.Now.Year.ToString();
            txtMonth.Text = DateTime.Now.Month.ToString();
            button2.Enabled = false;
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

        private void button1_Click(object sender, EventArgs e)
        {
            // 買掛残高表示
            dataShow();
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     買掛残高表示 </summary>
        ///--------------------------------------------------------
        private void dataShow()
        {
            if (Utility.strToInt(txtYear.Text) == 0)
            {
                MessageBox.Show("年が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return;
            }

            if (Utility.strToInt(txtMonth.Text) == 0 || Utility.strToInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("月が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            // データグリッドビューデータ表示
            gridShow(dataGridView1, 0, 0);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "買掛残高表");
        }

        private void txtMonth_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
    }
}

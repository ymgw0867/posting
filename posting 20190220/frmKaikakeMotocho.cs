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
    public partial class frmKaikakeMotocho : Form
    {
        public frmKaikakeMotocho()
        {
            InitializeComponent();
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.買掛元帳TableAdapter adp = new darwinDataSetTableAdapters.買掛元帳TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.外注支払TableAdapter sAdp = new darwinDataSetTableAdapters.外注支払TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        Utility.消費税率 cTax = new Utility.消費税率();

        #region グリッドビューカラム定義
        string colDt = "col6";
        string colID = "col1";
        string colGcd = "col2";
        string colMemo = "col3";
        string colAddDt = "col4";
        string colUpDt = "col5";
        string colUserID = "col7";
        string colClient = "col8";
        string colKingaku = "col9";
        string colKaikake = "col10";
        string colName = "col11";
        string colOrderNum = "col12";
        string colZandaka = "col13";
        #endregion

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビューの定義
            gridSetting(dataGridView1);

            //// 外注先コンボボックス
            //Utility.comboGaichu.itemLoad(comboBox1);

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
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                tempDGV.Columns.Add(colDt, "日付");
                tempDGV.Columns.Add(colClient, "外注先名称");
                tempDGV.Columns.Add(colName, "内容");
                tempDGV.Columns.Add(colKaikake, "買掛金額");
                tempDGV.Columns.Add(colKingaku, "支払金額");
                tempDGV.Columns.Add(colZandaka, "残高");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colID, "支払ID");
                tempDGV.Columns.Add(colOrderNum, "受注番号");
                //tempDGV.Columns.Add(colAddDt, "登録年月日");
                //tempDGV.Columns.Add(colUpDt, "更新年月日");
                //tempDGV.Columns.Add(colUserID, "ユーザーID");

                tempDGV.Columns[colID].Visible = false;
                tempDGV.Columns[colOrderNum].Visible = false;

                tempDGV.Columns[colDt].Width = 110;
                tempDGV.Columns[colClient].Width = 220;
                tempDGV.Columns[colName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colKaikake].Width = 110;
                tempDGV.Columns[colKingaku].Width = 110;
                tempDGV.Columns[colZandaka].Width = 110;
                tempDGV.Columns[colMemo].Width = 240;

                tempDGV.Columns[colDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colKaikake].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colZandaka].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
            // データ読み込み
            adp.Fill(dts.買掛元帳);
            gAdp.Fill(dts.外注先);
            jAdp.Fill(dts.受注1);
            sAdp.Fill(dts.外注支払);

            // グリッドビュー
            g.Rows.Clear();

            int iX = 0;
            decimal zan = 0;
            DateTime sDt = DateTime.Parse("1900/01/01");
            decimal mKaikake = 0;
            int mShiharai = 0;
            int tKaikake = 0;
            int tShiharai = 0;
            bool firstData = true;

            Cursor = Cursors.WaitCursor;

            // 指定期間以前の明細（繰り越し分）表示

            // グリッドに指定期間の明細表示
            var s = dts.買掛元帳.OrderBy(a => a.日付).ThenBy(a => a.外注先ID支払).ThenBy(a => a.区分);

            // 支払先指定
            if (txtsGaichu.Text.Trim() != string.Empty)
            {
                //s = s.Where(a => a.外注先Row.名称.Contains(txtsGaichu.Text)).OrderBy(a => a.日付).ThenBy(a => a.外注先ID支払).ThenBy(a => a.区分);
                // 2017/11/24
                s = s.Where(a => a.外注先Row != null && a.外注先Row.名称.Contains(txtsGaichu.Text)).OrderBy(a => a.日付).ThenBy(a => a.外注先ID支払).ThenBy(a => a.区分);
            }

            // 開始日付
            if (dateTimePicker1.Checked)
            {
                s = s.Where(a => a.日付 >= DateTime.Parse(dateTimePicker1.Value.ToShortDateString())).OrderBy(a => a.日付).ThenBy(a => a.外注先ID支払).ThenBy(a => a.区分);
            }

            // 終了日付
            if (dateTimePicker2.Checked)
            {
                s = s.Where(a => a.日付 <= DateTime.Parse(dateTimePicker2.Value.ToShortDateString())).OrderBy(a => a.日付).ThenBy(a => a.外注先ID支払).ThenBy(a => a.区分);
            }

            foreach (var t in s)
            {
                // 前月繰越
                if (firstData)
                {
                    mKaikake += setKurikoshi(dataGridView1, t.日付, txtsGaichu.Text);
                    zan += mKaikake;
                    firstData = false;
                }

                // 月が変わると月間合計
                if (sDt.Year != 1900 && sDt.Month != t.日付.Month)
                {
                    monthTotal(g, t.日付.Month, ref mKaikake, ref mShiharai, sDt);
                }

                // 明細
                g.Rows.Add();
                iX = g.RowCount - 1;
                g[colDt, iX].Value = t.日付.ToShortDateString();

                if (t.外注先Row == null)
                {
                    g[colClient, iX].Value = string.Empty;
                }
                else
                {
                    g[colClient, iX].Value = t.外注先Row.名称;
                }

                // 支払・調整入力のチラシ名を表示する : 2016/12/20
                string chName = string.Empty;

                if (t.区分 == 2 || t.区分 == 3)
                {
                    foreach (var it in dts.外注支払.Where(a => a.ID == t.支払ID))
                    {
                        if (dts.受注1.Any(a => a.外注支払ID == it.支払コード))
                        {
                            foreach (var g1 in dts.受注1.Where(a => a.外注支払ID == it.支払コード))
                            {
                                if (chName != string.Empty)
                                {
                                    chName += ":";
                                }

                                chName += g1.チラシ名;
                            }
                        }
                        
                        if (dts.受注1.Any(a => a.外注支払ID2 == it.支払コード))
                        {
                            foreach (var g2 in dts.受注1.Where(a => a.外注支払ID2 == it.支払コード))
                            {
                                if (chName != string.Empty)
                                {
                                    chName += ":";
                                }

                                chName += g2.チラシ名;
                            }
                        }

                        if (dts.受注1.Any(a => a.外注支払ID3 == it.支払コード))
                        {
                            foreach (var g3 in dts.受注1.Where(a => a.外注支払ID3 == it.支払コード))
                            {
                                if (chName != string.Empty)
                                {
                                    chName += ":";
                                }

                                chName += g3.チラシ名; ;
                            }
                        }
                    }

                    // 支払のとき：2016/12/20
                    if (t.区分 == 2)
                    {
                        g[colName, iX].Value = chName;
                    }

                    // 精算のとき：2016/12/20
                    if (t.区分 == 3)
                    {
                        g[colName, iX].Value = t.内容 + " " + chName.Trim();
                    }
                }
                else
                {
                    g[colName, iX].Value = t.内容;
                }


                //if (t.受注1Row == null)
                //{
                //    g[colName, iX].Value = t.内容;
                //}
                //else
                //{
                //    g[colName, iX].Value = t.受注1Row.チラシ名;
                //}

                // 2015/12/06 外注原価を単価から原価総額入力へ変更に伴う
                //decimal kaikake = (decimal)t.枚数 * t.外注原価支払;
                decimal kaikake = t.外注原価支払;

                // 税込・税抜
                if (t.区分 == 1 && cmbTax.SelectedIndex == 0)
                {
                    // 税率取得
                    cTax.Ritsu = Utility.GetTaxRT(t.日付);

                    // 消費税額計算 
                    decimal z = Utility.GetTax(kaikake, cTax.Ritsu);

                    // 税込金額
                    kaikake += z;
                }

                g[colKaikake, iX].Value = kaikake.ToString("#,###");
                g[colKingaku, iX].Value = t.支払金額.ToString("#,###");

                // 残高
                if (iX == 0)
                {
                    g[colZandaka, iX].Value = (kaikake + t.支払金額).ToString("#,##0");
                    zan += (decimal)(kaikake + t.支払金額);
                }
                else
                {
                    zan = ((decimal)zan + kaikake - t.支払金額);
                    g[colZandaka, iX].Value = zan.ToString("#,##0");
                }

                if (t.区分 == 1)
                {
                    if (t.受注1Row == null)
                    {
                        g[colMemo, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colMemo, iX].Value = t.受注1Row.特記事項;
                    }
                }
                else if (t.区分 == 2)
                {
                    if (t.外注支払Row == null)
                    {
                        g[colMemo, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colMemo, iX].Value = t.外注支払Row.備考;
                    }
                }

                g[colOrderNum, iX].Value = t.受注番号.ToString();
                g[colID, iX].Value = t.支払ID.ToString();

                // 月間要素
                sDt = t.日付;
                mKaikake += kaikake;
                mShiharai += t.支払金額;

                // 総計要素
                tKaikake += (int)kaikake;
                tShiharai += t.支払金額;
            }

            // 月合計
            if (g.RowCount > 0)
            {
                monthTotal(g, sDt.Month + 1, ref mKaikake, ref mShiharai, sDt);
            }

            //// 総計
            //g.Rows.Add();
            //iX = g.RowCount - 1;
            //g[colClient, iX].Value = "総合計";
            //g[colKaikake, iX].Value = tKaikake.ToString("#,###");
            //g[colKingaku, iX].Value = tShiharai.ToString("#,###");

            Cursor = Cursors.Default;

            if (g.RowCount == 0)
            {
                MessageBox.Show("対象となるデータがありませんでした", "買掛元帳", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
            }
            else
            {
                g.CurrentCell = g[col, row];
                g.CurrentCell = null;
                button2.Enabled = true;
            }
        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     明細開始時の前月繰越行を作成する </summary>
        /// <param name="gr">
        ///     データグリッドビューオブジェクト </param>
        /// <param name="dt">
        ///     データ日付 </param>
        /// -----------------------------------------------------------------
        private decimal setKurikoshi(DataGridView gr, DateTime dt, string sGaichu)
        {
            decimal sKaikake = 0;
            int sShiharai = 0;
            decimal sKurikoshi = 0;

            // 指定期間以前の繰越金額を求める
            if (dateTimePicker1.Checked)
            {
                var s = dts.買掛元帳.Where(a => a.日付 < DateTime.Parse(dateTimePicker1.Value.ToShortDateString()))
                        .GroupBy(a => new { a.外注先Row.名称 })
                        .Select(cg => new
                        {
                            cName = cg.Key,
                            cKbn = cg.GroupBy(a => new { a.区分 })
                                    .Select(g => new
                                    {
                                        g.Key.区分,
                                        //kaikake = g.Sum(a => (decimal)a.枚数 * a.外注原価支払),
                                        // 2015/12/06 外注原価を単価から原価総額入力へ変更に伴う
                                        kaikake = g.Sum(a => a.外注原価支払),
                                        shiharai = g.Sum(a => a.支払金額)
                                    })

                        });

                foreach (var t in s)
                {
                    // 指定以外の支払先はネグる
                    if (sGaichu != string.Empty)
                    {
                        if (!t.cName.名称.Contains(sGaichu))
                        {
                            continue;
                        }
                    }

                    foreach (var j in t.cKbn)
                    {
                        sKaikake += j.kaikake;
                        sShiharai += j.shiharai;
                    }
                }

                decimal z = 0;

                // 税込のとき指定期間以前の消費税を求める
                if (cmbTax.SelectedIndex == 0)
                {
                    // 指定期間以前の消費税を求める
                    var k = dts.買掛元帳.Where(a => a.日付 < DateTime.Parse(dateTimePicker1.Value.ToShortDateString()));

                    foreach (var t in k)
                    {
                        // 指定以外の支払先はネグる
                        if (sGaichu != string.Empty)
                        {
                            if (!t.外注先Row.名称.Contains(sGaichu))
                            {
                                continue;
                            }
                        }

                        // 税率取得
                        cTax.Ritsu = Utility.GetTaxRT(t.日付);

                        // 消費税額計算 
                        //decimal kaikake = (decimal)t.枚数 * t.外注原価支払;

                        // 2015/12/06 外注原価を単価から原価総額入力へ変更に伴う
                        decimal kaikake = t.外注原価支払;
                        z += Utility.GetTax(kaikake, cTax.Ritsu);
                    }
                }

                // 繰越金額
                sKurikoshi = sKaikake + z - sShiharai;
            }

            // 翌月繰越行
            gr.Rows.Add();
            int iX = gr.RowCount - 1;

            gr[colDt, iX].Value = dt.Year.ToString() + "/" + dt.Month.ToString().PadLeft(2, '0') + "/01";
            gr[colClient, iX].Value = "前月繰越";
            gr[colKaikake, iX].Value = sKurikoshi.ToString("#,##0");
            gr[colZandaka, iX].Value = sKurikoshi.ToString("#,##0");
            gr[colID, iX].Value = 0;
            gr[colOrderNum, iX].Value = 0;

            return sKurikoshi;
        }


        /// ------------------------------------------------------------------------
        /// <summary>
        ///     月間合計・翌月繰越行をデータグリッドビューに作成 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト </param>
        /// <param name="sMonth">
        ///     次データの月 </param>
        /// <param name="mKaikake">
        ///     月間買掛合計 </param>
        /// <param name="mShiharai">
        ///     月間支払合計 </param>
        /// <param name="dt">
        ///     現データ日付 </param>
        /// ------------------------------------------------------------------------
        private void monthTotal(DataGridView g, int sMonth, ref decimal mKaikake, ref int mShiharai, DateTime dt)
        {
            int iX = 0;
            decimal kurikoshi = 0;
            int mSpan = 0;
            DateTime bDt = DateTime.Parse(dt.Year.ToString() + "/" + dt.Month.ToString() + "/01");

            // 現在の月と次データの月のスパン
            if (dt.Month > sMonth)
            {
                mSpan = sMonth - dt.Month + 12;
            }
            else
            {
                mSpan = sMonth - dt.Month;
            }

            // 次のデータ月まで「月間合計」「前月繰越」行を作成する
            for (int i = 0; i < mSpan; i++)
            {
                // 次月繰越行
                g.Rows.Add();
                iX = g.RowCount - 1;
                DateTime tlDt = bDt.AddMonths(i + 1).AddDays(-1);       // 月末日
                g[colDt, iX].Value = tlDt.ToShortDateString();
                g[colName, iX].Value = "次月繰越";
                kurikoshi = mKaikake - (decimal)mShiharai;
                g[colKingaku, iX].Value = kurikoshi.ToString("#,##0");
                g[colID, iX].Value = 0;
                g[colOrderNum, iX].Value = 0;
                //g[colZandaka, iX].Value = kurikoshi.ToString("#,##0");
                g.Rows[iX].DefaultCellStyle.ForeColor = Color.Red;

                // 合計行
                g.Rows.Add();
                iX = g.RowCount - 1;
                //DateTime tlDt = bDt.AddMonths(i + 1).AddDays(-1);       // 月末日
                //g[colDt, iX].Value = tlDt.ToShortDateString();
                //g[colName, iX].Value = dt.AddMonths(i).Month.ToString() + "月合計";
                g[colKaikake, iX].Value = mKaikake.ToString("#,##0");
                g[colKingaku, iX].Value = (mShiharai + kurikoshi).ToString("#,##0");
                g[colID, iX].Value = 0;
                g[colOrderNum, iX].Value = 0;
                //kurikoshi = mKaikake - (decimal)mShiharai;
                //g[colZandaka, iX].Value = kurikoshi.ToString("#,##0");
                g.Rows[iX].DefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 翌月繰越行
                g.Rows.Add();
                iX = g.RowCount - 1;

                DateTime nDt = bDt.AddMonths(i + 1);        // 翌月の1日を取得
                g[colDt, iX].Value = nDt.ToShortDateString();
                g[colClient, iX].Value = "前月繰越";
                g[colKaikake, iX].Value = kurikoshi.ToString("#,##0");
                g[colZandaka, iX].Value = kurikoshi.ToString("#,##0");
                g[colID, iX].Value = 0;
                g[colOrderNum, iX].Value = 0;

                mKaikake = kurikoshi;
                mShiharai = 0;
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
            dateTimePicker1.Checked = false;
            dateTimePicker2.Checked = false;
            button2.Enabled = false;
            txtsGaichu.Text = string.Empty;
            cmbTax.SelectedIndex = 1;
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

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // 受注番号を取得
            Int64 orderNum = Int64.Parse(dataGridView1[colOrderNum, dataGridView1.SelectedRows[0].Index].Value.ToString());

            // 支払IDを取得
            int sID = int.Parse(dataGridView1[colID, dataGridView1.SelectedRows[0].Index].Value.ToString());

            // 受注確定書登録画面
            if (orderNum != 0)
            {
                frmOrder frm = new frmOrder(orderNum);
                frm.ShowDialog();
            }

            // 支払金額入力画面
            if (sID != 0)
            {
                frmShiharai frm = new frmShiharai(sID);
                frm.ShowDialog();
            }

            // データグリッドビューデータ表示
            gridShow(dataGridView1, 0, e.RowIndex);
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
            // データグリッドビューデータ表示
            gridShow(dataGridView1, 0, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "買掛元帳");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            frmKaikakeZan frm = new frmKaikakeZan();
            frm.ShowDialog();
            Show();
        }
    }
}

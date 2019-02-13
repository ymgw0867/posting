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
    public partial class frmUrikakeMotocho : Form
    {
        public frmUrikakeMotocho()
        {
            InitializeComponent();

            // データ読み込み
            uAdp.Fill(dts.得意先);
            rAdp.Fill(dts.新請求書);
            nAdp.Fill(dts.新入金);
            adp.Fill(dts.売掛元帳);
            jAdp.Fill(dts.受注1);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.売掛元帳TableAdapter adp = new darwinDataSetTableAdapters.売掛元帳TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter uAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.新請求書TableAdapter rAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.新入金TableAdapter nAdp = new darwinDataSetTableAdapters.新入金TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

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
        string colUrikake = "col10";
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
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                tempDGV.Columns.Add(colDt, "日付");
                tempDGV.Columns.Add(colClient, "請求先");
                tempDGV.Columns.Add(colName, "内容");
                tempDGV.Columns.Add(colUrikake, "売掛金額");
                tempDGV.Columns.Add(colKingaku, "入金額");
                tempDGV.Columns.Add(colZandaka, "残高");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colID, "入金ID");
                tempDGV.Columns.Add(colSeikyuNum, "請求書ID");
                tempDGV.Columns.Add(colKbn, "区分");
                //tempDGV.Columns.Add(colAddDt, "登録年月日");
                //tempDGV.Columns.Add(colUpDt, "更新年月日");
                //tempDGV.Columns.Add(colUserID, "ユーザーID");

                tempDGV.Columns[colID].Visible = false;
                tempDGV.Columns[colSeikyuNum].Visible = false;
                tempDGV.Columns[colKbn].Visible = false;

                tempDGV.Columns[colDt].Width = 100;
                tempDGV.Columns[colClient].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colName].Width = 260;
                tempDGV.Columns[colUrikake].Width = 110;
                tempDGV.Columns[colKingaku].Width = 110;
                tempDGV.Columns[colZandaka].Width = 110;
                tempDGV.Columns[colMemo].Width = 160;

                tempDGV.Columns[colDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colUrikake].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
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
            // グリッドビュー
            g.Rows.Clear();

            int iX = 0;
            decimal zan = 0;
            DateTime sDt = DateTime.Parse("1900/01/01");
            decimal mUrikake = 0;
            int mNyukin = 0;
            int tUrikake = 0;
            int tNyukin = 0;
            bool firstData = true;
            
            // グリッドに指定期間の明細表示
            //var s = dts.売掛元帳.OrderBy(a => a.日付); 

            // 2015/12/09
            var s = dts.売掛元帳
                .Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0)
                .OrderBy(a => a.日付);

            //s = s.Where(a => a.新請求書Row.Get受注1Rows().Count() > 0).OrderBy(a => a.日付);


            // 得意先指定
            if (txtsClient.Text.Trim() != string.Empty)
            {
                //s = s.Where(a => a.得意先Row != null && a.得意先Row.請求先名称.Contains(txtsClient.Text)).OrderBy(a => a.日付);
                
                // 売掛先の名称を請求先名称から得意先略称に変更 2016/03/11
                s = s.Where(a => a.得意先Row != null && a.得意先Row.略称.Contains(txtsClient.Text)).OrderBy(a => a.日付);
            }

            // 開始日付
            if (dateTimePicker1.Checked)
            {
                s = s.Where(a => a.日付 >= DateTime.Parse(dateTimePicker1.Value.ToShortDateString())).OrderBy(a => a.日付);
            }

            // 終了日付
            if (dateTimePicker2.Checked)
            {
                s = s.Where(a => a.日付 <= DateTime.Parse(dateTimePicker2.Value.ToShortDateString())).OrderBy(a => a.日付);
            }

            foreach (var t in s)
            {
                // 前月繰越
                if (firstData)
                {
                    mUrikake += setKurikoshi(dataGridView1, t.日付, txtsClient.Text);
                    zan += mUrikake;
                    firstData = false;
                }

                // 月が変わると月間合計
                if (sDt.Year != 1900 && sDt.Month != t.日付.Month)
                {
                    monthTotal(g, t.日付.Month, ref mUrikake, ref mNyukin, sDt);
                }

                // 明細
                g.Rows.Add();
                iX = g.RowCount - 1;

                g[colDt, iX].Value = t.日付.ToShortDateString();  // 日付

                // 請求先名称
                if (t.得意先Row == null)
                {
                    g[colClient, iX].Value = string.Empty;
                }
                else
                {
                    // 売掛先名称：略称に変更 2016/03/10
                    g[colClient, iX].Value = t.得意先Row.略称;

                    //if (t.得意先Row.請求先名称 == string.Empty)
                    //{
                    //    g[colClient, iX].Value = t.得意先Row.名称;
                    //}
                    //else
                    //{
                    //    g[colClient, iX].Value = t.得意先Row.請求先名称;
                    //}
                }
                
                // 内容
                if (t.区分 == 1)
                {
                    foreach (var n in t.新請求書Row.Get受注1Rows().OrderBy(a => a.ID))
                    {
                        g[colName, iX].Value = n.チラシ名;
                        break;
                    }
                }
                
                g[colUrikake, iX].Value = t.売掛金額.ToString("#,###");     // 売掛金
                g[colKingaku, iX].Value = t.入金額.ToString("#,###");       // 入金額

                // 残高
                if (iX == 0)
                {
                    g[colZandaka, iX].Value = (t.売掛金額 + t.入金額).ToString("#,##0");
                    zan += (decimal)(t.売掛金額 + t.入金額);
                }
                else
                {
                    zan = ((decimal)zan + t.売掛金額 - t.入金額);
                    g[colZandaka, iX].Value = zan.ToString("#,##0");
                }

                // 備考欄：売掛金
                if (t.区分 == 1)
                {
                    if (t.新請求書Row == null)
                    {
                        g[colMemo, iX].Value = string.Empty;
                    }
                    else if (t.新請求書Row.Is備考Null())
                    {
                        g[colMemo, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colMemo, iX].Value = t.新請求書Row.備考;
                    }
                }

                // 備考欄：入金
                if (t.区分 == 2)
                {
                    if (t.新入金Row == null)
                    {
                        g[colMemo, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colMemo, iX].Value = t.新入金Row.備考;
                    }
                }
                
                // 備考欄：精算
                if (t.区分 == 3)
                {
                    if (t.新請求書Row == null)
                    {
                        g[colMemo, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colMemo, iX].Value = t.新請求書Row.精算備考;
                    }
                }


                g[colSeikyuNum, iX].Value = t.請求書ID.ToString();
                g[colID, iX].Value = t.新入金ID.ToString();
                g[colKbn, iX].Value = t.区分.ToString();

                // 月間要素
                sDt = t.日付;
                mUrikake += t.売掛金額;
                mNyukin += t.入金額;

                // 総計要素
                tUrikake += t.売掛金額;
                tNyukin += t.入金額;
            }

            // 月合計
            if (g.RowCount > 0)
            {
                monthTotal(g, sDt.Month + 1, ref mUrikake, ref mNyukin, sDt);
            }

            //// 総計
            //g.Rows.Add();
            //iX = g.RowCount - 1;
            //g[colClient, iX].Value = "総合計";
            //g[colKaikake, iX].Value = tKaikake.ToString("#,###");
            //g[colKingaku, iX].Value = tShiharai.ToString("#,###");

            if (g.RowCount == 0)
            {
                MessageBox.Show("対象となるデータがありませんでした", "売掛元帳", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
            }
            else
            {
                g.CurrentCell = null;
                button2.Enabled = true;

                // カレント行
                g.CurrentCell = g[col, row];
                g.CurrentCell = null;
            }
        }
        
        /// -----------------------------------------------------------------
        /// <summary>
        ///     明細開始時の前月繰越行を作成する </summary>
        /// <param name="gr">
        ///     データグリッドビューオブジェクト </param>
        /// <param name="dt">
        ///     データ日付 </param>
        /// <param name="sClient">
        ///     指定請求先</param>
        /// -----------------------------------------------------------------
        private decimal setKurikoshi(DataGridView gr, DateTime dt, string sClient)
        {
            decimal sUrikake = 0;
            int sNyukin = 0;
            decimal sKurikoshi = 0;

            // 指定期間以前の繰越金額を求める : 得意先略称を使用 2016/03/11
            // 条件式：「a.得意先Row != null」を追加 2018/04/09
            if (dateTimePicker1.Checked)
            {
                var s = dts.売掛元帳.Where(a => a.日付 < DateTime.Parse(dateTimePicker1.Value.ToShortDateString()) && a.得意先Row != null)
                        .GroupBy(a => new { a.得意先Row.略称 })
                        .Select(cg => new
                        {
                            cName = cg.Key,
                            cKbn = cg.GroupBy(a => new { a.区分 })
                                    .Select(g => new
                                    {
                                        g.Key.区分,
                                        urikake = g.Sum(a => a.売掛金額),
                                        nyukin = g.Sum(a => a.入金額)
                                    })

                        });
                
                foreach (var t in s)
                {
                    // 指定以外の得意先はネグる
                    if (sClient != string.Empty)
                    {
                        if (!t.cName.略称.Contains(sClient)) // 得意先略称を使用 2016/03/11
                        {
                            continue;
                        }
                    }

                    foreach (var j in t.cKbn)
                    {
                        sUrikake += j.urikake;
                        sNyukin += j.nyukin;
                    }
                }

                sKurikoshi = sUrikake - sNyukin;
            }
            
            // 翌月繰越行
            gr.Rows.Add();
            int iX = gr.RowCount - 1;

            gr[colDt, iX].Value = dt.Year.ToString() + "/" + dt.Month.ToString().PadLeft(2, '0') + "/01";
            gr[colClient, iX].Value = "前月繰越";
            gr[colUrikake, iX].Value = sKurikoshi.ToString("#,##0");
            gr[colZandaka, iX].Value = sKurikoshi.ToString("#,##0");
            gr[colID, iX].Value = 0;
            gr[colSeikyuNum, iX].Value = 0;

            return sKurikoshi;
        }


        /// ------------------------------------------------------------------------
        /// <summary>
        ///     月間合計・翌月繰越行をデータグリッドビューに作成 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト </param>
        /// <param name="sMonth">
        ///     次データの月 </param>
        /// <param name="mUrikake">
        ///     月売掛合計 </param>
        /// <param name="mNyukin">
        ///     月間支払合計 </param>
        /// <param name="dt">
        ///     現データ日付 </param>
        /// ------------------------------------------------------------------------
        private void monthTotal(DataGridView g, int sMonth, ref decimal mUrikake, ref int mNyukin, DateTime dt)
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
                kurikoshi = mUrikake - (decimal)mNyukin;
                g[colKingaku, iX].Value = kurikoshi.ToString("#,##0");
                g[colID, iX].Value = 0;
                g[colSeikyuNum, iX].Value = 0;
                //g[colZandaka, iX].Value = kurikoshi.ToString("#,##0");
                g.Rows[iX].DefaultCellStyle.ForeColor = Color.Red;

                // 合計行
                g.Rows.Add();
                iX = g.RowCount - 1;
                //DateTime tlDt = bDt.AddMonths(i + 1).AddDays(-1);       // 月末日
                //g[colDt, iX].Value = tlDt.ToShortDateString();
                //g[colName, iX].Value = dt.AddMonths(i).Month.ToString() + "月合計";
                g[colUrikake, iX].Value = mUrikake.ToString("#,##0");
                g[colKingaku, iX].Value = (mNyukin + kurikoshi).ToString("#,##0");
                g[colID, iX].Value = 0;
                g[colSeikyuNum, iX].Value = 0;
                //kurikoshi = mKaikake - (decimal)mShiharai;
                //g[colZandaka, iX].Value = kurikoshi.ToString("#,##0");
                g.Rows[iX].DefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 翌月繰越行
                g.Rows.Add();
                iX = g.RowCount - 1;

                DateTime nDt = bDt.AddMonths(i + 1);        // 翌月の1日を取得
                g[colDt, iX].Value = nDt.ToShortDateString();
                g[colClient, iX].Value = "前月繰越";
                g[colUrikake, iX].Value = kurikoshi.ToString("#,##0");
                g[colZandaka, iX].Value = kurikoshi.ToString("#,##0");
                g[colID, iX].Value = 0;
                g[colSeikyuNum, iX].Value = 0;

                mUrikake = kurikoshi;
                mNyukin = 0;
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
            int r = e.RowIndex;
            
            if (dataGridView1[colKbn, dataGridView1.SelectedRows[0].Index].Value == null)
            {
                return;
            }

            // 区分取得
            int kbn = int.Parse(dataGridView1[colKbn, dataGridView1.SelectedRows[0].Index].Value.ToString());

            // 請求書IDを取得
            int seikyuNum = int.Parse(dataGridView1[colSeikyuNum, dataGridView1.SelectedRows[0].Index].Value.ToString());
            
            //// 入金IDを取得
            //int sID = int.Parse(dataGridView1[colID, dataGridView1.SelectedRows[0].Index].Value.ToString());

            // 請求詳細画面
            if (seikyuNum != 0)
            {
                if (kbn == 1)
                {
                    frmSeikyuItem2015 frm = new frmSeikyuItem2015(seikyuNum);
                    frm.ShowDialog();

                    // データ再読み込み
                    rAdp.Fill(dts.新請求書);
                }

                // 入金登録画面：精算額入力画面 (2016/05/23)
                if (kbn == 2 || kbn == 3)
                {
                    frmNyukinItem2015 frm = new frmNyukinItem2015(seikyuNum);
                    frm.ShowDialog();

                    // データ再読み込み
                    nAdp.Fill(dts.新入金);
                }
            }

            // データグリッドビューデータ再表示
            adp.Fill(dts.売掛元帳);
            gridShow(dataGridView1, 0, r);
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
            MyLibrary.CsvOut.GridView(dataGridView1, "売掛元帳");
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            frmUrikakeZan frm = new frmUrikakeZan();
            frm.ShowDialog();
            Show();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            frmUriageOrderKbn frm = new frmUriageOrderKbn();
            frm.ShowDialog();
            Show();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            frmUrikakeNen frm = new frmUrikakeNen();
            frm.ShowDialog();
            Show();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            frmUriageGaichu frm = new frmUriageGaichu();
            frm.ShowDialog();
            Show();
        }
    }
}

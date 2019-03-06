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
    public partial class frmUrikakeNen : Form
    {
        public frmUrikakeNen()
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
        string colID = "col1";
        string colClient = "col2";
        string colUrikake = "col8";
        string colTo = "col3";
        string col1 = "col4";
        string col2 = "col5";
        string col3 = "col6";
        string col4 = "col7";
        #endregion

        public class urikakeMonth
        {
            public int clientID = 0;
            public string clientName = "";
            public int uriZan = 0;
            public int uriTo = 0;
            public int nyuTo = 0;
            public int[] uri = new int [4];
            public int[] nyu = new int[4];
        }

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
                tempDGV.Columns.Add(colClient, "請求先");
                tempDGV.Columns.Add(colUrikake, "売掛残");
                tempDGV.Columns.Add(colTo, "当月");
                tempDGV.Columns.Add(col1, "１ヶ月経過");
                tempDGV.Columns.Add(col2, "２ヶ月経過");
                tempDGV.Columns.Add(col3, "３ヶ月経過");
                tempDGV.Columns.Add(col4, "４ヶ月以上");

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
                tempDGV.Columns[colUrikake].Width = 110;
                tempDGV.Columns[colTo].Width = 110;
                tempDGV.Columns[col1].Width = 110;
                tempDGV.Columns[col2].Width = 110;
                tempDGV.Columns[col3].Width = 110;
                tempDGV.Columns[col4].Width = 100;
                
                tempDGV.Columns[colID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colUrikake].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colTo].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[col1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[col2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[col3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[col4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
            g.RowCount = 0;

            urikakeMonth[] um = null;

            int iX = 0;
            DateTime sDt = DateTime.Parse("1900/01/01");
            DateTime sDate = DateTime.Parse(txtYear.Text + "/" + txtMonth.Text + "/" + txtDay.Text);
            decimal zenKurikoshi = 0;
            decimal _toTotal = 0;
            decimal _1Total = 0;
            decimal _2Total = 0;
            decimal _3Total = 0;
            decimal _4Total = 0;
            decimal zandakaTotal = 0;
            int yy = 0;
            int mm = 0;

            foreach (var it in dts.得意先.OrderBy(a => a.ID))
            {
                // 指定以外の得意先はネグる
                if (txtsClient.Text != string.Empty)
                {
                    if (!it.略称.Contains(txtsClient.Text)) // 得意先略称を使用 2016/03/11
                    {
                        continue;
                    }
                }

                DateTime zDt;   // 前月
                DateTime zDt2;  // 2ヶ月
                DateTime zDt3;  // 3ヶ月
                DateTime zDt4;  // 4ヶ月

                if (DateTime.DaysInMonth(sDate.Year, sDate.Month) == sDate.Day)
                {
                    // 基準日が月末日のとき
                    zDt = getDayInMonth(sDate, -1);
                    zDt2 = getDayInMonth(sDate, -2);
                    zDt3 = getDayInMonth(sDate, -3);
                    zDt4 = getDayInMonth(sDate, -4);
                }
                else
                {
                    zDt = sDate.AddMonths(-1);   // 前月
                    zDt2 = sDate.AddMonths(-2);  // 2ヶ月
                    zDt3 = sDate.AddMonths(-3);  // 3ヶ月
                    zDt4 = sDate.AddMonths(-4);  // 4ヶ月
                }                

                // 前月繰越
                zenKurikoshi = setKurikoshi(zDt, it.ID);

                // 当月売掛と入金
                //var s = dts.売掛元帳.Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0 && 
                //                          a.得意先ID == it.ID && a.日付 <= sDate && a.日付 > zDt)
                //            .GroupBy(a => new { a.得意先Row.略称 })
                //            .Select(cg => new
                //            {
                //                cName = cg.Key,
                //                urikake = cg.Sum(a => a.売掛金額),
                //                nyukin = cg.Sum(a => a.入金額)
                //            });

                var s = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0 && 
                                          a.得意先Row != null && a.得意先ID == it.ID &&
                                          a.請求書発行日 <= sDate && a.請求書発行日 > zDt)
                           .GroupBy(a => new { a.得意先Row.略称 })
                           .Select(cg => new
                           {
                               cName = cg.Key,
                               urikake = cg.Sum(a => a.請求金額) + cg.Sum(a => a.精算額) - cg.Sum(a => a.Get新入金Rows().Sum(n => n.金額))
                           });
                
                bool tougetsu = false;

                foreach (var t in s)
                {
                    tougetsu = true;

                    Array.Resize(ref um, iX + 1);
                    um[iX] = new urikakeMonth();

                    um[iX].clientID = it.ID;
                    um[iX].clientName = t.cName.略称;
                    um[iX].uriZan = (int)(t.urikake + zenKurikoshi);
                    um[iX].uriTo = t.urikake;

                    // 合計値加算
                    zandakaTotal += t.urikake + zenKurikoshi;
                }

                // 当月取引なし
                if (!tougetsu)
                {
                    // 前月繰越があるとき表示する
                    if (zenKurikoshi == 0)
                    {
                        continue;
                    }
                    else
                    {
                        Array.Resize(ref um, iX + 1);
                        um[iX] = new urikakeMonth();

                        um[iX].clientID = it.ID;
                        um[iX].clientName = it.略称;
                        um[iX].uriZan = (int)(zenKurikoshi);
                        um[iX].uriTo = 0;

                        // 残高合計加算
                        zandakaTotal += zenKurikoshi;
                    }
                }

                // 前月売掛と入金
                //var z1 = dts.売掛元帳.Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0 &&
                //                          a.得意先ID == it.ID && a.日付 <= zDt && a.日付 > zDt2)
                //            .GroupBy(a => new { a.得意先Row.略称 })
                //            .Select(cg => new
                //            {
                //                cName = cg.Key,
                //                urikake = cg.Sum(a => a.売掛金額),
                //                nyukin = cg.Sum(a => a.入金額)
                //            });

                var z1 = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0 && 
                                          a.得意先Row != null && a.得意先ID == it.ID &&
                                          a.請求書発行日 <= zDt && a.請求書発行日 > zDt2)
                           .GroupBy(a => new { a.得意先Row.略称 })
                           .Select(cg => new
                           {
                               cName = cg.Key,
                               urikake = cg.Sum(a => a.請求金額) + cg.Sum(a => a.精算額) - cg.Sum(a => a.Get新入金Rows().Sum(n => n.金額))
                           });

                foreach (var t in z1)
                {
                    um[iX].uri[0] = t.urikake;
                }
                
                // ２ヶ月経過売掛と入金
                //var z2 = dts.売掛元帳.Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0 &&
                //                          a.得意先ID == it.ID && a.日付 <= zDt2 && a.日付 > zDt3)
                //            .GroupBy(a => new { a.得意先Row.略称 })
                //            .Select(cg => new
                //            {
                //                cName = cg.Key,
                //                urikake = cg.Sum(a => a.売掛金額),
                //                nyukin = cg.Sum(a => a.入金額)
                //            });
                
                var z2 = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0 && 
                                          a.得意先Row != null && a.得意先ID == it.ID &&
                                          a.請求書発行日 <= zDt2 && a.請求書発行日 > zDt3)
                           .GroupBy(a => new { a.得意先Row.略称 })
                           .Select(cg => new
                           {
                               cName = cg.Key,
                               urikake = cg.Sum(a => a.請求金額) + cg.Sum(a => a.精算額) - cg.Sum(a => a.Get新入金Rows().Sum(n => n.金額))
                           });

                foreach (var t in z2)
                {
                    um[iX].uri[1] = t.urikake;
                }

                // ３ヶ月経過売掛と入金
                //var z3 = dts.売掛元帳.Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0 &&
                //                          a.得意先ID == it.ID && a.日付 <= zDt3 && a.日付 > zDt4)
                //            .GroupBy(a => new { a.得意先Row.略称 })
                //            .Select(cg => new
                //            {
                //                cName = cg.Key,
                //                urikake = cg.Sum(a => a.売掛金額),
                //                nyukin = cg.Sum(a => a.入金額)
                //            });

                var z3 = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0 && 
                                          a.得意先Row != null && a.得意先ID == it.ID &&
                                          a.請求書発行日 <= zDt3 && a.請求書発行日 > zDt4)
                           .GroupBy(a => new { a.得意先Row.略称 })
                           .Select(cg => new
                           {
                               cName = cg.Key,
                               urikake = cg.Sum(a => a.請求金額) + cg.Sum(a => a.精算額) - cg.Sum(a => a.Get新入金Rows().Sum(n => n.金額))
                           });

                foreach (var t in z3)
                {
                    um[iX].uri[2] = t.urikake;
                }


                // ４ヶ月以上経過売掛と入金
                //var z4 = dts.売掛元帳.Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0 &&
                //                          a.得意先ID == it.ID && a.日付 <= zDt4)
                //            .GroupBy(a => new { a.得意先Row.略称 })
                //            .Select(cg => new
                //            {
                //                cName = cg.Key,
                //                urikake = cg.Sum(a => a.売掛金額),
                //                nyukin = cg.Sum(a => a.入金額)
                //            });

                var z4 = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0 && 
                                          a.得意先Row != null && a.得意先ID == it.ID &&
                                          a.請求書発行日 <= zDt4)
                           .GroupBy(a => new { a.得意先Row.略称 })
                           .Select(cg => new
                           {
                               cName = cg.Key,
                               urikake = cg.Sum(a => a.請求金額) + cg.Sum(a => a.精算額) - cg.Sum(a => a.Get新入金Rows().Sum(n => n.金額))
                           });


                foreach (var t in z4)
                {
                    um[iX].uri[3] = t.urikake;
                }
                                
                iX++;
            }

            if (um != null)
            {
                foreach (var t in um)
                {
                    g.Rows.Add();
                    g[colID, g.RowCount - 1].Value = t.clientID;
                    g[colClient, g.RowCount - 1].Value = t.clientName;
                    g[colUrikake, g.RowCount - 1].Value = t.uriZan.ToString("#,0");
                    //g[colTo, g.RowCount - 1].Value = t.uriTo.ToString("#,0");
                    g[colTo, g.RowCount - 1].Value = t.uriTo.ToString("#,0");
                    g[col1, g.RowCount - 1].Value = t.uri[0].ToString("#,0");
                    g[col2, g.RowCount - 1].Value = t.uri[1].ToString("#,0");
                    g[col3, g.RowCount - 1].Value = t.uri[2].ToString("#,0");
                    g[col4, g.RowCount - 1].Value = t.uri[3].ToString("#,0");

                    _toTotal += t.uriTo;
                    _1Total += t.uri[0];
                    _2Total += t.uri[1];
                    _3Total += t.uri[2];
                    _4Total += t.uri[3];
                }
            }

            // 総計
            if (g.RowCount > 0)
            {
                g.Rows.Add();
                iX = g.RowCount - 1;
                g[colID, iX].Value = "";
                g[colClient, iX].Value = "総合計";
                g[colUrikake, iX].Value = zandakaTotal.ToString("#,0");
                g[colTo, g.RowCount - 1].Value = _toTotal.ToString("#,0");
                g[col1, g.RowCount - 1].Value = _1Total.ToString("#,0");
                g[col2, g.RowCount - 1].Value = _2Total.ToString("#,0");
                g[col3, g.RowCount - 1].Value = _3Total.ToString("#,0");
                g[col4, g.RowCount - 1].Value = _4Total.ToString("#,0");

                g.CurrentCell = null;
                button2.Enabled = true;

                // カレント行
                g.CurrentCell = g[col, row];
                g.CurrentCell = null;
            }
            else
            {
                MessageBox.Show("対象となるデータがありませんでした", "売掛月齢表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
            }

            // カーソルを戻す
            Cursor = Cursors.Default;
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     任意の日の任意の月数後の月末日を返す </summary>
        /// <param name="sDate">
        ///     任意の日付</param>
        /// <param name="n">
        ///     加減月数</param>
        /// <returns>
        ///     月末日付</returns>
        ///-----------------------------------------------------------------
        private DateTime getDayInMonth(DateTime sDate, int n)
        {
            // 基準日が月末日のとき
            DateTime dt = sDate.AddMonths(n);
            return DateTime.Parse(dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + (DateTime.DaysInMonth(dt.Year, dt.Month)).ToString());   // 前月

        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     明細開始時の前月繰越金額を取得する </summary>
        /// <param name="dt">
        ///     データ日付 </param>
        /// <param name="sClient">
        ///     指定請求先</param>
        /// -----------------------------------------------------------------
        private decimal setKurikoshi(DateTime dt, int sClient)
        {
            decimal sUrikake = 0;
            int sNyukin = 0;
            decimal sKurikoshi = 0;

            // 指定期間以前の繰越金額を求める : 得意先略称を使用 2016/03/11 : 2016/06/24
            //var s = dts.売掛元帳.Where(a => a.新請求書Row != null && a.新請求書Row.Get受注1Rows().Count() > 0 && a.日付 < dt && a.得意先Row.ID == sClient)
            //        .GroupBy(a => new { a.得意先Row.略称 })
            //        .Select(cg => new
            //        {
            //            cName = cg.Key,
            //            urikake = cg.Sum(a => a.売掛金額),
            //            nyukin = cg.Sum(a => a.入金額)
            //        });

            var s = dts.新請求書.Where(a => a.Get受注1Rows().Count() > 0 &&
                                      a.得意先Row != null && a.得意先ID == sClient &&
                                      a.請求書発行日 <= dt)
                       .GroupBy(a => new { a.得意先Row.略称 })
                       .Select(cg => new
                       {
                           cName = cg.Key,
                           urikake = cg.Sum(a => a.請求金額) + cg.Sum(a => a.精算額) - cg.Sum(a => a.Get新入金Rows().Sum(n => n.金額))
                       });

            foreach (var t in s)
            {
                sUrikake += t.urikake;
            }

            sKurikoshi = sUrikake;

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
            txtDay.Text = DateTime.Now.Day.ToString();
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
            // 売掛残高表示
            dataShow();
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     売掛残高表示 </summary>
        ///--------------------------------------------------------
        private void dataShow()
        {
            DateTime dt;
            if (!DateTime.TryParse(txtYear.Text + "/" + txtMonth.Text + "/" + txtDay.Text, out dt))
            {
                MessageBox.Show("基準日が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return;
            }

            // データグリッドビューデータ表示
            gridShow(dataGridView1, 0, 0);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "売掛月齢表");
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

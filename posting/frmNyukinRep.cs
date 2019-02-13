using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using MyLibrary;
using System.Linq;

namespace posting
{
    public partial class frmNyukinRep : Form
    {
        const string MESSAGE_CAPTION = "クライアント別請求一覧";

        public frmNyukinRep()
        {
            InitializeComponent();

            // データ読み込み
            rAdp.Fill(dts.新請求書);
            nAdp.Fill(dts.新入金);
            cAdp.Fill(dts.得意先);
            sAdp.Fill(dts.社員);
            jAdp.Fill(dts.受注1);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.新請求書TableAdapter rAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.新入金TableAdapter nAdp = new darwinDataSetTableAdapters.新入金TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();

        bool mukouStatus = false;

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            ////不要請求書データ削除 2010/05/31
            //DataDelete();

            //グリッド定義
            GridviewSet.Setting(dataGridView1);

            //画面クリア
            DispClear();
        }

        private void DataDelete()
        {
            string sqlString;

            Control.FreeSql fSql = new Control.FreeSql();
            
            //sqlSTRING += " and ((請求書.ID != 709) and (請求書.ID != 771) and (請求書.ID != 723) and 
            //(請求書.ID != 822) and (請求書.ID != 765) and (請求書.ID != 776) and (請求書.ID != 743) and
            //(請求書.ID != 899) and (請求書.ID != 1017) and (請求書.ID != 847) and (請求書.ID != 849) and
            //(請求書.ID != 990) and (請求書.ID != 1190)) ";

            //ソシエワールド
            sqlString = "delete from 請求書 where ID = 709";
            fSql.Execute(sqlString);

            //ごっぱち
            sqlString = "delete from 請求書 where ID = 771";
            fSql.Execute(sqlString);

            //キャメルリンク
            sqlString = "delete from 請求書 where ID = 723";
            fSql.Execute(sqlString);

            //ＤＡテクニカルサービス
            sqlString = "delete from 請求書 where ID = 822";
            fSql.Execute(sqlString);

            //
            sqlString = "delete from 請求書 where ID = 765";
            fSql.Execute(sqlString);

            //ＨＡＫＵ
            sqlString = "delete from 請求書 where ID = 776";
            fSql.Execute(sqlString);

            //表参道接骨院
            sqlString = "delete from 請求書 where ID = 743";
            fSql.Execute(sqlString);

            //レコプロ
            sqlString = "delete from 請求書 where ID = 899";
            fSql.Execute(sqlString);

            //レコプロ
            sqlString = "delete from 請求書 where ID = 1017";
            fSql.Execute(sqlString);

            //プリズミック
            sqlString = "delete from 請求書 where ID = 847";
            fSql.Execute(sqlString);

            //ＤＡテクニカルサービス
            sqlString = "delete from 請求書 where ID = 849";
            fSql.Execute(sqlString);

            //庭の音
            sqlString = "delete from 請求書 where ID = 990";
            fSql.Execute(sqlString);

            //早稲田ゼミナール
            sqlString = "delete from 請求書 where ID = 1190";
            fSql.Execute(sqlString);

            fSql.Close();
        }

        // データグリッドビュークラス
        private class GridviewSet
        {
            /// <summary>
            /// データグリッドビューの定義を行います
            /// </summary>
            /// <param name="tempDGV">データグリッドビューオブジェクト</param>
            public static void Setting(DataGridView tempDGV)
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
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);
                    
                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 470;

                    // 奇数行の色
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "番号");
                    tempDGV.Columns.Add("col2", "クライアント名");
                    tempDGV.Columns.Add("col13", "フリガナ");
                    tempDGV.Columns.Add("col3", "担当者");
                    tempDGV.Columns.Add("col4", "発行日");
                    tempDGV.Columns.Add("col5", "請求金額");
                    tempDGV.Columns.Add("col6", "入金予定日");
                    tempDGV.Columns.Add("col7", "入金日");
                    tempDGV.Columns.Add("col8", "入金額");
                    tempDGV.Columns.Add("col9", "入金残高");

                    DataGridViewCheckBoxColumn cbc = new DataGridViewCheckBoxColumn();
                    tempDGV.Columns.Add(cbc);
                    tempDGV.Columns[10].HeaderText = "入金済";
                    tempDGV.Columns[10].Name = "col10";

                    tempDGV.Columns.Add("col11", "請求書ID");
                    tempDGV.Columns.Add("col12", "備考");   // 2015/07/22

                    //tempDGV.Columns[1].Frozen = true;   // 2015/07/22

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].Width = 220;
                    tempDGV.Columns[2].Width = 160;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 60;
                    tempDGV.Columns[12].Width = 200;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    //tempDGV.Columns[3].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[5].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    //請求書IDを非表示とする
                    tempDGV.Columns[11].Visible = false;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // 編集不可とする
                    tempDGV.ReadOnly = false;

                    // 編集可・不可の指定
                    foreach (DataGridViewColumn d in tempDGV.Columns)
                    {
                        if (d.Name == "col10" || d.Name == "col12")
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
                    tempDGV.BorderStyle = BorderStyle.Fixed3D;
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void ShowData(DataGridView tempDGV, int tempKBN, TextBox tempTxt, TextBox tempTotal, DateTime tempSDate, DateTime tempEDate,string tempCstr)
            {
                string sqlSTRING = "";
                int sZan = 0;
                int sTotal = 0;

                try
                {
                    // データリーダーを取得する
                    OleDbDataReader dR;

                    sqlSTRING += "SELECT 請求書.ID,請求書.得意先ID,請求書.請求金額,請求書.消費税,";
                    sqlSTRING += "請求書.値引額,請求書.売上金額,請求書.税率,";
                    sqlSTRING += "請求書.入金予定日,請求書.発行日,請求書.入金残,請求書.完了区分,";
                    sqlSTRING += "請求書.振込口座ID1,請求書.振込口座ID2,請求書.備考,";
                    sqlSTRING += "請求書.登録年月日,請求書.変更年月日,得意先.略称,社員.氏名,n_tbl.入金日,n_tbl.入金額 ";
                    sqlSTRING += "from 請求書 LEFT OUTER JOIN 得意先 ";
                    sqlSTRING += "ON 請求書.得意先ID = 得意先.ID LEFT OUTER JOIN 社員 ";
                    sqlSTRING += "ON 得意先.担当社員コード = 社員.ID LEFT OUTER JOIN ";
                    sqlSTRING += "(SELECT 請求書ID, MAX(入金年月日) AS 入金日,sum(金額) as 入金額 ";
                    sqlSTRING += "FROM 入金 ";
                    sqlSTRING += "GROUP BY 請求書ID) AS n_tbl ON 請求書.ID = n_tbl.請求書ID ";
                    sqlSTRING += "where ";

                    if (tempKBN == 1)
                    {
                        sqlSTRING += "(請求書.完了区分 = 0) and ";
                    }

                    sqlSTRING += "(請求書.入金予定日 >= '" + tempSDate + "') and (請求書.入金予定日 <= '" + tempEDate + "') ";

                    if (tempCstr != "")
                    {
                        sqlSTRING += " and (得意先.略称 like '%" + tempCstr + "%') " ;
                    }

                    // 不要データ 2010/02/16
                    sqlSTRING += " and ((請求書.ID != 709) and (請求書.ID != 771) and (請求書.ID != 723) and (請求書.ID != 822) and (請求書.ID != 765) and (請求書.ID != 776) and (請求書.ID != 743) and (請求書.ID != 899) and (請求書.ID != 1017) and (請求書.ID != 847) and (請求書.ID != 849) and (請求書.ID != 990) and (請求書.ID != 1190)) ";
                    
                    sqlSTRING += "ORDER BY 請求書.得意先ID, 請求書.発行日 DESC ";

                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);

                    // グリッドビューに表示する
                    int iX = 0;

                    tempDGV.RowCount = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = int.Parse(dR["得意先ID"].ToString());
                        tempDGV[1, iX].Value = dR["略称"].ToString();
                        tempDGV[2, iX].Value = dR["氏名"].ToString();
                        tempDGV[3, iX].Value = DateTime.Parse(dR["発行日"].ToString()).ToShortDateString();
                        tempDGV[4, iX].Value = int.Parse(dR["請求金額"].ToString(),System.Globalization.NumberStyles.Any);

                        if (dR["入金予定日"] == DBNull.Value)
                        {
                            tempDGV[5, iX].Value = "";
                        }
                        else
                        {
                            tempDGV[5, iX].Value = DateTime.Parse(dR["入金予定日"].ToString()).ToShortDateString();
                        }

                        if (dR["入金日"] == DBNull.Value )
                        {
                            tempDGV[6, iX].Value = "";
                        }
                        else
                        {
                            tempDGV[6, iX].Value = DateTime.Parse(dR["入金日"].ToString()).ToShortDateString();
                        }

                        if (dR["入金額"] == DBNull.Value)
                        {
                            tempDGV[7, iX].Value = 0;
                        }
                        else
                        {
                            tempDGV[7, iX].Value = int.Parse(dR["入金額"].ToString(), System.Globalization.NumberStyles.Any);
                        }

                        tempDGV[8, iX].Value = int.Parse(dR["入金残"].ToString(), System.Globalization.NumberStyles.Any);

                        // 請求書ID 2010/2/16
                        tempDGV[9, iX].Value = dR["ID"].ToString();

                        // 備考 2015/07/22
                        tempDGV[10, iX].Value = dR["備考"].ToString();

                        // 入金合計
                        if (dR["入金額"] != DBNull.Value)
                        sTotal += int.Parse(dR["入金額"].ToString(), System.Globalization.NumberStyles.Any);
                        
                        // 入金残合計
                        sZan += int.Parse(dR["入金残"].ToString(), System.Globalization.NumberStyles.Any);
                        
                        iX++;
                    }

                    dR.Close();
                    fCon.Close();

                    //if (tempDGV.RowCount > 25)
                    //    //if (tempDGV.RowCount > 25)
                    //{
                    //    tempDGV.Columns[1].Width = 253;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 270;
                    //}

                    tempTotal.Text = sTotal.ToString("#,##0");
                    tempTxt.Text = sZan.ToString("#,##0");
                    tempDGV.CurrentCell = null;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }
            }
        }
        
        private void ShowDataLinq(DataGridView g)
        {
            int sZan = 0;
            int sTotal = 0;
            int sMinyu = 0;
            int sUrikake = 0;

            try
            {
                // グリッドビューに表示する
                int iX = 0;

                g.RowCount = 0;

                // 無効データは対象外
                //var s = dts.新請求書.Where(a => a.無効 == global.FLGOFF).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);

                // 2015/12/09  
                var s = dts.新請求書
                        .Where(a => a.無効 == global.FLGOFF && a.Get受注1Rows().Count() > 0)
                        .OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);

                // 支払期日
                if (tDate.Checked)
                {
                    s = s.Where(a => a.支払期日 >= DateTime.Parse(tDate.Value.ToShortDateString())).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);
                }

                if (tDate2.Checked)
                {
                    s = s.Where(a => a.支払期日 <= DateTime.Parse(tDate2.Value.ToShortDateString())).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);
                }

                // 請求書発行日
                if (hDate.Checked)
                {
                    s = s.Where(a => a.請求書発行日 >= DateTime.Parse(hDate.Value.ToShortDateString())).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);
                }

                if (hDate2.Checked)
                {
                    s = s.Where(a => a.請求書発行日 <= DateTime.Parse(hDate2.Value.ToShortDateString())).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);
                }

                // 入金完了のみ
                if (radioButton2.Checked)
                {
                    if (chkMishu.Checked)
                    {
                        s = s.Where(a => a.入金完了 == global.FLGOFF).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);
                    }
                    else
                    {
                        s = s.Where(a => a.入金完了 == global.FLGOFF || a.残金 > 0).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);
                    }
                }
                else if (radioButton3.Checked)  // 未収確定のみ
                {
                    s = s.Where(a => a.入金完了 == global.FLGON && a.残金 > 0).OrderBy(a => a.支払期日).ThenBy(a => a.請求書発行日);
                }

                foreach (var t in s)
                {
                    // クライアント名指定
                    if (txtsClientName.Text.Trim() != string.Empty)
                    {
                        // 得意先名
                        if (t.得意先Row == null)
                        {
                            continue;
                        }
                        else if (!t.得意先Row.略称.Contains(txtsClientName.Text.Trim()))
                        {
                            continue;
                        }
                    }

                    g.Rows.Add();

                    g[0, iX].Value = t.得意先ID.ToString();    // 得意先ID

                    // 得意先名・担当者
                    if (t.得意先Row == null)
                    {
                        g[1, iX].Value = string.Empty;
                        g[2, iX].Value = string.Empty;
                        g[3, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[1, iX].Value = t.得意先Row.略称;
                        g[2, iX].Value = t.得意先Row.フリガナ;

                        if (t.得意先Row.社員Row == null)
                        {
                            g[3, iX].Value = string.Empty;
                        }
                        else
                        {
                            g[3, iX].Value = t.得意先Row.社員Row.氏名;
                        }
                    }

                    g[4, iX].Value =  t.請求書発行日.ToShortDateString();     // 請求書発行日
                    g[5, iX].Value = t.請求金額.ToString("#,##0");            // 請求金額

                    // 入金予定日
                    if (t.Is支払期日Null())
                    {
                        g[6, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[6, iX].Value = t.支払期日.ToShortDateString();
                    }

                    // 入金額・入金日
                    DateTime nDt = DateTime.Parse("1900/01/01");
                    int nyukin = 0;

                    foreach (var n in t.Get新入金Rows())
                    {
                        nyukin += n.金額;

                        if (nDt < n.入金年月日)
                        {
                            nDt = n.入金年月日;
                        }
                    }

                    if (t.Get新入金Rows().Any())
                    {
                        g[7, iX].Value = nDt.ToShortDateString();
                        g[8, iX].Value = nyukin.ToString("#,##0");
                    }
                    else
                    {
                        g[7, iX].Value = string.Empty;
                        g[8, iX].Value = "0";
                    }

                    g[9, iX].Value = t.残金.ToString("#,##0");

                    if (t.入金完了 == global.FLGOFF)
                    {
                        g[10, iX].Value = false;
                    }
                    else
                    {
                        g[10, iX].Value = true;
                    }

                    g[11, iX].Value = t.ID.ToString();

                    if (t.Is備考Null())
                    {
                        g[12, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[12, iX].Value = t.備考;
                    }

                    sTotal += nyukin;   // 入金合計

                    //if (chkMishu.Checked)
                    //{
                    //    if ()
                    //}
                    //else
                    //{
                    //    sZan += t.残金;     // 入金残合計
                    //}

                    // 残金あり
                    if (t.残金 != 0)
                    {
                        if (chkMishu.Checked)　// 入金完了は残金なしとする
                        {                            
                            if (t.入金完了 == global.FLGOFF)
                            {
                                // 残金計
                                sZan += t.残金;

                                // 未収金 or 売掛金
                                if (DateTime.Today > t.支払期日)
                                {
                                    // 支払期日を過ぎていたら未収金
                                    sMinyu += t.残金;
                                }
                                else
                                {
                                    // 支払期日以前であれば売掛金
                                    sUrikake += t.残金;
                                }
                            }
                        }
                        else
                        {
                            // 残金計
                            sZan += t.残金;

                            // 未収金 or 売掛金
                            if (DateTime.Today > t.支払期日)
                            {
                                // 支払期日を過ぎていたら未収金
                                sMinyu += t.残金;
                            }
                            else
                            {
                                // 支払期日以前であれば売掛金
                                sUrikake += t.残金;
                            }
                        }
                    }

                    iX++;
                }
                    
                txtTotal.Text = sTotal.ToString("#,##0");
                txtZan.Text = sZan.ToString("#,##0");
                txtMinyu.Text = sMinyu.ToString("#,##0");
                txtUrikake.Text = sUrikake.ToString("#,##0");

                g.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }
        
        
        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {
            try
            {
                chkMishu.Checked = true;
                radioButton1.Checked = true;
                btnPrn.Enabled = false;
                button1.Enabled = false;    // 2017/08/14
                hDate.Checked = false;
                hDate2.Checked = false;
                tDate.Checked = false;
                tDate2.Checked = false;
                txtsClientName.Text = "";
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                //画面表示
                mukouStatus = false;
                ShowDataLinq(dataGridView1);
                mukouStatus = true;

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                    button1.Enabled = true;
                }
                else
                {
                    btnPrn.Enabled = false;
                    button1.Enabled = false;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"選択",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            EigyoReport(dataGridView1);
        }
        
        private void EigyoReport(DataGridView tempDGV)
        {
            const int S_GYO = 4;        // エクセルファイル明細印刷開始行
            const int S_ROWSMAX = 11;   // エクセルファイル列最大値

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセルクライアント別請求一覧, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    // 合計情報を印字
                    StringBuilder sb = new StringBuilder();
                    sb.Append("入金計：").Append(txtTotal.Text).Append("  ");
                    sb.Append("未収金：").Append(txtMinyu.Text).Append("  ");
                    sb.Append("売掛金：").Append(txtUrikake.Text).Append("  ");
                    sb.Append("入金残計：").Append(txtZan.Text);
                    oxlsSheet.Cells[1, 6] = sb.ToString();


                    // グリッド情報を印字
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        oxlsSheet.Cells[S_GYO - 3, S_ROWSMAX] = int.Parse(this.txtZan.Text, System.Globalization.NumberStyles.Any);
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[3, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[4, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[5, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 6] = tempDGV[6, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[7, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[8, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 9] = tempDGV[9, iX].Value.ToString();

                        if (tempDGV[10, iX].Value.ToString() == "True")
                        {
                            oxlsSheet.Cells[iX + S_GYO, 10] = "*";
                        }
                        else
                        {
                            oxlsSheet.Cells[iX + S_GYO, 10] = string.Empty;
                        }

                        oxlsSheet.Cells[iX + S_GYO, 11] = tempDGV[12, iX].Value.ToString();
                    }

                    ////////セル上部へ実線ヨコ罫線を引く
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //セル下部へ実線ヨコ罫線を引く
                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体に実線縦罫線を引く
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の左端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //表全体の右端縦罫線
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    
                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    //印刷
                    oxlsSheet.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //ダイアログボックスの初期設定
                    saveFileDialog1.Title = "クライアント別請求一覧";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = "クライアント別請求一覧";
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
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
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();

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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.BackColor = Color.White;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (mukouStatus)
            {
                // 入金済み、備考
                if (e.ColumnIndex == 10 || e.ColumnIndex == 12)
                {
                    // ID取得
                    int sID = Utility.strToInt(dataGridView1["col11", e.RowIndex].Value.ToString());

                    // データ取得
                    var s = dts.新請求書.Single(a => a.ID == sID);

                    // 入金済チェック
                    if (e.ColumnIndex == 10)
                    {
                        if (dataGridView1["col10", e.RowIndex].Value.ToString() == "True")
                        {
                            s.入金完了 = global.FLGON;
                        }
                        else
                        {
                            s.入金完了 = global.FLGOFF;
                        }
                    }
                    
                    // 備考
                    if (e.ColumnIndex == 12)
                    {
                        if (dataGridView1["col12", e.RowIndex].Value == null)
                        {
                            s.備考 = string.Empty;
                        }
                        else
                        {
                            s.備考 = dataGridView1["col12", e.RowIndex].Value.ToString();
                        }
                    }
                    
                    s.変更年月日 = DateTime.Now;
                    s.ユーザーID = global.loginUserID;

                    // 新請求書データ更新
                    rAdp.Update(dts.新請求書);

                    // データ読み込み
                    rAdp.Fill(dts.新請求書);
                }
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCellAddress.X == 10)
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
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
            sf.FileName = "勘定奉行汎用データ_請求";
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
                    //if (Utility.strToInt(g[8, i].Value.ToString()) <= 0)
                    //{
                    //    continue;
                    //}

                    //ヘッダファイル出力
                    if (pblFirstGyouFlg)
                    {
                        wrkOutputData = string.Empty;
                        wrkOutputData += Entity.OutPutHeader.dn01 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd01 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd02 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr02 + ",";

                        wrkOutputData += Entity.OutPutHeader.kr06 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks02 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks06 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr01 + ",";

                        wrkOutputData += Entity.OutPutHeader.kr55 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks52 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks53 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks54 + ",";

                        wrkOutputData += Entity.OutPutHeader.ks55 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks04 + ",";
                        wrkOutputData += Entity.OutPutHeader.tk01 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks01;

                        outFile.WriteLine(wrkOutputData);
                    }

                    //出力データ作成
                    if (pblFirstGyouFlg)
                    {
                        wrkOutputData = "*,";
                    }
                    else
                    {
                        wrkOutputData = ",";
                    }

                    wrkOutputData += "0,";          // 伝票部門コード
                    wrkOutputData += g[4, i].Value.ToString() + ",";  // 発行日
                    wrkOutputData += "135,";        // 借方勘定科目コード
                    wrkOutputData += Utility.strToInt(g[5, i].Value.ToString()) + ",";    // 借方本体金額
                    wrkOutputData += "501,";        // 貸方勘定科目コード
                    wrkOutputData += Utility.strToInt(g[5, i].Value.ToString()) + ",";    // 貸方本体金額
                    wrkOutputData += "20,";         // 借方部門コード
                    wrkOutputData += (Utility.strToInt(g[0, i].Value.ToString()) + 990000000) + ",";      // 取引先コード
                    wrkOutputData += ",";       // 貸方補助科目コード
                    wrkOutputData += "3,";      // 貸方税率区分コード
                    wrkOutputData += "8,";      // 貸方税率
                    wrkOutputData += "2,";      // 貸方消費税計算
                    wrkOutputData += "2,";      // 貸方端数処理
                    wrkOutputData += g[2, i].Value.ToString().Replace(",","") + ",";        // 摘要
                    wrkOutputData += "20";      // 貸方部門コード

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
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;

namespace posting
{
    public partial class frmHaifuShijiADD : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.配布エリア cArea = new Entity.配布エリア();
        Entity.配布指示 cMaster = new Entity.配布指示();

        const string MESSAGE_CAPTION = "配布指示データ登録";
        const int STATUS_RES = 1;
        const int STATUS_DEF = 0;
        const string UPDATE_OK = "0";
        const string UPDATE_NO = "1";
        const int FLG_ON = 1;
        const int FLG_OFF = 0;

        public frmHaifuShijiADD()
        {
            InitializeComponent();

            // データ読み込み
            jAdp.Fill(dts.受注1);
            aAdp.Fill(dts.配布エリア);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.配布エリアTableAdapter aAdp = new darwinDataSetTableAdapters.配布エリアTableAdapter();

        DateTime sDay;
        DateTime eDay;

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //配布フラグをONにする
            Utility.FlgOnOff(FLG_ON);

            // 日付カレンダー用の開始日から終了日を取得
            sDay = dts.配布エリア.Where(a => a.配布指示ID == 0 && a.ステータス == 0).Min(a => a.受注1Row.配布開始日);
            eDay = dts.配布エリア.Where(a => a.配布指示ID == 0 && a.ステータス == 0).Max(a => a.受注1Row.配布終了日);

            TimeSpan ts = eDay - sDay;
            double days = ts.TotalDays;

            // 最長１００日とします
            if (days > 100)
            {
                days = 100;
                eDay = sDay.AddDays(days);
            }

            //画面設定
            GridviewSet.Setting(dataGridView2, dts, jAdp, aAdp, sDay, days);
            GridviewSet.Setting2(dataGridView1);
            GridviewSet.ShowData(dataGridView2, sDay, eDay);
            dataGridView2.CurrentCell = null; //選択状態を回避する
        }

        // データグリッドビュークラス
        private class GridviewSet
        {
            /// <summary>
            /// データグリッドビューの定義を行います
            /// </summary>
            /// <param name="tempDGV">データグリッドビューオブジェクト</param>
            public static void Setting(DataGridView tempDGV, darwinDataSet dts, darwinDataSetTableAdapters.受注1TableAdapter jAdp, darwinDataSetTableAdapters.配布エリアTableAdapter aAdp, DateTime sDay, double days)
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
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 559;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "配布エリアID");
                    tempDGV.Columns.Add("col2", "受注番号");
                    tempDGV.Columns.Add("col3", "チラシ名");
                    tempDGV.Columns.Add("col4", "ｴﾘｱID");
                    tempDGV.Columns.Add("col5", "配布先住所");
                    tempDGV.Columns.Add("col6", "予定枚数");

                    tempDGV.Columns.Add("col7", "配布開始");
                    tempDGV.Columns.Add("col8", "配布終了");
                    tempDGV.Columns.Add("col9", "併配");
                    tempDGV.Columns.Add("col10", "配布条件");
                    tempDGV.Columns.Add("col11", "配布形態");
                    tempDGV.Columns.Add("col12", "単価");
                    tempDGV.Columns.Add("col13", "枚数");
                    tempDGV.Columns.Add("col14", "併配除外");

                    for (int i = 0; i <= days; i++)
                    {
                        string gDay = string.Empty;

                        if (sDay.AddDays(i).Day == 1)
                        {
                            gDay = sDay.AddDays(i).Month.ToString() + "/" + sDay.AddDays(i).Day.ToString();
                        }
                        else
                        {
                            gDay = sDay.AddDays(i).Day.ToString();
                        }

                        tempDGV.Columns.Add("day" + i.ToString(), gDay);
                        tempDGV.Columns["day" + i.ToString()].Width = 20;
                        tempDGV.Columns["day" + i.ToString()].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    tempDGV.Columns[4].Frozen = true;

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 100;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 230;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    //tempDGV.Columns[8].Width = 366;

                    //tempDGV.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[8].Width = 200;
                    tempDGV.Columns[13].Width = 80;

                    tempDGV.Columns[0].Visible = false;
                    tempDGV.Columns[9].Visible = false;
                    tempDGV.Columns[10].Visible = false;
                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;
                    //tempDGV.Columns[13].Visible = false; // 2014/11/26

                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

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

            public static void Setting2(DataGridView tempDGV)
            {
                try
                {
                    //フォームサイズ定義

                    // 列スタイルを変更する

                    tempDGV.EnableHeadersVisualStyles = false;

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", (float)9.5, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 559;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col0", "グループ");
                    tempDGV.Columns.Add("col1", "配布エリアID");
                    tempDGV.Columns.Add("col2", "受注番号");
                    tempDGV.Columns.Add("col3", "チラシ名");
                    tempDGV.Columns.Add("col4", "ｴﾘｱID");
                    tempDGV.Columns.Add("col5", "配布先住所");
                    tempDGV.Columns.Add("col6", "予定枚数");

                    tempDGV.Columns.Add("col7", "配布開始");
                    tempDGV.Columns.Add("col8", "配布終了");
                    tempDGV.Columns.Add("col9", "併配");
                    tempDGV.Columns.Add("col10", "配布条件");
                    tempDGV.Columns.Add("col11", "配布形態");
                    tempDGV.Columns.Add("col12", "単価");
                    tempDGV.Columns.Add("col13", "枚数");
                    tempDGV.Columns.Add("col14", "更新結果");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 70;
                    tempDGV.Columns[5].Width = 230;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 286;

                    tempDGV.Columns[1].Visible = false;
                    tempDGV.Columns[10].Visible = false;
                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;
                    tempDGV.Columns[13].Visible = false;
                    tempDGV.Columns[14].Visible = false;

                    //tempDGV.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[8].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[9].SortMode = DataGridViewColumnSortMode.NotSortable;

                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[0].SortMode = DataGridViewColumnSortMode.Automatic;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

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

            ///----------------------------------------------------------------
            /// <summary>
            ///     グリッドに受注データを表示する </summary>
            /// <param name="tempDGV">
            ///     データグリッドビューオブジェクト名</param>
            ///----------------------------------------------------------------
            public static void ShowData(DataGridView tempDGV, DateTime sDay, DateTime eDay)
            {
                string sqlSTRING = "";
                string strDate;
                int iX;

                try
                {
                    tempDGV.RowCount = 0;
                    
                    //データリーダーを取得する
        
                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select 配布エリア.ID,配布エリア.受注ID,受注.チラシ名,";
                    sqlSTRING += "配布エリア.町名ID,町名.名称 as 町名,受注.配布開始日,受注.配布終了日,";
                    sqlSTRING += "配布エリア.併配区分,配布エリア.予定枚数,";
                    sqlSTRING += "受注.配布条件,配布形態.名称 as 配布形態,受注.併配除外 ";
                    sqlSTRING += "from ((配布エリア left join 受注 on 配布エリア.受注ID = 受注.ID) ";
                    sqlSTRING += "left join 町名 on 配布エリア.町名ID = 町名.ID) ";
                    sqlSTRING += "left join 配布形態 on 受注.配布形態 = 配布形態.ID ";
                    sqlSTRING += "where (配布エリア.配布指示ID = 0) and (配布エリア.ステータス = 0) ";
                    sqlSTRING += "order by 配布エリア.受注ID,配布エリア.町名ID";
                                        
                    //配布指示データのデータリーダーを取得する
                    Control.FreeSql cArea = new Control.FreeSql();
                    dR = cArea.free_dsReader(sqlSTRING);

                    //グリッドビューに表示する
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = int.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = long.Parse(dR["受注ID"].ToString());
                            tempDGV[2, iX].Value = dR["チラシ名"].ToString();
                            tempDGV[3, iX].Value = int.Parse(dR["町名ID"].ToString());
                            tempDGV[4, iX].Value = dR["町名"].ToString();
                            tempDGV[5, iX].Value = int.Parse(dR["予定枚数"].ToString());

                            strDate = dR["配布開始日"].ToString() + "";
                            if (strDate == "")
                            {
                                tempDGV[6, iX].Value = "";
                            }
                            else
                            {
                                tempDGV[6, iX].Value = Convert.ToDateTime(dR["配布開始日"].ToString() + "");
                            }

                            strDate = dR["配布終了日"].ToString() + "";
                            if (strDate == "")
                            {
                                tempDGV[7, iX].Value = "";
                            }
                            else
                            {
                                tempDGV[7, iX].Value = Convert.ToDateTime(dR["配布終了日"].ToString() + "");
                            }

                            if (dR["併配区分"].ToString() == "1")
                            {
                                tempDGV[8, iX].Value = "○";
                            }
                            else
                            {
                                tempDGV[8, iX].Value = "";
                            }

                            //tempDGV[8, iX].Value = dR["配布条件"].ToString();
                            //tempDGV[9, iX].Value = dR["配布形態"].ToString() + "";
                            //tempDGV[10, iX].Value = dR["単価"].ToString();
                            //tempDGV[11, iX].Value = dR["枚数"].ToString();

                            // 併配除外 2014/11/26
                            tempDGV["col14", iX].Value = Utility.nullToInt(dR["併配除外"]).ToString();

                            // 配布日マーキング
                            for (int i = 0; sDay.AddDays(i) <= eDay; i++)
                            {
                                if (sDay.AddDays(i) >= DateTime.Parse(dR["配布開始日"].ToString()) &&
                                    sDay.AddDays(i) <= DateTime.Parse(dR["配布終了日"].ToString()))
                                {
                                    tempDGV["day" + i.ToString(), iX].Value = "●";
                                    tempDGV["day" + i.ToString(), iX].Style.BackColor = Color.LightPink;
                                    tempDGV["day" + i.ToString(), iX].Style.ForeColor = Color.LightPink;
                                }
                                else
                                {
                                    tempDGV["day" + i.ToString(), iX].Value = "";
                                    tempDGV["day" + i.ToString(), iX].Style.BackColor = Color.White;
                                    tempDGV["day" + i.ToString(), iX].Style.ForeColor = Color.White;
                                }
                            }

                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();
                    cArea.Close();

                    //if (tempDGV.RowCount <= 27)
                    //{
                    //    tempDGV.Columns[8].Width = 77;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[8].Width = 60;
                    //}

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //配布エリアを一括選択する
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("明細が選択されていません", "明細未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            } 

            if (MessageBox.Show("選択中のチラシ配布データを配布指示中データへ追加します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            int iY = 0;
 
            //選択データを指示中TABへ移動する
            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {

                dataGridView1.Rows.Add();
                iY = dataGridView1.Rows.Count;

                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    // 2014/11/26 「併配除外」以降列はスキップ
                    if (i < 13)
                    {
                        dataGridView1[i + 1, iY - 1].Value = dataGridView2[i, r.Index].Value;
                    }
                }

                dataGridView1[14, iY - 1].Value = UPDATE_OK; //更新結果
 
                //選択データのステータスを(1)に書き換える
                HaihuStatusUpdate(int.Parse(dataGridView2[0, r.Index].Value.ToString()),STATUS_RES);

                iY++;
            }

            ////配布未指示データ再表示
            //GridviewSet.ShowData(dataGridView2);

            //TAB2を表示
            tabControl1.SelectedIndex = 1;

            dataGridView1.CurrentCell = null; //選択状態を回避する

        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中のチラシデータ全てを選択状態にします。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.Rows[i].Selected = true;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("現在の選択状態を解除します。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                dataGridView2.Rows[r.Index].Selected = false;
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Ending();
        }

        private void Ending()
        {
            if (MessageBox.Show("終了します。現在、配布指示設定タブに表示中のデータは登録されません。" + Environment.NewLine + "よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            //予約データのステータスを0に戻す
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                HaihuStatusUpdate(int.Parse(dataGridView1[1, i].Value.ToString()), STATUS_DEF);
            }

            //配布フラグをOFFにする
            Utility.FlgOnOff(FLG_OFF);

            this.Close();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("併配表示を行います。よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            int toID;
            DateTime sDate;
            DateTime eDate;

            int targetID;
            DateTime tsDate;
            DateTime teDate;

            string sqlStr, msgStr;

            // 自分のループ
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                // 配布期間が登録されているものを対象とする
                // 併配除外案件は対象外とします 2014/11/26
                if (dataGridView2[6, i].Value.ToString().Trim() != "" &&
                    dataGridView2[7, i].Value.ToString().Trim() != "" && 
                    dataGridView2["col14", i].Value.ToString() != "1")
                {
                    // ID
                    toID = Int32.Parse(dataGridView2[3, i].Value.ToString());

                    //配布開始日
                    sDate = Convert.ToDateTime(dataGridView2[6, i].Value.ToString());

                    //配布終了日
                    eDate = Convert.ToDateTime(dataGridView2[7, i].Value.ToString());

                    //相手のループ
                    for (int iX = i + 1; iX < dataGridView2.Rows.Count; iX++)
                    {
                        // 配布期間が登録されているものを対象とする
                        // 併配除外案件は対象外とします 2014/11/26
                        if (dataGridView2[6, iX].Value.ToString().Trim() != "" &&
                            dataGridView2[7, iX].Value.ToString().Trim() != "" &&
                            dataGridView2["col14", iX].Value.ToString() != "1")
                        {
                            //相手のID
                            targetID = Int32.Parse(dataGridView2[3, iX].Value.ToString());

                            //相手の配布開始日
                            tsDate = Convert.ToDateTime(dataGridView2[6, iX].Value.ToString());

                            //相手の配布終了日
                            teDate = Convert.ToDateTime(dataGridView2[7, iX].Value.ToString());

                            if (toID == targetID)
                            {
                                //配布期間が重複する場合、併配扱い
                                //パターン①　期間の一部重複
                                if (((tsDate >= sDate) && (tsDate <= eDate)) || ((teDate >= sDate) && (teDate <= eDate)))
                                {
                                    dataGridView2[8, i].Value = "○";
                                    dataGridView2[8, iX].Value = "○";
                                    dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                                    dataGridView2.Rows[iX].DefaultCellStyle.ForeColor = Color.Red;
                                }

                                //パターン② 相手の期間中そのまま
                                if ((tsDate <= sDate) && (eDate <= teDate))
                                {
                                    dataGridView2[8, i].Value = "○";
                                    dataGridView2[8, iX].Value = "○";
                                    dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                                    dataGridView2.Rows[iX].DefaultCellStyle.ForeColor = Color.Red;
                                }
                            }
                        }
                    }

                    //配布指示済みで配布未完了のデータを対象に検索
                    // ･･･> 比較する配布日：受注データの配布開始日～配布終了日 2009/10/13

                    sqlStr = "";

                    //////sqlStr += "select 配布エリア.配布指示ID,配布エリア.町名ID,配布指示.配布日 ";
                    //////sqlStr += "from 配布エリア inner join 配布指示 ";
                    //////sqlStr += "on 配布エリア.配布指示ID = 配布指示.ID ";
                    //////sqlStr += "where ";
                    //////sqlStr += "(配布エリア.配布指示ID <> 0) and ";
                    //////sqlStr += "(配布エリア.ステータス = 2) and ";
                    //////sqlStr += "(配布エリア.完了区分 = 0) and ";
                    //////sqlStr += "(配布エリア.町名ID = " + toID.ToString() + ") ";
                    //////sqlStr += "order by 配布エリア.配布指示ID";

                    sqlStr += "select 配布エリア.配布指示ID,配布エリア.町名ID,受注.配布開始日,受注.配布終了日 ";
                    sqlStr += "from 配布エリア inner join 受注 ";
                    sqlStr += "on 配布エリア.受注ID = 受注.ID ";
                    sqlStr += "where ";
                    sqlStr += "(配布エリア.配布指示ID <> 0) and ";
                    sqlStr += "(配布エリア.ステータス = 2) and ";
                    sqlStr += "(配布エリア.完了区分 = 0) and ";
                    sqlStr += "(配布エリア.町名ID = " + toID.ToString() + ") ";
                    sqlStr += "order by 配布エリア.受注ID";

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlStr);

                    while (dR.Read())
                    {
                        //////配布指示書の配布日
                        ////tsDate = Convert.ToDateTime(dR["配布日"].ToString());

                        //////配布指示書の配布日が配布期間に該当するか
                        ////if ((tsDate >= sDate) && (tsDate <= eDate))
                        ////{
                        ////    msgStr = dataGridView2[8, i].Value.ToString();

                        ////    if (msgStr != "")
                        ////    {
                        ////        msgStr += "、";
                        ////    }

                        ////    msgStr += "指示№:" + dR["配布指示ID"].ToString();
                        ////    dataGridView2[8, i].Value = msgStr;
                        ////    dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                        ////}

                        //相手の配布開始日(受注データより)
                        tsDate = Convert.ToDateTime(dR["配布開始日"].ToString());

                        //相手の配布終了日(受注データより)
                        teDate = Convert.ToDateTime(dR["配布終了日"].ToString());

                        //配布期間が重複する場合、併配扱い
                        //パターン①　期間の一部重複
                        if (((tsDate >= sDate) && (tsDate <= eDate)) || ((teDate >= sDate) && (teDate <= eDate)))
                        {
                            msgStr = dataGridView2[8, i].Value.ToString();

                            if (msgStr != "")
                            {
                                msgStr += "、";
                            }

                            msgStr += "指示№:" + dR["配布指示ID"].ToString();
                            dataGridView2[8, i].Value = msgStr;
                            dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                        }
                        else
                        {
                            //パターン② 相手の期間中そのまま
                            if ((tsDate <= sDate) && (eDate <= teDate))
                            {
                                msgStr = dataGridView2[8, i].Value.ToString();

                                if (msgStr != "")
                                {
                                    msgStr += "、";
                                }

                                msgStr += "指示№:" + dR["配布指示ID"].ToString();
                                dataGridView2[8, i].Value = msgStr;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                            }
                        }
                    }

                    dR.Close();
                    fCon.Close();
                }
            }

            MessageBox.Show("終了しました", "併配表示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 配布エリアデータのステータスを初期状態に戻す
        /// </summary>
        /// <param name="tempID">配布エリアID</param>
        private void HaihuStatusUpdate(int tempID,int tempStatus)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update 配布エリア ";
            sqlStr += "set ";
            sqlStr += "配布指示ID = 0,";
            sqlStr += "ステータス = " + tempStatus.ToString() + ",";
            sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
            sqlStr += "where 配布エリア.ID = " + tempID.ToString();

            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("配布エリアデータのステータス更新に失敗しました(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        private void frmHaifuShijiADD_FormClosing(object sender, FormClosingEventArgs e)
        {
            Dispose();
        }

        private void button6_Click(object sender, EventArgs e)
        {

            const int DIALOG_OK = 1;
            //const int DIALOG_NO = 0;

            //選択中のデータをグループ化する
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("明細を選択してください", "明細未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            haihuGroup fSub = new haihuGroup();

            if (fSub.ShowDialog(this) == DialogResult.OK)
            {
                if (fSub.sStatus == DIALOG_OK)
                {
                    //選択データをグルーピングする
                    foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                    {
                        dataGridView1[0, r.Index].Value = fSub.sGroup;
                    }
                }

                //並び替える列を決める
                DataGridViewColumn sortColumn = dataGridView1.Columns[0];

                //並び替えの方向（昇順か降順か）を決める
                ListSortDirection sortDirection = ListSortDirection.Ascending;

                //並び替えを行う
                dataGridView1.Sort(sortColumn, sortDirection);

                //手動ソート禁止とする
                //dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;

            }

            fSub.Dispose();

            dataGridView1.CurrentCell = null;


        }

        private void button7_Click(object sender, EventArgs e)
        {
            //配布エリアを選択解除する
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("明細が選択されていません", "明細未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("選択中のチラシ配布データを未指示データに戻します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //選択データのステータスを0に戻す
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                HaihuStatusUpdate(int.Parse(dataGridView1[1, r.Index].Value.ToString()), STATUS_DEF);
            }
            
            //選択データをグリッドから削除する
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                if (!r.IsNewRow)
                {
                    dataGridView1.Rows.Remove(r);
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                GridviewSet.ShowData(dataGridView2, sDay, eDay);
            }

            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    GridviewSet.ShowData(dataGridView2, sDay, eDay);
                    dataGridView2.CurrentCell = null;
                    break;

                case 1:
                    dataGridView1.CurrentCell = null;
                    button6.Enabled = false;
                    button7.Enabled = false;

                    if (dataGridView1.RowCount > 0)
                    {
                        button8.Enabled = true;
                    }
                    else
                    {
                        button8.Enabled = false;
                    }

                    break;

                default:
                    break;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //if (dataGridView1.SelectedRows.Count == 0)
            //{
            //    MessageBox.Show("明細が選択されていません", "明細未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[0, i].Value == null)
                {
                    MessageBox.Show("グループ化されていない配布データがあります。全ての配布データをグループ化して再度実行してください。", "未グループ化", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[0, i].Value.ToString() == "")
                {
                    MessageBox.Show("グループ化されていない配布データがあります。全ての配布データをグループ化して再度実行してください。", "未グループ化", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            //並び替える列を決める
            DataGridViewColumn sortColumn = dataGridView1.Columns[0];

            //並び替えの方向（昇順か降順か）を決める
            ListSortDirection sortDirection = ListSortDirection.Ascending;

            //並び替えを行う
            dataGridView1.Sort(sortColumn, sortDirection);

            if (MessageBox.Show("配布指示データに登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //配布指示データ登録処理
            haifuAdd();

            //更新結果OKのデータはグリッドから削除する

            //セレクト
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[14, i].Value.ToString() == UPDATE_OK)
                {
                    dataGridView1.Rows[i].Selected = true;
                }
            }

            //削除
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                if (!r.IsNewRow)
                {
                    dataGridView1.Rows.Remove(r);
                }
            }

            //終了処理
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("全ての配布データの登録処理が終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                tabControl1.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("一部の配布データの登録処理が終了しませんでした", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void haifuAdd()
        {
            string sGroup = "";
            int sRow = 0, eRow = 0;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if ((sGroup != "") && (sGroup != dataGridView1[0, i].Value.ToString()))
                {
                    HaifuUpdate(sRow,eRow);
                    sRow = i;
                }

                sGroup = dataGridView1[0, i].Value.ToString();
                eRow = i;
            }

            HaifuUpdate(sRow, eRow);
        }

        private void HaifuUpdate(int sRow,int eRow)
        {
            try
            {
                HaifuShijiData();   //配布指示クラスにデータセット

                Control.DataControl Con;
                OleDbConnection cn;
                OleDbTransaction tran;
                OleDbCommand SCom;

                //IDを採番
                string sqlStr = "";
                int gID = (int)(1);

                sqlStr = "select max(ID) as ID from 配布指示 ";
                OleDbDataReader dR;
                Control.FreeSql fCon = new Control.FreeSql();
                dR = fCon.free_dsReader(sqlStr);

                while (dR.Read())
                {
                    if (dR["ID"] == DBNull.Value)
                    {
                        gID = (int)(1);
                    }
                    else
                    {
                        gID = Int32.Parse(dR["ID"].ToString()) + 1;
                    }
                }

                dR.Close();
                fCon.Close();

                //IDを設定
                cMaster.ID = gID;

                //登録処理
                Con = new Control.DataControl();
                cn = new OleDbConnection();

                cn = Con.GetConnection();

                //トランザクション開始
                tran = cn.BeginTransaction();

                SCom = new OleDbCommand();
                SCom.Connection = cn;
                SCom.Transaction = tran;

                try
                {
                    //配布指示データ登録処理
                    sqlStr = "";
                    sqlStr += "insert into 配布指示 ";
                    sqlStr += "(ID,配布日,入力日,配布員ID,交通費,交通区間開始,交通区間終了,";
                    sqlStr += "配布開始時刻,配布終了時刻,終了レポート,未配布区分,未配布理由,";
                    sqlStr += "登録年月日,変更年月日) ";
                    sqlStr += "values (";
                    sqlStr += cMaster.ID + ",";
                    sqlStr += "'" + cMaster.配布日 + "',";
                    sqlStr += "'" + cMaster.入力日 + "',";
                    sqlStr += cMaster.配布員ID + ",";
                    sqlStr += cMaster.交通費 + ",";
                    sqlStr += "'" + cMaster.交通区間開始 + "',";
                    sqlStr += "'" + cMaster.交通区間終了 + "',";
                    sqlStr += "'" + cMaster.配布開始時刻 + "',";
                    sqlStr += "'" + cMaster.配布終了時刻 + "',";
                    sqlStr += "'" + cMaster.終了レポート + "',";
                    sqlStr += "'" + cMaster.未配布区分 + "',";
                    sqlStr += "'" + cMaster.未配布理由 + "',";
                    sqlStr += "'" + cMaster.登録年月日 + "',";
                    sqlStr += "'" + cMaster.変更年月日 + "')";

                    SCom.CommandText = sqlStr;

                    SCom.ExecuteNonQuery();

                    //配布エリアデータ更新
                    string sID;
                    const string sSTATUS = "2"; 

                    for (int i = sRow; i <= eRow; i++)
                    {
                        sID = dataGridView1[1, i].Value.ToString();

                        sqlStr = "";
                        sqlStr += "update 配布エリア ";
                        sqlStr += "set ";
                        sqlStr += "配布指示ID = " + gID.ToString() + ",";
                        sqlStr += "実残数 = 予定枚数,";
                        sqlStr += "報告残数 = 予定枚数,";
                        sqlStr += "ステータス = " + sSTATUS + ",";
                        sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
                        sqlStr += "where (配布エリア.ID = " + sID + ") and ";
                        sqlStr += "(ステータス <> 0)";

                        //if (dataGridView1[0, i].Value.ToString() == "2")
                        //{
                        //    sqlStr += "(ステータス <> 0)";
                        //}


                        SCom.CommandText = sqlStr;

                        SCom.ExecuteNonQuery();
                    }

                    tran.Commit();

                    //更新結果書き込み
                    UpdateFlg(sRow, eRow, UPDATE_OK);

                    //MessageBox.Show("新規登録されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (Exception ex)
                {
                    tran.Rollback();

                    //更新結果書き込み
                    UpdateFlg(sRow, eRow, UPDATE_NO);

                    MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    MessageBox.Show("登録に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //未登録配布データのステータスを戻す
                    //StatusBack();
                }

                cn.Close();

                Con.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "更新処理", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void UpdateFlg(int sRow,int eRow, string sFlg)
        {
            //更新結果書き込み
            for (int i = sRow; i <= eRow; i++)
            {
                dataGridView1[14, i].Value = sFlg;
            }
        }

        private void HaifuShijiData()
        {
            //配布指示クラスにデータセット
            cMaster.配布日 = DateTime.Today;
            cMaster.入力日 = DateTime.Today;
            cMaster.配布員ID = 0;

            cMaster.交通費 = 0;
            cMaster.交通区間開始 = "";
            cMaster.交通区間終了 = "";

            cMaster.配布開始時刻 = "";
            cMaster.配布終了時刻 = "";

            cMaster.終了レポート = "";

            cMaster.未配布区分 = "";
            cMaster.未配布理由 = "";

            cMaster.注意事項 = "";

            cMaster.登録年月日 = DateTime.Today;
            cMaster.変更年月日 = DateTime.Today;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                button6.Enabled = true;
                button7.Enabled = true;
            }
            else
            {
                button6.Enabled = false;
                button7.Enabled = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Ending();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //選択中のデータの合計配布枚数を表示する
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("明細を選択してください", "明細未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int sMai = 0;

            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                if ((dataGridView1[6, r.Index].Value != null) && (Utility.NumericCheck(dataGridView1[6, r.Index].Value.ToString()) != false))
                {
                    sMai += int.Parse(dataGridView1[6, r.Index].Value.ToString(), System.Globalization.NumberStyles.Any);
                }
            }

            MessageBox.Show(sMai.ToString("#,##0") + " 枚です", "合計配布枚数", MessageBoxButtons.OK, MessageBoxIcon.Information);

            dataGridView1.CurrentCell = null;
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, "未指示データ");
        }

    }
}
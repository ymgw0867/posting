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

namespace posting
{
    public partial class frmHaifuShiji: Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.配布指示 cMaster = new Entity.配布指示();
        Entity.配布エリア cArea = new Entity.配布エリア();

        const string MESSAGE_CAPTION = "配布指示・報告書";
        const int COLUMN_KANRYO = 12;   //完了区分列  2015/07/09(11 → 12)
        const string KANRYO_STATUS = "True";    // 完了区分
        const string STATUS_KANRYO = "1";       // 完了
        const string STATUS_MIKANRYO = "0";     // 未完了

        bool STATUS_MAISU = false;      // 配布枚数計算ステータス

        public frmHaifuShiji()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            try
            {
                //グリッド設定
                GridviewSet.Setting(dataGridView1);
                GridviewSet.Setting2(dataGridView2);

                //配布員コンボ
                Utility.ComboStaff.load(cmbsStaff);

                DispClear();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }
            
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

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 166;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "指示番号");
                    tempDGV.Columns.Add("col2", "チラシ名");
                    tempDGV.Columns.Add("col3", "配布日");
                    tempDGV.Columns.Add("col4", "入力日");
                    tempDGV.Columns.Add("col5", "配布員");
                    tempDGV.Columns.Add("col6", "交通費");
                    tempDGV.Columns.Add("col7", "登録年月日");
                    tempDGV.Columns.Add("col8", "変更年月日");
                    tempDGV.Columns.Add("col9", "本日の注意事項");
                    tempDGV.Columns.Add("col10", "ユーザー");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 300;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 120;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 110;
                    tempDGV.Columns[7].Width = 110;
                    tempDGV.Columns[8].Width = 614;
                    tempDGV.Columns[9].Width = 100;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[3].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

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

            //配布指示明細データグリッド
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
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 252;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "PID");
                    tempDGV.Columns.Add("col2", "チラシ名");
                    tempDGV.Columns.Add("col3", "配布区分");
                    tempDGV.Columns.Add("col4", "配布形態");
                    tempDGV.Columns.Add("col5", "配布先住所");
                    tempDGV.Columns.Add("col6", "枝番記入");
                    tempDGV.Columns.Add("col7", "ID");
                    tempDGV.Columns.Add("col8", "単価");
                    tempDGV.Columns.Add("col9", "予定枚数");
                    tempDGV.Columns.Add("col10", "予定枚数計");
                    tempDGV.Columns.Add("col11", "配布枚数");
                    tempDGV.Columns.Add("col18", "配布枚数計");     // 2015/07/15

                    // 2016/10/31
                    DataGridViewCheckBoxColumn cbc2 = new DataGridViewCheckBoxColumn();
                    cbc2.Name = "col12";
                    cbc2.HeaderText = "配布完了";
                    cbc2.TrueValue = "1";
                    cbc2.FalseValue = "0";
                    tempDGV.Columns.Add(cbc2);

                    tempDGV.Columns.Add("col13", "報告枚数");
                    tempDGV.Columns.Add("col14", "delete");
                    tempDGV.Columns.Add("col15", "Add");
                    tempDGV.Columns.Add("col16", "未配布情報有無");
                    tempDGV.Columns.Add("col17", "枝番有無");

                    // 2016/11/14 未配布マンション入力画面表示ボタン
                    DataGridViewButtonColumn dbt = new DataGridViewButtonColumn();
                    dbt.UseColumnTextForButtonValue = true;
                    dbt.Text = "→";
                    dbt.Name = "col19";
                    dbt.HeaderText = "未配布";
                    tempDGV.Columns.Add(dbt);

                    // 2016/12/26 明細削除ボタン
                    DataGridViewButtonColumn delBtn = new DataGridViewButtonColumn();
                    delBtn.UseColumnTextForButtonValue = true;
                    delBtn.Text = "Del";
                    delBtn.Name = "col20";
                    delBtn.HeaderText = "削除";
                    tempDGV.Columns.Add(delBtn);
                    
                    //tempDGV.Columns.Add("col18", "番地号");
                    //tempDGV.Columns.Add("col19", "マンション名");
                    //tempDGV.Columns.Add("col20", "理由");
                    //tempDGV.Columns.Add("col21", "その他内容");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 120;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 200;
                    tempDGV.Columns[5].Width = 200;
                    tempDGV.Columns[6].Width = 60;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 100;
                    tempDGV.Columns[11].Width = 100;    // 2015/07/15
                    tempDGV.Columns[13].Width = 100;
                    tempDGV.Columns[16].Width = 200;
                    tempDGV.Columns[18].Width = 60;     // 2016/11/14
                    tempDGV.Columns[19].Width = 60;     // 2016/12/26

                    //tempDGV.Columns[17].Width = 120;
                    //tempDGV.Columns[18].Width = 160;
                    //tempDGV.Columns[19].Width = 70;
                    //tempDGV.Columns[20].Width = 200;

                    tempDGV.Columns[1].Frozen = true;       // 2016/10/27

                    tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[12].Visible = false;    // 2015/07/15
                    tempDGV.Columns[13].Visible = false;    // 2015/07/15
                    tempDGV.Columns[14].Visible = false;    // 2015/07/15
                    tempDGV.Columns[15].Visible = false;    // 2015/07/15
                    tempDGV.Columns[16].Visible = false;    // 2015/07/15
                    tempDGV.Columns[17].Visible = false;    // 2015/07/15

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[11].DefaultCellStyle.Format = "#,##0";    // 2015/07/15
                    tempDGV.Columns[13].DefaultCellStyle.Format = "#,##0";    // 2015/07/15

                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;    // 2015/07/15
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;    // 2015/07/15

                    //tempDGV.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    //tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //tempDGV.MultiSelect = true;

                    tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // 単価と配布枚数の編集可とする 2016/10/27
                    //tempDGV.ReadOnly = true;

                    foreach (DataGridViewColumn d in tempDGV.Columns)
                    {
                        if (d.Name == "col8" || d.Name == "col11" || d.Name == "col12" || d.Name == "col19" || d.Name == "col20")
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

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ///------------------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビューの指定行のデータを取得する </summary>
            /// <param name="dgv">
            ///     対象とするデータグリッドビューオブジェクト</param>
            ///------------------------------------------------------------------------------
            public static Boolean GetData(DataGridView dgv, ref Entity.配布指示 tempC, int iX)
            {
                string sqlStr;
                Control.配布指示 cShiji = new Control.配布指示();
                OleDbDataReader dr;

                sqlStr = " where 配布指示.ID = " + (int)dgv[0, iX].Value;
                dr = cShiji.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.配布日 = Convert.ToDateTime(dr["配布日"].ToString());
                        tempC.入力日 = Convert.ToDateTime(dr["入力日"].ToString());
                        tempC.配布員ID = Int32.Parse(dr["配布員ID"].ToString());
                        tempC.交通費 = Int32.Parse(dr["交通費"].ToString());
                        tempC.交通区間開始 = dr["交通区間開始"].ToString();
                        tempC.交通区間終了 = dr["交通区間終了"].ToString();
                        tempC.配布開始時刻 = dr["配布開始時刻"].ToString();
                        tempC.配布終了時刻 = dr["配布終了時刻"].ToString();
                        tempC.終了レポート = dr["終了レポート"].ToString();
                        tempC.未配布区分 = dr["未配布区分"].ToString();
                        tempC.未配布理由 = dr["未配布理由"].ToString();
                        tempC.注意事項 = dr["注意事項"].ToString();
                    }
                }
                else
                {
                    dr.Close();
                    cShiji.Close();
                    return false;
                }

                dr.Close();
                cShiji.Close();
                return true;
            }

            public static Boolean GetDataItem(int tempID, ref Entity.配布エリア tempC)
            {
                string sqlStr;

                Control.配布エリア cArea = new Control.配布エリア();
                OleDbDataReader dr;

                sqlStr = " where 配布エリア.ID = " + tempID;
                dr = cArea.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Int32.Parse(dr["ID"].ToString());
                        tempC.町名ID = Int32.Parse(dr["町名ID"].ToString());
                        tempC.予定枚数 = Int32.Parse(dr["予定枚数"].ToString());
                        tempC.受注ID = Int32.Parse(dr["受注ID"].ToString());
                        tempC.配布指示ID = Int32.Parse(dr["配布指示ID"].ToString());
                        tempC.配布単価 = double.Parse(dr["配布単価"].ToString());
                        tempC.配布日 = dr["配布日"].ToString();
                        tempC.実配布枚数 = Int32.Parse(dr["実配布枚数"].ToString());
                        tempC.実残数 = Int32.Parse(dr["実残数"].ToString());
                        tempC.報告枚数 = Int32.Parse(dr["報告枚数"].ToString());
                        tempC.報告残数 = Int32.Parse(dr["未配布区分"].ToString());
                        tempC.併配区分 = Int32.Parse(dr["併配区分"].ToString());
                        tempC.完了区分 = Int32.Parse(dr["完了区分"].ToString());
                        tempC.ステータス = Int32.Parse(dr["ステータス"].ToString());
                    }
                }
                else
                {
                    dr.Close();
                    cArea.Close();
                    return false;
                }

                dr.Close();
                cArea.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV,int tempID,string tempsID,string tempCName)
            {
                string sqlSTRING = "";
                int iX;
                string wID = "0";
                
                try
                {
                    tempDGV.RowCount = 0;

                    //配布指示データのデータリーダーを取得する

                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    //データリーダーを取得する
                    OleDbDataReader dR;

                    sqlSTRING += "select 配布指示.*,配布員.氏名,x.チラシ名, ログインユーザー.ログインユーザー ";
                    sqlSTRING += "from 配布指示 left join 配布員 on 配布指示.配布員ID = 配布員.ID ";

                    //チラシ名指定のとき
                    if (tempCName.Length == 0)
                    {
                        sqlSTRING += "left join ";
                    }
                    else
                    {
                        sqlSTRING += "inner join ";
                    }

                    sqlSTRING += "(SELECT DISTINCT 配布指示.ID, 受注.チラシ名 ";
                    sqlSTRING += "from 配布指示 inner join 配布エリア ";
                    sqlSTRING += "ON 配布指示.ID = 配布エリア.配布指示ID inner join ";
                    sqlSTRING += "受注 ON 配布エリア.受注ID = 受注.ID ";

                    //チラシ名指定のとき
                    if (tempCName.Length > 0)
                    {
                        sqlSTRING += "where (受注.チラシ名 like ?)";
                    }
                    
                    sqlSTRING += ") AS x ";
                    sqlSTRING += "ON 配布指示.ID = x.ID ";

                    // 2016/09/26
                    sqlSTRING += "left join ログインユーザー on ";
                    sqlSTRING += "配布指示.ユーザーID = ログインユーザー.ID ";

                    sqlSTRING += "where (1 = 1) ";

                    //////sqlSTRING += "(配布指示.配布日 >= ?) and ";
                    //////sqlSTRING += "(配布指示.配布日 <= ?) and ";
                    //////sqlSTRING += "(配布指示.入力日 >= ?) and ";
                    //////sqlSTRING += "(配布指示.入力日 <= ?) ";

                    //配布員指定のとき
                    if (tempID != -1)
                    {
                        sqlSTRING += "and (配布指示.配布員ID = ?) ";
                    }

                    //配布指示ID指定のとき
                    if (tempsID.Length > 0)
                    {
                        sqlSTRING += "and (配布指示.ID = ?) ";
                    }

                    // 配布日が1年以内の案件とする 2018/02/20
                    sqlSTRING += "and (配布指示.配布日 >= ?) ";

                    sqlSTRING += "order by 配布指示.ID desc ";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    // チラシ名指定パラメータ
                    if (tempCName.Length > 0)
                    {
                        SCom.Parameters.AddWithValue("@temCName","%" + tempCName + "%");
                    }

                    ////////配布日パラメータ
                    //////SCom.Parameters.AddWithValue("@sDate", tempsDate);
                    //////SCom.Parameters.AddWithValue("@eDate", tempeDate);

                    ////////入力日パラメータ
                    //////SCom.Parameters.AddWithValue("@nsDate", tempnsDate);
                    //////SCom.Parameters.AddWithValue("@neDate", tempneDate);

                    // 配布員IDパラメータ
                    if (tempID != -1)
                    {
                        SCom.Parameters.AddWithValue("@SID", tempID);
                    }

                    // 配布指示IDパラメータ
                    if (tempsID.Length > 0)
                    {
                        SCom.Parameters.AddWithValue("@tempsID",tempsID);
                    }

                    // 配布日パラメータ : 2018/02/20
                    DateTime hDt = DateTime.Today.AddYears(-1);
                    SCom.Parameters.AddWithValue("@hdt", hDt);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //グリッドビューに表示する
                    iX = 0;

                    //表示用データ領域のインスタンス作成
                    DataTemp d = new DataTemp();
                    dClear(d);

                    //データを読み込み
                    if (dR.HasRows == true)
                    {
                        while (dR.Read())
                        {
                            //最初のデータではなくIDでブレークが発生したとき
                            if ((wID != "0") && (wID != dR["ID"].ToString()))
                            {
                                //グリッドへ表示
                                AddDataGrid(tempDGV, d, iX);

                                //データ領域初期化
                                dClear(d);

                                iX++;
                            }

                            //表示項目の取得
                            d.ID = dR["ID"].ToString();

                            if (d.CName.Length > 0)
                            {
                                d.CName += ", " + dR["チラシ名"].ToString();
                            }
                            else
                            {
                                d.CName = dR["チラシ名"].ToString();
                            }

                            d.HDate = dR["配布日"].ToString();
                            d.IDate = dR["入力日"].ToString();
                            d.Name = dR["氏名"].ToString() + "";
                            d.Kotsuhi = dR["交通費"].ToString();
                            d.AddDate = dR["登録年月日"].ToString();
                            d.UppDate = dR["変更年月日"].ToString();
                            d.Memo = dR["注意事項"].ToString();

                            // 2016/09/26
                            if (dR["ログインユーザー"] == DBNull.Value)
                            {
                                d.loginUser = string.Empty;
                            }
                            else
                            {
                                d.loginUser = dR["ログインユーザー"].ToString();
                            }

                            wID = dR["ID"].ToString();
                        }

                        //最終データのグリッド表示
                        AddDataGrid(tempDGV, d, iX);
                    }
                    else
                    {
                        MessageBox.Show("該当するデータがありません", "配布指示データ検索",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    }

                    dR.Close();

                    Con.Close();

                    cn.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }

                tempDGV.CurrentCell = null;

            }

            private static void dClear(DataTemp d)
            {
                d.ID = "";
                d.CName = "";
                d.HDate = "";
                d.IDate = "";
                d.Kotsuhi = "";
                d.Name = "";
                d.AddDate = "";
                d.UppDate = "";
            }

            //グリッドビュー表示処理
            private static void AddDataGrid(DataGridView tempDGV, DataTemp d,int iX)
            {

                tempDGV.Rows.Add();

                tempDGV[0, iX].Value = int.Parse(d.ID);
                tempDGV[1, iX].Value = d.CName;
                tempDGV[2, iX].Value = Convert.ToDateTime(d.HDate);
                tempDGV[3, iX].Value = Convert.ToDateTime(d.IDate);
                tempDGV[4, iX].Value = d.Name + "";
                tempDGV[5, iX].Value = Int32.Parse(d.Kotsuhi);
                tempDGV[6, iX].Value = Convert.ToDateTime(d.AddDate);
                tempDGV[7, iX].Value = Convert.ToDateTime(d.UppDate);
                tempDGV[8, iX].Value = d.Memo;
                tempDGV[9, iX].Value = d.loginUser;     // 2016/09/26
            }

            ///--------------------------------------------------------------------
            /// <summary>
            ///     配布エリアデータ表示 </summary>
            /// <param name="tempDGV">
            ///     配布指示データグリッドビュー</param>
            /// <param name="tempDGV2">
            ///     配布エリアグリッドビュー</param>
            ///--------------------------------------------------------------------
            public static void ShowDataItem(DataGridView tempDGV, DataGridView tempDGV2,int iX)
            {
                //int iX = 0;

                try
                {
                    tempDGV2.RowCount = 0;

                    //配布エリアデータのデータリーダーを取得する
                    string sqlStr;

                    Control.配布エリア cArea = new Control.配布エリア();
                    OleDbDataReader dr;

                    sqlStr = " where 配布エリア.配布指示ID = " + (int)tempDGV[0, iX].Value + " ";
                    sqlStr += "order by 配布エリア.受注ID,配布エリア.町名ID";

                    dr = cArea.FillBy(sqlStr);

                    //グリッドビューに表示する
                    iX = 0;

                    while (dr.Read())
                    {
                        tempDGV2.Rows.Add();

                        tempDGV2[0, iX].Value = dr["ID"];

                        iX++;
                    }

                    if (iX > 0)
                    {
                        tempDGV2.CurrentCell = null;    // 2017/10/03
                    }

                    dr.Close();
                    cArea.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
                }

                tempDGV2.CurrentCell = null;

            }

            /// ----------------------------------------------------------------------
            /// <summary>
            ///     チラシ別の予定枚数と配布枚数表示 </summary>
            /// <param name="tempDGV">
            ///     データグリッド名</param>
            /// ----------------------------------------------------------------------
            public static void MaisuSubTotal(DataGridView tempDGV)
            {
                if (tempDGV.Rows.Count == 0) return; 
                
                string cName = "";
                int cSum = 0;
                int hSum = 0;
                int i;

                // 予定枚数と配布枚数の合計表示　2015/07/15
                for (i = 0; i < tempDGV.Rows.Count ; i++)
                {
                    if ((cName != "") && (cName != tempDGV[1, i].Value.ToString()))
                    {
                        tempDGV[9, i - 1].Value = cSum;     // 予定枚数
                        tempDGV[11, i - 1].Value = hSum;    // 配布枚数
                        cSum = 0;
                        hSum = 0;
                    }

                    cSum += Utility.nullToInt(tempDGV[8, i].Value);       // 予定枚数
                    hSum += Utility.nullToInt(tempDGV[10, i].Value);      // 配布枚数
                    cName = tempDGV[1, i].Value.ToString();
                }

                tempDGV[9, i - 1].Value = cSum;
                tempDGV[11, i - 1].Value = hSum;
            }

            ///--------------------------------------------------------
            /// <summary>
            ///     完了区分が[1]のデータは赤明細 </summary>
            /// <param name="tempDGV">
            ///     データグリッドビュー名</param>
            /// <param name="ex">
            ///     行番号</param>
            ///--------------------------------------------------------
            public static void KanryoColorShow(DataGridView tempDGV,int ex)
            {
                //完了区分
                if (tempDGV[COLUMN_KANRYO, ex].Value.ToString() == STATUS_KANRYO)
                {
                    tempDGV.Rows[ex].DefaultCellStyle.ForeColor = Color.Red;
                }
                else
                {
                    tempDGV.Rows[ex].DefaultCellStyle.ForeColor = Color.Black;
                }

                tempDGV.CurrentCell = null;
            }
        }

        private class DataTemp
        {
            //表示データ
            private string F_ID;        //ID

            public string ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }

            private string F_CName;     //チラシ名

            public string CName
            {
                get { return F_CName; }
                set { F_CName = value; }
            }

            private string F_HDate;     //配布日

            public string HDate
            {
                get { return F_HDate; }
                set { F_HDate = value; }
            }

            private string F_IDate;     //入力日

            public string IDate
            {
                get { return F_IDate; }
                set { F_IDate = value; }
            }

            private string F_Name;         //配布員氏名

            public string Name
            {
                get { return F_Name; }
                set { F_Name = value; }
            }

            private string F_Kotsuhi;   //交通費

            public string Kotsuhi
            {
                get { return F_Kotsuhi; }
                set { F_Kotsuhi = value; }
            }

            private string F_AddDate;   //登録年月日

            public string AddDate
            {
                get { return F_AddDate; }
                set { F_AddDate = value; }
            }
            private string F_UppDate;

            public string UppDate       //変更年月日
            {
                get { return F_UppDate; }
                set { F_UppDate = value; }
            }

            private string F_Memo;      //本日の注意事項

            public string Memo
            {
                get { return F_Memo; }
                set { F_Memo = value; }
            }

            // ログインユーザー
            public string loginUser { get; set; }
        }

        // グリッドからデータを選択
        private void GridEnter(int tempiX)
        {
            try
            {
                // 配布指示データを取得する
                if (!GridviewSet.GetData(dataGridView1, ref cMaster, tempiX))
                {
                    MessageBox.Show("該当するデータが登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // 配布指示データ値を表示
                jDate.Value = cMaster.配布日;

                // 配布員名表示
                txtStaffID.Text = cMaster.配布員ID.ToString();

                OleDbDataReader dR;
                Control.配布員 cStaff = new Control.配布員();
                dR = cStaff.FillBy("where ID = " + cMaster.配布員ID.ToString());
                while (dR.Read())
                {
                    label13.Text = dR["氏名"].ToString();                        
                }

                dR.Close();
                cStaff.Close();

                // 入力日
                nDate.Value = cMaster.入力日;
                txthID.Text = cMaster.ID.ToString();
                txtKoutsu.Text = cMaster.交通費.ToString("#,##0");
                txtChuui.Text = cMaster.注意事項;

                Control.受注 cOrder = new Control.受注();

                // ボタン状態
                btnDel.Enabled = true;
                btnClr.Enabled = true;

                fMode.Mode = 1;     // フォームモードステータス:変更削除

                jDate.Focus();

                GridViewEnable(dataGridView1, false);

                button1.Enabled = false;

                // 未登録チラシデータのステータスを戻す
                StatusBack();

                // 配布エリアデータ表示
                GridviewSet.ShowDataItem(dataGridView1, dataGridView2,tempiX);

                // チラシ別枚数表示
                GridviewSet.MaisuSubTotal(dataGridView2);

                // 配布指示書印刷ボタン
                if (dataGridView2.Rows.Count > 0)
                {
                    button4.Enabled = true;
                }

                // 支給控除明細登録ボタン
                button6.Enabled = true;

                // 天候表示
                tenkouUpdate();

                // 未完了明細がある時
                if (GetHaifuKanryo(dataGridView2) == false)
                {
                    // 入力日を今日の日付とする
                    nDate.Value = DateTime.Parse(DateTime.Today.ToShortDateString());
                     
                    // 配布日を前日の日付とする 2010/1/18
                    jDate.Value = DateTime.Parse(DateTime.Today.AddDays(-1).ToShortDateString());
                }

                // 受注@特記事項を注意事項に表示 2015/11/27
                txtChuui.Text = setTokkijikou(cMaster.ID);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "データ表示", MessageBoxButtons.OK);
            }
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     配布指示書に対応する受注データの特記事項を取得する </summary>
        /// <param name="sID">
        ///     配布指示ID</param>
        /// <returns>
        ///     特記事項文字列</returns>
        ///--------------------------------------------------------------------------------------
        private string setTokkijikou(int sID)
        {
            darwinDataSetTableAdapters.受注1TableAdapter dAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
            darwinDataSetTableAdapters.配布エリアTableAdapter aAdp = new darwinDataSetTableAdapters.配布エリアTableAdapter();
 
            darwinDataSet dts = new darwinDataSet();

            dAdp.Fill(dts.受注1);
            aAdp.Fill(dts.配布エリア);

            var s = dts.配布エリア.Where(a => a.配布指示ID == sID && a.受注1Row != null)
                                .Select(b => new
                                {
                                    sjID = b.配布指示ID,
                                    sjMemo = b.受注1Row.特記事項
                                });

            string wMemo = string.Empty;
            string w = string.Empty;

            foreach (var t in s.OrderBy(a => a.sjMemo))
            {
                if (w != t.sjMemo)
                {
                    if (wMemo == string.Empty)
                    {
                        wMemo = t.sjMemo;
                    }
                    else
                    {
                        ////複数の特記事項の場合は連結して追加します
                        //wMemo += ("　　" + t.sjMemo);

                        //複数の特記事項の場合は改行して追加します 2016/04/04
                        wMemo += Environment.NewLine + t.sjMemo;
                    }
                }

                w = t.sjMemo; 
            }

            // 特記事項を返す
            return wMemo;
        }
        
        // 配布エリアグリッドからデータを選択
        private void GridEnter_Haifu(int tempRow)
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView2[1, tempRow].Value + " " + dataGridView2[4, tempRow].Value + " が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "チラシデータ選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {

                    //チラシデータを取得する
                    //if (GridviewSet.GetDataItem(Int32.Parse(dataGridView2[1, dataGridView2.SelectedRows[iX].Index].Value.ToString()),ref cArea) == false)
                    //{
                    //    MessageBox.Show("該当するデータが登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    //    return;
                    //}

                    //チラシデータ編集画面を表示
                    frmHaifuShijiSUb2 frm = new frmHaifuShijiSUb2();
                    iX = tempRow;

                    if (txthID.Text.ToString() == "")
                    {
                        frm.SID = 0;
                    }
                    else
                    {
                        frm.SID = Int32.Parse(txthID.Text.ToString(), System.Globalization.NumberStyles.Any);
                    }

                    frm.hDate = jDate.Value.ToShortDateString();
                    frm.staffName = label13.Text;
                    frm.ID = Int32.Parse(dataGridView2[0, iX].Value.ToString());
                    frm.cName = dataGridView2[1, iX].Value.ToString();
                    frm.fJyouken = dataGridView2[2, iX].Value.ToString();
                    frm.fKeitai = dataGridView2[3, iX].Value.ToString();
                    frm.Add = dataGridView2[4, iX].Value.ToString();
                    frm.Tanka = double.Parse(dataGridView2[7, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                    frm.yMaisu = Int32.Parse(dataGridView2[8, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                    frm.hMaisu = Int32.Parse(dataGridView2[10, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                    frm.kanryo = Int32.Parse(dataGridView2[12, iX].Value.ToString(), System.Globalization.NumberStyles.Any);    // 2015/07/15
                    frm.Edaban = dataGridView2[5, iX].Value.ToString();
                    frm.Edaban_Status = int.Parse(dataGridView2[17, iX].Value.ToString());    // 2015/07/15

                    //未配布情報関連
                    frm.Mihaifu_Status = int.Parse(dataGridView2[16, iX].Value.ToString());    // 2015/07/15
                    //frm.Banchi = dataGridView2[17, iX].Value.ToString();
                    //frm.Manshon = dataGridView2[18, iX].Value.ToString();
                    //frm.Riyu =  int.Parse(dataGridView2[19, iX].Value.ToString());
                    //frm.Sonota = dataGridView2[20, iX].Value.ToString();

                    //編集画面
                    frm.ShowDialog();

                    //値を戻す
                    //dataGridView2[7, iX].Value = frm.Tanka;     //単価
                    //dataGridView2[10, iX].Value = frm.hMaisu;   //配布枚数
                    //dataGridView2[12, iX].Value = frm.kanryo;   //完了区分    // 2015/07/15
                    dataGridView2[5, iX].Value = frm.Edaban;    //枝番記入

                    //dataGridView2[17, iX].Value = frm.Banchi;   //番地・号
                    //dataGridView2[18, iX].Value = frm.Manshon;  //マンション名
                    //dataGridView2[19, iX].Value = frm.Riyu;     //理由
                    //dataGridView2[20, iX].Value = frm.Sonota;   //その他

                    //完了データは赤表示
                    GridviewSet.KanryoColorShow(dataGridView2, iX);

                    frm.Dispose();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "データ表示", MessageBoxButtons.OK);
                }
            }

            //登録画面終了後、次の行へカーソルを移動する　2009/10/29
            if (tempRow == dataGridView2.RowCount - 1)
            {
                dataGridView2.CurrentCell = dataGridView2[1, 0];　//最下行のときは最上行へ移動する
            }
            else
            {
                dataGridView2.CurrentCell = dataGridView2[1, tempRow + 1];
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            // Enterkey以外は対象外
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            //確認
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += "指示番号 " + dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value + " が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "配布指示・報告書選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }

            GridEnter(dataGridView1.SelectedRows[iX].Index);
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //GridEnter();
        }

        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {

            try
            {
                //StartDate.Value = DateTime.Today;
                //StartDate.Checked = false;
                //EndDate.Value = DateTime.Today;
                //EndDate.Checked = false;

                //nStartDate.Value = DateTime.Today;
                //nStartDate.Checked = false;
                //nEndDate.Value = DateTime.Today;
                //nEndDate.Checked = false;

                cmbsStaff.SelectedIndex = -1;

                button1.Enabled = true;

                txthID.Text = "";

                jDate.Value = DateTime.Today.AddDays(-1);   //デフォルトは前日 2010/1/18
                nDate.Value = DateTime.Today;

                //mStartTime.Text = "";
                //mEndTime.Text = "";

                //cmbStaff.SelectedIndex = -1;

                txtStaffID.Text = "";
                label13.Text = "";

                txtKoutsu.Text = "0";

                //cmbRiyu.SelectedIndex = -1;
                //txtMemo2.Text = "";

                txtChuui.Text = "";
                txtTenkou.Text = "";

                btnDel.Enabled = false;
                //btnClr.Enabled = false;

                dataGridView1.Enabled = true;

                //txtCode.Focus();
                jDate.Focus();

                dataGridView2.RowCount = 0;

                button4.Enabled = false;
                button6.Enabled = false;

                fMode.Mode = 0;

                GridViewEnable(dataGridView1, true);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        /// <summary>
        /// 未登録チラシデータのステータスを0に戻す
        /// </summary>
        private void StatusBack()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2[14, i].Value.ToString() == "1")
                {
                    //データベース更新
                    HaihuAreaUpdate(Int32.Parse(dataGridView2[0, i].Value.ToString()),
                     (int)(0), (int)(0), (int)(0), (int)(0), (int)(0), (int)(0),"");

                    //未配布情報削除
                    string sqlStr;
                    Control.FreeSql fCon = new Control.FreeSql();
                    sqlStr = "";
                    sqlStr += "delete from 未配布情報 ";
                    sqlStr += "where 配布エリアID = " + Int32.Parse(dataGridView2[0, i].Value.ToString());

                    if (fCon.Execute(sqlStr) == false)
                    {
                        MessageBox.Show("未配布情報の削除に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                    fCon.Close();
                }
            }

        }

        /// <summary>
        /// データグリッドビューステータス表示切替
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        /// <param name="tempBool">状態（true:false）</param>
        private void GridViewEnable(DataGridView tempDGV,bool tempBool)
        {
            if (tempBool == true)
            {
                for (int i = 0; i < tempDGV.Rows.Count; i++)
                {
                    tempDGV.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
                }

                tempDGV.Enabled = true;
            }
            else
            {
                for (int i = 0; i < tempDGV.Rows.Count; i++)
                {
                    tempDGV.Rows[i].DefaultCellStyle.ForeColor = Color.LightGray;
                }

                tempDGV.Enabled = false;
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("選択されているデータを変更しないで終了します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question)== DialogResult.No )
                return;

            //未登録チラシデータを未選択状態に戻す
            StatusBack();

            //画面表示を初期化
            DispClear();
        }

        //配布未完了エリアデータがあるか判断
        private bool GetHaifuKanryo(DataGridView dGv)
        {
            //int iX = 0;

            bool rtn = true;

            foreach (DataGridViewRow  r in dGv.Rows)
            {
                // 完了・未完了の判断を修正：2017/05/18
                if (dGv["col12", dGv.Rows[r.Index].Index].Value.ToString() == STATUS_MIKANRYO)
                {
                    rtn = false;
                    break;
                }
            }

            return rtn;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                //// 配布員は必須入力：2018/01/04
                //if (Utility.strToInt(txtStaffID.Text) == 0)
                //{
                //    MessageBox.Show("配布員が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    return;
                //}

                // 全ての明細が配布完了か調べる 2018/03/02
                bool kanryo = true;
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    if (dataGridView2[12, i].Value.ToString() != STATUS_KANRYO)
                    {
                        kanryo = false;
                        break;
                    }
                }

                // 全ての明細が配布完了で配布員が未入力のときアラート表示 2018/03/02
                if (kanryo && (Utility.strToInt(txtStaffID.Text) == 0))
                {
                    MessageBox.Show("配布員が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                
                // 枝番未入力チェックアラート：2018/01/11
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    if (dataGridView2[17, i].Value.ToString() == global.FLGON.ToString())
                    {
                        if (dataGridView2[5, i].Value.ToString() == string.Empty)
                        {
                            string cName = dataGridView2[1, i].Value.ToString();

                            if (MessageBox.Show("枝番未入力の配布明細があります。実行してよろしいですか？" + Environment.NewLine + Environment.NewLine + cName, "枝番未入力データあり", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                            {
                                dataGridView2.Focus();
                                dataGridView2.CurrentCell = dataGridView2[5, i];
                                return;
                            }
                        }
                    }
                }

                // 配布枚数０で配布完了チェック：2018/02/20
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    if (Utility.strToInt(dataGridView2[10, i].Value.ToString()) == global.FLGOFF)
                    {
                        if (dataGridView2[12, i].Value.ToString() == STATUS_KANRYO)
                        {
                            string cName = dataGridView2[1, i].Value.ToString();

                            if (MessageBox.Show("配布枚数０の町目があります。よろしいですか？" + Environment.NewLine + Environment.NewLine + cName, "配布枚数０町目あり", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                            {
                                dataGridView2.Focus();
                                dataGridView2.CurrentCell = dataGridView2[10, i];
                                return;
                            }
                        }
                    }
                }

                if (!GetHaifuKanryo(dataGridView2))
                {
                    if (MessageBox.Show("未完了の配布明細があります。実行してよろしいですか？", "未完了データあり", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) return;
                }

                if (MessageBox.Show("登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                            
                if (fDataCheck())
                {
                    Control.DataControl Con;
                    OleDbConnection cn;
                    OleDbTransaction tran;
                    OleDbCommand SCom;

                    switch (fMode.Mode)
                    {
                        case 0: 
                        
                            //IDを採番
                            string sqlStr = "";
                            int gID = (int)(1);

                            sqlStr = "select max(ID) as ID from 配布指示 ";
                            OleDbDataReader dR;
                            Control.FreeSql fCon = new Control.FreeSql();
                            dR = fCon.free_dsReader(sqlStr);

                            while (dR.Read())
                            {
                                if (dR["ID"] == DBNull.Value )
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
                                // 配布指示データ登録処理                                
                                sqlStr = "";
                                sqlStr += "insert into 配布指示 ";
                                sqlStr += "(ID,配布日,入力日,配布員ID,交通費,交通区間開始,交通区間終了,";
                                sqlStr += "配布開始時刻,配布終了時刻,終了レポート,未配布区分,未配布理由,";
                                sqlStr += "登録年月日,変更年月日,ユーザーID) ";
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
                                sqlStr += "'" + cMaster.変更年月日 + "',";
                                sqlStr += global.loginUserID.ToString() + ")";  // 2016/09/26 ユーザーID追加

                                SCom.CommandText = sqlStr;

                                SCom.ExecuteNonQuery();

                                //配布エリアデータ更新
                                string sID, sTanka, sMaisu, sMaisu2, sKanryo, sStatus, sEdaban;

                                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                                {
                                    if (dataGridView2[14, i].Value.ToString() != "1") //削除フラグ判断
                                    {
                                        sID = dataGridView2[0, i].Value.ToString();
                                        sTanka = Utility.nullToDouble(dataGridView2[7, i].Value).ToString();
                                        sMaisu = Utility.nullToInt(dataGridView2[10, i].Value).ToString();
                                        //sMaisu2 = dataGridView2[12, i].Value.ToString();
                                        //sKanryo = dataGridView2[11, i].Value.ToString();
                                        sMaisu2 = dataGridView2[13, i].Value.ToString();

                                        // 2016/10/31
                                        if (dataGridView2[12, i].Value.ToString() == STATUS_KANRYO)
                                        {
                                            sKanryo = STATUS_KANRYO;
                                        }
                                        else
                                        {
                                            sKanryo = STATUS_MIKANRYO;
                                        }

                                        sStatus = "2";
                                        sEdaban = dataGridView2[5, i].Value.ToString() + "";

                                        //mBanchi = dataGridView2[17, i].Value.ToString() + "";
                                        //mManshon = dataGridView2[18, i].Value.ToString() + "";
                                        //mRiyu = int.Parse(dataGridView2[19, i].Value.ToString(),System.Globalization.NumberStyles.Any);
                                        //mSonota = dataGridView2[20, i].Value.ToString() + ""; 

                                        sqlStr = "";
                                        sqlStr += "update 配布エリア ";
                                        sqlStr += "set ";
                                        sqlStr += "配布指示ID = " + gID.ToString() + ",";
                                        sqlStr += "配布単価 = " + sTanka + ",";
                                        sqlStr += "実配布枚数 = " + sMaisu + ",";
                                        sqlStr += "実残数 = 予定枚数 - " + sMaisu + ",";
                                        sqlStr += "報告枚数 = " + sMaisu2 + ",";
                                        sqlStr += "報告残数 = 予定枚数 - " + sMaisu2 + ",";
                                        sqlStr += "完了区分 = " + sKanryo + ",";
                                        sqlStr += "ステータス = " + sStatus + ",";
                                        sqlStr += "枝番記入 = '" + sEdaban + "',";

                                        //sqlStr += "番地号 = '" + mBanchi + "',";
                                        //sqlStr += "マンション名 = '" + mManshon + "',";
                                        //sqlStr += "理由 = " + mRiyu.ToString() + ",";
                                        //sqlStr += "その他内容 = '" + mSonota + "',";

                                        sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
                                        sqlStr += "where (配布エリア.ID = " + sID + ") and ";
                                        sqlStr += "(ステータス <> 0)";

                                        SCom.CommandText = sqlStr;

                                        SCom.ExecuteNonQuery();
                                    }
                                }

                                tran.Commit();
                                MessageBox.Show("新規登録されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();

                                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                MessageBox.Show("新規登録に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                //未登録配布データのステータスを戻す
                                StatusBack();
                            }

                            cn.Close();

                            Con.Close();

                            break;

                        case 1: //更新

                            //更新処理準備
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
                                //配布指示データ更新
                                sqlStr = "";
                                sqlStr += "update 配布指示 set ";
                                sqlStr += "配布日 = '" + cMaster.配布日 + "',";
                                sqlStr += "入力日 = '" + cMaster.入力日 + "',";
                                sqlStr += "配布員ID = " + cMaster.配布員ID + ",";
                                sqlStr += "交通費 = " + cMaster.交通費 + ",";
                                sqlStr += "交通区間開始 = '" + cMaster.交通区間開始 + "',";
                                sqlStr += "交通区間終了 = '" + cMaster.交通区間終了 + "',";
                                sqlStr += "配布開始時刻 = '" + cMaster.配布開始時刻 + "',";
                                sqlStr += "配布終了時刻 = '" + cMaster.配布終了時刻 + "',";
                                sqlStr += "終了レポート = '" + cMaster.終了レポート + "',";
                                sqlStr += "未配布区分 = '" + cMaster.未配布区分 + "',";
                                sqlStr += "未配布理由 = '" + cMaster.未配布理由 + "',";
                                sqlStr += "注意事項 = '" + cMaster.注意事項 + "',";
                                sqlStr += "変更年月日 = '" + DateTime.Today + "',";
                                sqlStr += "ユーザーID = " + global.loginUserID.ToString() + " "; // 2016/09/26
                                sqlStr += "where ID = " + cMaster.ID;

                                SCom.CommandText = sqlStr;

                                // SQLの実行
                                SCom.ExecuteNonQuery();

                                //配布エリアデータ更新
                                string sID, sTanka, sMaisu, sMaisu2, sKanryo, sStatus, sEdaban;

                                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                                {
                                    if (dataGridView2[14, i].Value.ToString() != "1") //削除フラグ判断
                                    {
                                        sID = dataGridView2[0, i].Value.ToString();
                                        sTanka = Utility.nullToDouble(dataGridView2[7, i].Value).ToString();
                                        sMaisu = Utility.nullToInt(dataGridView2[10, i].Value).ToString();
                                        //sMaisu2 = dataGridView2[12, i].Value.ToString();
                                        //sKanryo = dataGridView2[11, i].Value.ToString();
                                        sMaisu2 = dataGridView2[13, i].Value.ToString();

                                        // 2016/10/31
                                        if (dataGridView2[12, i].Value.ToString() == STATUS_KANRYO)
                                        {
                                            sKanryo = STATUS_KANRYO;
                                        }
                                        else
                                        {
                                            sKanryo = STATUS_MIKANRYO;
                                        }

                                        sStatus = "2";
                                        sEdaban = dataGridView2[5, i].Value.ToString() + "";

                                        //mBanchi = dataGridView2[17, i].Value.ToString() + "";
                                        //mManshon = dataGridView2[18, i].Value.ToString() + "";
                                        //mRiyu = int.Parse(dataGridView2[19, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                                        //mSonota = dataGridView2[20, i].Value.ToString() + ""; 

                                        sqlStr = "";
                                        sqlStr += "update 配布エリア ";
                                        sqlStr += "set ";
                                        sqlStr += "配布指示ID = " + txthID.Text + ",";
                                        sqlStr += "配布単価 = " + sTanka + ",";
                                        sqlStr += "実配布枚数 = " + sMaisu + ",";
                                        sqlStr += "実残数 = 予定枚数 - " + sMaisu + ",";
                                        sqlStr += "報告枚数 = " + sMaisu2 + ",";
                                        sqlStr += "報告残数 = 予定枚数 - " + sMaisu2 + ",";
                                        sqlStr += "完了区分 = " + sKanryo + ",";
                                        sqlStr += "ステータス = " + sStatus + ",";
                                        sqlStr += "枝番記入 = '" + sEdaban + "',";

                                        //sqlStr += "番地号 = '" + mBanchi + "',";
                                        //sqlStr += "マンション名 = '" + mManshon + "',";
                                        //sqlStr += "理由 = " + mRiyu.ToString() + ",";
                                        //sqlStr += "その他内容 = '" + mSonota + "',";

                                        sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
                                        sqlStr += "where (配布エリア.ID = " + sID + ") and ";
                                        sqlStr += "(ステータス <> 0)";

                                        SCom.CommandText = sqlStr;

                                        // SQLの実行
                                        SCom.ExecuteNonQuery();
                                    }
                                }

                                tran.Commit();
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();

                                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                MessageBox.Show("更新に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }

                            cn.Close();
                            Con.Close();
                            break;
                    }

                    DispClear();

                    //グリッド再表示
                    //GridDataShow(StartDate, EndDate,nStartDate,nEndDate, cmbsStaff);
                    dataGridView1.RowCount = 0;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"更新処理",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        //登録データチェック
        private Boolean fDataCheck()
        {

            try
            {

                //登録モードのとき、コードをチェック
                if (fMode.Mode == 0)
                {

                    //// 数字か？
                    //if (txtCode.Text == null)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("コードは数字で入力してください");
                    //}

                    //str = this.txtCode.Text;

                    //if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    //{
                    //}
                    //else
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("コードは数字で入力してください");
                    //}

                    //// 未入力またはスペースのみは不可
                    //if ((this.txtCode.Text).Trim().Length < 1)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("コードを入力してください");
                    //}

                    ////ゼロは不可
                    //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("ゼロは登録できません");
                    //}

                    ////登録済みコードか調べる
                    //string sqlStr;
                    //Control.受注 Order = new Control.受注();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Order.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Order.Close();
                    //    throw new Exception("既に登録済みのコードです");
                    //}

                    //dr.Close();
                    //Order.Close();

                }

                //クラスにデータセット
                cMaster.配布日 = jDate.Value;
                cMaster.入力日 = nDate.Value;

                if (txtStaffID.Text == "")
                {
                    cMaster.配布員ID = 0;
                }
                else
                {
                    cMaster.配布員ID = int.Parse(txtStaffID.Text);
                }

                cMaster.交通費 = Int32.Parse(txtKoutsu.Text, System.Globalization.NumberStyles.Any);
                cMaster.交通区間開始 = "";
                cMaster.交通区間終了 = "";

                cMaster.配布開始時刻 = "";
                cMaster.配布終了時刻 = "";

                cMaster.終了レポート = "";

                cMaster.未配布区分 = "";
                cMaster.未配布理由 = "";

                cMaster.注意事項 = txtChuui.Text;

                if (fMode.Mode == 0) cMaster.登録年月日 = DateTime.Today;
                cMaster.変更年月日 = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void txtEnter(object sender, EventArgs e)
        {
            TextBox objtxt = new TextBox();
            MaskedTextBox objMtxt = new MaskedTextBox();

            if (sender == txtStaffID)
            {
                objtxt = txtStaffID;
                if (txtStaffID.Text == "0")
                {
                    txtStaffID.Text = "";
                }
            }

            if (sender == txtKoutsu)
            {
                objtxt = txtKoutsu;

                dataGridView2.CurrentCell = null;   // 2017/10/03
            }

            if (sender == txtsID) objtxt = txtsID;
            if (sender == txtsCName) objtxt = txtsCName;

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

            objMtxt.SelectAll();
            objMtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {
            
            TextBox objtxt = new TextBox();
            MaskedTextBox objMtxt = new MaskedTextBox();

            if (sender == txtStaffID) objtxt = txtStaffID;            
            if (sender == txtKoutsu) objtxt = txtKoutsu;
            if (sender == txtsID) objtxt = txtsID;
            if (sender == txtsCName) objtxt = txtsCName;

            objtxt.BackColor = Color.White;
            objMtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //未登録チラシデータのステータスを戻す
            StatusBack();

            //データ削除
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();
            cn = Con.GetConnection();

            OleDbTransaction tran;

            //トランザクション開始
            tran = cn.BeginTransaction();

            OleDbCommand SCom = new OleDbCommand();

            SCom.Connection = cn;
            SCom.Transaction = tran;

            string sqlSTR;

            try
            {
                //配布指示データ削除
                sqlSTR = "";
                sqlSTR += "delete from 配布指示 ";
                sqlSTR += "where ID = " + cMaster.ID.ToString();

                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                //未配布情報の削除
                sqlSTR = "";
                sqlSTR += "delete a FROM 未配布情報 as a inner join 配布エリア as b ";
                sqlSTR += "on a.配布エリアID = b.ID ";
                sqlSTR += "where b.配布指示ID = " + cMaster.ID.ToString();
                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                //配布エリアデータの初期化
                sqlSTR = "";
                sqlSTR += "update 配布エリア ";
                sqlSTR += "set ";
                sqlSTR += "配布指示ID = 0,";
                //sqlSTR += "配布単価 = 0,";
                sqlSTR += "実配布枚数 = 0,";
                sqlSTR += "実残数 = 予定枚数,";
                sqlSTR += "報告枚数 = 0,";
                sqlSTR += "報告残数 = 予定枚数,";
                sqlSTR += "完了区分 = 0,";
                sqlSTR += "ステータス = 0,";
                sqlSTR += "枝番記入 = '',";

                //sqlSTR += "番地号 = '',";
                //sqlSTR += "マンション名 = '',";
                //sqlSTR += "理由 = 0,";
                //sqlSTR += "その他内容 = '',";

                sqlSTR += "変更年月日 = '" + DateTime.Today + "' ";
                sqlSTR += "where 配布エリア.配布指示ID = " + cMaster.ID.ToString();

                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                tran.Commit();

                MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                tran.Rollback();

                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MessageBox.Show("削除に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            cn.Close();

            Con.Close();

            DispClear();

            //グリッド再表示
            //GridDataShow(StartDate, EndDate, nStartDate, nEndDate, cmbsStaff);
            dataGridView1.RowCount = 0;
        }

        private void btnCsv_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            GridEnter(0);
        }

        //private void cmbRiyuSet()
        //{
        //    cmbRiyu.Items.Clear();
        //    cmbRiyu.Items.Add("管理人");
        //    cmbRiyu.Items.Add("住民");
        //    cmbRiyu.Items.Add("厳重貼紙");
        //    cmbRiyu.Items.Add("その他");
        //}


        private void button1_Click(object sender, EventArgs e)
        {
            if ((txtsID.Text.Trim().Length > 0) && (Utility.NumericCheck(txtsID.Text) == false))
            {
                MessageBox.Show("検索用配布指示IDは数字で入力してください", "検索配布指示ID", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            GridDataShow(cmbsStaff,txtsID,txtsCName);
        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            // 確認
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += "指示番号 " + dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value + " が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "配布指示・報告書選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }

            Cursor = Cursors.WaitCursor;
            GridEnter(dataGridView1.SelectedRows[iX].Index);
            Cursor = Cursors.Default;
        }

        //////private void GridDataShow(DateTimePicker d1, DateTimePicker d2, DateTimePicker d3, DateTimePicker d4, ComboBox tempCmb, TextBox tempID,TextBox tempCName)

        private void GridDataShow(ComboBox tempCmb, TextBox tempID,TextBox tempCName)
        {
            //DateTime Date1;
            //DateTime Date2;
            //DateTime Date3;
            //DateTime Date4;

            int SID;

            ////////配布開始日
            //////if (d1.Checked == true)
            //////{
            //////    Date1 = d1.Value;
            //////}
            //////else
            //////{
            //////    Date1 = Convert.ToDateTime("1900/01/01");
            //////}

            ////////配布終了日
            //////if (d2.Checked == true)
            //////{
            //////    Date2 = d2.Value;
            //////}
            //////else
            //////{
            //////    Date2 = Convert.ToDateTime("2999/12/31");
            //////}

            ////////入力開始日
            //////if (d3.Checked == true)
            //////{
            //////    Date3 = d3.Value;
            //////}
            //////else
            //////{
            //////    Date3 = Convert.ToDateTime("1900/01/01");
            //////}

            ////////入力終了日
            //////if (d4.Checked == true)
            //////{
            //////    Date4 = d4.Value;
            //////}
            //////else
            //////{
            //////    Date4 = Convert.ToDateTime("2999/12/31");
            //////}

            //配布員ID
            if (tempCmb.SelectedIndex != -1)
            {
                Utility.ComboStaff cmb1 = new Utility.ComboStaff();
                cmb1 = (Utility.ComboStaff)tempCmb.SelectedItem;
                SID = cmb1.ID;
            }
            else
            {
                SID = (int)(-1);
            }

            this.Cursor = Cursors.WaitCursor;

            //グリッド表示
            GridviewSet.ShowData(dataGridView1,SID,tempID.Text,tempCName.Text);
            this.Cursor = Cursors.Default;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //他のマシンで配布指示実行中か検証する
            int cFlg = 0;

            OleDbDataReader dRflg;
            Control.会社情報 cSystem = new Control.会社情報();
            dRflg = cSystem.Fill();

            while (dRflg.Read())
            {
                cFlg = int.Parse(dRflg["配布フラグ"].ToString());
            }

            dRflg.Close();
            cSystem.Close();

            if (cFlg == 1)
            {
                MessageBox.Show("現在、他のマシンで配布指示登録中です。", "起動チェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //未指示配布データの登録
            frmHaifuShijiSUb fSub = new frmHaifuShijiSUb();

            long jID;
            int cID;
            string msgHD = "";
            string msgSTR = "";
            string d1, d2;

            if (fSub.ShowDialog(this) == DialogResult.OK)
            {

                for (int i = 0; i < fSub.Count; i++)
                {
                    //当日、同じ受注ID、同じ町名IDの登録のときメッセージ表示
                    //①受注IDと町名IDを取得
                    OleDbDataReader dR,dR2;
                    Control.配布エリア cArea = new Control.配布エリア();
                    dR = cArea.FillBy("where ID = " + fSub[i].ToString());

                    while (dR.Read())
                    {
                        jID = long.Parse(dR["受注ID"].ToString());
                        cID = int.Parse(dR["町名ID"].ToString());

                        string sqlStr;
                        Control.FreeSql fCon = new Control.FreeSql();

                        sqlStr = "";
                        sqlStr += "select 受注.チラシ名,町名.名称,配布エリア.配布指示ID,配布エリア.予定枚数,";
                        sqlStr += "配布指示.配布日,配布エリア.町名ID ";
                        sqlStr += "from 配布エリア ";
                        sqlStr += "inner join 配布指示 on 配布エリア.配布指示ID = 配布指示.ID ";
                        sqlStr += "left join 受注 on 配布エリア.受注ID = 受注.ID ";
                        sqlStr += "left join 町名 on 配布エリア.町名ID = 町名.ID ";
                        sqlStr += "where ";
                        sqlStr += "(配布エリア.受注ID = " + jID.ToString() + ") and ";
                        sqlStr += "(配布エリア.町名ID = " + cID.ToString() + ") and ";
                        sqlStr += "(配布エリア.配布指示ID != 0) ";

                        dR2 = fCon.free_dsReader(sqlStr);

                        msgHD = "";
                        msgSTR = "";

                        while(dR2.Read())
                        {
                            if (dR2["チラシ名"] == DBNull.Value )
                            {
                                d1 = "チラシ名：　";
                            }
                            else
                            {
                                d1 = "チラシ名：" + dR2["チラシ名"].ToString() + "　";
                            }

                            if (dR2["名称"] == DBNull.Value)
                            {
                                d2 = dR2["町名ID"].ToString() + "　";
                            }
                            else
                            {
                                d2 = dR2["町名ID"].ToString() + "：" + dR2["名称"].ToString();
                            }

                            msgHD = d1 + d2 + Environment.NewLine + Environment.NewLine;
                            
                            msgSTR += "指示番号：" + dR2["配布指示ID"].ToString() + "　";
                            msgSTR += "配布日：" + DateTime.Parse(dR2["配布日"].ToString()).ToShortDateString() + "　";
                            msgSTR += "配布枚数：" + dR2["予定枚数"].ToString() + Environment.NewLine;
                        }

                        dR2.Close();
                        fCon.Close();

                        //②メッセージ表示
                        if (msgSTR != "")
                        {
                            MessageBox.Show(msgHD + msgSTR, "過去の同様の配布指示",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        }

                    }

                    dR.Close();
                    cArea.Close();



                    //グリッド明細行追加
                    HaifuItemAdd(int.Parse(fSub[i].ToString()));

                    //配布エリアデータステータスを[1]に書き換え
                    HaihuStatusUpdate(int.Parse(fSub[i].ToString()));   
                }

                button4.Enabled = false;
                
            }
            else
            {
            }

            fSub.Dispose();

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            string sqlSTRING;

            if (e.ColumnIndex == 0)
            {

                //dataGridView2[2, e.RowIndex].Value = GetTownName(dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString());

                OleDbDataReader dR;

                sqlSTRING = "";
                sqlSTRING += "select 配布エリア.ID,配布エリア.受注ID,受注.チラシ名,";
                sqlSTRING += "配布エリア.町名ID,町名.名称 as 町名,受注.配布開始日,受注.配布終了日,";
                sqlSTRING += "配布エリア.併配区分,配布エリア.配布単価,配布エリア.完了区分,";
                sqlSTRING += "受注.配布条件,配布形態.名称 as 配布形態,配布エリア.予定枚数,";
                sqlSTRING += "配布エリア.実配布枚数,配布エリア.報告枚数,配布エリア.枝番記入,";
                sqlSTRING += "受注.未配布情報有無,受注.枝番有無 ";

                //sqlSTRING += "配布エリア.番地号,配布エリア.マンション名,配布エリア.理由,配布エリア.その他内容 ";
            
                sqlSTRING += "from ((配布エリア left join 受注 on 配布エリア.受注ID = 受注.ID) ";
                sqlSTRING += "left join 町名 on 配布エリア.町名ID = 町名.ID) ";
                sqlSTRING += "left join 配布形態 on 受注.配布形態 = 配布形態.ID ";
                sqlSTRING += "where 配布エリア.ID = " + dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString();

                //配布指示データのデータリーダーを取得する
                Control.FreeSql cArea = new Control.FreeSql();
                dR = cArea.free_dsReader(sqlSTRING);

                //グリッドビューに表示する
                while (dR.Read())
                {
                    try
                    {
                        // チラシ別枚数表示ステータス
                        STATUS_MAISU = false;

                        // グリッドにデータ表示
                        dataGridView2[1, e.RowIndex].Value = dR["チラシ名"].ToString();
                        dataGridView2[2, e.RowIndex].Value = dR["配布条件"].ToString();
                        dataGridView2[3, e.RowIndex].Value = dR["配布形態"].ToString();
                        dataGridView2[4, e.RowIndex].Value = dR["町名"].ToString();
                        dataGridView2[5, e.RowIndex].Value = dR["枝番記入"].ToString();
                        dataGridView2[6, e.RowIndex].Value = dR["町名ID"].ToString();
                        dataGridView2[7, e.RowIndex].Value = Double.Parse(dR["配布単価"].ToString(),System.Globalization.NumberStyles.Any).ToString("#,##0.00");
                        dataGridView2[8, e.RowIndex].Value = int.Parse(dR["予定枚数"].ToString());
                        dataGridView2[10, e.RowIndex].Value = int.Parse(dR["実配布枚数"].ToString());
                        dataGridView2[12, e.RowIndex].Value = int.Parse(dR["完了区分"].ToString());    // 2015/07/15
                        dataGridView2[13, e.RowIndex].Value = int.Parse(dR["報告枚数"].ToString());    // 2015/07/15
                        dataGridView2[14, e.RowIndex].Value = "0"; //削除フラグ    // 2015/07/15
                        dataGridView2[15, e.RowIndex].Value = "0"; //追加フラグ    // 2015/07/15
                        dataGridView2[16, e.RowIndex].Value = int.Parse(dR["未配布情報有無"].ToString());    // 2015/07/15
                        dataGridView2[17, e.RowIndex].Value = int.Parse(dR["枝番有無"].ToString());    // 2015/07/15

                        //dataGridView2[18, e.RowIndex].Value = dR["番地号"].ToString();
                        //dataGridView2[19, e.RowIndex].Value = dR["マンション名"].ToString();
                        //dataGridView2[20, e.RowIndex].Value = int.Parse(dR["理由"].ToString(),System.Globalization.NumberStyles.Any);
                        //dataGridView2[21, e.RowIndex].Value = dR["その他内容"].ToString();

                        // チラシ別枚数表示ステータス
                        STATUS_MAISU = true;

                        // 枝番セル編集：2017/10/03
                        if (Utility.strToInt(dR["枝番有無"].ToString()) != 0)
                        {
                            dataGridView2[5, e.RowIndex].ReadOnly = false;
                        }
                        else
                        {
                            dataGridView2[5, e.RowIndex].ReadOnly = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }

                    //完了データは赤表示
                    GridviewSet.KanryoColorShow(dataGridView2, e.RowIndex);
                }

                dR.Close();
                cArea.Close();
            }
            else if (e.ColumnIndex == 12)
            {
                //完了データは赤表示
                GridviewSet.KanryoColorShow(dataGridView2, e.RowIndex);
            }
            else if (e.ColumnIndex == 10)
            {
                // チラシ別枚数表示
                if (STATUS_MAISU)
                {
                    GridviewSet.MaisuSubTotal(dataGridView2);
                }
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     配布明細行追加 </summary>
        /// <param name="tempID">
        ///     配布エリアID </param>
        ///----------------------------------------------------------------
        private void HaifuItemAdd(int tempID)
        {
            int iX;
            dataGridView2.Rows.Add();
            iX = dataGridView2.Rows.Count;

            dataGridView2[0, iX - 1].Value = tempID.ToString();
            dataGridView2[15, iX - 1].Value = "1";  //追加フラグ    // 2015/07/15

        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     配布エリアデータのステータスを"登録中"(1)にする </summary>
        /// <param name="tempID">
        ///     配布エリアID</param>
        ///----------------------------------------------------------------
        private void HaihuStatusUpdate(int tempID)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update 配布エリア ";
            sqlStr += "set ステータス = 1, ";
            sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
            sqlStr += "where 配布エリア.ID = " + tempID.ToString();
            
            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("配布エリアデータのステータス更新に失敗しました(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     配布指示データを更新する </summary>
        /// <param name="tempID">
        ///     ID</param>
        /// <param name="tempShijiID">
        ///     配布指示ID</param>
        /// <param name="tempTanka">
        ///     単価</param>
        /// <param name="temphMaisu">
        ///     配布枚数</param>
        /// <param name="temphMaisu2">
        ///     報告枚数</param>
        /// <param name="tempKanryo">
        ///     完了区分</param>
        /// <param name="tempStatus">
        ///     ステータス</param>
        ///-----------------------------------------------------------------
        private void HaihuAreaUpdate(int tempID,int tempShijiID,double tempTanka,int temphMaisu,
            　　　　　　　　　　　　 int temphMaisu2,int tempKanryo,int tempStatus,string tempEdaban)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update 配布エリア ";
            sqlStr += "set ";
            sqlStr += "配布指示ID = " + tempShijiID + ",";
            //sqlStr += "配布単価 = " + tempTanka + ",";
            sqlStr += "実配布枚数 = " + temphMaisu + ",";
            sqlStr += "実残数 = 予定枚数 - " + temphMaisu + ",";
            sqlStr += "報告枚数 = " + temphMaisu2 + ",";
            sqlStr += "報告残数 = 予定枚数 - " + temphMaisu2 + ",";
            sqlStr += "完了区分 = " + tempKanryo + ",";
            sqlStr += "ステータス = " + tempStatus + ",";
            sqlStr += "枝番記入 = '" + tempEdaban + "',";

            //sqlStr += "番地号 = '',";
            //sqlStr += "マンション名 = '',";
            //sqlStr += "理由 = 0,";
            //sqlStr += "その他内容 = '',";

            sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
            sqlStr += "where (配布エリア.ID = " + tempID.ToString() + ") and ";
            sqlStr += "(ステータス <> 0)";

            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("配布エリアデータの更新に失敗しました(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     配布エリアデータを未選択状態に戻す </summary>
        /// <param name="tempID">
        ///     配布指示ID</param>
        ///----------------------------------------------------------------
        private void HaihuAreaClear(int tempID)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update 配布エリア ";
            sqlStr += "set ";
            sqlStr += "配布指示ID = 0,";
            sqlStr += "配布単価 = 0,";
            sqlStr += "実配布枚数 = 0,";
            sqlStr += "実残数 = 予定枚数,";
            sqlStr += "報告枚数 = 0,";
            sqlStr += "報告残数 = 予定枚数,";
            sqlStr += "完了区分 = 0,";
            sqlStr += "ステータス = 0,";
            sqlStr += "枝番記入 = '',";

            //sqlStr += "番地号 = '',";
            //sqlStr += "マンション名 = '',";
            //sqlStr += "理由 = 0,";
            //sqlStr += "その他内容 = '',";

            sqlStr += "変更年月日 = '" + DateTime.Today + "' ";
            sqlStr += "where (配布エリア.配布指示ID = " + tempID.ToString() + ") ";

            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("配布エリアデータの初期化に失敗しました(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter_Haifu(dataGridView2.SelectedRows[0].Index);
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView2.Rows.Count == 0) return;
            if (dataGridView2.SelectedRows.Count == 0) return;

            //EnterKey押下後の行移動を禁止する
            e.Handled = true;

            GridEnter_Haifu(dataGridView2.SelectedRows[0].Index);

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     配布エリア削除 </summary>
        /// <param name="eRow">
        ///     行インデックス</param>
        ///--------------------------------------------------------------------
        private void areaDel(int eRow)
        {
            //配布エリアを削除する
            //if (dataGridView2.SelectedRows.Count == 0) return;

            if (MessageBox.Show("選択中のチラシ配布データを配布指示書から除外します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            dataGridView2[14, eRow].Value = "1";    // 2015/07/15
            dataGridView2.Rows[eRow].DefaultCellStyle.ForeColor = Color.LightGray;

            //データベース更新
            HaihuAreaUpdate(Int32.Parse(dataGridView2[0, eRow].Value.ToString()),
             (int)(0), (int)(0), (int)(0), (int)(0), (int)(0), (int)(0), "");

            //未配布情報削除
            string sqlStr;
            Control.FreeSql fCon = new Control.FreeSql();
            sqlStr = "";
            sqlStr += "delete from 未配布情報 ";
            sqlStr += "where 配布エリアID = " + Int32.Parse(dataGridView2[0, eRow].Value.ToString());

            if (fCon.Execute(sqlStr) == false)
            {
                MessageBox.Show("未配布情報の削除に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            fCon.Close();

            dataGridView2.CurrentCell = null;

            //配布指示書印刷ボタン
            button4.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            const int G_COUNT = 15; //配布指示書の明細行数

            int iX = 0;
    
            if (fMode.Mode == 0)     //グリッドから選択した配布指示書を印刷する
            {
                //複数の配布指示書を連続印刷する
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("配布指示データが選択されていません", "データ未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (MessageBox.Show("選択されている配布指示書を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                //注意事項を書き込み
                foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                {
                    ShijiMemoUpdate(dataGridView1[0, r.Index].Value.ToString(), txtChuui.Text);
                }

                //選択されている行を取得
                iX = 0;
                foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                {
                    GridEnter(r.Index);  //データ画面表示

                    if (dataGridView2.Rows.Count > 0)
                    {
                        ShijiReportSet(G_COUNT);    //印刷処理
                    }

                    StatusBack();   //未登録チラシデータを未選択状態に戻す

                    DispClear();    //画面表示を初期化

                    iX++;
                }
            }
            else
            {
                //画面表示されている配布指示書を印刷する
                if (MessageBox.Show("配布指示書を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                ShijiReportSet(G_COUNT);    //印刷処理
            }

        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     配布指示書発行データの注意書きを書き換える </summary>
        /// <param name="tempID">
        ///     配布指示ID</param>
        /// <param name="tempMemo">
        ///     注意書き文字列</param>
        ///----------------------------------------------------------------
        private void ShijiMemoUpdate(string tempID,string tempMemo)
        {
            string sqlStr;
            Control.FreeSql fCon = new Control.FreeSql();

            sqlStr = "";
            sqlStr += "update 配布指示 set ";
            sqlStr += "注意事項 = '" + tempMemo + "' ";
            sqlStr += " where 配布指示.ID = " + tempID;
            
            fCon.Execute(sqlStr);

            fCon.Close();
        }

        private void ShijiReportSet(int G_COUNT)
        {
            int pCnt;

            //ページカウント
            //pCnt = dataGridView2.Rows.Count / G_COUNT + 1;

            // 2015/11/18
            pCnt = dataGridView2.Rows.Count / G_COUNT;

            if ((dataGridView2.Rows.Count % G_COUNT) > 0)
            {
                pCnt++;
            }

            for (int i = 1; i <= pCnt; i++)
            {
                ShijiReport(pCnt, i, G_COUNT);
            }
        }

        private void ShijiReport(int tempPage, int tempCurrentPage, int tempFixRows)
        {

            const int S_GYO = 5;    //エクセルファイル明細は5行目から印字
            int dgvIndex;
            int i;

            try
            {

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル配布指示報告書 , Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    oxlsSheet.Cells[1, 1] = "指示番号：" + String.Format("{0:0000000}", int.Parse(txthID.Text, System.Globalization.NumberStyles.Any));         //指示番号
                    oxlsSheet.Cells[1, 17] = "発行日：" + DateTime.Today.ToLongDateString() + "　P." + tempCurrentPage.ToString() + "/" + tempPage.ToString();  //発行日

                    //配布エリア明細
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //データグリッドビューの行インデックスを求める

                        oxlsSheet.Cells[i + S_GYO, 1] = dataGridView2[1, dgvIndex].Value.ToString();   //チラシ名
                        oxlsSheet.Cells[i + S_GYO, 4] = dataGridView2[2, dgvIndex].Value.ToString();   //配布区分
                        oxlsSheet.Cells[i + S_GYO, 6] = dataGridView2[3, dgvIndex].Value.ToString();   //配布形態
                        oxlsSheet.Cells[i + S_GYO, 7] = dataGridView2[4, dgvIndex].Value.ToString();   //配布先住所
                        oxlsSheet.Cells[i + S_GYO, 14] = dataGridView2[6, dgvIndex].Value.ToString();   //エリアID
                        oxlsSheet.Cells[i + S_GYO, 15] = dataGridView2[7, dgvIndex].Value.ToString();   //単価
                        oxlsSheet.Cells[i + S_GYO, 16] = dataGridView2[8, dgvIndex].Value.ToString();   //予定枚数

                        if (dataGridView2[9, dgvIndex].Value != null)
                        {
                            oxlsSheet.Cells[i + S_GYO, 17] = dataGridView2[9, dgvIndex].Value.ToString();   //予定枚数計
                        }
                        
                        //本日の注意事項
                        oxlsSheet.Cells[25, 1] = txtChuui.Text;

                        //グリッド最終行のとき終了
                        if (dgvIndex == (dataGridView2.Rows.Count - 1)) break;

                        //印刷明細最大行のとき終了
                        if (i == (tempFixRows - 1)) break;

                        i++;
                    }

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

                    //////DialogResult ret;

                    ////////ダイアログボックスの初期設定
                    //////saveFileDialog1.Title = "配布指示書保存";
                    //////saveFileDialog1.OverwritePrompt = true;
                    //////saveFileDialog1.RestoreDirectory = true;
                    //////saveFileDialog1.FileName = "配布指示書_" + String.Format("{0:0000000}", int.Parse(txthID.Text, System.Globalization.NumberStyles.Any));

                    ////////複数ページのとき、ページ数も付与
                    //////if (tempPage > 1)
                    //////{
                    //////    saveFileDialog1.FileName += "_" + tempCurrentPage.ToString();
                    //////}

                    //////saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

                    ////////ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    //////string fileName;
                    //////ret = saveFileDialog1.ShowDialog();

                    //////if (ret == System.Windows.Forms.DialogResult.OK)
                    //////{
                    //////    fileName = saveFileDialog1.FileName;
                    //////    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                    //////                    Type.Missing, Type.Missing, Type.Missing,
                    //////                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    //////                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //////}

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "配布指示書", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
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
                MessageBox.Show(e.Message, "配布指示書", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            
            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void frmHaifuShiji_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("終了します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }
                
            //未登録のチラシデータを戻す
            StatusBack();

        }

        private void txtStaffID_Validating(object sender, CancelEventArgs e)
        {
            int d;
            string str;

            // 未入力またはスペースのみは可
            if ((this.txtStaffID.Text).Trim().Length < 1)
            {
                label13.Text = "";
                return;
            }

            // 数字か？
            if (txtStaffID.Text == null)
            {
                MessageBox.Show("数字で入力してください", "配布員コード", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtStaffID.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", "配布員コード", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            //コード検証
            string sqlStr;
            Control.配布員 cStaff = new Control.配布員();
            OleDbDataReader dR;

            sqlStr = " where ID = " + txtStaffID.Text.ToString();
            dR = cStaff.FillBy(sqlStr);

            if (dR.HasRows == false)
            {
                MessageBox.Show("未登録コードです", "配布員コード", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                label13.Text = "";
                dR.Close();
                cStaff.Close();
                return;
            }
            else
            {
                while (dR.Read())
                {
                    label13.Text = dR["氏名"].ToString();
                }
            }

            dR.Close();
            cStaff.Close();
        }

        /// <summary>
        /// 配布日の天候を表示する
        /// </summary>
        private void tenkouUpdate()
        {

            //天候表示
            string sqlStr;
            OleDbDataReader dRt;
            Control.天候 cTenkou = new Control.天候();

            sqlStr = "";
            sqlStr += "where 日付 = '" + jDate.Value.ToShortDateString() + "'";

            dRt = cTenkou.FillBy(sqlStr);

            while (dRt.Read())
            {
                txtTenkou.Text = dRt["天候"].ToString();
            }

            dRt.Close();
            cTenkou.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //天候登録
            frmTenkou frm = new frmTenkou();
            frm.ShowDialog();

            //天候表示
            tenkouUpdate();
        }

        private void jDate_ValueChanged(object sender, EventArgs e)
        {
            //天候表示
            tenkouUpdate();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if ((txtStaffID.Text != null) && (txtStaffID.Text != "") && (txtStaffID.Text != "0"))
            {
                frmTesuuryouMeisai frm = new frmTesuuryouMeisai(1);

                frm.配布日 = jDate.Value;
                frm.配布員ID = int.Parse(txtStaffID.Text, System.Globalization.NumberStyles.Any);
                frm.配布員名 = label13.Text;

                frm.ShowDialog();
            }
            else
            {
                MessageBox.Show("配布員を入力してください",MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //ポスティングエリア登録
            frmPosting frm = new frmPosting();
            frm.ShowDialog();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            button4.Enabled = true;
        }

        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t' && e.KeyChar != '.')
                e.Handled = true;
        }

        private void dataGridView2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            string cName = dataGridView2.CurrentCell.OwningColumn.Name;
            if (cName == "col8" || cName == "col11")
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);

                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
            else
            {
                //イベントハンドラを削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void dataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentCellAddress.X == 12)
            {
                if (dataGridView2.IsCurrentCellDirty)
                {
                    dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 18)    // 編集画面表示
            {
                GridEnter_Haifu(e.RowIndex);
            }
            else if (e.ColumnIndex == 19)   // 配布エリア削除
            {
                areaDel(e.RowIndex);
            }
        }

        private void dataGridView2_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //dataGridView2.CurrentCell = null;
        }

        private void nDate_Enter(object sender, EventArgs e)
        {
            dataGridView2.CurrentCell = null;   // 2017/10/03
        }

        private void txtStaffID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
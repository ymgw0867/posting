using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace posting
{
    public partial class frmHaifuShijiSUb : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.配布エリア cMaster = new Entity.配布エリア();

        const string MESSAGE_CAPTION = "配布指示データ登録";
        const int FLG_ON = 1;
        const int FLG_OFF = 0;

        public frmHaifuShijiSUb()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //配布フラグON
            Utility.FlgOnOff(FLG_ON);

            //画面設定
            GridviewSet.Setting(dataGridView2);
            GridviewSet.ShowData(dataGridView2);
            dataGridView2.CurrentCell = null; //選択状態を回避する
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
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", (float)9.5, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 505;

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

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 100;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 230;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 349;

                    tempDGV.Columns[0].Visible = false;
                    tempDGV.Columns[9].Visible = false;
                    tempDGV.Columns[10].Visible = false;
                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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

            public static void ShowData(DataGridView tempDGV)
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
                    sqlSTRING += "受注.配布条件,配布形態.名称 as 配布形態 ";
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

            //配布エリアを一括登録する
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("明細を選択してください", "明細未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            } 

            if (MessageBox.Show("選択中のチラシ配布データを配布指示書へ追加登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            int iX;
            int iY = 0;

            iX = dataGridView2.SelectedRows.Count;

            _Count = dataGridView2.SelectedRows.Count;
            F_ID = new int[iX];
 
            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                F_ID[iY] = Int32.Parse(dataGridView2[0, r.Index].Value.ToString());
        
                iY++;
            }
        }

        //選択された配布エリアデータの行数
        private int _Count;
        public int Count
        {
            get 
            {
                return this._Count; 
            }
            set
            {
                this._Count = value;
            }
        }

        //選択された配布エリアデータIDのインデクサ
        private int[] F_ID;
        public int this[int iX]
        {
            set
            {
                this.F_ID[iX] = value;
            }
            get
            {
                return this.F_ID[iX];
            }
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
            if (MessageBox.Show("終了します。現在、選択状態のデータは登録されません。" + Environment.NewLine + "よろしいですか？", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            button4.DialogResult = DialogResult.No;
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

            string sqlStr,msgStr;

            //自分のループ
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                //配布期間が登録されているものを対象とする
                if ((dataGridView2[6, i].Value.ToString().Trim() != "") &&
                    (dataGridView2[7, i].Value.ToString().Trim() != ""))
                {

                    //ID
                    toID = Int32.Parse(dataGridView2[3, i].Value.ToString());

                    //配布開始日
                    sDate = Convert.ToDateTime(dataGridView2[6, i].Value.ToString());

                    //配布終了日
                    eDate = Convert.ToDateTime(dataGridView2[7, i].Value.ToString());

                    //相手のループ
                    for (int iX = i + 1; iX < dataGridView2.Rows.Count; iX++)
                    {
                        //配布期間が登録されているものを対象とする
                        if ((dataGridView2[6, iX].Value.ToString().Trim() != "") &&
                            (dataGridView2[7, iX].Value.ToString().Trim() != ""))
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
                    sqlStr = "";
                    sqlStr += "select 配布エリア.配布指示ID,配布エリア.町名ID,配布指示.配布日 ";
                    sqlStr += "from 配布エリア inner join 配布指示 ";
                    sqlStr += "on 配布エリア.配布指示ID = 配布指示.ID ";
                    sqlStr += "where ";
                    sqlStr += "(配布エリア.配布指示ID <> 0) and ";
                    sqlStr += "(配布エリア.ステータス = 2) and ";
                    sqlStr += "(配布エリア.完了区分 = 0) and ";
                    sqlStr += "(配布エリア.町名ID = " + toID.ToString() + ") ";
                    sqlStr += "order by 配布エリア.配布指示ID";

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlStr);

                    while(dR.Read())
                    {
                        //配布指示書の配布日
                        tsDate = Convert.ToDateTime(dR["配布日"].ToString());

                        //配布指示書の配布日が配布期間に該当するか
                        if ((tsDate >= sDate) && (tsDate <= eDate))
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

                    dR.Close();
                    fCon.Close();

                }
            }

            MessageBox.Show("終了しました", "併配表示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, "未指示データ");
        }

        private void frmHaifuShijiSUb_FormClosing(object sender, FormClosingEventArgs e)
        {
            //配布フラグOFF
            Utility.FlgOnOff(FLG_OFF);
        }
    }
}
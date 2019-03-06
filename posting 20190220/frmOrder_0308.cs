using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;

namespace posting
{
    public partial class frmOrder_0308 : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Utility.areaMode aMode = new Utility.areaMode();
        Utility.消費税率 cTax = new Utility.消費税率();
        Entity.受注 cMaster = new Entity.受注();

        const string MESSAGE_CAPTION = "受注確定書";

        public frmOrder_0308()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            // TODO: このコード行はデータを 'darwinDataSet.受注' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);
            this.darwinDataSet.Clear();
            this.darwinDataSet.EnforceConstraints = false;
            this.受注TableAdapter.Fill(this.darwinDataSet.受注);

            //ポスティングエリア
            GridviewSet.AriaSetting(dataGridView2);

            //町名マスタ
            GridviewSet.TownSetting(dataGridView3);

            //受注区分コンボ
            cmbJkbnSet();

            //事業所コンボ
            Utility.ComboOffice.load(cmbOffice);

            //得意先コンボ
            Utility.ComboClient.load(cmbClient);

            //受注種別コンボ
            Utility.ComboJshubetsu.load(cmbNaiyou);

            //配布形態コンボ
            Utility.ComboFkeitai.load(cmbFkeitai);

            //配布条件コンボ
            cmbFjyoukenSet();

            //判型コンボ
            Utility.ComboSize.load(cmbSize);

            //配布条件コンボ
            cmbFyuyoSet();

            //納品形態コンボ
            cmbNkeitaiSet();

            //入金方法コンボ
            Utility.ComboShimebi.load(cmbNyukin);

            //振込口座コンボ
            Utility.ComboFuri.load(cmbFuri);

            //報告時期コンボ
            cmbHjikiSet();

            //報告精度コンボ
            cmbHseidoSet();

            //報告方法コンボ
            cmbHhouhouSet();

            //税率取得
            cTax.Ritsu = GetTaxRT(DateTime.Today);

            DispClear();

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
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 16;
                    tempDGV.RowTemplate.Height = 16;

                    // 全体の高さ
                    tempDGV.Height = 164;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    //tempDGV.Columns.Add("col1", "ｺｰﾄﾞ");
                    //tempDGV.Columns.Add("col2", "名称");
                    //tempDGV.Columns.Add("col3", "備考");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 90;
                    tempDGV.Columns[11].Width = 90;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void AriaSetting(DataGridView tempDGV)
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
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 16;
                    tempDGV.RowTemplate.Height = 16;

                    // 全体の高さ
                    tempDGV.Height = 230;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "エリアID");
                    tempDGV.Columns.Add("col2", "配布エリア");
                    tempDGV.Columns.Add("col3", "配布枚数");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 80;

                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // 編集制御
                    tempDGV.Columns[0].ReadOnly = true;
                    tempDGV.Columns[1].ReadOnly = true;
                    tempDGV.Columns[2].ReadOnly = true;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void TownSetting(DataGridView tempDGV)
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
                    tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 16;
                    tempDGV.RowTemplate.Height = 16;

                    // 全体の高さ
                    tempDGV.Height = 140;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns.Add("col1", "エリアID");
                    tempDGV.Columns.Add("col2", "町名");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 200;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// データグリッドビューの指定行のデータを取得する
            /// </summary>
            /// <param name="dgv">対象とするデータグリッドビューオブジェクト</param>
            public static Boolean GetData(DataGridView dgv,ref Entity.受注 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.受注 Order = new Control.受注();
                OleDbDataReader dr;

                sqlStr = " where 受注.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Order.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.事業所ID = Int32.Parse(dr["事業所ID"].ToString());
                        tempC.受注日 = Convert.ToDateTime(dr["受注日"].ToString());
                        tempC.受注区分 = dr["受注区分"].ToString() + "";
                        tempC.得意先ID = Int32.Parse(dr["得意先ID"].ToString());

                        //tempC.社員ID = Int32.Parse(dr["社員ID"].ToString());

                        tempC.チラシ名 = dr["チラシ名"].ToString() + "";
                        tempC.受注種別ID = Int32.Parse(dr["受注種別ID"].ToString());
                        tempC.単価 = Convert.ToDouble(dr["単価"].ToString());
                        tempC.枚数 = Int32.Parse(dr["枚数"].ToString());
                        tempC.金額 = Int32.Parse(dr["金額"].ToString(),System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.消費税 = Int32.Parse(dr["消費税"].ToString(),System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.税込金額 = Int32.Parse(dr["税込金額"].ToString(),System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.値引額 = Int32.Parse(dr["値引額"].ToString(), System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.売上金額 = Int32.Parse(dr["売上金額"].ToString(), System.Globalization.NumberStyles.AllowDecimalPoint);                        
                        tempC.税率 = Int32.Parse(dr["税率"].ToString());
                        tempC.判型 = Int32.Parse(dr["判型"].ToString());
                        tempC.依頼先 = dr["依頼先"].ToString();
                        tempC.原価 = Convert.ToDouble(dr["原価"].ToString());
                        tempC.配布形態 = Int32.Parse(dr["配布形態"].ToString());
                        tempC.配布条件 = dr["配布条件"].ToString() + "";
                        tempC.配布開始日 = dr["配布開始日"].ToString();
                        tempC.配布終了日 = dr["配布終了日"].ToString();
                        tempC.配布猶予 = dr["配布猶予"].ToString() + "";
                        tempC.納品予定日 = dr["納品予定日"].ToString();
                        tempC.納品形態 = dr["納品形態"].ToString() + "";
                        tempC.請求書 = Int32.Parse(dr["請求書"].ToString());
                        tempC.入金方法 = dr["入金方法"].ToString() + "";
                        tempC.入金予定日 = dr["入金予定日"].ToString();
                        tempC.報告時期 = dr["報告時期"].ToString() + "";
                        tempC.報告精度 = dr["報告精度"].ToString() + "";
                        tempC.報告方法 = dr["報告方法"].ToString() + "";
                        tempC.メールアドレス = dr["メールアドレス"].ToString() + "";
                        tempC.振込口座ID = Int32.Parse(dr["振込口座ID"].ToString());
                        tempC.未配布情報有無 = Int32.Parse(dr["未配布情報有無"].ToString());
                        tempC.枝番有無 = Int32.Parse(dr["枝番有無"].ToString());
                        tempC.特記事項 = dr["特記事項"].ToString() + "";
                        tempC.エリア備考 = dr["エリア備考"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    Order.Close();
                    return false;
                }

                dr.Close();
                Order.Close();
                return true;
            }

            //public static void ShowData(DataGridView tempDGV)
            //{
            //    string sqlSTRING = "";

            //    try
            //    {
            //        tempDGV.RowCount = 0;

            //        //受注データのデータリーダーを取得する
            //        Control.FreeSql cOrder = new Control.FreeSql();
            //        cOrder.Execute("");
            //        Control.DataControl dCon = new Control.DataControl();

            //        sqlSTRING = "select * from m_Costname " +
            //                    "order by ID";

            //        dR = dCon.FreeReader(sqlSTRING);

            //        iX = 0;

            //        while (dR.Read())
            //        {
            //            tempDGV.Rows.Add();

            //            tempDGV[0, iX].Value = dR["ID"];
            //            tempDGV[1, iX].Value = NullConvert.Noth(dR["原価名"]);
            //            tempDGV[2, iX].Value = NullConvert.Noth(dR["備考"]);
            //            //tempDGV[1, iX].Value = dR["原価名"];
            //            //tempDGV[2, iX].Value = dR["備考"];
            //            iX++;
            //        }

            //        dR.Close();

            //        dCon.Close();

            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //    }

            //}

        }

        //グリッドからデータを選択
        private void GridEnter()
        {

            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[2, dataGridView1.SelectedRows[iX].Index].Value + " " + dataGridView1[3, dataGridView1.SelectedRows[iX].Index].Value +"が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "受注確定書選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {

                    //データを取得する
                    if (GridviewSet.GetData(dataGridView1,ref cMaster) == false)
                    {
                        MessageBox.Show("該当するデータが登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'データ値を取得
                    jDate.Value = Convert.ToDateTime(cMaster.受注日.ToString());
                    comboBox1.Text = cMaster.受注区分;

                    Utility.ComboOffice.selectedIndex(cmbOffice, Int32.Parse(cMaster.事業所ID.ToString()));
                    
                    //クライアント情報表示
                    Utility.ComboClient.selectedIndex(cmbClient, Int32.Parse(cMaster.得意先ID.ToString()));

                    txtCzipcode.Text = "";
                    txtName2.Text = "";
                    txtCbusho.Text = "";
                    txtTantou.Text = "";
                    txtCtel.Text = "";
                    txtCfax.Text  = "";
                    txtCbusho.Text = "";
                    txtCtantou.Text = "";

                    //クライアント情報表示
                    ClientShow(cMaster.得意先ID);

                    txtChirashi.Text = cMaster.チラシ名;
                    Utility.ComboJshubetsu.selectedIndex(cmbNaiyou, Int32.Parse(cMaster.受注種別ID.ToString()));

                    txtTanka.Text = cMaster.単価.ToString("#,##0.0");
                    txtMai.Text  = cMaster.枚数.ToString("#,##0");
                    txtUri.Text = cMaster.金額.ToString("#,##0");
                    txtTax.Text = cMaster.消費税.ToString("#,##0");
                    txtZeikomi.Text = cMaster.税込金額.ToString("#,##0");
                    txtNebiki.Text = cMaster.値引額.ToString("#,##0");
                    txtUriTL.Text = cMaster.売上金額.ToString("#,##0");

                    Utility.ComboFkeitai.selectedIndex(cmbFkeitai, Int32.Parse(cMaster.配布形態.ToString()));
                    cmbFjyouken.Text = cMaster.配布条件;
                    Utility.ComboSize.selectedIndex(cmbSize, Int32.Parse(cMaster.判型.ToString()));

                    txtIraisaki.Text = cMaster.依頼先;
                    txtGenka.Text = cMaster.原価.ToString("#,##0.0");

                    if (cMaster.配布開始日.ToString() == "")
                    {
                        StartDate.Checked = false;
                    }
                    else
                    {
                        StartDate.Value = Convert.ToDateTime(cMaster.配布開始日.ToString());
                    }

                    if (cMaster.配布終了日.ToString() == "")
                    {
                        EndDate.Checked = false;
                    }
                    else
                    {
                        EndDate.Value = Convert.ToDateTime(cMaster.配布終了日.ToString());
                    }

                    cmbFyuyo.Text = cMaster.配布猶予;

                    if (Int32.Parse(cMaster.請求書.ToString()) == 1)
                    {
                        checkBox1.Checked = true;
                    }
                    else
                    {
                        checkBox1.Checked = false;
                    }

                    if (cMaster.納品予定日.ToString() == "")
                    {
                        NouhinDate.Checked = false;
                    }
                    else
                    {
                        NouhinDate.Value = Convert.ToDateTime(cMaster.納品予定日.ToString());
                    }

                    cmbNkeitai.Text = cMaster.納品形態;
                    cmbNyukin.Text = cMaster.入金方法;
                    
                    if (cMaster.入金予定日.ToString() == "")
                    {
                        NyukinDate.Checked = false;
                    }
                    else
                    {
                        NyukinDate.Value = Convert.ToDateTime(cMaster.入金予定日.ToString());
                    }

                    Utility.ComboFuri.selectedIndex(cmbFuri, cMaster.振込口座ID);
                    cmbHjiki.Text = cMaster.報告時期;
                    cmbHseido.Text = cMaster.報告精度;
                    cmbHhouhou.Text = cMaster.報告方法;

                    txtEmail.Text = cMaster.メールアドレス;
                    txtMemo.Text = cMaster.特記事項;
                    txtMemo2.Text = cMaster.エリア備考;

                    if (Int32.Parse(cMaster.未配布情報有無.ToString()) == 1)
                    {
                        checkBox2.Checked = true;
                    }
                    else
                    {
                        checkBox2.Checked = false;
                    }

                    if (Int32.Parse(cMaster.枝番有無.ToString()) == 1)
                    {
                        checkBox3.Checked = true;
                    }
                    else
                    {
                        checkBox3.Checked = false;
                    }

                    //IDテキストボックスは編集不可とする
                    //txtCode.Enabled = false;

                    //ボタン状態
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //フォームモードステータス:変更削除

                    jDate.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "データ表示", MessageBoxButtons.OK);
                }
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            // Enterkey以外は対象外
            if (e.KeyCode.ToString() != "Return")
            {
                return;
            }

            GridEnter();
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
                fMode.Mode = 0;
                aMode.Mode = 0;

                jDate.Value = DateTime.Today;
                comboBox1.SelectedIndex = -1;
                cmbOffice.SelectedIndex = -1;
                cmbClient.SelectedIndex = -1;
                txtCzipcode.Text = "";
                txtName2.Text = "";
                txtCbusho.Text = "";
                txtTantou.Text = "";
                txtCtel.Text = "";
                txtCfax.Text = "";
                txtCbusho.Text = "";
                txtCtantou.Text = "";

                txtChirashi.Text = "";
                cmbNaiyou.SelectedIndex=-1;
                txtTanka.Text = "0";
                txtMai.Text = "0";
                txtUri.Text = "0";
                txtTax.Text = "0";
                txtZeikomi.Text = "0";
                txtNebiki.Text = "0";
                txtUriTL.Text = "0";

                cmbFkeitai.SelectedIndex=-1;
                cmbFjyouken.SelectedIndex = -1;
                cmbSize.SelectedIndex = -1;

                txtIraisaki.Text = "";
                txtGenka.Text = "0";

                StartDate.Checked = false;
                EndDate.Checked = false;
                cmbFyuyo.SelectedIndex = -1;
                checkBox1.Checked = false;
                NouhinDate.Checked = false;
                cmbNkeitai.SelectedIndex = -1;
                cmbNyukin.SelectedIndex = -1;
                NyukinDate.Checked = false;
                cmbFuri.SelectedIndex = -1;
                cmbHjiki.SelectedIndex = -1;
                cmbHseido.SelectedIndex = -1;
                cmbHhouhou.SelectedIndex = -1;

                txtEmail.Text = "";
                txtMemo.Text = "";
                txtMemo2.Text = "";
                checkBox2.Checked = false;
                checkBox3.Checked = false;

                btnDel.Enabled = false;
                btnClr.Enabled = false;

                txtAreaID.Text = "";
                txtAreaName.Text = "";
                txtHaihuMaisu.Text = "";

                dataGridView2.Rows.Clear();
                dataGridView3.Rows.Clear();

                textBox5.Text = "";
                txtAdel.Enabled = false;

                //txtCode.Focus();
                jDate.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("選択されているデータを破棄します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question)== DialogResult.No )
                return;
                
            DispClear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {

                if (fDataCheck() == true)
                {
                    Control.受注 Order = new Control.受注();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                break;

                            if (Order.DataInsert(cMaster) == true)
                            {
                                MessageBox.Show("新規登録されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("新規登録に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                break;

                            if (Order.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Order.Close();

                    DispClear();

                    //データを 'darwinDataSet.受注' テーブルに読み込みます。
                    this.受注TableAdapter.Fill(this.darwinDataSet.受注);
                    dataGridView1.DataSource = this.darwinDataSet.受注;

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

            string str;
            double d;

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
                //受注区分チェック
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("受注区分を選択してください");
                }

                //事業所ＩＤチェック
                if (cmbOffice.SelectedIndex == -1)
                {
                    cmbOffice.Focus();
                    throw new Exception("事業所を選択してください");
                }

                //クライアントチェック
                if (cmbClient.SelectedIndex == -1)
                {
                    cmbClient.Focus();
                    throw new Exception("クライアントを選択してください");
                }

                //チラシ名チェック
                if (txtChirashi.Text.Trim().Length < 1)
                {
                    txtChirashi.Focus();
                    throw new Exception("チラシ名を入力してください");
                }

                //受注内容チェック
                if (cmbNaiyou.SelectedIndex == -1)
                {
                    cmbNaiyou.Focus();
                    throw new Exception("受注内容を選択してください");
                }

                //単価：数字か？
                if (txtTanka.Text == null)
                {
                    this.txtTanka.Focus();
                    throw new Exception("単価は数字で入力してください");
                }

                str = this.txtTanka.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtTanka.Focus();
                    throw new Exception("単価は数字で入力してください");
                }

                //枚数：数字か？
                if (txtMai.Text == null)
                {
                    this.txtMai.Focus();
                    throw new Exception("枚数は数字で入力してください");
                }

                str = this.txtMai.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtMai.Focus();
                    throw new Exception("枚数は数字で入力してください");
                }

                //配布形態チェック
                if (cmbFkeitai.SelectedIndex == -1)
                {
                    cmbFkeitai.Focus();
                    throw new Exception("配布形態を選択してください");
                }

                //判型チェック
                if (cmbSize.SelectedIndex == -1)
                {
                    cmbSize.Focus();
                    throw new Exception("サイズを選択してください");
                }

                //原価：数字か？
                if (txtGenka.Text == null)
                {
                    this.txtGenka.Focus();
                    throw new Exception("原価は数字で入力してください");
                }

                str = this.txtGenka.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtGenka.Focus();
                    throw new Exception("原価は数字で入力してください");
                }

                if (cmbNaiyou.Text == "ポスティング")
                {
                    //配布条件チェック
                    if (cmbFjyouken.SelectedIndex == -1)
                    {
                        cmbFjyouken.Focus();
                        throw new Exception("配布条件を選択してください");
                    }

                    //配布期間
                    if (StartDate.Value > EndDate.Value)
                    {
                        StartDate.Focus();
                        throw new Exception("配布期間が正しくありません");
                    }

                    //配布猶予チェック
                    if (cmbFyuyo.SelectedIndex == -1)
                    {
                        cmbFyuyo.Focus();
                        throw new Exception("配布猶予を選択してください");
                    }

                    //報告時期チェック
                    if (cmbHjiki.SelectedIndex == -1)
                    {
                        cmbHjiki.Focus();
                        throw new Exception("報告時期を選択してください");
                    }

                    //報告精度チェック
                    if (cmbHseido.SelectedIndex == -1)
                    {
                        cmbHseido.Focus();
                        throw new Exception("報告精度を選択してください");
                    }
                    //報告方法チェック
                    if (cmbHhouhou.SelectedIndex == -1)
                    {
                        cmbHhouhou.Focus();
                        throw new Exception("報告方法を選択してください");
                    }
                }

                //クラスにデータセット
                Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                cmb1 = (Utility.ComboOffice)cmbOffice.SelectedItem;
                cMaster.事業所ID = cmb1.ID;

                cMaster.受注日 = jDate.Value;
                cMaster.受注区分 = comboBox1.Text;

                Utility.ComboClient cmb2 = new Utility.ComboClient();
                cmb2 = (Utility.ComboClient)cmbClient.SelectedItem;
                cMaster.得意先ID = cmb2.ID;

                cMaster.チラシ名 = txtChirashi.Text.ToString();

                Utility.ComboJshubetsu cmb3 = new Utility.ComboJshubetsu();
                cmb3 = (Utility.ComboJshubetsu)cmbNaiyou.SelectedItem;
                cMaster.受注種別ID = cmb3.ID;

                cMaster.単価 = Convert.ToDouble(txtTanka.Text.ToString());
                cMaster.枚数 = Int32.Parse(txtMai.Text.ToString(),System.Globalization.NumberStyles.AllowThousands);
                cMaster.金額 = Int32.Parse(txtUri.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.消費税 = Int32.Parse(txtTax.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.税込金額 = Int32.Parse(txtZeikomi.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.値引額 = Int32.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.売上金額 = Int32.Parse(txtUriTL.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.税率 = cTax.Ritsu;

                Utility.ComboSize cmb4 = new Utility.ComboSize();
                cmb4 = (Utility.ComboSize)cmbSize.SelectedItem;
                cMaster.判型 = cmb4.ID;

                cMaster.依頼先 = txtIraisaki.Text.ToString();
                cMaster.原価 = Convert.ToDouble(txtGenka.Text.ToString());

                Utility.ComboFkeitai cmb5 = new Utility.ComboFkeitai();
                cmb5 = (Utility.ComboFkeitai)cmbFkeitai.SelectedItem;
                cMaster.配布形態 = cmb5.ID;

                cMaster.配布条件 = cmbFjyouken.Text.ToString();

                if (StartDate.Checked == true)
                {
                    cMaster.配布開始日 = StartDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.配布開始日 = "";
                }

                if (EndDate.Checked == true)
                {
                    cMaster.配布終了日 = EndDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.配布終了日 = "";
                }

                cMaster.配布猶予 = cmbFyuyo.Text.ToString();

                if (NouhinDate.Checked == true)
                {
                    cMaster.納品予定日 = NouhinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.納品予定日 = "";
                }

                cMaster.納品形態 = cmbNkeitai.Text.ToString();

                if (checkBox1.Checked == true)
                {
                    cMaster.請求書 = 1;
                }
                else
                {
                    cMaster.請求書 = 0;
                }

                if (cmbNyukin.SelectedIndex == -1)
                {
                    cMaster.入金方法 = "";
                }
                else
                {
                    Utility.ComboShimebi cmb6 = new Utility.ComboShimebi();
                    cmb6 = (Utility.ComboShimebi)cmbNyukin.SelectedItem;
                    cMaster.入金方法 = cmb6.Name;
                }

                if (NyukinDate.Checked == true)
                {
                    cMaster.入金予定日 = NyukinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.入金予定日 = "";
                }

                cMaster.報告時期 = cmbHjiki.Text;
                cMaster.報告精度 = cmbHseido.Text;
                cMaster.報告方法 = cmbHhouhou.Text;
                cMaster.メールアドレス = txtEmail.Text;

                if (cmbFuri.SelectedIndex == -1)
                {
                    cMaster.振込口座ID = 0;
                }
                else
                {
                    Utility.ComboFuri cmb7 = new Utility.ComboFuri();
                    cmb7 = (Utility.ComboFuri)cmbFuri.SelectedItem;
                    cMaster.振込口座ID = cmb7.ID;
                }

                if (checkBox2.Checked == true)
                {
                    cMaster.未配布情報有無 = 1;
                }
                else
                {
                    cMaster.未配布情報有無 = 0;
                }

                if (checkBox3.Checked == true)
                {
                    cMaster.枝番有無 = 1;
                }
                else
                {
                    cMaster.枝番有無 = 0;
                }

                cMaster.特記事項 = txtMemo.Text.ToString();
                cMaster.エリア備考 = txtMemo2.Text.ToString();

                if (fMode.Mode == 0) cMaster.登録年月日 = DateTime.Today;
                cMaster.変更年月日 = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void txtEnter(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtChirashi)
            {
                objtxt = txtChirashi;
            }

            if (sender == txtTanka)
            {
                objtxt = txtTanka;
            }

            if (sender == txtMai)
            {
                objtxt = txtMai;
            }

            if (sender == txtNebiki)
            {
                objtxt = txtNebiki;
            }

            if (sender == txtIraisaki)
            {
                objtxt = txtIraisaki;
            }

            if (sender == txtGenka)
            {
                objtxt = txtGenka;
            }

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == txtMemo2)
            {
                objtxt = txtMemo2;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == cmbClient)
            {
                cmbClient.BackColor = Color.LightGray;
            }

            if (sender == txtAreaID)
            {
                objtxt = txtAreaID;
            }

            if (sender == txtHaihuMaisu)
            {
                objtxt = txtHaihuMaisu;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {

           TextBox objtxt = new TextBox();

           double Kingaku;
           int KingakuTax;
           int KingakuZeikomi;
           int KingakuTL;
           string str;
           double d;

           try
           {


               if (sender == txtChirashi)
               {
                   objtxt = txtChirashi;
               }

               if (sender == txtTanka)
               {
                   objtxt = txtTanka;

                   if (txtTanka.Text == null) txtTanka.Text = "0";

                   str = txtTanka.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtTanka.Text = "0";

                   UriageSum(Convert.ToDateTime(jDate.Value), Convert.ToDouble(txtTanka.Text),
                             int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             int.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             out Kingaku, out KingakuTax, out KingakuZeikomi,out KingakuTL);


                   //売上金額
                   txtUri.Text = Kingaku.ToString("#,##0");

                   //消費税額計算               
                   txtTax.Text = KingakuTax.ToString("#,##0");

                   //税込金額
                   txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                   //売上金額
                   txtUriTL.Text  = KingakuTL.ToString("#,##0");
               }

               if (sender == txtMai)
               {
                   objtxt = txtMai;

                   if (txtMai.Text == null) txtMai.Text = "0";

                   str = txtMai.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtMai.Text = "0";

                   UriageSum(Convert.ToDateTime(jDate.Value), Convert.ToDouble(txtTanka.Text),
                             int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             int.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             out Kingaku, out KingakuTax, out KingakuZeikomi,out KingakuTL);

                   //売上金額
                   txtUri.Text = Kingaku.ToString("#,##0");

                   //消費税額計算               
                   txtTax.Text = KingakuTax.ToString("#,##0");

                   //税込金額
                   txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                   //売上金額
                   txtUriTL.Text = KingakuTL.ToString("#,##0");
               }

               if (sender == txtNebiki)
               {
                   objtxt = txtNebiki;

                   if (txtNebiki.Text == null) txtNebiki.Text = "0";

                   str = txtNebiki.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtNebiki.Text = "0";

                   UriageSum(Convert.ToDateTime(jDate.Value), Convert.ToDouble(txtTanka.Text),
                             int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             int.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             out Kingaku, out KingakuTax, out KingakuZeikomi, out KingakuTL);

                   //売上金額
                   txtUri.Text = Kingaku.ToString("#,##0");

                   //消費税額計算               
                   txtTax.Text = KingakuTax.ToString("#,##0");

                   //税込金額
                   txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                   //売上金額
                   txtUriTL.Text = KingakuTL.ToString("#,##0");

               }

               if (sender == txtIraisaki)
               {
                   objtxt = txtIraisaki;
               }

               if (sender == txtGenka)
               {
                   objtxt = txtGenka;
               }

               if (sender == txtEmail)
               {
                   objtxt = txtEmail;
               }

               if (sender == txtMemo)
               {
                   objtxt = txtMemo;
               }

               if (sender == txtMemo2)
               {
                   objtxt = txtMemo2;
               }

               if (sender == textBox1)
               {
                   objtxt = textBox1;
               }

               if (sender == cmbClient)
               {
                   cmbClient.BackColor = Color.White;

                   //クライアント情報表示
                   Utility.ComboClient cmbC = new Utility.ComboClient();
                   cmbC = (Utility.ComboClient)cmbClient.SelectedItem;
                   ClientShow(cmbC.ID);

               }

               //配布エリアID
               if (sender == txtAreaID)
               {
                   objtxt = txtAreaID;

                   if (txtAreaID.Text == null) txtAreaID.Text = "0";

                   str = txtAreaID.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtAreaID.Text = "0";

                   txtAreaName.Text = GetTownName(txtAreaID.Text.ToString());
               }

               if (sender == txtHaihuMaisu)
               {
                   objtxt = txtHaihuMaisu;
               }

               objtxt.BackColor = Color.White;
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message,"エラーメッセージ");
           }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.受注 Order = new Control.受注();
            if (Order.DataDelete(cMaster.ID) == true)
            {
                MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("受注データの削除に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            Order.Close();

            //配布エリアデータ削除
            string strSql;
            strSql = "";
            strSql += "delete from 配布エリア ";
            strSql += "where 配布エリア.受注ID = " + cMaster.ID.ToString();

            Control.FreeSql fsql = new Control.FreeSql();
            if (fsql.Execute(strSql) == true)
            {
                MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("配布エリアデータの削除に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            fsql.Close();

            DispClear();

            //データを 'darwinDataSet.受注' テーブルに読み込みます。
            this.受注TableAdapter.Fill(this.darwinDataSet.受注);
            dataGridView1.DataSource = this.darwinDataSet.受注;

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

            GridEnter();
        }

        private void cmbJkbnSet()
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("新規");
            comboBox1.Items.Add("リピート");
        }

        private void cmbFjyoukenSet()
        {
            cmbFjyouken.Items.Clear();
            cmbFjyouken.Items.Add("予定枚数どおり");
            cmbFjyouken.Items.Add("詰めて配布 ＯＫ");
        }

        private void cmbFyuyoSet()
        {
            cmbFyuyo.Items.Clear();
            cmbFyuyo.Items.Add("厳守");
            cmbFyuyo.Items.Add("前後ＯＫ");
        }

        private void cmbNkeitaiSet()
        {
            cmbNkeitai.Items.Clear();
            cmbNkeitai.Items.Add("宅急便");
            cmbNkeitai.Items.Add("持込");
            cmbNkeitai.Items.Add("集荷：バック");
            cmbNkeitai.Items.Add("集荷：営業");
        }

        private void cmbHjikiSet()
        {
            cmbHjiki.Items.Clear();
            cmbHjiki.Items.Add("デイリー");
            cmbHjiki.Items.Add("週単位");
            cmbHjiki.Items.Add("終了後");
        }

        private void cmbHseidoSet()
        {
            cmbHseido.Items.Clear();
            cmbHseido.Items.Add("実配布数");
            cmbHseido.Items.Add("予定枚数");
        }

        private void cmbHhouhouSet()
        {
            cmbHhouhou.Items.Clear();
            cmbHhouhou.Items.Add("FAX");
            cmbHhouhou.Items.Add("メール");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            darwinDataSet ds = new darwinDataSet();
            ds.Clear();
            ds.EnforceConstraints = false;
            this.受注TableAdapter.FillByName(ds.受注, "%" + textBox1.Text.ToString() + "%");
            dataGridView1.DataSource = ds.受注;
        }

        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmOffice();

            frm.ShowDialog();
            Utility.ComboOffice.load(cmbOffice);

        }

        /// <summary>
        /// 消費税率取得
        /// </summary>
        /// <param name="tempDate">基準日付</param>
        /// <returns>税率</returns>
        private int GetTaxRT(DateTime tempDate)
        {
            //税率取得
            string Sqlstr;
            int Ritsu = 0;

            OleDbDataReader drTax;
            Control.税率 cR = new Control.税率();
            Sqlstr = "where 税率.開始年月日 <= '" + jDate.Value.ToShortDateString() + "' order by 開始年月日";
            drTax = cR.FillBy(Sqlstr);

            while (drTax.Read())
            {
                Ritsu = Int32.Parse(drTax["税率"].ToString());
            }

            drTax.Close();

            return Ritsu;
            
        }

        /// <summary>
        /// 消費税計算
        /// </summary>
        /// <param name="tempKin">対象金額</param>
        /// <param name="tempTax">税率</param>
        /// <returns>消費税額</returns>
        private int GetTax(double tempKin,int tempTax)
        {
            double GakuD;
            int GakuI;

            GakuD = tempKin * tempTax / 100;
            GakuD += 0.5;
            GakuI = (int)GakuD;

            return GakuI;
        }

        private void UriageSum(DateTime tempDate,double tempTanka,int tempMai,int tempNebiki,
                               out double Kingaku,out int KingakuTax,out int KingakuZeikomi,
                               out int KingakuTL)
        {

            //売上金額
            Kingaku = tempTanka * tempMai;

            //税率再取得
            cTax.Ritsu = GetTaxRT(tempDate);

            //消費税額計算               
            KingakuTax = GetTax(Kingaku, cTax.Ritsu);

            //税込金額
            KingakuZeikomi = (int)Kingaku + KingakuTax;

            //売上金額
            KingakuTL = KingakuZeikomi - tempNebiki;
        }

        private void ClientShow(int tempID)
        {
            OleDbDataReader drt;
            Control.得意先 Client = new Control.得意先();
            drt = Client.FillBy("where ID = " + tempID.ToString());

            while (drt.Read())
            {
                txtCzipcode.Text = drt["郵便番号"].ToString();
                txtName2.Text = drt["住所1"].ToString() + "" + " " + drt["住所2"].ToString() + "";
                txtCbusho.Text = drt["部署名"].ToString() + "";
                txtCtantou.Text = drt["担当者名"].ToString() + "";
                txtCtel.Text = drt["電話番号"].ToString() + "";
                txtCfax.Text = drt["FAX番号"].ToString() + "";

                OleDbDataReader drS;
                Control.社員 Shain = new Control.社員();
                drS = Shain.FillBy("where ID = " + drt["担当社員コード"].ToString());

                while (drS.Read())
                {
                    txtTantou.Text = drS["氏名"].ToString();
                }

                drS.Close();

            }

            drt.Close();

        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void cmbClient_Click(object sender, EventArgs e)
        {

        }

        private void 受注BindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void txtZeikomi_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void cmbFkeitai_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtAdd_Click(object sender, EventArgs e)
        {
            //配布エリア登録

            string str;
            double d;
            int iX;

            try
            {
                if (txtHaihuMaisu.Text == null)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("配布枚数は数字で入力してください");
                }

                str = txtHaihuMaisu.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("配布枚数は数字で入力してください");
                }

                if (Int32.Parse(txtHaihuMaisu.Text.ToString()) < 0)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("配布枚数が正しくありません");
                }

                switch (aMode.Mode)
                {
                    case 0: //登録
                        dataGridView2.Rows.Add();
                        iX = dataGridView2.Rows.Count - 1;
                        dataGridView2[0, iX].Value = Int32.Parse(txtAreaID.Text.ToString());
                        //dataGridView2[1, iX].Value = txtAreaName.Text;
                        dataGridView2[2, iX].Value = Int32.Parse(txtHaihuMaisu.Text.ToString());
                        break;

                    case 1: //更新

                        break;
                }

                textAreaClear();
                txtAreaID.Focus();

                txtTotal.Text = GetMaisuTotal().ToString("#,##0");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void textAreaClear()
        {
            aMode.Mode = 0;

            txtAreaID.Text = "";
            txtAreaName.Text = "";
            txtHaihuMaisu.Text = "";
            txtAdel.Enabled = false;

            //ソート使用可
            dataGridView2.Columns[0].SortMode = DataGridViewColumnSortMode.Automatic;
            dataGridView2.Columns[1].SortMode = DataGridViewColumnSortMode.Automatic;
            dataGridView2.Columns[2].SortMode = DataGridViewColumnSortMode.Automatic;
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex != 0) return;

            dataGridView2[1, e.RowIndex].Value = GetTownName(dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString());
        }

        private string GetTownName(string tempID)
        {
            //配布エリア町名検索
            string strName = ""; 
            OleDbDataReader dr;

            Control.町名 cTown = new Control.町名();
            dr = cTown.FillBy("where ID = " + tempID);

            while (dr.Read())
            {
                strName = dr["名称"].ToString();
            }

            dr.Close();
            return strName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                //配布エリア町名検索
                OleDbDataReader dr;
                int iX = 0;

                Control.町名 cTown = new Control.町名();
                dr = cTown.FillBy("where 名称 like '%" + textBox5.Text.ToString() + "%' order by ID");
                dataGridView3.Rows.Clear();

                while (dr.Read())
                {
                    dataGridView3.Rows.Add();
                    dataGridView3[0, iX].Value = dr["ID"];
                    dataGridView3[1, iX].Value = dr["名称"];
                    iX++;
                }

                dr.Close();
                cTown.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        //配布エリアの町名を一括登録する
        {
            int iX = 0;
            foreach (DataGridViewRow r in dataGridView3.SelectedRows)
            {
                dataGridView2.Rows.Add();
                iX = dataGridView2.Rows.Count - 1;
                dataGridView2[0, iX].Value = Int32.Parse(dataGridView3[0, r.Index].Value.ToString());
                dataGridView2[2, iX].Value = 0;
                iX++;
            }

        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int iX = 0;

            txtAreaID.Text = dataGridView2[0, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtAreaName.Text = dataGridView2[1, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtHaihuMaisu.Text  = dataGridView2[2, dataGridView2.SelectedRows[iX].Index].Value.ToString();

            //編集中はソート禁止
            dataGridView2.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;

            aMode.Mode = 1;
            aMode.RowIndex = dataGridView2.SelectedRows[iX].Index;
            txtAdel.Enabled = true;
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            txtAreaID.Text = dataGridView3[0, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtAreaName.Text = dataGridView3[1, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtHaihuMaisu.Focus();
            aMode.Mode = 0;
        }

        private int GetMaisuTotal()
        {
            int Total = 0;

            for (int i = 0; i <= dataGridView2.Rows.Count - 1; i++)
            {
                Total += Int32.Parse(dataGridView2[2, i].Value.ToString(), System.Globalization.NumberStyles.AllowThousands);
            }

            return Total;
        }

        private void txtAdel_Click(object sender, EventArgs e)
        {
            //行削除
            if (MessageBox.Show(dataGridView2[1,aMode.RowIndex].Value.ToString() + " を削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            dataGridView2.Rows.RemoveAt(aMode.RowIndex);
        }

        private void txtAclear_Click(object sender, EventArgs e)
        {
            textAreaClear();
        }

    }
}
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
    public partial class frmOrder: Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Utility.消費税率 cTax = new Utility.消費税率();
        Entity.受注 cMaster = new Entity.受注();

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.新請求書TableAdapter rAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.配布エリアTableAdapter hAdp = new darwinDataSetTableAdapters.配布エリアTableAdapter();

        // 2016/07/07
        darwinDataSetTableAdapters.受注編集制限TableAdapter lAdp = new darwinDataSetTableAdapters.受注編集制限TableAdapter();

        const string MESSAGE_CAPTION = "受注確定書";
        const string J_NAIYOU_POSTING = "ポスティング";

        Int64 _orderNum = 0;

        int sLoginTag = 0;                    // 受注確定書編集設定@編集可能ログインタグ
        DateTime dtLock = DateTime.Today;     // 請求書発行日期限
        bool orderEditStatus = true;          // 受注確定書編集設定

        Utility.comboGaichuUser[] cg = null;

        public frmOrder(Int64 orderNum)
        {
            InitializeComponent();

            //// 受注データ読み込み
            //jAdp.Fill(dts.受注1);

            // 得意先データ読み込み
            cAdp.Fill(dts.得意先);

            // 社員データ読み込み
            sAdp.Fill(dts.社員);

            // 2016/07/07
            lAdp.Fill(dts.受注編集制限);

            // 2016/07/19
            rAdp.Fill(dts.新請求書);

            //// 2017/01/27
            //hAdp.Fill(dts.配布エリア);

            // 指定受注番号
            _orderNum = orderNum;
        }

        private void form_Load(object sender, EventArgs e)
        {
            // 2018/01/05
            txtTanka.AutoSize = false;
            txtTanka.Height = 25;

            txtMai.AutoSize = false;
            txtMai.Height = 25;

            txtNebikigo.AutoSize = false;
            txtNebikigo.Height = 24;

            txtZeikomi.AutoSize = false;
            txtZeikomi.Height = 24;

            txtTax.AutoSize = false;
            txtTax.Height = 24;

            cmbeGaichu.AutoSize = false;
            cmbeGaichu.Height = 24;
            
            // ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: このコード行はデータを 'darwinDataSet.受注' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);
            //this.darwinDataSet.Clear();
            //this.darwinDataSet.EnforceConstraints = false;
            //this.受注TableAdapter.Fill(this.darwinDataSet.受注);

            //////受注区分コンボ
            ////cmbJkbnSet();

            // 外注先コンボボックス
            Utility.comboGaichu.itemLoad(cmbeGaichu);   // 営業用
            Utility.comboGaichu.itemLoad(cmbpGaichu);   // 支払用
            Utility.comboGaichu.itemLoad(cmbpGaichu2);  // 支払用2 : 2016/10/14
            Utility.comboGaichu.itemLoad(cmbpGaichu3);  // 支払用3 : 2016/10/14

            cg = Utility.comboGaichu.getArrayGaichu();  // 2018/01/03

            // 事業所コンボ
            Utility.ComboOffice.load(cmbOffice);

            // 得意先コンボ
            Utility.ComboClient.itemsLoad(cmbClient);

            // 受注種別コンボ
            Utility.ComboJshubetsu.load(cmbNaiyou);
            Utility.ComboJshubetsu.load(cmbsNaiyou);

            // 配布形態コンボ
            Utility.ComboFkeitai.load(cmbFkeitai);

            // 配布条件コンボ
            cmbFjyoukenSet();

            // 判型コンボ
            Utility.ComboSize.load(cmbSize);
            Utility.ComboSize.load(cmbsSize);   // 検索サイズ 2019/02/14

            // 案件種別コンボ
            Utility.ComboAnshu.load(cmbAnShu);

            //配布猶予コンボ
            //cmbFyuyoSet();    // 2015/06/23

            //納品形態コンボ
            //cmbNkeitaiSet();  // 2015/06/23

            ////////入金方法コンボ
            //////Utility.ComboShimebi.load(cmbNyukin);

            //振込口座コンボ
            //Utility.ComboFuri.load(cmbFuri);

            //報告時期コンボ
            //cmbHjikiSet();    // 2015/06/23

            //報告精度コンボ
            //cmbHseidoSet();   // 2015/06/23

            //報告方法コンボ
            //cmbHhouhouSet();  // 2015/06/23

            // 税率取得
            cTax.Ritsu = GetTaxRT(DateTime.Today);

            // 画面初期化
            DispClear();

            sDt.Value = DateTime.Today.AddYears(-1);
            eDt.Value = DateTime.Today;

            // 受注番号指定のとき
            if (_orderNum != 0)
            {
                // フォームモードステータス/受注ID:変更削除
                fMode.Mode = 1;
                fMode.jID = _orderNum;
                cMaster.ID = fMode.jID;

                // データ表示
                dataShow();
            }

            // 受注確定書編集設定を取得
            sLoginTag = getOrderLock(out dtLock);            
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     受注確定書編集設定を取得 </summary>
        /// <param name="dt">
        ///     請求書発行日</param>
        /// <returns>
        ///     ログインタグ</returns>
        ///-----------------------------------------------------------
        private int getOrderLock(out DateTime dt)
        {
            int sTag = 0;
            dt = DateTime.Parse("2999/01/01");

            if (dts.受注編集制限.Any(a => a.ID == global.lockKey))
            {
                var s = dts.受注編集制限.Single(a => a.ID == global.lockKey);
                sTag = s.ログイングループ;
                dt = s.請求書発行日;
            }

            return sTag;
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
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    tempDGV.EnableHeadersVisualStyles = false;

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 144;

                    // 奇数行の色
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    tempDGV.Columns[0].Width = 110;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 140;
                    tempDGV.Columns[11].Width = 140;
                    tempDGV.Columns[12].Width = 140;

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

            ///--------------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビューの指定行のデータを取得する </summary>
            /// <param name="dgv">
            ///     対象とするデータグリッドビューオブジェクト</param>
            ///--------------------------------------------------------------------------
            public static Boolean GetData(DataGridView dgv, ref Entity.受注 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.受注 Order = new Control.受注();
                OleDbDataReader dr;

                sqlStr = " where 受注.ID = " + (long)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Order.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = long.Parse(dr["ID"].ToString());
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
                        tempC.配布単価 = double.Parse(dr["配布単価"].ToString());
                        tempC.依頼先 = dr["依頼先"].ToString();
                        tempC.原価 = Convert.ToDouble(dr["原価"].ToString());
                        tempC.配布形態 = Int32.Parse(dr["配布形態"].ToString());
                        tempC.配布条件 = dr["配布条件"].ToString() + "";
                        tempC.配布開始日 = dr["配布開始日"].ToString();
                        tempC.配布終了日 = dr["配布終了日"].ToString();
                        //tempC.配布猶予 = dr["配布猶予"].ToString() + "";  // 2015/06/23
                        tempC.納品予定日 = dr["納品予定日"].ToString();
                        //tempC.納品形態 = dr["納品形態"].ToString() + "";  // 2015/06/23
                        tempC.請求書 = Int32.Parse(dr["請求書"].ToString());
                        tempC.請求書ID = int.Parse(dr["請求書ID"].ToString());
                        tempC.請求書発行日 = dr["請求書発行日"].ToString();
                        tempC.入金方法 = dr["入金方法"].ToString() + "";
                        tempC.入金予定日 = dr["入金予定日"].ToString();
                        //tempC.報告時期 = dr["報告時期"].ToString() + "";  // 2015/06/23
                        //tempC.報告精度 = dr["報告精度"].ToString() + "";  // 2015/06/23
                        //tempC.報告方法 = dr["報告方法"].ToString() + "";  // 2015/06/23
                        //tempC.メールアドレス = dr["メールアドレス"].ToString() + "";    // 2015/06/23
                        tempC.振込口座ID = Int32.Parse(dr["振込口座ID"].ToString());
                        tempC.未配布情報有無 = Int32.Parse(dr["未配布情報有無"].ToString());
                        tempC.枝番有無 = Int32.Parse(dr["枝番有無"].ToString());
                        tempC.特記事項 = dr["特記事項"].ToString() + "";
                        tempC.エリア備考 = dr["エリア備考"].ToString() + "";
                        tempC.完了区分 = int.Parse(dr["完了区分"].ToString());
                        tempC.併配除外 = Int32.Parse(dr["併配除外"].ToString());
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
            msgStr += dataGridView1[3, dataGridView1.SelectedRows[iX].Index].Value +"が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "受注確定書選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    // フォームモードステータス/受注ID:変更削除
                    fMode.Mode = 1;
                    fMode.jID = (long)dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value;
                    cMaster.ID = fMode.jID;

                    // 入力開始可能日付：2017/01/27
                    jDate.MinDate = DateTime.Parse("1900/01/01");
                    dtSeikyu.MinDate = DateTime.Parse("1900/01/01");
                    NyukinDate.MinDate = DateTime.Parse("1900/01/01");

                    // データ表示
                    dataShow();

                    //// フォームモードステータス/受注ID:変更削除
                    //fMode.Mode = 1; 
                    //fMode.jID = (long)dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value;
                    //cMaster.ID = fMode.jID;

                    //// データを取得する
                    //if (!GetData(fMode.jID))
                    //{
                    //    MessageBox.Show("該当するデータが登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    //    return;
                    //}

                    //// ボタン状態
                    //btnDel.Enabled = true;
                    //btnClr.Enabled = true;
                    //button2.Enabled = true;
                    //button3.Enabled = true;

                    //// 受注番号 2015/07/13
                    //label45.Visible = false;
                    //groupBox3.Visible = false;

                    //jDate.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "データ表示", MessageBoxButtons.OK);
                }
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     データ表示 </summary>
        ///------------------------------------------------------------
        private void dataShow()
        {
            int hCnt = 0;

            // データを取得する
            if (!GetData(fMode.jID, ref hCnt))
            {
                MessageBox.Show("該当するデータが登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            // ボタン状態
            if (orderEditStatus)
            {
                btnUpdate.Enabled = true;
                btnDel.Enabled = true;
                button2.Enabled = true;
                label52.Visible = false;
            }
            else
            {
                btnUpdate.Enabled = false;
                btnDel.Enabled = false;
                button2.Enabled = false;
                label52.Visible = true;
            }

            // 配布エリア登録済み受注データのとき削除不可とする：2017/01/27
            if (hCnt > 0)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }

            btnClr.Enabled = true;
            button3.Enabled = true;

            // 受注番号 2015/07/13
            label45.Visible = false;
            groupBox3.Visible = false;

            // 受注データコピーリンクボタン：2018/01/05
            lnkCopy.Visible = true;
            
            jDate.Focus();
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの指定行のデータを取得する </summary>
        /// <param name="dgv">
        ///     対象とするデータグリッドビューオブジェクト</param>
        /// -----------------------------------------------------------------------------
        private Boolean GetData(long sID, ref int hCnt)
        {
            // 受注データ読み込み
            DateTime sd = DateTime.Parse(sDt.Value.ToShortDateString());
            DateTime ed = DateTime.Parse(eDt.Value.ToShortDateString());
            jAdp.FillByFromYMDToYMD(dts.受注1, sd, ed);

            //jAdp.Fill(dts.受注1);

            foreach (var t in dts.受注1.Where(a => a.ID == sID))
            {
                cmbNaiyou.SelectedValue = t.受注種別ID;
                jDate.Value = DateTime.Parse(t.受注日.ToShortDateString());
                cmbOffice.SelectedValue = t.事業所ID;

                cmbClient.SelectedValue = t.得意先ID;
                txtCzipcode.Text = "";
                txtName2.Text = "";
                txtCbusho.Text = "";
                txtTantou.Text = "";
                txtCtel.Text = "";
                txtCfax.Text = "";
                txtCbusho.Text = "";
                txtCtantou.Text = "";

                //クライアント情報表示
                ClientShow(t.得意先ID);

                txtChirashi.Text = t.チラシ名;

                txtUri.Text = t.金額.ToString("#,##0");  // 単価、枚数より先に表示 2019/07/30
                txtTanka.Text = t.単価.ToString("#,##0.00");
                txtMai.Text = t.枚数.ToString("#,##0");
                txtNebiki.Text = t.値引額.ToString("#,##0");

                // 値引後金額
                int nebikigoKin = Utility.strToInt(txtUri.Text) - Utility.strToInt(txtNebiki.Text);
                txtNebikigo.Text = nebikigoKin.ToString("#,##0");

                txtTax.Text = t.消費税.ToString("#,##0");
                txtZeikomi.Text = t.税込金額.ToString("#,##0");

                cmbFkeitai.SelectedValue = t.配布形態;
                cmbFjyouken.Text = t.配布条件;
                cmbSize.SelectedValue = t.判型;
                txtHTanka.Text = t.配布単価.ToString("#,##0.00");

                if (t.Is配布開始日Null())
                {
                    StartDate.Checked = false;
                }
                else
                {
                    StartDate.Checked = true;
                    StartDate.Value = DateTime.Parse(t.配布開始日.ToShortDateString());
                }

                if (t.Is配布終了日Null())
                {
                    EndDate.Checked = false;
                }
                else
                {
                    EndDate.Checked = true;
                    EndDate.Value = DateTime.Parse(t.配布終了日.ToShortDateString());
                }

                if (t.Is納品予定日Null())
                {
                    NouhinDate.Checked = false;
                }
                else
                {
                    NouhinDate.Checked = true;
                    NouhinDate.Value = DateTime.Parse(t.納品予定日.ToShortDateString());
                }
                
                if (t.Is入金予定日Null())
                {
                    NyukinDate.Checked = false;
                }
                else
                {
                    NyukinDate.Checked = true;
                    NyukinDate.Value = DateTime.Parse(t.入金予定日.ToShortDateString());
                }

                txtMemo.Text = t.特記事項;
                //txtMemo2.Text = t.エリア備考;

                txtSalesMemo.Text = t.営業備考;     // 2019/03/01

                if (int.Parse(t.未配布情報有無.ToString()) == 1)
                {
                    checkBox2.Checked = true;
                }
                else
                {
                    checkBox2.Checked = false;
                }

                if (int.Parse(t.枝番有無.ToString()) == 1)
                {
                    checkBox3.Checked = true;
                }
                else
                {
                    checkBox3.Checked = false;
                }

                // 請求書№ 2010/05/30
                txtSeikyuNumber.Text = t.請求書ID.ToString();

                // 併配除外 2014/11/26
                if (Utility.nullToInt(t.併配除外) == 1)
                {
                    chkHeihai.Checked = true;
                }
                else
                {
                    chkHeihai.Checked = false;
                }

                //請求書発行日 2015/06/30
                if (t.Is請求書発行日Null())
                {
                    dtSeikyu.Checked = false;
                }
                else
                {
                    dtSeikyu.Checked = true;
                    dtSeikyu.Value = DateTime.Parse(t.請求書発行日.ToShortDateString());
                }

                // 案件種別 2015/06/30
                cmbAnShu.SelectedValue = t.案件種別;

                // 外注先・営業用 2015/06/30
                cmbeGaichu.SelectedValue = t.外注先ID営業;

                // 外注支払日・営業 2015/06/30
                if (t.Is外注支払日営業Null())
                {
                    dteGaichuPay.Checked = false;
                }
                else
                {
                    dteGaichuPay.Checked = true;
                    dteGaichuPay.Value = DateTime.Parse(t.外注支払日営業.ToShortDateString());
                }

                // 外注原価・営業 2015/06/30
                txteGaichuGenka.Text = t.外注原価営業.ToString("#,##0");

                // 外注先・支払用 2015/06/30
                cmbpGaichu.SelectedValue = t.外注先ID支払;
                

                // 外注支払日・支払 2015/06/30
                if (t.Is外注支払日支払Null())
                {
                    dtpGaichuPay.Checked = false;
                }
                else
                {
                    dtpGaichuPay.Checked = true;
                    dtpGaichuPay.Value = DateTime.Parse(t.外注支払日支払.ToShortDateString());
                }

                // 外注原価・支払 2015/06/30
                txtpGaichuGenka.Text = t.外注原価支払.ToString("#,##0");

                // 外注先・支払用2 2016/10/14
                cmbpGaichu2.SelectedValue = t.外注先ID支払2;

                // 外注支払日・支払2 2016/10/14
                if (t.Is外注支払日支払2Null())
                {
                    dtpGaichuPay2.Checked = false;
                }
                else
                {
                    dtpGaichuPay2.Checked = true;
                    dtpGaichuPay2.Value = DateTime.Parse(t.外注支払日支払2.ToShortDateString());
                }

                // 外注原価・支払2 2016/10/14
                txtpGaichuGenka2.Text = t.外注原価支払2.ToString("#,##0");

                // 外注先・支払用3 2016/10/14
                cmbpGaichu3.SelectedValue = t.外注先ID支払3;

                // 外注支払日・支払3 2016/10/14
                if (t.Is外注支払日支払3Null())
                {
                    dtpGaichuPay3.Checked = false;
                }
                else
                {
                    dtpGaichuPay3.Checked = true;
                    dtpGaichuPay3.Value = DateTime.Parse(t.外注支払日支払3.ToShortDateString());
                }

                // 外注原価・支払3 2016/10/14
                txtpGaichuGenka3.Text = t.外注原価支払3.ToString("#,##0");

                // 外注依頼日 2015/08/11
                if (t.Is外注依頼日支払Null())
                {
                    dtGaichuIrai.Checked = false;
                }
                else
                {
                    dtGaichuIrai.Checked = true;
                    dtGaichuIrai.Value = DateTime.Parse(t.外注依頼日支払.ToShortDateString());
                }

                // 2016/10/15
                if (t.Is外注依頼日支払2Null())
                {
                    dtGaichuIrai2.Checked = false;
                }
                else
                {
                    dtGaichuIrai2.Checked = true;
                    dtGaichuIrai2.Value = DateTime.Parse(t.外注依頼日支払2.ToShortDateString());
                }

                if (t.Is外注依頼日支払3Null())
                {
                    dtGaichuIrai3.Checked = false;
                }
                else
                {
                    dtGaichuIrai3.Checked = true;
                    dtGaichuIrai3.Value = DateTime.Parse(t.外注依頼日支払3.ToShortDateString());
                }

                // 外注渡し日 2015/08/11
                if (t.Is外注渡し日Null())
                {
                    dtGaichuWatashi.Checked = false;
                }
                else
                {
                    dtGaichuWatashi.Checked = true;
                    dtGaichuWatashi.Value = DateTime.Parse(t.外注渡し日.ToShortDateString());
                }

                // 2016/10/15
                if (t.Is外注渡し日2Null())
                {
                    dtGaichuWatashi2.Checked = false;
                }
                else
                {
                    dtGaichuWatashi2.Checked = true;
                    dtGaichuWatashi2.Value = DateTime.Parse(t.外注渡し日2.ToShortDateString());
                }

                if (t.Is外注渡し日3Null())
                {
                    dtGaichuWatashi3.Checked = false;
                }
                else
                {
                    dtGaichuWatashi3.Checked = true;
                    dtGaichuWatashi3.Value = DateTime.Parse(t.外注渡し日3.ToShortDateString());
                }

                // 外注受け渡し担当者 2015/08/11
                txtGaichuUkeTan.Text = t.外注受け渡し担当者;
                txtGaichuUkeTan2.Text = t.外注受け渡し担当者2;    // 2016/10/15
                txtGaichuUkeTan3.Text = t.外注受け渡し担当者3;    // 2016/10/15

                // 外注委託枚数
                txtGaichuMaisu.Text = t.外注委託枚数.ToString();
                txtGaichuMaisu2.Text = t.外注委託枚数2.ToString();    // 2016/10/15
                txtGaichuMaisu3.Text = t.外注委託枚数3.ToString();    // 2016/10/15

                // 業種
                if (t.業種 != null)
                {
                    txtGyoushu.Text = t.業種;
                }
                else
                {
                    txtGyoushu.Text = string.Empty;
                }

                // 入力完了の受注確定書は編集ロックする 2016/07/19
                if (t.新請求書Row != null && t.新請求書Row.入金完了 == global.FLGON)
                {
                    orderEditStatus = false;
                    label52.Text = "既に入金が完了しているため編集は出来ません";
                }
                else if (sLoginTag != global.FLGOFF)
                {
                    // 受注確定書編集制限を検証 2016/07/07
                    if (global.loginType == sLoginTag)
                    {
                        // 編集可能なログインタイプ
                        orderEditStatus = true;
                    }
                    else
                    {
                        if (!t.Is請求書発行日Null() && (t.請求書発行日 <= dtLock))
                        {
                            // 請求書発行日が編集ロック期間中
                            orderEditStatus = false;
                            label52.Text = "請求書発行日が " + dtLock.ToShortDateString() + " 以前のためこの受注確定書の編集は制限されています";
                        }
                        else
                        {
                            // 請求書発行日が編集ロック期間以降
                            orderEditStatus = true;
                        }
                    }

                    //// 配布エリア取り込み済みか？　2017/01/27
                    //hCnt = t.Get配布エリアRows().Count();


                    // 配布エリア取り込み済みか？　2018/01/27
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();

                    cn = Con.GetConnection();
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.Connection = cn;

                    //配布エリア取り込み済みか？
                    StringBuilder sb = new StringBuilder();
                    sb.Append("Select count(*) as cnt from 配布エリア ");
                    sb.Append("where 受注ID = ").Append(t.ID);

                    SCom.CommandText = sb.ToString();
                    OleDbDataReader dr = SCom.ExecuteReader();

                    string cc = string.Empty;
                    while (dr.Read())
                    {
                        cc = dr["cnt"].ToString();
                    }

                    dr.Close();
                    SCom.Connection.Close();

                    hCnt = Utility.strToInt(cc);

                }
                else
                {
                    orderEditStatus = true;
                }

                // 確実な粗利金額：2019/07/30
                txteGaichuArari.Text = (Utility.strToInt(txtUri.Text) - Utility.strToInt(txteGaichuGenka.Text)).ToString("#,##0");

            }

            return true;
        }


        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            // Enterkey以外は対象外
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            GridEnter();
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //GridEnter();
        }

        ///----------------------------------------------------
        /// <summary>
        ///     画面をクリアする </summary>
        ///----------------------------------------------------
        private void DispClear()
        {
            try
            {
                // 検索項目 : 2015/06/24
                //sDt.Checked = false;  2018/01/27
                //eDt.Checked = false;  2018/01/27
                textBox1.Text = string.Empty;
                cmbsNaiyou.SelectedIndex = -1;
                txtsClient.Text = string.Empty;
                txtsGaichu.Text = string.Empty;
                txtsOrderNum.Text = string.Empty;

                // 検索項目追加 2019/02/14
                txtsMaisu.Text = string.Empty;  // 枚数
                cmbsSize.SelectedIndex = -1;    // サイズ
                seDt.Checked = false;           // 請求締日

                // 表示項目
                fMode.Mode = 0;

                // 新規入力時にのみ入力日付制限：2017/08/28
                DateTime minDate = DateTime.Today.AddMonths(-1);
                if (_orderNum == global.FLGOFF)
                {
                    // 入力日付制限 2016/01/27
                    minDate = DateTime.Parse(minDate.Year.ToString() + "/" + minDate.Month.ToString() + "/01");

                    // 受注日入力可能開始日：2017/01/27
                    jDate.Value = DateTime.Today;
                    jDate.MinDate = minDate;
                }

                //////comboBox1.SelectedIndex = -1;
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
                textBox4.Text = "";

                txtChirashi.Text = "";
                cmbNaiyou.SelectedIndex = -1;
                txtTanka.Text = "0";
                txtMai.Text = "0";
                txtUri.Text = "0";
                txtTax.Text = "0";
                txtZeikomi.Text = "0";
                txtNebiki.Text = "0";
                txtNebikigo.Text = "0";

                cmbFkeitai.SelectedIndex = -1;
                cmbFjyouken.SelectedIndex = -1;
                cmbSize.SelectedIndex = -1;
                //cmbSize.Text = string.Empty;

                txtHTanka.Text = "0";

                //txtIraisaki.Text = "";
                //txtGenka.Text = "0";

                StartDate.Checked = false;
                EndDate.Checked = false;

                //cmbFyuyo.SelectedIndex = -1;      // 2015/06/23
                //checkBox1.Checked = false;

                NouhinDate.Checked = false;
                NouhinDate.Enabled = true;  // 2019/02/14

                //cmbNkeitai.SelectedIndex = -1;    // 2015/06/23

                //////cmbNyukin.SelectedIndex = -1;

                NyukinDate.Checked = false;

                // 新規入力時にのみ入力日付制限：2017/08/28
                if (_orderNum == global.FLGOFF)
                {
                    NyukinDate.MinDate = minDate;   // 支払期日入力開始可能日：2017/01/27
                }

                //cmbFuri.SelectedIndex = -1;

                //cmbHjiki.SelectedIndex = -1;      // 2015/06/23
                //cmbHseido.SelectedIndex = -1;     // 2015/06/23
                //cmbHhouhou.SelectedIndex = -1;    // 2015/06/23

                //txtEmail.Text = "";               // 2015/06/23

                txtMemo.Text = "";
                txtSalesMemo.Text = "";             // 2019/03/01
                //txtMemo2.Text = "";
                checkBox2.Checked = false;
                checkBox3.Checked = false;

                txtSeikyuNumber.Text = "0";             // 2010/05/30

                chkHeihai.Checked = false;              // 2014/11/26 併配チェック

                dtSeikyu.Checked = false;               // 請求書発行日 2015/06/30

                // 新規入力時にのみ入力日付制限：2017/08/28
                if (_orderNum == global.FLGOFF)
                {
                    dtSeikyu.MinDate = minDate;         // 請求書発行日入力可能開始日：2017/01/27
                }

                cmbAnShu.SelectedIndex = -1;            // 案件種別コンボボックス 2015/06/30

                cmbeGaichu.SelectedIndex = -1;          // 営業・外注先 2015/06/30
                dteGaichuPay.Checked = false;           // 営業・外注支払日 2015/06/30
                txteGaichuGenka.Text = string.Empty;    // 営業・外注原価 2015/06/30

                cmbpGaichu.SelectedIndex = -1;          // 支払・外注先 2015/06/30
                dtpGaichuPay.Checked = false;           // 支払・外注支払日 2015/06/30
                txtpGaichuGenka.Text = string.Empty;    // 支払・外注原価 2015/06/30
                lblSD1.Text = string.Empty;             // 支払日 2018/01/04

                cmbpGaichu2.SelectedIndex = -1;         // 支払・外注先2 2016/10/14
                dtpGaichuPay2.Checked = false;          // 支払・外注支払日2 2016/10/14
                txtpGaichuGenka2.Text = string.Empty;   // 支払・外注原価2 2016/10/14
                lblSD2.Text = string.Empty;             // 支払日 2018/01/04

                cmbpGaichu3.SelectedIndex = -1;         // 支払・外注先3 2016/10/14
                dtpGaichuPay3.Checked = false;          // 支払・外注支払日3 2016/10/14
                txtpGaichuGenka3.Text = string.Empty;   // 支払・外注原価3 2016/10/14
                lblSD3.Text = string.Empty;             // 支払日 2018/01/04

                dtGaichuIrai.Checked = false;           // 外注依頼日 2015/08/11
                dtGaichuWatashi.Checked = false;        // 外注渡し日 2015/08/11
                txtGaichuUkeTan.Text = string.Empty;    // 外注受け渡し担当者 2015/08/11
                txtGaichuMaisu.Text = string.Empty;     // 外注に渡した枚数 2015/09/20

                dtGaichuIrai2.Checked = false;          // 外注依頼日 2016/10/15
                dtGaichuWatashi2.Checked = false;       // 外注渡し日 2016/10/15
                txtGaichuUkeTan2.Text = string.Empty;   // 外注受け渡し担当者 2016/10/15
                txtGaichuMaisu2.Text = string.Empty;    // 外注に渡した枚数 2016/10/15

                dtGaichuIrai3.Checked = false;          // 外注依頼日 2016/10/15
                dtGaichuWatashi3.Checked = false;       // 外注渡し日 2016/10/15
                txtGaichuUkeTan3.Text = string.Empty;   // 外注受け渡し担当者 2016/10/15
                txtGaichuMaisu3.Text = string.Empty;    // 外注に渡した枚数 2016/10/15

                txtGyoushu.Text = string.Empty;         // 業種            2015/09/20

                tabControl1.SelectTab(0);
                tabControl1.TabPages[3].Text = "粗利";    // 2019/02/21

                // 2019/02/21
                txtaUri.Text = string.Empty;
                txtaGai1.Text = string.Empty;
                txtaGai2.Text = string.Empty;
                txtaGai3.Text = string.Empty;
                txtaGaiGenka1.Text = string.Empty;
                txtaGaiGenka2.Text = string.Empty;
                txtaGaiGenka3.Text = string.Empty;
                txtaArari.Text = string.Empty;

                btnUpdate.Enabled = true;
                btnDel.Enabled = false;
                btnClr.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;

                // 受注番号 2015/07/13
                label45.Visible = true;
                rBtnOrderNum1.Checked = true;
                groupBox3.Visible = true;
                lblOrderNum.Text = string.Empty;
                
                //txtCode.Focus();
                jDate.Focus();

                // 2016/07/07
                label52.Visible = false;

                lblClientShimebi.Text = string.Empty;   // 2018/01/04
                lnkCopy.Visible = false;    // 2018/01/05
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("選択されているデータを変更しないで終了します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
                
            DispClear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                if (fDataCheck())
                {
                    Control.受注 Order = new Control.受注();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Order.Close();
                                return;
                            }

                            // 以下、2015/07/15 Utility.getOrderNum()に置き換え

                            ////受注日取得
                            //DateTime sJDate = jDate.Value;

                            //IDを採番
                            //string sqlStr = "";
                            //long gID = 0;
                            //long sWk;
                            //DateTime sJDate;
                            
                            ////伝票番号最小値
                            //sWk = sJDate.Year;
                            //cMaster.IDTemplateS = sWk * 100000000;

                            //sWk = sJDate.Month;
                            //cMaster.IDTemplateS += sWk * 1000000;

                            //sWk = sJDate.Day;
                            //cMaster.IDTemplateS += sWk * 10000;

                            //cMaster.IDTemplateS++;

                            ////伝票番号最大値
                            //cMaster.IDTemplateE = cMaster.IDTemplateS + 9998;

                            ////受注日の伝票があるか？
                            //sqlStr += "select max(ID) as ID from 受注 ";
                            //sqlStr += "where (ID >= " + cMaster.IDTemplateS.ToString() + ") and ";
                            //sqlStr += "(ID <= " + cMaster.IDTemplateE.ToString() + ")";

                            //OleDbDataReader dR;
                            //Control.FreeSql fCon = new Control.FreeSql();
                            //dR = fCon.free_dsReader(sqlStr);

                            //while (dR.Read())
                            //{
                            //    if (dR["ID"] == DBNull.Value)
                            //    {
                            //        gID = cMaster.IDTemplateS;　//なければテンプレートの伝票番号最小値をセット
                            //    }
                            //    else
                            //    {
                            //        gID = long.Parse(dR["ID"].ToString()) + 1;　//あれば1インクリメント
                            //    }
                            //}

                            //dR.Close();
                            //fCon.Close();

                            // 受注番号採番 2015/07/15
                            if (rBtnOrderNum1.Checked)
                            {
                                // 自動採番で受注番号を取得
                                cMaster.ID = Utility.getOrderNum(jDate.Value);
                            }
                            else if (rBtnOrderNum2.Checked)
                            {
                                // 取得済み番号と紐付け
                                cMaster.ID = long.Parse(lblOrderNum.Text);
                            }

                            //請求書ID
                            //cMaster.請求書ID = 0;

                            if (Order.DataInsert(cMaster) == true)
                            {
                                if (rBtnOrderNum2.Checked)
                                {
                                    // 受注番号採番データ更新
                                    orderNumUpdate(cMaster.ID, cMaster.得意先ID);
                                }

                                MessageBox.Show("新規登録されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("新規登録に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Order.Close();
                                return;
                            }

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
                    //this.受注TableAdapter.Fill(this.darwinDataSet.受注);
                    //dataGridView1.DataSource = this.darwinDataSet.受注;
                    dataGridView1.DataSource = null;

                    // 受注データセット読み込み
                    //jAdp.Fill(dts.受注1);

                    DateTime sd = DateTime.Parse(sDt.Value.ToShortDateString()); // 2018/01/27
                    DateTime ed = DateTime.Parse(eDt.Value.ToShortDateString()); // 2018/01/27
                    jAdp.FillByFromYMDToYMD(dts.受注1, sd, ed); // 2018/01/27
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"更新処理",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     受注番号採番紐付情報更新 </summary>
        /// <param name="jID">
        ///     受注番号</param>
        /// -------------------------------------------------------------------
        private void orderNumUpdate(long jID, int clientID)
        {
            darwinDataSetTableAdapters.受注番号採番TableAdapter adp = new darwinDataSetTableAdapters.受注番号採番TableAdapter();
            adp.Fill(dts.受注番号採番);

            // 受注番号採番紐付情報更新
            if (dts.受注番号採番.Any(a => a.受注番号 == jID))
            {
                var s = dts.受注番号採番.Single(a => a.受注番号 == jID);
                s.得意先ID = clientID;
                s.確定書入力 = 1;
                s.確定書入力日付 = DateTime.Now;
                s.確定書入力ユーザーID = global.loginUserID;
                s.更新年月日 = DateTime.Now;

                // データベース更新
                adp.Update(dts.受注番号採番);
            }
        }


        // 登録データチェック
        private Boolean fDataCheck()
        {
            string str;
            double d;
            int dInt;

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

                    // 紐付け受注番号のとき
                    if (rBtnOrderNum2.Checked && lblOrderNum.Text == string.Empty)
                    {
                        btnOrderNum.Focus();
                        throw new Exception("取得済みの受注番号を選択してください");
                    }

                }

                //受注区分チェック
                //////if (comboBox1.SelectedIndex == -1)
                //////{
                //////    comboBox1.Focus();
                //////    throw new Exception("受注区分を選択してください");
                //////}

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
                else
                {
                    // 受注内容と案件種別 2015/06/30
                    if ((cmbNaiyou.SelectedIndex == 0 && cmbAnShu.SelectedIndex == 2) || 
                        (cmbNaiyou.SelectedIndex > 0 && cmbAnShu.SelectedIndex < 2))
                    {
                        cmbNaiyou.Focus();
                        throw new Exception("受注内容と案件種別が一致していません");
                    }
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

                if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out dInt))
                {
                }
                else
                {
                    this.txtMai.Focus();
                    throw new Exception("枚数が正しくありません");
                }

                ////////原価：数字か？
                //////if (txtGenka.Text == null)
                //////{
                //////    this.txtGenka.Focus();
                //////    throw new Exception("原価は数字で入力してください");
                //////}

                //////str = this.txtGenka.Text;

                //////if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                //////{
                //////}
                //////else
                //////{
                //////    this.txtGenka.Focus();
                //////    throw new Exception("原価は数字で入力してください");
                //////}

                //ポスティングのときのみチェック対象とする項目
                if (cmbNaiyou.Text == J_NAIYOU_POSTING)
                {
                    //配布形態チェック
                    if (cmbFkeitai.SelectedIndex == -1)
                    {
                        cmbFkeitai.Focus();
                        throw new Exception("配布形態を選択してください");
                    }

                    //配布条件チェック
                    if (cmbFjyouken.SelectedIndex == -1)
                    {
                        cmbFjyouken.Focus();
                        throw new Exception("配布条件を選択してください");
                    }

                    //判型チェック
                    if (cmbSize.SelectedIndex == -1)
                    {
                        cmbSize.Focus();
                        throw new Exception("サイズを選択してください");
                    }

                    //配布単価：数字か？
                    if (Utility.NumericCheck(txtHTanka.Text) == false)
                    {
                        this.txtHTanka.Focus();
                        throw new Exception("配布単価は数字で入力してください");
                    }

                    ////配布開始日
                    //if (StartDate.Checked == false)
                    //{
                    //    StartDate.Focus();
                    //    throw new Exception("配布開始日を入力してください");
                    //}

                    ////配布終了日
                    //if (EndDate.Checked == false)
                    //{
                    //    EndDate.Focus();
                    //    throw new Exception("配布終了日を入力してください");
                    //}

                    //配布期間
                    if (StartDate.Value > EndDate.Value)
                    {
                        StartDate.Focus();
                        throw new Exception("配布期間が正しくありません");
                    }
                    
                    // 納品日必須とする 2016/01/17
                    if (!NouhinDate.Checked)
                    {
                        NouhinDate.Focus();
                        throw new Exception("納品日が未入力です");
                    }

                    ////配布猶予チェック    // 2015/06/23
                    //if (cmbFyuyo.SelectedIndex == -1)
                    //{
                    //    cmbFyuyo.Focus();
                    //    throw new Exception("配布猶予を選択してください");
                    //}

                    ////報告時期チェック    // 2015/06/23
                    //if (cmbHjiki.SelectedIndex == -1)
                    //{
                    //    cmbHjiki.Focus();
                    //    throw new Exception("報告時期を選択してください");
                    //}

                    ////報告精度チェック    // 2015/06/23
                    //if (cmbHseido.SelectedIndex == -1)
                    //{
                    //    cmbHseido.Focus();
                    //    throw new Exception("報告精度を選択してください");
                    //}

                    ////報告方法チェック    // 2015/06/23
                    //if (cmbHhouhou.SelectedIndex == -1)
                    //{
                    //    cmbHhouhou.Focus();
                    //    throw new Exception("報告方法を選択してください");
                    //}
                }
                
                // 配布開始日：2016/08/03 受注内容がポスティングに限らず必須入力項目とする
                if (StartDate.Checked == false)
                {
                    StartDate.Focus();
                    throw new Exception("配布開始日を入力してください");
                }

                //配布終了日：2017/01/27 受注内容がポスティングに限らず必須入力項目とする
                if (EndDate.Checked == false)
                {
                    EndDate.Focus();
                    throw new Exception("配布終了日を入力してください");
                }

                // 外注依頼日と外注渡し枚数　2015/09/20
                if (!dtGaichuIrai.Checked && Utility.strToInt(txtGaichuMaisu.Text) > 0)
                {
                    tabControl1.SelectTab(0);   // 2016/10/15
                    dtGaichuIrai.Focus();
                    throw new Exception("外注依頼日が未入力で渡し枚数が入力されています");
                }

                // 2016/10/15
                if (!dtGaichuIrai2.Checked && Utility.strToInt(txtGaichuMaisu2.Text) > 0)
                {
                    tabControl1.SelectTab(1);
                    dtGaichuIrai2.Focus();
                    throw new Exception("外注依頼日が未入力で渡し枚数が入力されています");
                }

                // 2016/10/15
                if (!dtGaichuIrai3.Checked && Utility.strToInt(txtGaichuMaisu3.Text) > 0)
                {
                    tabControl1.SelectTab(2);
                    dtGaichuIrai3.Focus();
                    throw new Exception("外注依頼日が未入力で渡し枚数が入力されています");
                }

                // 営業原価　2015/11/18
                // 外注先が選択済みで
                if (cmbeGaichu.SelectedIndex != -1)
                {
                    // 支払日が未入力のとき
                    if (!dteGaichuPay.Checked)
                    {
                        dteGaichuPay.Focus();
                        throw new Exception("営業原価の支払日が未入力です");
                    }

                    // 原価が未入力のとき
                    if (txteGaichuGenka.Text == string.Empty)
                    {
                        txteGaichuGenka.Focus();
                        throw new Exception("営業原価が未入力です");
                    }
                }
                else
                {
                    // 2016/01/17
                    cmbeGaichu.Focus();
                    throw new Exception("外注先が選択されていません");
                }

                // 支払日が入力済みで
                if (dteGaichuPay.Checked)
                {
                    // 外注先が未選択のとき
                    if (cmbeGaichu.SelectedIndex == -1)
                    {
                        cmbeGaichu.Focus();
                        throw new Exception("営業原価の外注先を選択してください");
                    }

                    // 原価が未入力のとき
                    if (txteGaichuGenka.Text == string.Empty)
                    {
                        txteGaichuGenka.Focus();
                        throw new Exception("営業原価が未入力です");
                    }
                }

                // 原価が入力済みで
                if (Utility.strToDouble(txteGaichuGenka.Text) != 0)
                {
                    // 外注先が未選択のとき
                    if (cmbeGaichu.SelectedIndex == -1)
                    {
                        cmbeGaichu.Focus();
                        throw new Exception("営業原価の外注先を選択してください");
                    }

                    // 支払日が未入力のとき
                    if (!dteGaichuPay.Checked)
                    {
                        dteGaichuPay.Focus();
                        throw new Exception("支払日が未入力です");
                    }
                }

                // 外注費支払用　2015/11/18
                // 支払用外注先が選択済みで
                if (cmbpGaichu.SelectedIndex != -1)
                {
                    // 支払日が未入力のとき
                    if (!dtpGaichuPay.Checked)
                    {
                        tabControl1.SelectTab(0);
                        dtpGaichuPay.Focus();
                        throw new Exception("外注費支払用の支払日が未入力です");
                    }

                    // 原価が未入力のとき
                    if (txtpGaichuGenka.Text == string.Empty)
                    {
                        tabControl1.SelectTab(0);
                        txtpGaichuGenka.Focus();
                        throw new Exception("外注費支払用原価が未入力です");
                    }
                }
                else
                {
                    tabControl1.SelectTab(0);

                    // 2016/01/17
                    cmbpGaichu.Focus();
                    throw new Exception("外注先が選択されていません");
                }

                // 外注費支払日が入力済みで
                if (dtpGaichuPay.Checked)
                {
                    // 支払用外注先が未選択のとき
                    if (cmbpGaichu.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(0);
                        cmbpGaichu.Focus();
                        throw new Exception("外注費支払用外注先を選択してください");
                    }

                    // 原価が未入力のとき
                    if (txtpGaichuGenka.Text == string.Empty)
                    {
                        tabControl1.SelectTab(0);
                        txtpGaichuGenka.Focus();
                        throw new Exception("外注費支払用原価が未入力です");
                    }
                }

                // 原価が入力済みで
                if (Utility.strToDouble(txtpGaichuGenka.Text) != 0)
                {
                    // 外注費支払用外注先が未選択のとき
                    if (cmbpGaichu.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(0);
                        cmbpGaichu.Focus();
                        throw new Exception("外注費支払用外注先を選択してください");
                    }

                    // 外注費支払日が未入力のとき
                    if (!dtpGaichuPay.Checked)
                    {
                        tabControl1.SelectTab(0);
                        dtpGaichuPay.Focus();
                        throw new Exception("外注費支払日が未入力です");
                    }
                }


                // 外注費支払用２　2016/10/14
                // 支払用外注先２が選択済みで
                if (cmbpGaichu2.SelectedIndex != -1)
                {
                    // 支払日が未入力のとき
                    if (!dtpGaichuPay2.Checked)
                    {
                        tabControl1.SelectTab(1);
                        dtpGaichuPay2.Focus();
                        throw new Exception("外注費支払用２の支払日が未入力です");
                    }

                    // 原価が未入力のとき
                    if (txtpGaichuGenka2.Text == string.Empty)
                    {
                        tabControl1.SelectTab(1);
                        txtpGaichuGenka2.Focus();
                        throw new Exception("外注費支払用原価２が未入力です");
                    }
                }

                // 外注費支払日2が入力済みで
                if (dtpGaichuPay2.Checked)
                {
                    // 支払用外注先2が未選択のとき
                    if (cmbpGaichu2.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(1);
                        cmbpGaichu2.Focus();
                        throw new Exception("外注費支払用外注先２を選択してください");
                    }

                    // 原価2が未入力のとき
                    if (txtpGaichuGenka2.Text == string.Empty)
                    {
                        tabControl1.SelectTab(1);
                        txtpGaichuGenka2.Focus();
                        throw new Exception("外注費支払用原価２が未入力です");
                    }
                }

                // 原価２が入力済みで
                if (Utility.strToDouble(txtpGaichuGenka2.Text) != 0)
                {
                    // 外注費支払用外注先２が未選択のとき
                    if (cmbpGaichu2.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(1);
                        cmbpGaichu2.Focus();
                        throw new Exception("外注費支払用外注先２を選択してください");
                    }

                    // 外注費支払日２が未入力のとき
                    if (!dtpGaichuPay2.Checked)
                    {
                        tabControl1.SelectTab(1);
                        dtpGaichuPay2.Focus();
                        throw new Exception("外注費支払日２が未入力です");
                    }
                }


                // 外注費支払用３　2016/10/14
                // 支払用外注先３が選択済みで
                if (cmbpGaichu3.SelectedIndex != -1)
                {
                    // 支払日が未入力のとき
                    if (!dtpGaichuPay3.Checked)
                    {
                        tabControl1.SelectTab(2);
                        dtpGaichuPay3.Focus();
                        throw new Exception("外注費支払用３の支払日が未入力です");
                    }

                    // 原価が未入力のとき
                    if (txtpGaichuGenka3.Text == string.Empty)
                    {
                        tabControl1.SelectTab(2);
                        txtpGaichuGenka3.Focus();
                        throw new Exception("外注費支払用原価３が未入力です");
                    }
                }

                // 外注費支払日３が入力済みで
                if (dtpGaichuPay3.Checked)
                {
                    // 支払用外注先3が未選択のとき
                    if (cmbpGaichu3.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(2);
                        cmbpGaichu3.Focus();
                        throw new Exception("外注費支払用外注先３を選択してください");
                    }

                    // 原価３が未入力のとき
                    if (txtpGaichuGenka3.Text == string.Empty)
                    {
                        tabControl1.SelectTab(2);
                        txtpGaichuGenka3.Focus();
                        throw new Exception("外注費支払用原価３が未入力です");
                    }
                }

                // 原価３が入力済みで
                if (Utility.strToDouble(txtpGaichuGenka3.Text) != 0)
                {
                    // 外注費支払用外注先３が未選択のとき
                    if (cmbpGaichu3.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(2);
                        cmbpGaichu3.Focus();
                        throw new Exception("外注費支払用外注先３を選択してください");
                    }

                    // 外注費支払日３が未入力のとき
                    if (!dtpGaichuPay3.Checked)
                    {
                        tabControl1.SelectTab(2);
                        dtpGaichuPay3.Focus();
                        throw new Exception("外注費支払日３が未入力です");
                    }
                }
                
                // 請求書発行日必須とする 2016/01/17
                if (!dtSeikyu.Checked)
                {
                    dtSeikyu.Focus();
                    throw new Exception("請求書発行日が未入力です");
                }

                // 支払期日必須とする 2016/01/17
                if (!NyukinDate.Checked)
                {
                    NyukinDate.Focus();
                    throw new Exception("支払期日が未入力です");
                }

                //クラスにデータセット
                if (fMode.Mode == 1)
                {
                    cMaster.ID = fMode.jID;
                }

                Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                cmb1 = (Utility.ComboOffice)cmbOffice.SelectedItem;
                cMaster.事業所ID = cmb1.ID;

                cMaster.受注日 = jDate.Value;

                //cMaster.受注区分 = comboBox1.Text;
                cMaster.受注区分 = "";

                Utility.ComboClient cmb2 = new Utility.ComboClient();
                cmb2 = (Utility.ComboClient)cmbClient.SelectedItem;
                cMaster.得意先ID = cmb2.ID;

                cMaster.チラシ名 = txtChirashi.Text.ToString();

                Utility.ComboJshubetsu cmb3 = new Utility.ComboJshubetsu();
                cmb3 = (Utility.ComboJshubetsu)cmbNaiyou.SelectedItem;
                cMaster.受注種別ID = cmb3.ID;

                cMaster.単価 = Utility.strToDouble(txtTanka.Text);
                cMaster.枚数 = Utility.strToInt(txtMai.Text);
                cMaster.金額 = Utility.strToInt(txtUri.Text);
                cMaster.消費税 = Utility.strToInt(txtTax.Text);
                cMaster.税込金額 = Utility.strToInt(txtZeikomi.Text);
                cMaster.値引額 = Utility.strToInt(txtNebiki.Text);
                cMaster.売上金額 = Utility.strToInt(txtZeikomi.Text);
                cMaster.税率 = cTax.Ritsu;

                //サイズ(判型)ID取得
                if (cmbSize.SelectedIndex != -1)
                {
                    Utility.ComboSize cmb4 = new Utility.ComboSize();
                    cmb4 = (Utility.ComboSize)cmbSize.SelectedItem;
                    cMaster.判型 = cmb4.ID;
                }
                else
                {
                    cMaster.判型 = 0;
                }

                cMaster.配布単価 = Utility.strToDouble(txtHTanka.Text);
                //cMaster.依頼先 = txtIraisaki.Text.ToString();
                //cMaster.原価 = Convert.ToDouble(txtGenka.Text.ToString());
                cMaster.依頼先 = "";
                cMaster.原価 = (double)0;

                //配布形態ID取得
                if (cmbFkeitai.SelectedIndex != -1)
                {
                    Utility.ComboFkeitai cmb5 = new Utility.ComboFkeitai();
                    cmb5 = (Utility.ComboFkeitai)cmbFkeitai.SelectedItem;
                    cMaster.配布形態 = cmb5.ID;
                }
                else
                {
                    cMaster.配布形態 = 0;
                }

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

                cMaster.配布猶予 = string.Empty;    // 2015/07/01

                if (NouhinDate.Checked == true)
                {
                    cMaster.納品予定日 = NouhinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.納品予定日 = "";
                }

                cMaster.納品形態 = string.Empty;    // 2015/07/01
                
                cMaster.請求書 = 1;

                //if (checkBox1.Checked == true)
                //{
                //    cMaster.請求書 = 1;
                //}
                //else
                //{
                //    cMaster.請求書 = 0;
                //}

                //請求書№
                cMaster.請求書ID = Int32.Parse(txtSeikyuNumber.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);

                cMaster.入金方法 = "";

                //////if (cmbNyukin.SelectedIndex == -1)
                //////{
                //////    cMaster.入金方法 = "";
                //////}
                //////else
                //////{
                //////    Utility.ComboShimebi cmb6 = new Utility.ComboShimebi();
                //////    cmb6 = (Utility.ComboShimebi)cmbNyukin.SelectedItem;
                //////    cMaster.入金方法 = cmb6.Name;
                //////}

                if (NyukinDate.Checked == true)
                {
                    cMaster.入金予定日 = NyukinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.入金予定日 = "";
                }

                cMaster.報告時期 = string.Empty;        // 2015/07/01
                cMaster.報告精度 = string.Empty;        // 2015/07/01
                cMaster.報告方法 = string.Empty;        // 2015/07/01
                cMaster.メールアドレス = string.Empty;  // 2015/07/01

                cMaster.振込口座ID = 0;

                //if (cmbFuri.SelectedIndex == -1)
                //{
                //    cMaster.振込口座ID = 0;
                //}
                //else
                //{
                //    Utility.ComboFuri cmb7 = new Utility.ComboFuri();
                //    cmb7 = (Utility.ComboFuri)cmbFuri.SelectedItem;
                //    cMaster.振込口座ID = cmb7.ID;
                //}

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
                cMaster.エリア備考 = string.Empty;       // 2015/11/11
                cMaster.営業備考 = txtSalesMemo.Text;   // 2019/03/01
                cMaster.完了区分 = 0;

                // 2014/11/26 併配除外
                if (chkHeihai.Checked == true)  
                {
                    cMaster.併配除外 = 1;
                }
                else
                {
                    cMaster.併配除外 = 0;
                }

                if (fMode.Mode == 0)
                {
                    cMaster.登録年月日 = DateTime.Now;
                }

                cMaster.変更年月日 = DateTime.Now;

                // 2015/06/30
                if (dtSeikyu.Checked == true)
                {
                    cMaster.請求書発行日 = dtSeikyu.Value.ToShortDateString();
                }
                else
                {
                    cMaster.請求書発行日 = "";
                }

                // 2015/06/30
                if (cmbeGaichu.SelectedIndex != -1)
                { 
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbeGaichu.SelectedItem;
                    cMaster.外注先ID営業 = cmb.ID;
                }
                else
                {
                    cMaster.外注先ID営業 = 0;
                }

                // 2015/06/30
                if (dteGaichuPay.Checked)
                {
                    cMaster.外注先支払日営業 = dteGaichuPay.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注先支払日営業 = string.Empty;
                }

                // 2015/06/30
                cMaster.外注先原価営業 = Utility.strToDouble(txteGaichuGenka.Text);

                // 2015/06/30
                if (cmbpGaichu.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbpGaichu.SelectedItem;
                    cMaster.外注先ID支払 = cmb.ID;
                }
                else
                {
                    cMaster.外注先ID支払 = 0;
                }

                // 2015/06/30
                if (dtpGaichuPay.Checked)
                {
                    cMaster.外注先支払日支払 = dtpGaichuPay.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注先支払日支払 = string.Empty;
                }

                // 2015/06/30
                cMaster.外注先原価支払 = Utility.strToDouble(txtpGaichuGenka.Text);

                // 2016/10/14
                if (cmbpGaichu2.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbpGaichu2.SelectedItem;
                    cMaster.外注先ID支払2 = cmb.ID;
                }
                else
                {
                    cMaster.外注先ID支払2 = 0;
                }

                // 2016/10/14
                if (dtpGaichuPay2.Checked)
                {
                    cMaster.外注先支払日支払2 = dtpGaichuPay2.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注先支払日支払2 = string.Empty;
                }

                // 2016/10/14
                cMaster.外注先原価支払2 = Utility.strToDouble(txtpGaichuGenka2.Text);

                // 2016/10/14
                if (cmbpGaichu3.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbpGaichu3.SelectedItem;
                    cMaster.外注先ID支払3 = cmb.ID;
                }
                else
                {
                    cMaster.外注先ID支払3 = 0;
                }

                // 2016/10/14
                if (dtpGaichuPay3.Checked)
                {
                    cMaster.外注先支払日支払3 = dtpGaichuPay3.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注先支払日支払3 = string.Empty;
                }

                // 2016/10/14
                cMaster.外注先原価支払3 = Utility.strToDouble(txtpGaichuGenka3.Text);

                // 2015/06/30
                if (cmbAnShu.SelectedIndex != -1)
                {
                    Utility.ComboAnshu cmb = (Utility.ComboAnshu)cmbAnShu.SelectedItem;
                    cMaster.案件種別 = cmb.ID;
                }
                else
                {
                    cMaster.案件種別 = 0;
                }

                // 2015/06/30, 2015/07/10
                cMaster.ユーザーID = global.loginUserID;

                // 2015/08/11
                cMaster.外注先依頼日営業 = string.Empty;

                // 2015/08/11
                if (dtGaichuIrai.Checked)
                {
                    cMaster.外注先依頼日支払 = dtGaichuIrai.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注先依頼日支払 = string.Empty;
                }

                // 2016/10/15
                if (dtGaichuIrai2.Checked)
                {
                    cMaster.外注先依頼日支払2 = dtGaichuIrai2.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注先依頼日支払2 = string.Empty;
                }

                if (dtGaichuIrai3.Checked)
                {
                    cMaster.外注先依頼日支払3 = dtGaichuIrai3.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注先依頼日支払3 = string.Empty;
                }

                // 2015/08/11
                if (dtGaichuWatashi.Checked)
                {
                    cMaster.外注渡し日 = dtGaichuWatashi.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注渡し日 = string.Empty;
                }

                // 2016/10/15
                if (dtGaichuWatashi2.Checked)
                {
                    cMaster.外注渡し日2 = dtGaichuWatashi2.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注渡し日2 = string.Empty;
                }

                if (dtGaichuWatashi3.Checked)
                {
                    cMaster.外注渡し日3 = dtGaichuWatashi3.Value.ToShortDateString();
                }
                else
                {
                    cMaster.外注渡し日3 = string.Empty;
                }
                                
                cMaster.外注受け渡し担当者 = txtGaichuUkeTan.Text;     // 2015/08/11
                cMaster.外注受け渡し担当者2 = txtGaichuUkeTan2.Text;   // 2016/10/15
                cMaster.外注受け渡し担当者3 = txtGaichuUkeTan3.Text;   // 2016/10/15

                // 2015/09/20
                cMaster.外注委託枚数 = Utility.strToInt(txtGaichuMaisu.Text);
                cMaster.外注委託枚数2 = Utility.strToInt(txtGaichuMaisu2.Text);   // 2016/10/15
                cMaster.外注委託枚数3 = Utility.strToInt(txtGaichuMaisu3.Text);   // 2016/10/15

                cMaster.業種 = txtGyoushu.Text;

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

            //if (sender == txtIraisaki)
            //{
            //    objtxt = txtIraisaki;
            //}

            //if (sender == txtGenka)
            //{
            //    objtxt = txtGenka;
            //}

            //if (sender == txtEmail)   // 2015/06/23
            //{
            //    objtxt = txtEmail;
            //}

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            //if (sender == txtMemo2)
            //{
            //    objtxt = txtMemo2;
            //}

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == txtHTanka)
            {
                objtxt = txtHTanka;
            }

            if (sender == cmbClient)
            {
                cmbClient.BackColor = Color.LightSteelBlue;
            }

            // 2015/07/01
            if (sender == cmbeGaichu)
            {
                cmbeGaichu.BackColor = Color.LightSteelBlue;
            }

            if (sender == txteGaichuGenka)
            {
                objtxt = txteGaichuGenka;
            }

            if (sender == cmbpGaichu)
            {
                cmbpGaichu.BackColor = Color.LightSteelBlue;
            }

            if (sender == txtpGaichuGenka)
            {
                objtxt = txtpGaichuGenka;
            }

            if (sender == txtsClient)
            {
                objtxt = txtsClient;
            }

            if (sender == txtsGaichu)
            {
                objtxt = txtsGaichu;
            }

            if (sender == cmbAnShu)
            {
                cmbAnShu.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbNaiyou)
            {
                cmbNaiyou.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbOffice)
            {
                cmbOffice.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbsNaiyou)
            {
                cmbsNaiyou.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbFkeitai)
            {
                cmbFkeitai.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbFjyouken)
            {
                cmbFjyouken.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbSize)
            {
                cmbSize.BackColor = Color.LightSteelBlue;
            }

            if (sender == txtsOrderNum)
            {
                txtsOrderNum.BackColor = Color.LightSteelBlue;
            }


            objtxt.SelectAll();
            objtxt.BackColor = Color.LightSteelBlue;
        }

        private void txtLeave(object sender, EventArgs e)
        {
            TextBox objtxt = new TextBox();

            decimal Kingaku = 0;
            decimal KingakuTax = 0;
            decimal KingakuZeikomi = 0;
            decimal KingakuTL = 0;
            DateTime dt = DateTime.Today;
            string str;
            double d;
            decimal dTanka = 0;
            decimal dHTanka = 0;
            int dMai = 0;
            int dNebiki = 0;

            try
            {
                // 内容のとき
                if (sender == txtChirashi)
                {
                    objtxt = txtChirashi;
                }

                // 単価または枚数のとき 2015/07/01
                if (sender == txtTanka || sender == txtMai || sender == txtNebiki)
                {
                    // 単価
                    if (sender == txtTanka)
                    {
                        objtxt = txtTanka;
                    }

                    // 枚数
                    if (sender == txtMai)
                    {
                        objtxt = txtMai;
                    }
                    
                    // 日付取得
                    //DateTime.TryParse(jDate.Value.ToShortDateString(), out dt);  // 2019/08/26 コメント化
                    DateTime.TryParse(dtSeikyu.Value.ToShortDateString(), out dt);  // 請求締日 2019/08/26

                    // 単価取得
                    str = Utility.strToDouble(txtTanka.Text).ToString();

                    if (!decimal.TryParse(str, out dTanka))
                    {
                        dTanka = 0m;
                    }

                    // 枚数取得
                    dMai = Utility.strToInt(txtMai.Text);

                    // 値引額取得
                    dNebiki = Utility.strToInt(txtNebiki.Text);

                    // 売上金額
                    Kingaku = Math.Floor(dTanka * (decimal)dMai);

                    // 値引後金額
                    decimal nebikigoKin = Kingaku - dNebiki;
                    txtNebikigo.Text = nebikigoKin.ToString("#,##0");

                    if (dtSeikyu.Checked)
                    {
                        // 金額計算
                        UriageSum(dt, nebikigoKin, out KingakuTax, out KingakuZeikomi, out KingakuTL);
                    }

                    // 売上金額
                    txtUri.Text = Kingaku.ToString("#,##0");

                    // 営業原価・粗利金額：2019/02/21
                    txteGaichuArari.Text = (Utility.strToInt(txtUri.Text) - Utility.strToInt(txteGaichuGenka.Text)).ToString("#,##0");

                    // 外注粗利：2019/02/21
                    int gai = Utility.strToInt(txtpGaichuGenka.Text) + Utility.strToInt(txtpGaichuGenka2.Text) + Utility.strToInt(txtpGaichuGenka3.Text);
                    int arari = Utility.strToInt(txtUri.Text) - gai;
                    txtaUri.Text = Kingaku.ToString("#,##0");
                    txtaArari.Text = arari.ToString("#,##0");

                    tabControl1.TabPages[3].Text = "粗利：" + arari.ToString("#,##0");

                    // 消費税額計算               
                    txtTax.Text = KingakuTax.ToString("#,##0");

                    // 税込金額
                    txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                    // 営業原価：ポスティングのとき 2018/01/03
                    if (cmbNaiyou.Text == J_NAIYOU_POSTING)
                    {
                        // 配布単価取得
                        str = Utility.strToDouble(txtHTanka.Text).ToString();

                        if (!decimal.TryParse(str, out dHTanka))
                        {
                            dHTanka = 0m;
                        }

                        decimal genka = Math.Floor(dHTanka * (decimal)dMai);
                        txteGaichuGenka.Text = genka.ToString("#,##0");
                    }
                }

                // 値引のとき 2015/07/01
                if (sender == txtNebiki)
                {
                    objtxt = txtNebiki;

                    //売上金額
                    //txtNebikigo.Text = (Utility.strToInt(txtZeikomi.Text) - Utility.strToInt(txtNebiki.Text)).ToString("#,##0");
                }

                // 配布単価のとき  2018/01/03
                if (sender == txtHTanka)
                {
                    objtxt = txtHTanka;

                    // 営業原価：ポスティングのとき
                    if (cmbNaiyou.Text == J_NAIYOU_POSTING)
                    {
                        // 配布単価取得
                        str = Utility.strToDouble(txtHTanka.Text).ToString();

                        if (!decimal.TryParse(str, out dHTanka))
                        {
                            dHTanka = 0m;
                        }

                        // 枚数取得
                        dMai = Utility.strToInt(txtMai.Text);

                        decimal genka = Math.Floor(dHTanka * (decimal)dMai);
                        txteGaichuGenka.Text = genka.ToString("#,##0");
                    }
                }

                //if (sender == txtIraisaki)
                //{
                //    objtxt = txtIraisaki;
                //}

                //if (sender == txtGenka)
                //{
                //    objtxt = txtGenka;
                //}

                //if (sender == txtEmail)    // 2015/06/23
                //{
                //    objtxt = txtEmail;
                //}

                if (sender == txtMemo)
                {
                    objtxt = txtMemo;
                }

                //if (sender == txtMemo2)
                //{
                //    objtxt = txtMemo2;
                //}

                if (sender == textBox1)
                {
                    objtxt = textBox1;
                }

                if (sender == txtHTanka)
                {
                    objtxt = txtHTanka;
                }

                if (sender == cmbClient)
                {
                    cmbClient.BackColor = Color.White;

                    if (cmbClient.SelectedIndex != -1)
                    {
                        //クライアント情報表示
                        Utility.ComboClient cmbC = new Utility.ComboClient();
                        cmbC = (Utility.ComboClient)cmbClient.SelectedItem;
                        ClientShow(cmbC.ID);
                    }
                }

                // 2015/07/01
                if (sender == cmbsNaiyou)
                {
                    cmbsNaiyou.BackColor = Color.White;
                }

                if (sender == cmbNaiyou)
                {
                    cmbNaiyou.BackColor = Color.White;
                }

                if (sender == cmbOffice)
                {
                    cmbOffice.BackColor = Color.White;
                }

                if (sender == cmbFkeitai)
                {
                    cmbFkeitai.BackColor = Color.White;
                }

                if (sender == cmbFjyouken)
                {
                    cmbFjyouken.BackColor = Color.White;
                }

                if (sender == cmbSize)
                {
                    cmbSize.BackColor = Color.White;
                }

                if (sender == cmbAnShu)
                {
                    cmbAnShu.BackColor = Color.White;
                }

                if (sender == cmbeGaichu)
                {
                    cmbeGaichu.BackColor = Color.White;
                }

                if (sender == cmbpGaichu)
                {
                    cmbpGaichu.BackColor = Color.White;
                }

                if (sender == txteGaichuGenka)
                {
                    objtxt = txteGaichuGenka;
                }

                if (sender == txtpGaichuGenka)
                {
                    objtxt = txtpGaichuGenka;
                }

                if (sender == txtsClient)
                {
                    objtxt = txtsClient;
                }

                if (sender == txtsGaichu)
                {
                    objtxt = txtsGaichu;
                }

                if (sender == txtsOrderNum)
                {
                    objtxt = txtsOrderNum;
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
            // 新請求書データとなったため以下コメント化 2015/12/02
            //請求書発行済みか？  2010/05/30
            //if (txtSeikyuNumber.Text.ToString() != "0")
            //{
            //    MessageBox.Show("既に請求書が発行済みです。(№" + txtSeikyuNumber.Text + ")" + Environment.NewLine + "受注データを削除するには請求書データを削除した後に実行してください", "請求書発行済み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}

            // 2015/12/02 請求書番号取得
            int seikyuNum = Utility.strToInt(txtSeikyuNumber.Text);

            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            
            string strSql;

            //データ削除
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();
            OleDbTransaction tran;

            cn = Con.GetConnection();

            //トランザクションの開始
            tran = cn.BeginTransaction();

            OleDbCommand SCom = new OleDbCommand();

            SCom.Connection = cn;
            SCom.Transaction = tran;

            try
            {
                //受注データの削除
                strSql = "";
                strSql += "delete from 受注 ";
                strSql += "where ID = " + cMaster.ID.ToString();

                SCom.CommandText = strSql;

                SCom.ExecuteNonQuery();

                //配布エリアデータの削除
                strSql = "";
                strSql += "delete from 配布エリア ";
                strSql += "where 配布エリア.受注ID = " + cMaster.ID.ToString();

                SCom.CommandText = strSql;

                SCom.ExecuteNonQuery();

                //コミット
                tran.Commit();

                MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // 受注データ読み込み
                jAdp.Fill(dts.受注1);
                
                // 請求書データ取得
                if (dts.受注1.Count(a => a.請求書ID == seikyuNum) > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("削除した受注データを含む請求書データがあります。").Append(Environment.NewLine);
                    sb.Append("請求締め処理を行って請求金額を再計算してください。");

                    MessageBox.Show(sb.ToString(), "請求書データの再作成が必要です", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                //ロールバック
                tran.Rollback();
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MessageBox.Show("削除に失敗しました。ロールバックしました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            cn.Close();
            Con.Close();

            DispClear();

            //データを 'darwinDataSet.受注' テーブルに読み込みます。
            //this.受注TableAdapter.Fill(this.darwinDataSet.受注);
            //dataGridView1.DataSource = this.darwinDataSet.受注;

            dataGridView1.DataSource = null;

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

        //////private void cmbJkbnSet()
        //////{
        //////    comboBox1.Items.Clear();
        //////    comboBox1.Items.Add("新規");
        //////    comboBox1.Items.Add("リピート");
        //////}

        private void cmbFjyoukenSet()
        {
            cmbFjyouken.Items.Clear();
            cmbFjyouken.Items.Add("予定枚数どおり");
            cmbFjyouken.Items.Add("詰めて配布 ＯＫ");
        }

        //private void cmbFyuyoSet()    // 2015/06/23
        //{
        //    cmbFyuyo.Items.Clear();
        //    cmbFyuyo.Items.Add("厳守");
        //    cmbFyuyo.Items.Add("前後ＯＫ");
        //}

        //private void cmbNkeitaiSet()  // 2015/06/23
        //{
        //    cmbNkeitai.Items.Clear();
        //    cmbNkeitai.Items.Add("宅急便");
        //    cmbNkeitai.Items.Add("持込");
        //    cmbNkeitai.Items.Add("集荷：バック");
        //    cmbNkeitai.Items.Add("集荷：営業");
        //}

        //private void cmbHjikiSet()    // 2015/06/23
        //{
        //    cmbHjiki.Items.Clear();
        //    cmbHjiki.Items.Add("デイリー");
        //    cmbHjiki.Items.Add("週単位");
        //    cmbHjiki.Items.Add("終了後");
        //}

        //private void cmbHseidoSet()   // 2015/06/23
        //{
        //    cmbHseido.Items.Clear();
        //    cmbHseido.Items.Add("実配布数");
        //    cmbHseido.Items.Add("予定枚数");
        //}

        //private void cmbHhouhouSet()  // 2015/06/23
        //{
        //    cmbHhouhou.Items.Clear();
        //    cmbHhouhou.Items.Add("FAX");
        //    cmbHhouhou.Items.Add("メール");
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            //darwinDataSet ds = new darwinDataSet();
            //ds.Clear();
            //ds.EnforceConstraints = false;
            //this.受注TableAdapter.FillByName(ds.受注, "%" + textBox1.Text.ToString() + "%");
            //dataGridView1.DataSource = ds.受注;

            dataSerach();
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     受注データ検索 </summary>
        /// ------------------------------------------------------------------
        private void dataSerach()
        {
            this.Cursor = Cursors.WaitCursor;
            
            darwinDataSet dts = new darwinDataSet();
            dts.EnforceConstraints = false;

            darwinDataSetTableAdapters.受注TableAdapter adp = new darwinDataSetTableAdapters.受注TableAdapter();
            darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();
            darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
            darwinDataSetTableAdapters.社員TableAdapter eAdp = new darwinDataSetTableAdapters.社員TableAdapter();

            //adp.Fill(dts.受注);

            DateTime sd = DateTime.Parse(sDt.Value.ToShortDateString()); // 2018/01/27
            DateTime ed = DateTime.Parse(eDt.Value.ToShortDateString()); // 2018/01/27
            adp.FillByFromYMDToYMD(dts.受注, sd, ed); // 2018/01/27

            gAdp.Fill(dts.外注先);
            cAdp.Fill(dts.得意先);
            eAdp.Fill(dts.社員);

            var s = dts.受注.Where(a => a.ID >= 0);

            // ログインユーザーの受注データ保守制限
            if (global.loginOrderMntType == 0)
            {
                s = s.Where(a => a.登録ユーザーID == global.loginUserID);
            }

            // 受注日
            if (sDt.Checked)
            {
                s = s.Where(a => a.受注日 >= DateTime.Parse(sDt.Value.ToShortDateString()));
            }

            if (eDt.Checked)
            {
                s = s.Where(a => a.受注日 <= DateTime.Parse(eDt.Value.ToShortDateString()));
            }

            // チラシ名
            if (textBox1.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.チラシ名.Contains(textBox1.Text.Trim()));
            }

            // 受注番号
            if (txtsOrderNum.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.ID == Utility.strToLong(txtsOrderNum.Text));
            }

            // 受注内容（受注種別ID）
            if (cmbsNaiyou.SelectedIndex != -1)
            {
                Utility.ComboJshubetsu cmbJ = (Utility.ComboJshubetsu)cmbsNaiyou.SelectedItem;
                s = s.Where(a => a.受注種別ID == cmbJ.ID); 
            }

            // 外注先
            if (txtsGaichu.Text != string.Empty)
            {
                s = s.Where(a => a.外注先Row != null && a.外注先Row.名称.Contains(txtsGaichu.Text));
            }

            // 得意先
            if (txtsClient.Text != string.Empty)
            {
                s = s.Where(a => (!a.Is略称Null() && a.略称.Contains(txtsClient.Text)) || (!a.IsフリガナNull() && a.フリガナ.Contains(txtsClient.Text)));
            }

            // 営業担当者
            if (txtsTantou.Text != string.Empty)
            {
                s = s.Where(a => (!a.Is得意先IDNull() && a.得意先Row.社員Row != null && a.得意先Row.社員Row.氏名.Contains(txtsTantou.Text)));
            }
            
            //サイズ(判型)ID取得 2019/02/14
            if (cmbsSize.SelectedIndex != -1)
            {
                Utility.ComboSize cmbs = new Utility.ComboSize();
                cmbs = (Utility.ComboSize)cmbsSize.SelectedItem;
                int sSize = cmbs.ID;

                s = s.Where(a => a.判型 == sSize);
            }


            // 枚数 2019/02/14
            if (txtsMaisu.Text != string.Empty)
            {
                int sMai = Utility.strToInt(txtsMaisu.Text);
                if (sMai != global.FLGOFF)
                {
                    s = s.Where(a => a.枚数 == sMai);
                }
            }

            // 請求締日 2019/02/14
            if (seDt.Checked)
            {
                s = s.Where(a => a.請求書発行日 == DateTime.Parse(seDt.Value.ToShortDateString()));
            }


            // データグリッドビューにデータソースを設定
            dataGridView1.DataSource = s.ToList();

            if (dataGridView1.RowCount > 0)
            {
                dataGridView1.CurrentCell = null;
            }

            this.Cursor = Cursors.Default;
        }
        
        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmOffice();

            frm.ShowDialog();
            Utility.ComboOffice.load(cmbOffice);
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     消費税率取得 </summary>
        /// <param name="tempDate">
        ///     基準日付</param>
        /// <returns>
        ///     税率</returns>
        ///  
        /// 2015/06/24
        ///------------------------------------------------------------------
        private int GetTaxRT(DateTime tempDate)
        {
            //税率取得
            int Ritsu = 0;

            darwinDataSet dts = new posting.darwinDataSet();
            darwinDataSetTableAdapters.税率TableAdapter adp = new darwinDataSetTableAdapters.税率TableAdapter();
            adp.Fill(dts.税率);

            foreach (var t in dts.税率.Where(a => a.開始年月日 <= tempDate).OrderByDescending(a => a.開始年月日))
            {
                Ritsu = t.税率;
                break;
            }

            return Ritsu;            
        }

        /// -------------------------------------------------------------
        /// <summary>
        ///     消費税計算 </summary>
        /// <param name="tempKin">
        ///     対象金額</param>
        /// <param name="tempTax">
        ///     税率</param>
        /// <returns>
        ///     消費税額</returns>
        /// -------------------------------------------------------------
        private decimal GetTax(decimal tempKin, int tempTax)
        {
            decimal GakuD;
            //int GakuI;

            // 端数切捨て 2015/07/01
            GakuD = Math.Floor(tempKin * tempTax / 100);

            //GakuD += (decimal)0.5;
            //GakuI = (int)GakuD;

            return GakuD;
        }

        /// ---------------------------------------------------------------------------------------
        /// <summary>
        ///     売上金額を計算する </summary>
        /// <param name="tempDate">
        ///     売上日付（請求締日とする 2019/08/26）</param>
        /// <param name="nebikigo">
        ///     値引後金額 </param>
        /// <param name="KingakuTax">
        ///     消費税</param>
        /// <param name="KingakuZeikomi">
        ///     税込売上</param>
        /// <param name="KingakuTL">
        ///     売上金額</param>
        /// -----------------------------------------------------------------------------------------
        private void UriageSum(DateTime tempDate, decimal nebikigo, 
                                out decimal KingakuTax, out decimal KingakuZeikomi, out decimal KingakuTL)
        {
            //税率再取得
            cTax.Ritsu = Utility.GetTaxRT(tempDate);

            //消費税額計算
            KingakuTax = Utility.GetTax(nebikigo, cTax.Ritsu);

            //税込金額
            KingakuZeikomi = (int)nebikigo + KingakuTax;

            //売上金額
            KingakuTL = KingakuZeikomi;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     クライアント情報表示 </summary>
        /// <param name="tempID">
        ///     得意先ID</param>
        /// -------------------------------------------------------------------------
        private void ClientShow(int tempID)
        {
            foreach (var t in dts.得意先.Where(a => a.ID == tempID))
            {
                txtCzipcode.Text = t.郵便番号;
                txtName2.Text = (t.住所1 + " " + t.住所2).Trim();

                // 請求先情報
                txtCbusho.Text = t.部署名;
                //txtCtantou.Text = t.担当者名;
                txtCtel.Text = t.電話番号; 
                txtCfax.Text = t.FAX番号;

                txtCtantou.Text = t.請求先担当者名;
                textBox4.Text = t.請求先名称;

                if (t.社員Row != null)
                {
                    txtTantou.Text = t.社員Row.氏名;
                }
                else
                {
                    txtTantou.Text = string.Empty;
                }

                // 締日 : 2018/01/03
                lblClientShimebi.Text = t.締日.ToString() + "日";
            }
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

        private void button2_Click(object sender, EventArgs e)
        {
            OrderReport();
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     発注書印刷 </summary>
        ///--------------------------------------------------------------------
        private void OrderReport()
        {
            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル発注書シート名, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {

                    if (cmbClient.SelectedIndex != -1)
                    {
                        Utility.ComboClient cmbC = new Utility.ComboClient();
                        cmbC = (Utility.ComboClient)cmbClient.SelectedItem;
                        oxlsSheet.Cells[5, 2] = cmbC.NameShow; //クライアント名
                    }
                    else
                    {
                        oxlsSheet.Cells[5, 2] = "";
                    }

                    oxlsSheet.Cells[6, 2] = txtName2.Text;  //住所
                    oxlsSheet.Cells[9, 3] = txtCtel.Text;   //電話番号
                    oxlsSheet.Cells[9, 6] = txtCfax.Text;   //FAX番号

                    if (StartDate.Checked == true)
                    {
                        oxlsSheet.Cells[16, 1] = StartDate.Value.Month.ToString() + "月" + StartDate.Value.Day.ToString() + "日～"; //配布日
                    }
                    else
                    {
                        oxlsSheet.Cells[16, 1] = "";
                    }

                    oxlsSheet.Cells[16, 2] = cmbSize.Text;  //サイズ
                    oxlsSheet.Cells[16, 3] = cmbFkeitai.Text;   //配布形態
                    oxlsSheet.Cells[16, 5] = int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);   //配布枚数
                    oxlsSheet.Cells[16, 6] = double.Parse(txtTanka.Text.ToString(), System.Globalization.NumberStyles.Any);           //単価
                    oxlsSheet.Cells[22, 6] = Utility.strToInt(txtNebiki.Text) * (-1);   // 値引 
                    oxlsSheet.Cells[23, 6] = Utility.strToInt(txtTax.Text);   // 消費税 

                    oxlsSheet.Cells[39, 6] = txtTantou.Text;    //営業担当者

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
                    saveFileDialog1.Title = "発注書保存";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = "発注書_" + cmbClient.Text + " " + jDate.Value.ToLongDateString();
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName,Type.Missing,Type.Missing,
                                        Type.Missing,Type.Missing,Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange,Type.Missing,
                                        Type.Missing,Type.Missing,Type.Missing,Type.Missing);
                    }

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "発注書", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                MessageBox.Show(e.Message, "発注書", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            NyukoReport();
        }

        private void NyukoReport()
        {
            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル入庫管理表シート名, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    oxlsSheet.Cells[2, 2] = txtChirashi.Text; //ちらし名
                    oxlsSheet.Cells[3, 2] = int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);   //配布枚数

                    if (StartDate.Checked == true)
                    {
                        oxlsSheet.Cells[4, 2] = StartDate.Value.Month.ToString() + "/" + StartDate.Value.Day.ToString(); //配布開始日
                    }
                    else
                    {
                        oxlsSheet.Cells[4, 2] = "";
                    }

                    if (EndDate.Checked == true)
                    {
                        oxlsSheet.Cells[5, 2] = EndDate.Value.Month.ToString() + "/" + EndDate.Value.Day.ToString(); //配布終了日
                    }
                    else
                    {
                        oxlsSheet.Cells[5, 2] = "";
                    }

                    if (NouhinDate.Checked == true)
                    {
                        oxlsSheet.Cells[6, 2] = NouhinDate.Value.Month.ToString() + "/" + NouhinDate.Value.Day.ToString(); //納品予定日
                    }
                    else
                    {
                        oxlsSheet.Cells[6, 2] = "";
                    }

                    oxlsSheet.Cells[7, 2] = txtTantou.Text;         // 営業担当者
                    //oxlsSheet.Cells[8, 2] = cmbHseido.Text;       // 報告精度 2015/06/23 撤廃
                    oxlsSheet.Cells[8, 2] = cmbOffice.Text;         // 事業所 
                    oxlsSheet.Cells[9, 2] = cMaster.ID.ToString();  // 受注ID 2015/06/23 

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
                    saveFileDialog1.Title = "入庫管理表保存";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = "入庫管理表_" + txtChirashi.Text + " " + jDate.Value.ToLongDateString();
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

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "入庫管理表", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                MessageBox.Show(e.Message, "入庫管理表", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void label3_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            frmClient frm = new frmClient();
            frm.ShowDialog();

            //得意先コンボ再ロード : 2015/06/24
            //Utility.ComboClient.load(cmbClient);
            int cIdx = cmbClient.SelectedIndex;
            Utility.ComboClient.itemsLoad(cmbClient);
            cmbClient.SelectedIndex = cIdx;

            // 得意先データ読み込み
            cAdp.Fill(dts.得意先);
        }

        private void cmbSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSize.SelectedIndex != -1)
            {
                Utility.ComboSize Cmbs = new Utility.ComboSize();
                Cmbs = (Utility.ComboSize)cmbSize.SelectedItem;

                txtHTanka.Text = Cmbs.Tanka.ToString("#,##0.0");
            }
        }

        private void cmbNaiyou_SelectedIndexChanged(object sender, EventArgs e)
        {
            //ポスティングかそれ以外かで入力項目を制御する 2010/1/27
            if (cmbNaiyou.SelectedIndex != -1)
            {
                // ポスティング以外のとき
                if (cmbNaiyou.SelectedIndex != 0)
                {
                    this.cmbFkeitai.SelectedIndex = -1;
                    this.cmbFkeitai.Enabled = false;

                    this.cmbFjyouken.SelectedIndex = -1;
                    this.cmbFjyouken.Enabled = false;

                    this.cmbSize.SelectedIndex = -1;

                    // 2015/11/10 印刷、デザインもサイズ入力可能とする
                    // 2016/01/19 新聞折込もサイズ入力可能とする
                    if (cmbNaiyou.SelectedIndex == 1 || cmbNaiyou.SelectedIndex == 3 || cmbNaiyou.Text == "新聞折込")
                    {
                        this.cmbSize.Enabled = true;
                    }
                    else
                    {
                        this.cmbSize.Enabled = false;
                    }

                    this.txtHTanka.Text = "0";
                    this.txtHTanka.Enabled = false;

                    //this.StartDate.Checked = false;
                    //this.StartDate.Enabled = false;

                    //this.EndDate.Checked = false;
                    //this.EndDate.Enabled = false;

                    // 2015/06/23
                    //this.cmbFyuyo.SelectedIndex = -1;
                    //this.cmbFyuyo.Enabled = false;

                    // 2019/02/14 全て入力可能に変更 コメント化
                    //this.NouhinDate.Checked = false;
                    //this.NouhinDate.Enabled = false;

                    // 2015/06/23
                    //this.cmbNkeitai.SelectedIndex = -1;
                    //this.cmbNkeitai.Enabled = false;

                    // 2015/06/23
                    //this.cmbHjiki.SelectedIndex = -1;
                    //this.cmbHjiki.Enabled = false;

                    // 2015/06/23
                    //this.cmbHseido.SelectedIndex = -1;
                    //this.cmbHseido.Enabled = false;

                    // 2015/06/23
                    //this.cmbHhouhou.SelectedIndex = -1;
                    //this.cmbHhouhou.Enabled = false;

                    // 2015/06/23
                    //this.txtEmail.Text = "";
                    //this.txtEmail.Enabled = false;

                    //未配布情報要チェック欄
                    this.checkBox2.Checked = false;
                    this.checkBox2.Enabled = false;

                    //枝番要チェック欄
                    this.checkBox3.Checked = false;
                    this.checkBox3.Enabled = false;

                    // 案件種別 2015/06/30
                    if (cmbAnShu.Items.Count > 0)
                    {
                        cmbAnShu.SelectedIndex = 2;
                    }

                    // 併配除外入力不可とする　2015/11/27
                    chkHeihai.Enabled = false;
                }
                else
                {
                    this.cmbFkeitai.SelectedIndex = -1;
                    this.cmbFkeitai.Enabled = true;

                    this.cmbFjyouken.SelectedIndex = -1;
                    this.cmbFjyouken.Enabled = true;

                    this.cmbSize.SelectedIndex = -1;
                    this.cmbSize.Enabled = true;

                    this.txtHTanka.Text = "0";
                    this.txtHTanka.Enabled = true;

                    this.StartDate.Checked = false;
                    this.StartDate.Enabled = true;

                    this.EndDate.Checked = false;
                    this.EndDate.Enabled = true;

                    // 2015/06/23
                    //this.cmbFyuyo.SelectedIndex = -1;
                    //this.cmbFyuyo.Enabled = true;

                    // 2019/02/14 全て入力可能に変更 コメント化
                    //this.NouhinDate.Checked = false;
                    //this.NouhinDate.Enabled = true;

                    // 2015/06/23
                    //this.cmbNkeitai.SelectedIndex = -1;
                    //this.cmbNkeitai.Enabled = true;

                    // 2015/06/23
                    //this.cmbHjiki.SelectedIndex = -1;
                    //this.cmbHjiki.Enabled = true;

                    // 2015/06/23
                    //this.cmbHseido.SelectedIndex = -1;
                    //this.cmbHseido.Enabled = true;

                    // 2015/06/23
                    //this.cmbHhouhou.SelectedIndex = -1;
                    //this.cmbHhouhou.Enabled = true;

                    // 2015/06/23
                    //this.txtEmail.Text = "";
                    //this.txtEmail.Enabled = true;

                    //未配布情報要チェック欄
                    this.checkBox2.Checked = false;
                    this.checkBox2.Enabled = true;

                    //枝番要チェック欄
                    this.checkBox3.Checked = false;
                    this.checkBox3.Enabled = true;

                    // 案件種別 2015/06/30
                    if (cmbAnShu.Items.Count > 0)
                    {
                        cmbAnShu.SelectedIndex = 0;
                    }

                    // 併配除外入力可とする　2015/11/27
                    chkHeihai.Enabled = true;
                }
            }
        }

        private void cmbeGaichu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbpGaichu.Items.Count > 0)
            {
                if (fMode.Mode == 0)
                {
                    //cmbeGaichu.SelectedIndex = cmbeGaichu.SelectedIndex;
                    cmbpGaichu.SelectedIndex = cmbeGaichu.SelectedIndex;
                }
            }
        }

        private void dteGaichuPay_ValueChanged(object sender, EventArgs e)
        {
            if (fMode.Mode == 0)
            {
                dtpGaichuPay.Value = dteGaichuPay.Value;
            }
        }

        private void txteGaichuGenka_TextChanged(object sender, EventArgs e)
        {
            if (fMode.Mode == 0)
            {
                txtpGaichuGenka.Text = txteGaichuGenka.Text;
            }

            // 営業原価と外注先原価はイコールとする（外注先原価２、３がないとき） 2019/02/21
            if ((Utility.strToInt(txtpGaichuGenka2.Text) == global.FLGOFF) &&
                (Utility.strToInt(txtpGaichuGenka3.Text) == global.FLGOFF))
            {
                txtpGaichuGenka.Text = txteGaichuGenka.Text;
            }

            // 粗利金額：2019/02/21
            txteGaichuArari.Text = (Utility.strToInt(txtUri.Text) - Utility.strToInt(txteGaichuGenka.Text)).ToString("#,##0");

        }

        private void txtTanka_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void txtMai_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void rBtnOrderNum1_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnOrderNum1.Checked)
            {
                btnOrderNum.Enabled = false;
                lblOrderNum.Enabled = false;
            }
            else if (rBtnOrderNum2.Checked)
            {
                btnOrderNum.Enabled = true;
                lblOrderNum.Enabled = true;
            }
        }

        private void btnOrderNum_Click(object sender, EventArgs e)
        {
            frmGetOrderNumber frm = new frmGetOrderNumber();
            frm.ShowDialog();

            if (frm._orderNum != string.Empty)
            {
                lblOrderNum.Text = frm._orderNum;
            }
            frm.Dispose();
        }

        private void label13_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // 外注先マスター保守
            frmGaichu frm = new frmGaichu();
            frm.ShowDialog();

            // 外注先コンボボックスアイテムロード
            int idx = cmbeGaichu.SelectedIndex;
            Utility.comboGaichu.itemLoad(cmbeGaichu);
            cmbeGaichu.SelectedIndex = idx;
        }

        private void label36_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // 外注先マスター保守
            frmGaichu frm = new frmGaichu();
            frm.ShowDialog();

            // 外注先コンボボックスアイテムロード
            int idx = cmbpGaichu.SelectedIndex;
            Utility.comboGaichu.itemLoad(cmbpGaichu);
            cmbpGaichu.SelectedIndex = idx;

            //// 検索用外注先コンボボックスアイテムロード
            //Utility.comboGaichu.itemLoad(cmbsGaichu);
        }

        private void label50_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void dtGaichuIrai_EnabledChanged(object sender, EventArgs e)
        {

        }

        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.受注TableAdapter.FillBy(this.darwinDataSet.受注);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillBy1ToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.受注TableAdapter.FillBy1(this.darwinDataSet.受注);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void cmbpGaichu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu.SelectedIndex >= 0)
                {
                    lblSD1.Text = cg[cmbpGaichu.SelectedIndex].shiharaibi.ToString() + "日";
                }
                else
                {
                    lblSD1.Text = string.Empty;
                }

                txtaGai1.Text = cmbpGaichu.Text;
            }
        }

        private void cmbpGaichu2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu2.SelectedIndex >= 0)
                {
                    lblSD2.Text = cg[cmbpGaichu2.SelectedIndex].shiharaibi.ToString() + "日";
                }
                else
                {
                    lblSD2.Text = string.Empty;
                }

                txtaGai2.Text = cmbpGaichu2.Text;
            }
        }

        private void cmbpGaichu3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu3.SelectedIndex >= 0)
                {
                    lblSD3.Text = cg[cmbpGaichu3.SelectedIndex].shiharaibi.ToString() + "日";
                }
                else
                {
                    lblSD3.Text = string.Empty;
                }

                txtaGai3.Text = cmbpGaichu3.Text;
            }
        }

        private void cmbpGaichu3_TextChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu3.Text == string.Empty)
                {
                    lblSD3.Text = string.Empty;
                }
            }
        }

        private void cmbpGaichu_TextChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu.Text == string.Empty)
                {
                    lblSD1.Text = string.Empty;
                }
            }
        }

        private void cmbpGaichu2_TextChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu2.Text == string.Empty)
                {
                    lblSD2.Text = string.Empty;
                }
            }
        }

        private void lnkCopy_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmOrderCopy frm = new frmOrderCopy(dts, jAdp, fMode.jID);
            frm.ShowDialog();

            bool fs = frm.fStatus;
            frm.Dispose();

            DispClear();

            if (frm.fStatus)
            {
                Cursor = Cursors.WaitCursor;
                jAdp.Update(dts.受注1);
                dataGridView1.DataSource = null;
                jAdp.Fill(dts.受注1);
                Cursor = Cursors.Default;
            }
        }

        private void txtpGaichuGenka3_TextChanged(object sender, EventArgs e)
        {
            // 売上金額
            txtaUri.Text = txtUri.Text;

            // 外注粗利：2019/02/21
            int gai = Utility.strToInt(txtpGaichuGenka.Text) + Utility.strToInt(txtpGaichuGenka2.Text) + Utility.strToInt(txtpGaichuGenka3.Text);
            int arari = Utility.strToInt(txtUri.Text) - gai;
            txtaArari.Text = arari.ToString("#,##0");

            tabControl1.TabPages[3].Text = "粗利：" + arari.ToString("#,##0");

            // 外注原価
            txtaGaiGenka1.Text = txtpGaichuGenka.Text;
            txtaGaiGenka2.Text = txtpGaichuGenka2.Text;
            txtaGaiGenka3.Text = txtpGaichuGenka3.Text;
        }

        private void txtMai_TextChanged(object sender, EventArgs e)
        {
            // 営業原価：ポスティングのとき 2019/03/16
            if (cmbNaiyou.Text == J_NAIYOU_POSTING)
            {
                txteGaichuGenka.Text = getGenka().ToString("#,##0");
            }
        }

        private void txtHTanka_TextChanged(object sender, EventArgs e)
        {
            // 営業原価：ポスティングのとき 2019/03/16
            if (cmbNaiyou.Text == J_NAIYOU_POSTING)
            {
                txteGaichuGenka.Text = getGenka().ToString("#,##0");
            }
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     営業原価を求める：2019/03/16 </summary>
        /// <returns>
        ///     営業原価</returns>
        ///-------------------------------------------------------
        private decimal getGenka()
        {
            // 配布単価取得
            String str = Utility.strToDouble(txtHTanka.Text).ToString();

            decimal dHTanka = 0;

            if (!decimal.TryParse(str, out dHTanka))
            {
                dHTanka = 0m;
            }

            decimal genka = Math.Floor(dHTanka * Utility.strToDecimal(txtMai.Text.Replace(",","")));

            return genka;
        }

        private void dtSeikyu_ValueChanged(object sender, EventArgs e)
        {
            //
            //  消費税算定基準日を請求締日とした処理
            //      日付の変更に伴い消費税額、税込み金額を自動変更する
            //      2019-08-26
            //

            DateTime dt;

            if (dtSeikyu.Checked)
            {
                if (!DateTime.TryParse(dtSeikyu.Value.ToShortDateString(), out dt))
                {
                    return;
                }

                decimal nebikigoKin = Utility.strToDecimal(txtNebikigo.Text);
                decimal KingakuTax = 0;
                decimal KingakuZeikomi = 0;
                decimal KingakuTL = 0;

                // 金額計算
                UriageSum(dt, nebikigoKin, out KingakuTax, out KingakuZeikomi, out KingakuTL);

                // 消費税額計算               
                txtTax.Text = KingakuTax.ToString("#,##0");

                // 税込金額
                txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");
            }
            else
            {
                // 消費税額計算               
                txtTax.Text = global.FLGOFF.ToString();

                // 税込金額
                txtZeikomi.Text = Utility.strToDecimal(txtNebikigo.Text).ToString("#,##0");
            }
        }
    }
}
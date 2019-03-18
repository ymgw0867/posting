using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MyLibrary;

namespace posting
{
    public partial class frmOrderExcel : Form
    {
        public frmOrderExcel()
        {
            InitializeComponent();

            // データ読み込み
            cAdp.Fill(dts.得意先);
            eAdp.Fill(dts.社員);
            gAdp.Fill(dts.外注先);
            hAdp.Fill(dts.判型);
            rAdp.Fill(dts.受注確定書発行記録);
            pAdp.Fill(dts.会社情報);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter eAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();
        darwinDataSetTableAdapters.判型TableAdapter hAdp = new darwinDataSetTableAdapters.判型TableAdapter();
        darwinDataSetTableAdapters.受注確定書発行記録TableAdapter rAdp = new darwinDataSetTableAdapters.受注確定書発行記録TableAdapter();
        darwinDataSetTableAdapters.会社情報TableAdapter pAdp = new darwinDataSetTableAdapters.会社情報TableAdapter();

        #region グリッドビューカラム定義
        // 受注案件
        string colID = "col1";
        string colDt = "col2";
        string colClient = "col3";
        string colName = "col4";
        string colKikan = "col5";
        string colTantou = "col6";
        string colSel = "col7";
        string colClientCode = "col23";

        // 選択案件
        string colCmb = "col8";
        string colDtHaifuS = "col9";
        string colDtHaifuE = "col10";
        string colSize = "col11";
        string colCmbOri = "col12";
        string colTanka = "col13";
        string colMaisu = "col14";
        string colKingaku = "col15";
        string colGaichu = "col16";
        string colDtShiharai = "col17";
        string colGenka = "col18";
        string colArari = "col19";
        string colArariRT = "col20";
        string colNaiyou = "col21";
        string colDel = "col22";
        string colGyoushu = "col23";
        string colGaichu2 = "col24";
        string colDtShiharai2 = "col25";
        string colGenka2 = "col26";
        string colArari2 = "col27";
        string colArariRT2 = "col28";
        string colNebiki = "col29";
        string colGaichu3 = "col30";
        string colDtShiharai3 = "col31";
        string colGenka3 = "col32";
        string colGaichu4 = "col33";
        string colDtShiharai4 = "col34";
        string colGenka4 = "col35";

        #endregion

        // 受注確定書入力シートパス：2019/03/06
        string sheetPath = string.Empty;

        /// -------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// -------------------------------------------------------------------
        private void gridSetting(DataGridView dgv)
        {
            try
            {
                dgv.EnableHeadersVisualStyles = false;
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // 列スタイルを変更する

                dgv.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                dgv.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                dgv.ColumnHeadersHeight = 20;
                dgv.RowTemplate.Height = 20;

                // 全体の高さ
                //dgv.Height = 221;

                // 奇数行の色
                dgv.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "選択";
                dgv.Columns.Add(btn);
                dgv.Columns[0].HeaderText = "";
                dgv.Columns[0].Name = colSel;                

                dgv.Columns.Add(colID, "受注番号");
                dgv.Columns.Add(colDt, "受注日");
                dgv.Columns.Add(colClient, "クライアント");
                dgv.Columns.Add(colName, "内容");
                dgv.Columns.Add(colKikan, "配布期間");
                dgv.Columns.Add(colTantou, "担当");
                dgv.Columns.Add(colClientCode, "");
                dgv.Columns[colClientCode].Visible = false;
                
                dgv.Columns[colSel].Width = 50;
                dgv.Columns[colID].Width = 120;
                dgv.Columns[colDt].Width = 120;
                dgv.Columns[colClient].Width = 300;
                dgv.Columns[colName].Width = 300;
                dgv.Columns[colKikan].Width = 160;
                dgv.Columns[colTantou].Width = 100;

                // 行ヘッダを表示しない
                dgv.RowHeadersVisible = false;

                // 選択モード
                //dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgv.MultiSelect = false;

                // 編集不可とする
                dgv.ReadOnly = true;

                // 追加行表示しない
                dgv.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                dgv.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                dgv.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                dgv.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                dgv.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 罫線
                dgv.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                //tempDGV.BorderStyle = BorderStyle.None;
                dgv.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// -------------------------------------------------------------------
        private void gridSetting2(DataGridView dgv)
        {
            try
            {
                dgv.EnableHeadersVisualStyles = false;
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // 列スタイルを変更する

                dgv.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                dgv.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                dgv.ColumnHeadersHeight = 20;
                dgv.RowTemplate.Height = 20;

                // 全体の高さ
                //dgv.Height = 221;

                // 奇数行の色
                dgv.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列定義
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "削除";
                dgv.Columns.Add(btn);
                dgv.Columns[0].HeaderText = "";
                dgv.Columns[0].Name = colDel;

                DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
                cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                dgv.Columns.Add(cmb);
                dgv.Columns[1].HeaderText = "受注内容";
                dgv.Columns[1].Name = colCmb;

                cmb.Items.Add("ポスティング");
                Utility.ComboJshubetsu.load(cmb);

                dgv.Columns.Add(colNaiyou, "内容");
                dgv.Columns.Add(colGyoushu, "業種");
                dgv.Columns.Add(colDtHaifuS, "配布期間");
                dgv.Columns.Add(colSize, "サイズ");

                DataGridViewComboBoxColumn cmb2 = new DataGridViewComboBoxColumn();
                cmb2.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                dgv.Columns.Add(cmb2);
                dgv.Columns[6].HeaderText = "折";
                dgv.Columns[6].Name = colCmbOri;
                Utility.comboOri.load(cmb2);

                dgv.Columns.Add(colTanka, "単価");
                dgv.Columns.Add(colMaisu, "枚数");
                dgv.Columns.Add(colKingaku, "金額");
                dgv.Columns.Add(colNebiki, "値引");    // 2015/12/10
                dgv.Columns.Add(colGaichu, "外注先１");
                dgv.Columns.Add(colDtShiharai, "支払日");
                dgv.Columns.Add(colGenka, "原価");
                dgv.Columns.Add(colArari, "粗利益");
                dgv.Columns.Add(colArariRT, "利益率");
                dgv.Columns.Add(colGaichu2, "外注先２");    // 2015/12/10
                dgv.Columns.Add(colDtShiharai2, "支払日");    // 2015/12/10
                dgv.Columns.Add(colGenka2, "原価");           // 2015/12/10
                dgv.Columns.Add(colGaichu3, "外注先３");      // 2016/10/26
                dgv.Columns.Add(colDtShiharai3, "支払日");    // 2016/10/26
                dgv.Columns.Add(colGenka3, "原価");    // 2016/10/26
                dgv.Columns.Add(colGaichu4, "外注先４");    // 2016/10/26
                dgv.Columns.Add(colDtShiharai4, "支払日");    // 2016/10/26
                dgv.Columns.Add(colGenka4, "原価");    // 2016/10/26
                dgv.Columns.Add(colArari2, "粗利益");    // 2015/12/10
                dgv.Columns.Add(colArariRT2, "利益率");    // 2015/12/10
                dgv.Columns.Add(colID, "受注番号");
                dgv.Columns[colID].Visible = false;

                // 各列幅指定
                dgv.Columns[colDel].Width = 50;
                dgv.Columns[colCmb].Width = 100;
                dgv.Columns[colNaiyou].Width = 220;
                dgv.Columns[colDtHaifuS].Width = 180;
                dgv.Columns[colSize].Width = 80;
                dgv.Columns[colCmbOri].Width = 50;
                dgv.Columns[colTanka].Width = 70;
                dgv.Columns[colMaisu].Width = 70;
                dgv.Columns[colKingaku].Width = 100;
                dgv.Columns[colNebiki].Width = 100;    // 2015/12/10
                dgv.Columns[colGaichu].Width = 180;
                dgv.Columns[colDtShiharai].Width = 110;
                dgv.Columns[colGenka].Width = 70;
                dgv.Columns[colArari].Width = 100;
                dgv.Columns[colArariRT].Width = 100;

                dgv.Columns[colGaichu2].Width = 180;    // 2015/12/10
                dgv.Columns[colDtShiharai2].Width = 110;    // 2015/12/10
                dgv.Columns[colGenka2].Width = 70;    // 2015/12/10

                dgv.Columns[colGaichu3].Width = 180;    // 2016/10/26
                dgv.Columns[colDtShiharai3].Width = 110;    // 2016/10/26
                dgv.Columns[colGenka3].Width = 70;    // 20116/10/26

                dgv.Columns[colGaichu4].Width = 180;    // 2016/10/26
                dgv.Columns[colDtShiharai4].Width = 110;    // 2016/10/26
                dgv.Columns[colGenka4].Width = 70;    // 2016/10/26

                dgv.Columns[colArari2].Width = 100;    // 2015/12/10
                dgv.Columns[colArariRT2].Width = 100;    // 2015/12/10

                dgv.Columns[colDtHaifuS].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dgv.Columns[colSize].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dgv.Columns[colCmbOri].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dgv.Columns[colTanka].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                dgv.Columns[colMaisu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                dgv.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                dgv.Columns[colNebiki].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;    // 2015/12/10
                dgv.Columns[colDtShiharai].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dgv.Columns[colGenka].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                dgv.Columns[colArari].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                dgv.Columns[colArariRT].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                dgv.Columns[colDtShiharai2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;    // 2015/12/10
                dgv.Columns[colGenka2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;    // 2015/12/10

                dgv.Columns[colDtShiharai3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;    // 2016/10/26
                dgv.Columns[colGenka3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;    // 2016/10/26

                dgv.Columns[colDtShiharai4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;    // 2016/10/26
                dgv.Columns[colGenka4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;    // 2016/10/26

                dgv.Columns[colArari2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;    // 2015/12/10
                dgv.Columns[colArariRT2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;    // 2015/12/10
                
                // 行ヘッダを表示しない
                dgv.RowHeadersVisible = false;

                // 選択モード
                //dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgv.MultiSelect = false;

                // 編集可とする
                dgv.ReadOnly = false;
                //dgv.EditMode = DataGridViewEditMode.EditProgrammatically;
                dgv.EditMode = DataGridViewEditMode.EditOnEnter;

                // 追加行表示しない
                dgv.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                dgv.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                dgv.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                dgv.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                dgv.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 罫線
                dgv.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                //tempDGV.BorderStyle = BorderStyle.None;
                dgv.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmOrderExcel_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void frmOrderExcel_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // データグリッドビューの定義
            gridSetting(dataGridView1);
            gridSetting2(dataGridView2);

            // 配布形態
            Utility.ComboFkeitai.load(cmbHaifukeitai);

            // 配布条件
            cmbFjyoukenSet();            
            
            // 納品場所コンボ設定
            cmbNouhinBashoSet();
            
            // 事前請求書コンボ設定
            //cmbJizenSeikyushoSet();   // 2015/11/27 コメント化

            //// データグリッドビューデータ表示
            //gridShow(dataGridView1);

            // 画面初期化
            dispClear();

            // 受注確定書入力シートパスを取得：2019/03/06
            var s = dts.会社情報.Where(a => a.ID == 1);
            foreach (var t in s)
            {
                sheetPath = t.受注確定書入力シートパス;
            }

            // 2019/03/06
            if (sheetPath == string.Empty)
            {
                string msg = "出力先シートが未登録のため、受注確定書入力シートにデータ出力は行われません" + Environment.NewLine + "会社情報画面で出力先となる受注確定書入力シートパスを登録してください";
                MessageBox.Show(msg, "出力シート未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        // 配布条件コンボ設定
        private void cmbFjyoukenSet()
        {
            cmbHaifuJyoken.Items.Clear();
            cmbHaifuJyoken.Items.Add("予定枚数どおり");
            cmbHaifuJyoken.Items.Add("詰めて配布 ＯＫ");
            cmbHaifuJyoken.SelectedIndex = -1;
        }

        // 納品場所コンボ設定
        private void cmbNouhinBashoSet()
        {
            cmbNouhinBasho.Items.Clear();
            cmbNouhinBasho.Items.Add("新宿");
            cmbNouhinBasho.Items.Add("上野");
            cmbNouhinBasho.Items.Add("他社へ直納");
            cmbNouhinBasho.SelectedIndex = 0;
        }

        // 事前請求書コンボ設定
        private void cmbJizenSeikyushoSet()
        {
            cmbJizenseikyusho.Items.Clear();
            cmbJizenseikyusho.Items.Add("要");
            cmbJizenseikyusho.Items.Add("不要");
            cmbJizenseikyusho.SelectedIndex = 0;
        }

        /// ----------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// ----------------------------------------------------
        private void dispClear()
        {
            //dataGridView1.RowCount = 0;

            dtOrder.Checked = false;
            cmbNewRep.SelectedIndex = -1;
            txtClient.Text = string.Empty;
            txtName.Text = string.Empty;
            txtZipCode.Text = string.Empty;
            txtAdd.Text = string.Empty;
            txtTel.Text = string.Empty;
            txtFax.Text = string.Empty;
            txtBusho.Text = string.Empty;
            txtTantou.Text = string.Empty;
            txtsName.Text = string.Empty;
            txtsBusho.Text = string.Empty;
            txtsTantou.Text = string.Empty;
            txtsZipCode.Text = string.Empty;
            txtsAdd.Text = string.Empty;

            txtZeinuki.Text = string.Empty;
            txtArari.Text = string.Empty;
            txtEiTantou.Text = string.Empty;

            //dataGridView2.RowCount = 0;

            cmbHaifuJyoken.SelectedIndex = -1;
            cmbHaifukeitai.SelectedIndex = -1;
            dtNouhin.Checked = false;
            dtNyukin.Checked = false;
            dtSeikyuShime.Checked = false;  // 2015/12/10
            cmbNouhinBasho.SelectedIndex = -1;
            //txtHaifuhoukoku.Text = string.Empty; //  2015/11/27 コメント化
            //cmbJizenseikyusho.SelectedIndex = -1; // 2015/11/27 コメント化
            txtMemo.Text = string.Empty;
            txtSalesMemo.Text = string.Empty;   // 2019/03/04

            chkKaishuFlg.Checked = false;   // 2019/02/18

            btnRep.Enabled = false;

            // 検索欄 2019/02/18
            txtSNum.Text = string.Empty;
            sDate.Checked = false;
            txtSClient.Text = string.Empty;
            chk1Year.Checked = true;
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     受注データをグリッドに表示する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------
        private void gridShow(DataGridView g)
        {
            Cursor = Cursors.WaitCursor;

            g.Rows.Clear();
            int iX = 0;
            DateTime dt = DateTime.Today.AddYears(-1);

            // 過去1年間 : 2019/02/16
            if (chk1Year.Checked)
            {
                jAdp.FillByFromYMDToYMD(dts.受注1, dt, DateTime.Parse("2999/12/31"));
            }
            else
            {
                jAdp.Fill(dts.受注1);
            }

            var d = dts.受注1.OrderByDescending(a => a.ID);

            // ログインユーザーの受注データ保守区分が「自分が登録したデータのみ(1)」のとき
            if (global.loginOrderMntType == 0)
            {
                d = d.Where(a => a.登録ユーザーID == global.loginUserID).OrderByDescending(a => a.ID);
            }

            // 過去1年間 : 2019/02/16
            //if (chk1Year.Checked)
            //{
            //    d = d.Where(a => a.受注日 >= dt).OrderByDescending(a => a.ID);
            //}

            // 受注番号検索 : 2019/02/18
            if (txtSNum.Text != string.Empty)
            {
                d = d.Where(a => a.ID.ToString().Contains(txtSNum.Text)).OrderByDescending(a => a.ID);
            }

            // 受注日検索 : 2019/02/18
            if (sDate.Checked)
            {
                d = d.Where(a => !a.Is受注日Null() && a.受注日.ToShortDateString() == sDate.Value.ToShortDateString()).OrderByDescending(a => a.ID);
            }

            // クライアント名 : 2019/02/18
            if (txtSClient.Text != string.Empty)
            {
                d = d.Where(a => a.得意先Row != null && a.得意先Row.略称.Contains(txtSClient.Text)).OrderByDescending(a => a.ID);
            }



            // グリッドに表示
            foreach (var t in d)
            {
                g.Rows.Add();
                g[colSel, iX].Value = "選択";
                g[colID, iX].Value = t.ID.ToString();
                g[colDt, iX].Value = t.受注日.ToShortDateString();
                
                if (t.得意先Row == null)
                {
                    g[colClient, iX].Value = string.Empty;
                    g[colClientCode, iX].Value = 0;
                }
                else
                {
                    g[colClient, iX].Value = t.得意先Row.略称;
                    g[colClientCode, iX].Value = t.得意先ID;
                }

                g[colName, iX].Value = t.チラシ名;

                string sDt = string.Empty;
                string eDt = string.Empty;

                if (t.Is配布開始日Null())
                {
                    sDt = string.Empty;
                }
                else
                {
                    sDt = t.配布開始日.ToShortDateString();
                }

                if (t.Is配布終了日Null())
                {
                    eDt = string.Empty;
                }
                else
                {
                    eDt = t.配布終了日.ToShortDateString();
                }

                if (sDt == string.Empty && eDt == string.Empty)
                {
                    g[colKikan, iX].Value = string.Empty;
                }
                else
                {
                    g[colKikan, iX].Value = sDt + "～" + eDt;
                }
                
                if (t.得意先Row == null)
                {
                    g[colTantou, iX].Value = string.Empty;
                }
                else
                {
                    if (t.得意先Row.社員Row == null)
                    {
                        g[colTantou, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colTantou, iX].Value = t.得意先Row.社員Row.氏名;
                    }
                }

                // 受注確定書発行済みのとき
                if (t.受注確定書発行 == 1)
                {
                    g.Rows[iX].DefaultCellStyle.ForeColor = Color.LightGray;
                }
                else
                {
                    g.Rows[iX].DefaultCellStyle.ForeColor = Color.Black;
                }

                iX++;
            }

            g.CurrentCell = null;

            Cursor = Cursors.Default;

            if (g.RowCount == 0)
            {
                MessageBox.Show("受注確定書作成可能な受注案件がありません。", "対象データなし", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 選択ボタンのとき
            if (e.ColumnIndex == 0)
            {
                //string msg = "受注確定書に追加しますか？" + Environment.NewLine + Environment.NewLine + "受注番号：" + dataGridView1[colID, e.RowIndex].Value.ToString();
                //if (MessageBox.Show(msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                //{
                //    return;
                //}

                // 受注番号取得
                Int64 jNum = Int64.Parse(dataGridView1[colID, e.RowIndex].Value.ToString());

                // クライアント番号取得
                int cNum = int.Parse(dataGridView1[colClientCode, e.RowIndex].Value.ToString());

                // 異なる支払日のチェック：2018/01/04
                if (dtSeikyuShime.Checked)
                {
                    if (!getShimebi(jNum))
                    {
                        MessageBox.Show("請求締日が一致しません", "エラー確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                // 異なる支払日のチェック：2018/01/04
                if (dtNyukin.Checked)
                {
                    if (!getShiharaibi(jNum))
                    {
                        MessageBox.Show("支払日が一致しません", "エラー確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                // 追加済みの受注案件を調べる
                if (!getOrderNumDGV2(dataGridView2, jNum))
                {
                    MessageBox.Show("この受注案件は既に追加済です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // 確定済み受注案件のチェック
                if (dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor == Color.LightGray)
                {
                    if (MessageBox.Show("この案件は受注確定書発行済です。このまま追加しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }

                // 異なるクライアントのチェック
                if (!getClientDGV2(dataGridView2, cNum))
                {
                    if (MessageBox.Show("クライアントが一致しません。このまま追加しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }
                
                // 受注確定書に追加する
                addDataGridView2(dataGridView2, jNum);

                // グリッドの合計税抜き、粗利を求める
                int tlZeinuki = 0;
                int tlArari = 0;
                int tlArari2 = 0;

                if (getTotalKingaku(dataGridView2, out tlZeinuki, out tlArari, out tlArari2))
                {
                    txtZeinuki.Text = tlZeinuki.ToString("#,##0");
                    txtArari.Text = tlArari.ToString("#,##0");
                    txtArari2.Text = tlArari2.ToString("#,##0");
                }
            }
        }
        
        /// ----------------------------------------------------------------
        /// <summary>
        ///     追加済みの受注番号を検索する </summary>
        /// <param name="g">
        ///     データグリッドビュー</param>
        /// <param name="jNum">
        ///     受注番号</param>
        /// <returns>
        ///     true:検索結果なし、false:検索結果あり</returns>
        /// ----------------------------------------------------------------
        private bool getOrderNumDGV2(DataGridView g, Int64 jNum)
        {
            bool res = true;

            for (int i = 0; i < g.RowCount; i++)
            {
                if (g[colID, i].Value.ToString().Equals(jNum.ToString()))
                {
                    res = false;
                    break;
                }
            }

            return res;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     追加済みの受注データのクライアントコードを調べる </summary>
        /// <param name="g">
        ///     データグリッドビュー </param>
        /// <param name="cID">
        ///     得意先コード </param>
        /// <returns>
        ///     true:検索結果一致、false:検索結果不一致あり</returns>
        /// -------------------------------------------------------------------------
        private bool getClientDGV2(DataGridView g, int cID)
        {
            bool res = true;

            for (int i = 0; i < g.RowCount; i++)
            {
                Int64 jNum = Int64.Parse(g[colID, i].Value.ToString());

                if (!dts.受注1.Any(a => a.ID == jNum && a.得意先ID == cID))
                {
                    res = false;
                    break;
                }
            }
            return res;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     追加済みの支払日と一致しているかチェックする </summary>
        /// <param name="jNum">
        ///     受注番号</param>
        /// <returns>
        ///     true:検索結果一致、false:検索結果不一致</returns>
        /// -------------------------------------------------------------------------
        private bool getShiharaibi(Int64 jNum)
        {
            bool res = true;
            DateTime sDt = DateTime.Today;

            if (dts.受注1.Any(a => a.ID == jNum))
            {
                var s = dts.受注1.Single(a => a.ID == jNum);
                //sDt = s.外注支払日営業;

                // 2018/02/20
                if (!s.Is入金予定日Null())
                {
                    sDt = s.入金予定日;
                }
                else
                {
                    sDt = DateTime.Today;
                }


                if (DateTime.Parse(dtNyukin.Value.ToShortDateString()) != DateTime.Parse(sDt.ToShortDateString()))
                {
                    res = false;
                }
            }

            return res;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     追加済みの請求締日と一致しているかチェックする </summary>
        /// <param name="jNum">
        ///     受注番号</param>
        /// <returns>
        ///     true:検索結果一致、false:検索結果不一致</returns>
        /// -------------------------------------------------------------------------
        private bool getShimebi(Int64 jNum)
        {
            bool res = true;
            DateTime sDt = DateTime.Today;

            if (dts.受注1.Any(a => a.ID == jNum))
            {
                var s = dts.受注1.Single(a => a.ID == jNum);
                sDt = s.請求書発行日;

                if (DateTime.Parse(dtSeikyuShime.Value.ToShortDateString()) != DateTime.Parse(sDt.ToShortDateString()))
                {
                    res = false;
                }
            }

            return res;
        }

        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     確定書に受注データを追加する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sID">
        ///     受注番号</param>
        /// -----------------------------------------------------------------------------------
        private void addDataGridView2(DataGridView g, Int64 sID)
        {
            // グリッドに表示
            foreach (var t in dts.受注1.Where(a => a.ID == sID))
            {
                g.Rows.Add();

                int iX = g.RowCount - 1;

                g[colCmb, iX].Value = t.受注種別ID;
                g[colNaiyou, iX].Value = t.チラシ名;

                if (t.業種 != null)
                {
                    g[colGyoushu, iX].Value = t.業種;
                }
                else
                {
                    g[colGyoushu, iX].Value = string.Empty;
                }

                string sDt = string.Empty;
                string eDt = string.Empty;

                // 配布期間
                if (t.Is配布開始日Null())
                {
                    sDt = string.Empty;
                }
                else
                {
                    sDt = t.配布開始日.ToShortDateString();
                }

                if (t.Is配布終了日Null())
                {
                    eDt = string.Empty;
                }
                else
                {
                    eDt = t.配布終了日.ToShortDateString();
                }

                if (sDt == string.Empty && eDt == string.Empty)
                {
                    g[colDtHaifuS, iX].Value = string.Empty;
                }
                else
                {
                    g[colDtHaifuS, iX].Value = sDt + "～" + eDt;
                }

                g[colSize, iX].Value = t.受注日.ToShortDateString();

                if (t.判型 == 0)
                {
                    g[colSize, iX].Value = string.Empty;
                }
                else
                {
                    g[colSize, iX].Value = t.判型Row.名称;
                }

                //g[colCmbOri, iX].Value = 2;
                g[colTanka, iX].Value = t.単価.ToString("#,##0.00");
                g[colMaisu, iX].Value = t.枚数.ToString("#,##0");
                g[colKingaku, iX].Value = t.金額.ToString("#,##0");
                g[colNebiki, iX].Value = t.値引額.ToString("#,##0");

                int sGenka = 0;
                int sGenka2 = 0;
                int sGenka3 = 0;

                // 請求締日表示：2018/01/04
                dtSeikyuShime.Checked = true;    // 明示的にチェックをオン 2019/03/16   
                dtSeikyuShime.Value = t.請求書発行日;

                // 支払期日欄に表示：2018/02/20
                if (!t.Is入金予定日Null())
                {
                    dtNyukin.Checked = true;
                    dtNyukin.Value = t.入金予定日;
                }
                else
                {
                    dtNyukin.Checked = false;
                }

                // 外注先１ 2015/12/10
                if (t.外注先RowBy外注先_受注11 != null)
                {
                    // 外注
                    g[colGaichu, iX].Value = t.外注先RowBy外注先_受注11.名称;

                    // 2015/11/19
                    if (t.Is外注支払日営業Null())
                    {
                        g[colDtShiharai, iX].Value = string.Empty;

                        //// 2018/01/04
                        //dtNyukin.Checked = false;
                    }
                    else
                    {
                        g[colDtShiharai, iX].Value = t.外注支払日営業.ToShortDateString();

                        //// 支払期日欄に表示：2018/01/04
                        //dtNyukin.Checked = true;
                        //dtNyukin.Value = t.外注支払日営業;
                    }

                    // 2015/12/06 外注原価を「単価から原価総額入力へ変更」に伴う
                    sGenka = (int)(t.外注原価営業);
                    g[colGenka, iX].Value = sGenka.ToString("#,##0");

                    // 粗利に値引額を反映 2016/01/04
                    g[colArari, iX].Value = (t.金額 - t.値引額 - (decimal)sGenka).ToString("#,##0");

                    if ((t.金額 - t.値引額) <= 0)
                    {
                        g[colArariRT, iX].Value = 0;
                    }
                    else
                    {
                        g[colArariRT, iX].Value = ((t.金額 - t.値引額 - (decimal)sGenka) / (t.金額 - t.値引額) * 100).ToString("##0.00") + "%";
                    }
                }
                else
                {
                    g[colArari, iX].Value = 0;
                }

                // 外注先２ 2015/12/10
                if (t.外注先RowBy外注先_受注1 == null)
                {
                    // 自社ポス
                    g[colGaichu, iX].Value = "自社";
                    g[colDtShiharai, iX].Value = string.Empty;
                    sGenka = (int)(t.配布単価 * (decimal)t.枚数);
                    g[colGenka, iX].Value = sGenka.ToString("#,##0");

                    // 粗利に値引額を反映 2016/01/04
                    g[colArari, iX].Value = (t.金額 - t.値引額 - (decimal)sGenka).ToString("#,##0");

                    if ((t.金額 - t.値引額) <= 0)
                    {
                        g[colArariRT, iX].Value = 0;
                    }
                    else
                    {
                        g[colArariRT, iX].Value = ((t.金額 - t.値引額 - (decimal)sGenka) / (t.金額 - t.値引額) * 100).ToString("##0.00") + "%";                
                    }
                }
                else
                {
                    // 外注先
                    g[colGaichu2, iX].Value = t.外注先RowBy外注先_受注1.名称;

                    // 2015/11/19
                    if (t.Is外注支払日支払Null())
                    {
                        g[colDtShiharai2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colDtShiharai2, iX].Value = t.外注支払日支払.ToShortDateString();
                    }

                    // 2015/12/06 外注原価を「単価から原価総額入力へ変更」に伴う
                    sGenka = (int)(t.外注原価支払);
                    g[colGenka2, iX].Value = sGenka.ToString("#,##0");
                }

                // 外注先３
                if (t.外注先RowBy外注先_受注12 != null)
                {
                    // 外注先
                    g[colGaichu3, iX].Value = t.外注先RowBy外注先_受注12.名称;

                    // 支払日
                    if (t.Is外注支払日支払2Null())
                    {
                        g[colDtShiharai3, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colDtShiharai3, iX].Value = t.外注支払日支払2.ToShortDateString();
                    }

                    // 原価
                    sGenka2 = (int)(t.外注原価支払2);
                    g[colGenka3, iX].Value = sGenka2.ToString("#,##0");
                }

                // 外注先４
                if (t.外注先RowBy外注先_受注13 != null)
                {
                    // 外注先
                    g[colGaichu4, iX].Value = t.外注先RowBy外注先_受注13.名称;

                    // 支払日
                    if (t.Is外注支払日支払3Null())
                    {
                        g[colDtShiharai4, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colDtShiharai4, iX].Value = t.外注支払日支払3.ToShortDateString();
                    }

                    // 原価
                    sGenka3 = (int)(t.外注原価支払3);
                    g[colGenka4, iX].Value = sGenka3.ToString("#,##0");
                }

                // 粗利に値引額を反映 2016/01/04
                // 外注先４つに対応 2016/10/26
                g[colArari2, iX].Value = (t.金額 - t.値引額 - (decimal)sGenka - (decimal)sGenka2 - (decimal)sGenka3).ToString("#,##0");

                if ((t.金額 - t.値引額) <= 0)
                {
                    g[colArariRT2, iX].Value = 0;
                }
                else
                {
                    g[colArariRT2, iX].Value = ((t.金額 - t.値引額 - (decimal)sGenka - (decimal)sGenka2 - (decimal)sGenka3) / (t.金額 - t.値引額) * 100).ToString("##0.00") + "%";
                }
                
                g[colID, iX].Value = t.ID.ToString();
            }

            g.CurrentCell = null;

            btnRep.Enabled = true;
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     グリッドの合計税抜き、粗利を求める </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="tlZeinuki">
        ///     合計（税抜）</param>
        /// <param name="tlArari">
        ///     合計粗利１</param>
        /// <param name="tlArari">
        ///     合計粗利２</param>
        /// <returns>
        ///     true, false</returns>
        /// -------------------------------------------------------------------------------
        private bool getTotalKingaku(DataGridView g, out int tlZeinuki, out int tlArari, out int tlArari2)
        {
            tlZeinuki = 0;
            tlArari = 0;
            tlArari2 = 0;

            for (int i = 0; i < g.RowCount; i++)
            {
                tlZeinuki += Utility.strToInt(g[colKingaku, i].Value.ToString().Replace(",", ""));
                tlZeinuki -= Utility.strToInt(g[colNebiki, i].Value.ToString().Replace(",", ""));   // 値引額を税抜金額から控除 2016/01/04
                tlArari += Utility.nullToInt(g[colArari, i].Value);
                tlArari2 += Utility.nullToInt(g[colArari2, i].Value);
            }

            return true;
        }


        /// -------------------------------------------------------------------------
        /// <summary>
        ///     クライアント情報表示 </summary>
        /// <param name="sID">
        ///     受注番号 </param>
        /// -------------------------------------------------------------------------
        private void showHeader(Int64 sID)
        {            
            foreach (var t in dts.受注1.Where(a => a.ID == sID))
            {
                dtOrder.Value = t.受注日;

                if (t.得意先Row != null)
                {
                    txtClient.Text = t.得意先Row.略称;
                    txtZipCode.Text = t.得意先Row.郵便番号;
                    txtAdd.Text = t.得意先Row.都道府県 + " " + t.得意先Row.住所1 + " " + t.得意先Row.住所2;
                    txtTel.Text = t.得意先Row.電話番号;
                    txtFax.Text = t.得意先Row.FAX番号;
                    txtBusho.Text = t.得意先Row.部署名;

                    if (t.得意先Row.担当者名 != string.Empty)
                    {
                        txtTantou.Text = t.得意先Row.担当者名 + " 様";
                    }
                    else
                    {
                        txtTantou.Text = string.Empty;
                    }

                    if ((t.得意先Row.略称.Trim().Equals(t.得意先Row.請求先名称.Trim())) || t.得意先Row.請求先名称.Trim() == string.Empty)
                    {
                        txtsName.Text = "同上";
                        txtsBusho.Text = string.Empty;
                        txtsTantou.Text = string.Empty;
                        txtsAdd.Text = string.Empty;
                    }
                    else
                    {
                        txtsName.Text = t.得意先Row.請求先名称;
                        //txtsBusho.Text = string.Empty;    // 2019/02/21
                        txtsBusho.Text = t.得意先Row.請求先部署名;   // 2019/02/21

                        if (t.得意先Row.請求先担当者名 != string.Empty)
                        {
                            //txtsTantou.Text = t.得意先Row.請求先担当者名 + " 様";    // 2019/02/21
                            txtsTantou.Text = t.得意先Row.請求先担当者名 + " " + t.得意先Row.請求先敬称;  // 2019/02/21
                        }
                        else
                        {
                            txtsTantou.Text = string.Empty;
                        }

                        txtsZipCode.Text = t.得意先Row.請求先郵便番号;
                        txtsAdd.Text = t.得意先Row.請求先住所1 + " " + t.得意先Row.請求先住所2;
                    }

                    if (t.得意先Row.社員Row != null)
                    {
                        txtEiTantou.Text = t.得意先Row.社員Row.氏名;
                    }
                    else
                    {
                        txtEiTantou.Text = string.Empty;
                    }
                }
                else
                {
                    txtClient.Text = string.Empty;
                    txtZipCode.Text = string.Empty;
                    txtAdd.Text = string.Empty;
                    txtTel.Text = string.Empty;
                    txtFax.Text = string.Empty;
                    txtBusho.Text = string.Empty;
                    txtTantou.Text = string.Empty;
                    txtEiTantou.Text = string.Empty;

                    txtsName.Text = string.Empty;
                    txtsBusho.Text = string.Empty;
                    txtsTantou.Text = string.Empty;
                    txtsAdd.Text = string.Empty;
                }

                txtName.Text = t.チラシ名;

                // 配布形態
                cmbHaifukeitai.SelectedValue = t.配布形態;

                // 配布条件
                for (int i = 0; i < cmbHaifuJyoken.Items.Count; i++)
                {
                    if (cmbHaifuJyoken.Items[i].ToString() == t.配布条件)
                    {
                        cmbHaifuJyoken.SelectedIndex = i;
                        break;
                    }
                }

                // 納品日
                if (t.Is納品予定日Null())
                {
                    dtNouhin.Checked = false;
                }
                else
                {
                    dtNouhin.Checked = true;
                    dtNouhin.Value = t.納品予定日;
                }

                // 特記事項
                //txtMemo.Text = t.特記事項;    // 受注データの特記事項を採用しない 2015/11/27

                // 営業備考 2019/03/04
                txtSalesMemo.Text = t.営業備考;
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 削除ボタンのとき
            if (e.ColumnIndex == 0)
            {
                if (MessageBox.Show("削除してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }

                // 行削除
                dataGridView2.Rows.RemoveAt(e.RowIndex);

                // グリッドの合計税抜き、粗利を求める
                int tlZeinuki =0;
                int tlArari = 0;
                int tlArari2 = 0;

                if (getTotalKingaku(dataGridView2, out tlZeinuki, out tlArari, out tlArari2))
                {
                    txtZeinuki.Text = tlZeinuki.ToString("#,##0");
                    txtArari.Text = tlArari.ToString("#,##0");
                    txtArari2.Text = tlArari2.ToString("#,##0");
                }
            }
        }

        private void dataGridView2_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (dataGridView2.RowCount > 0)
            {
                // 受注番号取得
                Int64 jNum = Int64.Parse(dataGridView2[colID, 0].Value.ToString());

                // クライアント情報表示
                showHeader(jNum);
            }
            else
            {
                // 受注案件を全て削除したとき
                dispClear();
            }
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // 1行目の受注番号のとき
            if (e.RowIndex == 0 && dataGridView2.Columns[e.ColumnIndex].Name == colID)
            {
                // 受注番号取得
                Int64 jNum = Int64.Parse(dataGridView2[colID, 0].Value.ToString());

                // クライアント情報表示
                showHeader(jNum);
            }
        }

        private void btnRep_Click(object sender, EventArgs e)
        {
            if (cmbNewRep.SelectedIndex == -1)
            {
                MessageBox.Show("新規／リピートを選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbNewRep.Focus();
                return;
            }

            // 折の有無を選択：2018/01/04
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                // 「ポスティング」で判型が「B4」以上のサイズのとき、折の有無を必須とする：2018/04/13
                if (dataGridView2[colCmbOri, i].Value == null && 
                    Utility.nullToStr(dataGridView2[colCmb, i].Value) == "1" && 
                   (Utility.nullToStr(dataGridView2[colSize, i].Value) == "B4" || 
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B４" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ4" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ４" || 
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B3" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B３" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ3" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ３" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B2" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B２" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ2" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ２" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B1" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B１" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ1" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ１" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B0" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "B０" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ0" ||
                    Utility.nullToStr(dataGridView2[colSize, i].Value) == "Ｂ０"))

                {
                    MessageBox.Show("折の有無を選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            if (MessageBox.Show("受注確定書を発行します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 受注確定書作成
            kakuteishoReport();

            // 受注確定書入力シートへの書き込み：2019/03/07
            if (sheetPath != string.Empty)
            {
                if (MessageBox.Show("受注確定書入力シートへの書き込みを行いますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    setData2Sheet();
                }
            }

            // 受注データに受注確定書発行フラグを書き込む
            if (MessageBox.Show("受注データに受注確定書発行フラグを書き込みます。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                updateOrderData();
            }

            // 画面初期化
            dispClear();
            dataGridView2.Rows.Clear();
        }
        
        /// ------------------------------------------------------------------------
        /// <summary>
        ///     受注@受注確定書発行：更新</summary>
        /// ------------------------------------------------------------------------
        private void updateOrderData()
        {
            // 受注確定書を発行した受注データにフラグを書き込む
            for (int iX = 0; iX < dataGridView2.RowCount; iX++)
            {
                Int64 jNum = Utility.strToLong(dataGridView2[colID, iX].Value.ToString());

                if (dts.受注1.Any(a => a.ID == jNum))
                {
                    var s = dts.受注1.Single(a => a.ID == jNum);
                    s.受注確定書発行 = 1;
                    s.変更年月日 = DateTime.Now;
                    s.ユーザーID = global.loginUserID;
                }
            }

            // データベース更新
            jAdp.Update(dts.受注1);
        }

        private void kakuteishoReport()
        {
            int _tlPages = 0;
            int _Pages = 0;

            // 受注案件明細
            int mRow = 0;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                // 受注確定書テンプレートBOOK
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル受注確定書, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                // 受注確定書印刷用BOOK
                Excel.Workbook oXlsBookPrn = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル受注確定書印刷, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                // 受注確定書テンプレートシート
                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                // 受注確定書印刷シート
                Excel.Worksheet oxlsSheetPrn = null;

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    // 受注確定書の総ページ数を求める
                    _tlPages = dataGridView2.RowCount / 4;
                    if ((dataGridView2.RowCount % 4) > 0)
                    {
                        _tlPages++;
                    }
                    
                    // 1ページから総ページ数まで
                    for (int i = 1; i <= _tlPages; i++)
                    {
                        _Pages = i;

                        // ページシートを追加する
                        oxlsSheet.Copy(Type.Missing, oXlsBookPrn.Sheets[_Pages]);

                        // 印刷用カレントシート
                        oxlsSheetPrn = (Excel.Worksheet)oXlsBookPrn.Sheets[_Pages + 1];

                        // 発行日・ページ数
                        oxlsSheetPrn.Cells[1, 25] = "発行： " + DateTime.Today.ToShortDateString() + "   P." + _Pages.ToString() + "/" + _tlPages.ToString();

                        // 営業担当者
                        if (txtEiTantou.Text != string.Empty)
                        {
                            oxlsSheetPrn.Cells[2, 30] = "担当： " + txtEiTantou.Text;
                        }
                        else
                        {
                            oxlsSheetPrn.Cells[2, 30] = "担当： ";
                        }

                        oxlsSheetPrn.Cells[3, 7] = dtOrder.Value.ToShortDateString();  // 受注日
                        oxlsSheetPrn.Cells[3, 19] = cmbNewRep.Text;    // 新規orリピート
                        oxlsSheetPrn.Cells[4, 7] = txtClient.Text;     // クライアント名
                        oxlsSheetPrn.Cells[4, 25] = txtName.Text;      // 商品名
                        oxlsSheetPrn.Cells[5, 8] = txtZipCode.Text;    // 郵便番号
                        oxlsSheetPrn.Cells[6, 7] = txtAdd.Text;        // 住所
                        oxlsSheetPrn.Cells[7, 7] = txtTel.Text;        // 電話番号
                        oxlsSheetPrn.Cells[7, 25] = txtFax.Text;       // ＦＡＸ
                        oxlsSheetPrn.Cells[8, 7] = txtBusho.Text;      // 部署名
                        oxlsSheetPrn.Cells[8, 25] = txtTantou.Text;    // 担当者
                        oxlsSheetPrn.Cells[9, 10] = txtsName.Text;    // 御社名
                        oxlsSheetPrn.Cells[10, 10] = txtsBusho.Text;   // 部署名
                        oxlsSheetPrn.Cells[10, 25] = txtsTantou.Text;  // 担当者

                        //string zmk = string.Empty;
                        //if (txtsZipCode.Text.Trim() != string.Empty)
                        //{
                        //    zmk = "〒";
                        //}

                        oxlsSheetPrn.Cells[11, 11] = txtsZipCode.Text;  // 郵便番号
                        oxlsSheetPrn.Cells[11, 14] = txtsAdd.Text;      // 住所

                        // 以下、明細印刷
                        int sR = (i - 1) * 4;   // グリッド開始行
                        int eR = sR + 3;        // グリッド終了行
                        int _m = 0;

                        if ((dataGridView2.RowCount - 1) < eR)
                        {
                            eR = dataGridView2.RowCount - 1;
                        }

                        for (int iX = sR; iX <= eR; iX++)
                        {
                            mRow = _m * 6 + 13;

                            string gy = string.Empty;
                            if (dataGridView2[colGyoushu, iX].Value.ToString() != string.Empty)
                            {
                                gy = "・" + dataGridView2[colGyoushu, iX].Value.ToString();
                            }

                            oxlsSheetPrn.Cells[mRow, 1] = dataGridView2[colCmb, iX].FormattedValue.ToString() + gy;  // 受注内容
                            oxlsSheetPrn.Cells[mRow, 7] = dataGridView2[colDtHaifuS, iX].Value.ToString();     // 配布日
                            oxlsSheetPrn.Cells[mRow, 14] = dataGridView2[colSize, iX].Value.ToString();        // サイズ
                            oxlsSheetPrn.Cells[mRow, 17] = dataGridView2[colCmbOri, iX].FormattedValue.ToString();      // 折り
                            oxlsSheetPrn.Cells[mRow, 19] = Utility.strToDecimal(dataGridView2[colTanka, iX].Value.ToString());   // 単価
                            oxlsSheetPrn.Cells[mRow, 24] = Utility.strToDecimal(dataGridView2[colMaisu, iX].Value.ToString());   // 枚数

                            oxlsSheetPrn.Cells[mRow + 1, 1] = Utility.strToInt64(dataGridView2[colID, iX].Value.ToString());    // 受注番号　2016/10/26

                            // 2015/12/10
                            oxlsSheetPrn.Cells[mRow + 1, 19] = Utility.nullToInt(dataGridView2[colNebiki, iX].Value).ToString();    // 値引

                            // 2015/12/10
                            oxlsSheetPrn.Cells[mRow + 2, 4] = Utility.nullToStr(dataGridView2[colGaichu, iX].Value);         // 外注先１
                            oxlsSheetPrn.Cells[mRow + 2, 13] = Utility.nullToStr(dataGridView2[colDtShiharai, iX].Value);    // 支払日
                            oxlsSheetPrn.Cells[mRow + 2, 19] = Utility.nullToInt(dataGridView2[colGenka, iX].Value).ToString();       // 原価
                            oxlsSheetPrn.Cells[mRow + 2, 26] = Utility.nullToInt(dataGridView2[colArari, iX].Value).ToString();       // 粗利

                            // 2015/12/10
                            oxlsSheetPrn.Cells[mRow + 3, 4] = Utility.nullToStr(dataGridView2[colGaichu2, iX].Value);         // 外注先２
                            oxlsSheetPrn.Cells[mRow + 3, 13] = Utility.nullToStr(dataGridView2[colDtShiharai2, iX].Value);    // 支払日
                            oxlsSheetPrn.Cells[mRow + 3, 19] = Utility.nullToInt(dataGridView2[colGenka2, iX].Value).ToString();       // 原価
                            oxlsSheetPrn.Cells[mRow + 3, 26] = Utility.nullToInt(dataGridView2[colArari2, iX].Value).ToString();       // 粗利

                            // 2016/10/25
                            oxlsSheetPrn.Cells[mRow + 4, 4] = Utility.nullToStr(dataGridView2[colGaichu3, iX].Value);         // 外注先３
                            oxlsSheetPrn.Cells[mRow + 4, 13] = Utility.nullToStr(dataGridView2[colDtShiharai3, iX].Value);    // 支払日

                            if (Utility.nullToInt(dataGridView2[colGenka3, iX].Value) != 0)
                            {
                                oxlsSheetPrn.Cells[mRow + 4, 19] = Utility.nullToInt(dataGridView2[colGenka3, iX].Value).ToString();       // 原価
                            }

                            // 2016/10/25
                            oxlsSheetPrn.Cells[mRow + 5, 4] = Utility.nullToStr(dataGridView2[colGaichu4, iX].Value);         // 外注先４
                            oxlsSheetPrn.Cells[mRow + 5, 13] = Utility.nullToStr(dataGridView2[colDtShiharai4, iX].Value);    // 支払日

                            if (Utility.nullToInt(dataGridView2[colGenka4, iX].Value) != 0)
                            {
                                oxlsSheetPrn.Cells[mRow + 5, 19] = Utility.nullToInt(dataGridView2[colGenka4, iX].Value).ToString();       // 原価
                            }

                            _m++;
                        }


                        // 売上金額は最後のページにのみ印字する 2016/10/26
                        if (i == _tlPages)
                        {
                            oxlsSheetPrn.Cells[38, 4] = cmbHaifukeitai.Text;                        // 配布形態
                            oxlsSheetPrn.Cells[38, 14] = Utility.strToInt(txtZeinuki.Text);         // 税抜合計
                            oxlsSheetPrn.Cells[38, 23] = Utility.strToInt(txtArari.Text);         // 粗利1 2016/01/04

                            // 納品日
                            if (dtNouhin.Checked)
                            {
                                oxlsSheetPrn.Cells[39, 4] = dtNouhin.Value.ToShortDateString();
                            }
                            else
                            {
                                oxlsSheetPrn.Cells[39, 4] = string.Empty;
                            }

                            oxlsSheetPrn.Cells[39, 13] = cmbNouhinBasho.Text;                      // 納品場所

                            // 請求締日 2015/12/10
                            if (dtSeikyuShime.Checked)
                            {
                                oxlsSheetPrn.Cells[39, 22] = dtSeikyuShime.Value.ToShortDateString();
                            }
                            else
                            {
                                oxlsSheetPrn.Cells[39, 22] = string.Empty;
                            }

                            // 支払期日
                            if (dtNyukin.Checked)
                            {
                                oxlsSheetPrn.Cells[39, 29] = dtNyukin.Value.ToShortDateString();
                            }
                            else
                            {
                                oxlsSheetPrn.Cells[39, 29] = string.Empty;
                            }
                        }
                        else
                        {
                            oxlsSheetPrn.Cells[38, 4] = "＊＊＊＊＊";                        // 配布形態
                            oxlsSheetPrn.Cells[38, 14] = "＊＊＊＊＊＊＊";
                            oxlsSheetPrn.Cells[38, 23] = "＊＊＊＊＊＊＊";
                            oxlsSheetPrn.Cells[39, 4] = "＊＊＊＊＊";
                            oxlsSheetPrn.Cells[39, 13] = "＊＊＊＊＊";                      // 納品場所
                            oxlsSheetPrn.Cells[39, 22] = "＊＊＊＊＊";
                            oxlsSheetPrn.Cells[39, 29] = "＊＊＊＊＊";
                        }

                        oxlsSheetPrn.Cells[40, 4] = txtMemo.Text;           // 特記事項
                        oxlsSheetPrn.Cells[40, 22] = txtSalesMemo.Text;     // 営業備考 2019/03/06

                        // 注文書回収フラグ : 2019/02/18
                        if (chkKaishuFlg.Checked)
                        {
                            oxlsSheetPrn.Cells[43, 20] = "済";
                        }
                        else
                        {
                            oxlsSheetPrn.Cells[43, 20] = string.Empty;
                        }
                    }

                    // 印刷用BOOKの1番目のシートは削除する
                    ((Excel.Worksheet)oXlsBookPrn.Sheets[1]).Delete();

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    //印刷
                    oXlsBookPrn.PrintOutEx(1, Type.Missing, 1, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //oXlsBookPrn.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "受注確定書作成";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = txtClient.Text + " 受注確定書";
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        // エクセル保存
                        fileName = saveFileDialog1.FileName;
                        oXlsBookPrn.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        // 受注確定書発行記録テーブル登録
                        setOrderRecord(fileName);
                    }
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "受注確定書作成", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                finally
                {
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);
                    oXlsBookPrn.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheetPrn);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBookPrn);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "受注確定書作成", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     受注確定書発行記録登録 </summary>
        /// <param name="sPath">
        ///     受注確定書パス </param>
        /// ----------------------------------------------------------------------
        private void setOrderRecord(string sPath)
        {
            darwinDataSet.受注確定書発行記録Row r = (darwinDataSet.受注確定書発行記録Row)dts.受注確定書発行記録.NewRow();

            r.発行日 = DateTime.Now;
            r.クライアント名 = txtClient.Text;
            r.商品名 = txtName.Text;
            r.受注確定書パス = sPath;
            r.発行者 = global.loginUserID;
            r.確認1 = global.FLGOFF;
            r.確認者1 = global.FLGOFF;
            r.確認2 = global.FLGOFF;
            r.確認者2 = global.FLGOFF;
            r.確認3 = global.FLGOFF;
            r.確認者3 = global.FLGOFF;
            r.確認4 = global.FLGOFF;
            r.確認者4 = global.FLGOFF;
            r.確認5 = global.FLGOFF;
            r.確認者5 = global.FLGOFF;
            r.登録年月日 = DateTime.Now;
            r.更新年月日 = DateTime.Now;

            dts.受注確定書発行記録.Add受注確定書発行記録Row(r);

            rAdp.Update(dts.受注確定書発行記録);
        }

        private void dataGridView2_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.EditMode == DataGridViewEditMode.EditProgrammatically)
            {
                dataGridView2.BeginEdit(false);
            }
        }

        private void dataGridView2_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView2_Leave(object sender, EventArgs e)
        {
            if (dataGridView2.EditMode == DataGridViewEditMode.EditProgrammatically)
            {
                dataGridView2.EndEdit();
            }
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
            // データグリッドビューデータ表示
            gridShow(dataGridView1);
        }
    }
}

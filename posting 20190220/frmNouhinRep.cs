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
    public partial class frmNouhinRep : Form
    {
        public frmNouhinRep()
        {
            InitializeComponent();

            // データ読み込み
            jAdp.Fill(dts.受注1);
            cAdp.Fill(dts.得意先);
            eAdp.Fill(dts.社員);
            gAdp.Fill(dts.外注先);
            sAdp.Fill(dts.受注種別);
            hAdp.Fill(dts.判型);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter eAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();
        darwinDataSetTableAdapters.受注種別TableAdapter sAdp = new darwinDataSetTableAdapters.受注種別TableAdapter();
        darwinDataSetTableAdapters.判型TableAdapter hAdp = new darwinDataSetTableAdapters.判型TableAdapter();

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
        string colJyuchu = "col36";

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
        string colTax = "col37";
        string colUriage = "col38";

        #endregion

        private decimal tlUriage = 0;
        const int NOUHIN_HAKKO = 1;     // 納品書発行済みフラグ

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
                dgv.Columns.Add(colJyuchu, "受注");
                dgv.Columns.Add(colTantou, "担当");
                dgv.Columns.Add(colClientCode, "");
                dgv.Columns[colClientCode].Visible = false;
                
                dgv.Columns[colSel].Width = 50;
                dgv.Columns[colID].Width = 120;
                dgv.Columns[colDt].Width = 120;
                dgv.Columns[colClient].Width = 300;
                dgv.Columns[colName].Width = 300;
                dgv.Columns[colJyuchu].Width = 100;
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
                btn.Name = colDel;
                btn.HeaderText = "";
                dgv.Columns.Add(btn);

                DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
                cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                cmb.Name = colCmb;
                cmb.HeaderText = "受注内容";
                dgv.Columns.Add(cmb);

                cmb.Items.Add("ポスティング");
                Utility.ComboJshubetsu.load(cmb);

                dgv.Columns.Add(colNaiyou, "内容");
                dgv.Columns.Add(colGyoushu, "業種");
                dgv.Columns.Add(colJyuchu, "受注");
                dgv.Columns.Add(colDtHaifuS, "日付");
                dgv.Columns.Add(colSize, "サイズ");

                //DataGridViewComboBoxColumn cmb2 = new DataGridViewComboBoxColumn();
                //cmb2.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                //cmb2.HeaderText = "折";
                //cmb2.Name = colCmbOri;
                //dgv.Columns.Add(cmb2);
                //Utility.comboOri.load(cmb2);

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
                dgv.Columns.Add(colTax, "消費税");
                dgv.Columns.Add(colUriage, "売上金額");
                dgv.Columns[colID].Visible = false;
                dgv.Columns[colTax].Visible = false;
                dgv.Columns[colUriage].Visible = false;

                // 各列幅指定
                dgv.Columns[colDel].Width = 50;
                dgv.Columns[colCmb].Width = 100;
                dgv.Columns[colNaiyou].Width = 220;
                dgv.Columns[colDtHaifuS].Width = 110;
                dgv.Columns[colJyuchu].Width = 110;
                dgv.Columns[colSize].Width = 80;
                //dgv.Columns[colCmbOri].Width = 50;
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
                //dgv.Columns[colCmbOri].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
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

                // 編集不可とする
                dgv.ReadOnly = false;

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
       
            
            //// 納品場所コンボ設定
            //cmbNouhinBashoSet();
            
            // 事前請求書コンボ設定
            //cmbJizenSeikyushoSet();   // 2015/11/27 コメント化

            // データグリッドビューデータ表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();
        }

        // 納品場所コンボ設定
        //private void cmbNouhinBashoSet()
        //{
        //    cmbNouhinBasho.Items.Clear();
        //    cmbNouhinBasho.Items.Add("新宿");
        //    cmbNouhinBasho.Items.Add("上野");
        //    cmbNouhinBasho.Items.Add("他社へ直納");
        //    cmbNouhinBasho.SelectedIndex = 0;
        //}

        //// 事前請求書コンボ設定
        //private void cmbJizenSeikyushoSet()
        //{
        //    cmbJizenseikyusho.Items.Clear();
        //    cmbJizenseikyusho.Items.Add("要");
        //    cmbJizenseikyusho.Items.Add("不要");
        //    cmbJizenseikyusho.SelectedIndex = 0;
        //}

        /// ----------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// ----------------------------------------------------
        private void dispClear()
        {
            //dataGridView1.RowCount = 0;

            dtOrder.Checked = false;
            //cmbNewRep.SelectedIndex = -1;
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
            txtArari2.Text = string.Empty;
            txtEiTantou.Text = string.Empty;

            //dataGridView2.RowCount = 0;

            //cmbHaifuJyoken.SelectedIndex = -1;
            //cmbHaifukeitai.SelectedIndex = -1;
            dtNouhin.Checked = false;
            dtNyukin.Checked = false;
            dtSeikyuShime.Checked = false;  // 2015/12/10
            //cmbNouhinBasho.SelectedIndex = -1;
            //txtHaifuhoukoku.Text = string.Empty; //  2015/11/27 コメント化
            //cmbJizenseikyusho.SelectedIndex = -1; // 2015/11/27 コメント化
            txtMemo.Text = string.Empty;

            btnRep.Enabled = false;
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     受注データをグリッドに表示する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------
        private void gridShow(DataGridView g)
        {
            g.Rows.Clear();
            int iX = 0;

            var d = dts.受注1.Where(a => a.受注種別ID > 1).OrderByDescending(a => a.ID);

            // ログインユーザーの受注データ保守区分が「自分が登録したデータのみ(1)」のとき
            if (global.loginOrderMntType == 0)
            {
                d = d.Where(a => a.登録ユーザーID == global.loginUserID).OrderByDescending(a => a.ID);
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

                //if (t.Is配布開始日Null())
                //{
                //    sDt = string.Empty;
                //}
                //else
                //{
                //    sDt = t.配布開始日.ToShortDateString();
                //}

                //if (t.Is配布終了日Null())
                //{
                //    eDt = string.Empty;
                //}
                //else
                //{
                //    eDt = t.配布終了日.ToShortDateString();
                //}

                if (t.受注種別Row == null)
                {
                    g[colJyuchu, iX].Value = string.Empty;
                }
                else
                {
                    g[colJyuchu, iX].Value = t.受注種別Row.名称;
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

                // 納品書発行済みのとき
                if (t.納品書発行 == NOUHIN_HAKKO)
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

            if (g.RowCount == 0)
            {
                MessageBox.Show("納品書作成可能な受注案件がありません。", "対象データなし", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

                // 追加済みの受注案件を調べる
                if (!getOrderNumDGV2(dataGridView2, jNum))
                {
                    MessageBox.Show("この受注案件は既に追加済です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // 確定済み受注案件のチェック
                if (dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor == Color.LightGray)
                {
                    if (MessageBox.Show("この案件は納品書発行済です。このまま追加しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
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
                
                // 納品書に追加する
                addDataGridView2(dataGridView2, jNum);

                // グリッドの合計税抜き、粗利を求める
                int tlZeinuki = 0;
                int tlArari = 0;
                int tlArari2 = 0;

                if (getTotalKingaku(dataGridView2, out tlZeinuki, out tlArari, out tlArari2, out tlUriage))
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

                if (t.受注種別Row != null)
                {
                    g[colJyuchu, iX].Value = t.受注種別Row.名称;
                }
                else
                {
                    g[colJyuchu, iX].Value = string.Empty;
                }

                string sDt = string.Empty;
                string eDt = string.Empty;

                //// 配布期間
                //if (t.Is配布開始日Null())
                //{
                //    sDt = string.Empty;
                //}
                //else
                //{
                //    sDt = t.配布開始日.ToShortDateString();
                //}

                //if (t.Is配布終了日Null())
                //{
                //    eDt = string.Empty;
                //}
                //else
                //{
                //    eDt = t.配布終了日.ToShortDateString();
                //}

                //if (sDt == string.Empty && eDt == string.Empty)
                //{
                //    g[colDtHaifuS, iX].Value = string.Empty;
                //}
                //else
                //{
                //    g[colDtHaifuS, iX].Value = sDt + "～" + eDt;
                //}

                //g[colSize, iX].Value = t.受注日.ToShortDateString();

                if (t.判型Row == null)
                {
                    g[colSize, iX].Value = string.Empty;
                }
                else
                {
                    g[colSize, iX].Value = t.判型Row.名称;
                }

                if (t.Is配布開始日Null())
                {
                    g[colDtHaifuS, iX].Value = string.Empty;
                }
                else
                {
                    g[colDtHaifuS, iX].Value = t.配布開始日.ToShortDateString();
                }

                g[colTanka, iX].Value = t.単価.ToString("#,##0.00");
                g[colMaisu, iX].Value = t.枚数.ToString("#,##0");
                g[colKingaku, iX].Value = t.金額.ToString("#,##0");
                g[colNebiki, iX].Value = t.値引額.ToString("#,##0");
                g[colTax, iX].Value = t.消費税.ToString();
                g[colUriage, iX].Value = t.売上金額.ToString();

                int sGenka = 0;
                int sGenka2 = 0;
                int sGenka3 = 0;

                // 外注先１ 2015/12/10
                if (t.外注先RowBy外注先_受注11 != null)
                {
                    // 外注
                    g[colGaichu, iX].Value = t.外注先RowBy外注先_受注11.名称;

                    // 2015/11/19
                    if (t.Is外注支払日営業Null())
                    {
                        g[colDtShiharai, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[colDtShiharai, iX].Value = t.外注支払日営業.ToShortDateString();
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
        /// <param name="tlArari2">
        ///     合計粗利２</param>
        /// <param name="tlUriage">
        ///     税込売上金額</param>
        /// <returns>
        ///     true, false</returns>
        /// -------------------------------------------------------------------------------
        private bool getTotalKingaku(DataGridView g, out int tlZeinuki, out int tlArari, out int tlArari2, out decimal tlUriage)
        {
            tlZeinuki = 0;
            tlArari = 0;
            tlArari2 = 0;
            tlUriage = 0;

            for (int i = 0; i < g.RowCount; i++)
            {
                tlZeinuki += Utility.strToInt(g[colKingaku, i].Value.ToString().Replace(",", ""));
                tlZeinuki -= Utility.strToInt(g[colNebiki, i].Value.ToString().Replace(",", ""));   // 値引額を税抜金額から控除 2016/01/04
                tlArari += Utility.nullToInt(g[colArari, i].Value);
                tlArari2 += Utility.nullToInt(g[colArari2, i].Value);
                tlUriage += Utility.strToDecimal(g[colUriage, i].Value.ToString());
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

                    // 2016/11/15 コメント化
                    //if ((t.得意先Row.略称.Trim().Equals(t.得意先Row.請求先名称.Trim())) || t.得意先Row.請求先名称.Trim() == string.Empty)
                    //{
                    //    txtsName.Text = "同上";
                    //    txtsBusho.Text = string.Empty;
                    //    txtsTantou.Text = string.Empty;
                    //    txtsAdd.Text = string.Empty;
                    //}
                    //else
                    //{
                    //    txtsName.Text = t.得意先Row.請求先名称;
                    //    txtsBusho.Text = string.Empty;

                    //    if (t.得意先Row.請求先担当者名 != string.Empty)
                    //    {
                    //        txtsTantou.Text = t.得意先Row.請求先担当者名 + " 様";
                    //    }
                    //    else
                    //    {
                    //        txtsTantou.Text = string.Empty;
                    //    }

                    //    txtsZipCode.Text = t.得意先Row.請求先郵便番号;
                    //    txtsAdd.Text = t.得意先Row.請求先住所1 + " " + t.得意先Row.請求先住所2;
                    //}

                    // 2016/11/15
                    txtsName.Text = t.得意先Row.請求先名称;
                    txtsBusho.Text = string.Empty;

                    if (t.得意先Row.請求先担当者名 != string.Empty)
                    {
                        txtsTantou.Text = t.得意先Row.請求先担当者名 + " 様";
                    }
                    else
                    {
                        txtsTantou.Text = string.Empty;
                    }

                    txtsZipCode.Text = t.得意先Row.請求先郵便番号;
                    txtsAdd.Text = t.得意先Row.請求先住所1 + " " + t.得意先Row.請求先住所2;


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

                //// 配布形態
                //cmbHaifukeitai.SelectedValue = t.配布形態;

                //// 配布条件
                //for (int i = 0; i < cmbHaifuJyoken.Items.Count; i++)
                //{
                //    if (cmbHaifuJyoken.Items[i].ToString() == t.配布条件)
                //    {
                //        cmbHaifuJyoken.SelectedIndex = i;
                //        break;
                //    }
                //}

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

                if (getTotalKingaku(dataGridView2, out tlZeinuki, out tlArari, out tlArari2, out tlUriage))
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
            //if (cmbNewRep.SelectedIndex == -1)
            //{
            //    MessageBox.Show("新規／リピートを選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    cmbNewRep.Focus();
            //    return;
            //}

            if (MessageBox.Show("納品書を発行します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 納品書作成
            kakuteishoReport();

            // 受注データに受注確定書発行フラグを書き込む
            if (MessageBox.Show("受注データに納品書発行フラグを書き込みます。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                updateOrderData();
            }

            // 画面初期化
            dispClear();
            dataGridView2.Rows.Clear();
        }
        
        /// ------------------------------------------------------------------------
        /// <summary>
        ///     受注@納品書発行：更新</summary>
        /// ------------------------------------------------------------------------
        private void updateOrderData()
        {
            // 納品書を発行した受注データにフラグを書き込む
            for (int iX = 0; iX < dataGridView2.RowCount; iX++)
            {
                Int64 jNum = Utility.strToLong(dataGridView2[colID, iX].Value.ToString());

                if (dts.受注1.Any(a => a.ID == jNum))
                {
                    var s = dts.受注1.Single(a => a.ID == jNum);
                    s.納品書発行 = NOUHIN_HAKKO;
                    s.変更年月日 = DateTime.Now;
                    s.ユーザーID = global.loginUserID;
                }
            }

            // データベース更新
            jAdp.Update(dts.受注1);
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     納品書発行 </summary>
        ///-----------------------------------------------------------
        private void kakuteishoReport()
        {
            int _Pages = 0;
            int _Rows = 0;
            int sIr = 26;
            int _RowsPrn = 21;

            decimal _nebikigo = 0;      // 売上金額 - 値引額
            decimal _tax = 0;           // 消費税
            decimal _seikyu = 0;        // 請求金額

            bool openStatus = true;

            string seikyuNum = string.Empty;    // 請求書番号
            string corp = string.Empty;         // 請求先

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                // 納品書テンプレートBOOK
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル納品書, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                // 納品書印刷用BOOK
                Excel.Workbook oXlsBookPrn = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.エクセル受注確定書印刷, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                // 納品書テンプレートシート
                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                // 納品書印刷シート
                Excel.Worksheet oxlsSheetPrn = null;

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    int iX = 0;

                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        if (iX == 0 || _Rows >= _RowsPrn)
                        {
                            // ページ計印字
                            if (!openStatus)
                            {
                                //oxlsSheetPrn.Cells[47, 16] = _nebikigo.ToString();  // 小計
                                //oxlsSheetPrn.Cells[48, 16] = _tax.ToString();       // 消費税
                                //oxlsSheetPrn.Cells[49, 16] = _seikyu.ToString();    // 合計

                                oxlsSheetPrn.Cells[47, 16] = "*****";   // 小計
                                oxlsSheetPrn.Cells[48, 16] = "*****";   // 消費税
                                oxlsSheetPrn.Cells[49, 16] = "*****";   // 合計

                                // 備考
                                oxlsSheetPrn.Cells[47, 4] = txtMemo.Text;

                                // 営業担当
                                oxlsSheetPrn.Cells[57, 15] = txtEiTantou.Text;

                                //// ページ計初期化 2016/11/16 コメント化
                                //_nebikigo = 0;
                                //_tax = 0;
                                //_seikyu = 0;
                            }

                            // ページ数加算
                            iX++;
                            _Pages = iX;

                            // ページシートを追加する
                            oxlsSheet.Copy(Type.Missing, oXlsBookPrn.Sheets[_Pages]);

                            // 印刷用カレントシート
                            oxlsSheetPrn = (Excel.Worksheet)oXlsBookPrn.Sheets[_Pages + 1];

                            // 対象行をゼロにしページを初期化
                            _Rows = 0;

                            // 納品書№
                            seikyuNum = dataGridView2[colID, i].Value.ToString();
                            oxlsSheetPrn.Cells[1, 16] = "№ " + seikyuNum;

                            // 発行日
                            if (dtNouhin.Checked)
                            {
                                oxlsSheetPrn.Cells[2, 15] = dtNouhin.Value.ToShortDateString();
                            }
                            else
                            {
                                oxlsSheetPrn.Cells[2, 15] = string.Empty;
                            }

                            //// ページ数
                            //oxlsSheetPrn.Cells[3, 16] = "(" + iX.ToString() + "/" + it.明細数.ToString() + ")";

                            corp = string.Empty;

                            // 請求先住所
                            oxlsSheetPrn.Cells[2, 3] = "〒 " + txtZipCode.Text;
                            oxlsSheetPrn.Cells[3, 3] = txtsAdd.Text;
                            oxlsSheetPrn.Cells[4, 3] = "";

                            // 請求先名
                            corp = txtsName.Text;

                            // 部署・担当者名
                            string tn = (txtsBusho.Text + " " + txtsTantou.Text).Trim();

                            if (tn != string.Empty)
                            {
                                oxlsSheetPrn.Cells[8, 3] = tn + " 様";
                                oxlsSheetPrn.Cells[6, 3] = corp;
                            }
                            else
                            {
                                oxlsSheetPrn.Cells[8, 3] = string.Empty;
                                oxlsSheetPrn.Cells[6, 3] = corp + " 御中";
                            }

                            //// 1ページ目 2016/11/16 コメント化
                            //if (_Pages == 1)
                            //{
                            //    // 合計請求金額
                            //    oxlsSheetPrn.Cells[22, 5] = tlUriage.ToString("#,0");
                            //}
                            //else
                            //{
                            //    // 合計請求金額
                            //    oxlsSheetPrn.Cells[22, 5] = "********************";
                            //}

                            //// 支払期日 2016/11/16 コメント化
                            //if (dtNyukin.Checked)
                            //{
                            //    oxlsSheetPrn.Cells[55, 4] = dtNyukin.Value.ToShortDateString();
                            //}
                            //else
                            //{
                            //    oxlsSheetPrn.Cells[55, 4] = string.Empty;
                            //}

                            // 開始ステータス
                            openStatus = false;
                        }

                        // 日付
                        if (dataGridView2[colDtHaifuS, i].Value != null && dataGridView2[colDtHaifuS, i].Value.ToString() != string.Empty)
                        {
                            oxlsSheetPrn.Cells[sIr + _Rows, 1] = dataGridView2[colDtHaifuS, i].Value.ToString().Substring(5, 2) + "." + dataGridView2[colDtHaifuS, i].Value.ToString().Substring(8, 2);
                        }
                        else
                        {
                            oxlsSheetPrn.Cells[sIr + _Rows, 1] = "";
                        }

                        // 受注内容
                        oxlsSheetPrn.Cells[sIr + _Rows, 3] = dataGridView2[colJyuchu, i].Value.ToString();

                        // サイズ
                        oxlsSheetPrn.Cells[sIr + _Rows, 5] = dataGridView2[colSize, i].Value.ToString();

                        // チラシ名
                        oxlsSheetPrn.Cells[sIr + _Rows, 6] = dataGridView2[colNaiyou, i].Value.ToString();
                        
                        // 単価
                        oxlsSheetPrn.Cells[sIr + _Rows, 14] = dataGridView2[colTanka, i].Value.ToString();

                        // 数量
                        oxlsSheetPrn.Cells[sIr + _Rows, 15] = Utility.strToInt(dataGridView2[colMaisu, i].Value.ToString()).ToString("#,0");

                        // 税込売上金額
                        decimal kingaku = Utility.strToDecimal(dataGridView2[colKingaku, i].Value.ToString());
                        oxlsSheetPrn.Cells[sIr + _Rows, 16] = kingaku.ToString("#,0");

                        // 行加算
                        _Rows++;

                        // 値引額があるとき値引行を印字します
                        decimal nebiki = Utility.strToDecimal(dataGridView2[colNebiki, i].Value.ToString());

                        if (nebiki > 0)
                        {
                            // 明細内容
                            oxlsSheetPrn.Cells[sIr + _Rows, 6] = "値引";
                            oxlsSheetPrn.Cells[sIr + _Rows, 14] = (nebiki * (-1)).ToString("#,0");
                            oxlsSheetPrn.Cells[sIr + _Rows, 15] = "1";
                            oxlsSheetPrn.Cells[sIr + _Rows, 16] = (nebiki * (-1)).ToString();

                            // 行加算
                            _Rows++;
                        }

                        // 合計加算
                        _nebikigo += (kingaku - nebiki);    // 小計
                        _tax += Utility.strToDecimal(dataGridView2[colTax, i].Value.ToString());     // 消費税
                        _seikyu += Utility.strToDecimal(dataGridView2[colUriage, i].Value.ToString());   // 合計
                    }

                    // 合計印字
                    oxlsSheetPrn.Cells[47, 16] = _nebikigo.ToString();  // 小計
                    oxlsSheetPrn.Cells[48, 16] = _tax.ToString();       // 消費税
                    oxlsSheetPrn.Cells[49, 16] = _seikyu.ToString();    // 合計

                    // 備考
                    oxlsSheetPrn.Cells[47, 4] = txtMemo.Text;

                    // 営業担当
                    oxlsSheetPrn.Cells[57, 15] = txtEiTantou.Text;

                    // 印刷用BOOKの1番目のシートは削除する
                    ((Excel.Worksheet)oXlsBookPrn.Sheets[1]).Delete();

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    // 印刷
                    oXlsBookPrn.PrintOutEx(1, Type.Missing, 1, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    // ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "納品書発行";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = "納品書";
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    // ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBookPrn.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    // 処理終了メッセージ
                    MessageBox.Show("終了しました", "納品書発行", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "納品書発行", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                MessageBox.Show(e.Message, "納品書作成", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }
    }
}

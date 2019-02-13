using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmSeikyuItem2015 : Form
    {
        public frmSeikyuItem2015(int sID)
        {
            InitializeComponent();

            _sID = sID;

            // データ読み込み
            jAdp.Fill(dts.新請求書);
            cAdp.Fill(dts.得意先);
            sAdp.Fill(dts.社員);
            aAdp.Fill(dts.受注1);
        }

        // 請求書№
        int _sID;

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.新請求書TableAdapter jAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter aAdp = new darwinDataSetTableAdapters.受注1TableAdapter();

        darwinDataSet.新請求書Row nr;

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        #region グリッドビューカラム定義
        string colID = "col8";          // ID
        string colNaiyo = "col0";       // 内容
        string colTanka = "col1";       // 単価
        string colSuu = "col2";         // 数量
        string colKingaku = "col4";     // 小計
        string colTax = "col5";         // 消費税
        string colNebiki = "col7";      // 値引
        string colZeikomi = "col6";     // 税込金額
        #endregion
        
        const string CHKON = "1";
        const string CHKOFF = "0";

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // データグリッドビューの定義
            gridSetting(dataGridView1);

            // データグリッドビューデータ表示
            gridShow(dataGridView1, _sID);

            // 画面初期化
            //dispClear();
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
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

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
                //tempDGV.Height = 582;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列指定
                tempDGV.Columns.Add(colID, "受注番号");
                tempDGV.Columns.Add(colNaiyo, "内容");
                tempDGV.Columns.Add(colTanka, "単価");
                tempDGV.Columns.Add(colSuu, "数量");
                tempDGV.Columns.Add(colKingaku, "小計");
                tempDGV.Columns.Add(colNebiki, "値引");
                tempDGV.Columns.Add(colTax, "消費税");
                tempDGV.Columns.Add(colZeikomi, "税込金額");

                // 各列幅指定
                tempDGV.Columns[colNaiyo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colTanka].Width = 80;
                tempDGV.Columns[colSuu].Width = 80;
                tempDGV.Columns[colKingaku].Width = 100;
                tempDGV.Columns[colTax].Width = 80;
                tempDGV.Columns[colNebiki].Width = 80;
                tempDGV.Columns[colZeikomi].Width = 100;

                tempDGV.Columns[colTanka].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colSuu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colTax].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colNebiki].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colZeikomi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                // 列固定
                //tempDGV.Columns[colClient].Frozen = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                //foreach (DataGridViewColumn d in tempDGV.Columns)
                //{
                //    if (d.Name == colSel)
                //    {
                //        d.ReadOnly = false;
                //    }
                //    else
                //    {
                //        d.ReadOnly = true;
                //    }
                //}

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
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                //tempDGV.BorderStyle = BorderStyle.Fixed3D;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
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
        /// <param name="sID">
        ///         新請求書ID</param>
        /// ------------------------------------------------------------------
        private void gridShow(DataGridView g, int sID)
        {
            if (!dts.新請求書.Any(a => a.ID == sID))
            {
                return;
            }

            nr = dts.新請求書.Single(a => a.ID == sID);

            lblNum.Text = nr.ID.ToString();
            lblClientCode.Text = nr.得意先ID.ToString();

            if (nr.得意先Row != null)
            {
                lblClientName.Text = nr.得意先Row.略称;
            }
            else
            {
                lblClientName.Text = string.Empty;
            }

            lblHDt.Text = nr.請求書発行日.ToShortDateString();
            lblSDt.Text = nr.支払期日.ToShortDateString();
            lblSeikyu.Text = nr.請求金額.ToString("#,##0");
            lblUriage.Text = nr.売上金額.ToString("#,##0");
            lblTax.Text = nr.消費税.ToString("#,##0");
            lblNebiki.Text = nr.値引額.ToString("#,##0");

            if (nr.Is備考Null())
            {
                txtMemo.Text = string.Empty;
            }
            else
            {
                txtMemo.Text = nr.備考;
            }
            
            g.Rows.Clear();
            int iX = 0;

            foreach (var t in nr.Get受注1Rows())
            {
                g.Rows.Add();

                g[colID, iX].Value = t.ID;
                g[colNaiyo, iX].Value = t.チラシ名;
                g[colTanka, iX].Value = t.単価.ToString("#,##0.00");
                g[colSuu, iX].Value = t.枚数.ToString("#,##0");
                g[colKingaku, iX].Value = t.金額.ToString("#,##0");
                g[colTax, iX].Value = t.消費税.ToString("#,##0");
                g[colNebiki, iX].Value = t.値引額.ToString("#,##0");
                g[colZeikomi, iX].Value = t.売上金額.ToString("#,##0");

                iX++;
            }
            
            g.CurrentCell = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダ、タグデータ登録 </summary>
        /// <param name="sMode">
        ///     処理モード</param>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// -------------------------------------------------------------------------
        private void dataUpdate(int sMode, int sID)
        {
            //// 新規登録
            //if (sMode == 0)
            //{
            //    darwinDataSet.受注番号採番Row r = dts.受注番号採番.New受注番号採番Row();
            //    r.受注番号 =  Int64.Parse(lblOrderNum.Text);
            //    r.入庫日 = DateTime.Parse(dtNyuko.Value.ToShortDateString());

            //    if (cmbClient.SelectedIndex == -1)
            //    {
            //        r.得意先ID = 0;
            //    }
            //    else
            //    {
            //        Utility.ComboClient cmb = (Utility.ComboClient)cmbClient.SelectedItem;
            //        r.得意先ID = cmb.ID;
            //    }

            //    r.確定書入力 = 0;
            //    r.確定書入力日付 = DateTime.Parse("1900/01/01");
            //    r.確定書入力ユーザーID = 0;
            //    r.備考 = txtMemo.Text;
            //    r.登録年月日 = DateTime.Now;
            //    r.更新年月日 = DateTime.Now;
            //    r.ユーザーID = global.loginUserID;

            //    dts.受注番号採番.Add受注番号採番Row(r);
            //}
            
            //// データベース更新
            //adp.Update(dts.受注番号採番);

            //// データ読み込み
            //adp.Fill(dts.受注番号採番);
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmLoginType_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 備考書き換え
            nr.備考 = txtMemo.Text;
            jAdp.Update(dts.新請求書);

            // 後片付け
            this.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 画面初期化
            //dispClear();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        
    }
}

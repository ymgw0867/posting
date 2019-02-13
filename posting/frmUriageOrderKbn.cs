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
    public partial class frmUriageOrderKbn : Form
    {
        public frmUriageOrderKbn()
        {
            InitializeComponent();

            // データ読み込み
            uAdp.Fill(dts.受注種別);
            adp.Fill(dts.受注内容別売上);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注内容別売上TableAdapter adp = new darwinDataSetTableAdapters.受注内容別売上TableAdapter();
        darwinDataSetTableAdapters.受注種別TableAdapter uAdp = new darwinDataSetTableAdapters.受注種別TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        #region グリッドビューカラム定義
        string colID = "col1";
        string colClient = "col2";
        #endregion

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);
                        
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

                // 行・列を初期化
                tempDGV.Rows.Clear();
                tempDGV.Columns.Clear();

                // 各列幅指定
                tempDGV.Columns.Add(colID, "ID");
                tempDGV.Columns.Add(colClient, "受注内容");

                tempDGV.Columns[colID].Width = 40;
                tempDGV.Columns[colClient].Width = 160;
                
                tempDGV.Columns[colID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colClient].Frozen = true;

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
            // データグリッドビューの定義
            gridSetting(dataGridView1);
            gridloadOrderKbn(dataGridView1);

            // グリッドビュー
            DateTime sDt = DateTime.Parse("1900/01/01");
            DateTime sDate = DateTime.Parse(txtYear.Text + "/" + txtMonth.Text + "/01");

            int dYY = 0;
            int dMM = 0;
            decimal tl = 0;
            string colnm = "";


            var s = dts.受注内容別売上.OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.受注種別ID);
            
            // 範囲開始年月
            if (txtYear.Text != string.Empty && txtMonth.Text != string.Empty)
            {
                int yymm = Utility.strToInt(txtYear.Text) * 100 + Utility.strToInt(txtMonth.Text);
                s = s.Where(a => a.年 * 100 + a.月 >= yymm).OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.受注種別ID);
            }

            // 範囲終了年月
            if (txtYear2.Text != string.Empty && txtMonth2.Text != string.Empty)
            {
                int yymm = Utility.strToInt(txtYear2.Text) * 100 + Utility.strToInt(txtMonth2.Text);
                s = s.Where(a => a.年 * 100 + a.月 <= yymm).OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.受注種別ID);
            }

            // グリッドに指定期間の明細表示
            foreach (var t in s)
            {
                // 
                if (dYY != t.年 || dMM != t.月)
                {
                    // 新しい列
                    colnm = "col" + t.年.ToString() + t.月.ToString();
                    g.Columns.Add(colnm, t.年.ToString() + "年" + t.月.ToString().PadLeft(2, '0') + "月");
                    g.Columns[colnm].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    g.Columns[colnm].Width = 110;
                    tl = 0;
                }
                
                foreach (DataGridViewRow j in g.Rows)
                {
                    if (j.Cells[colID].Value.ToString() == t.受注種別ID.ToString())
                    {
                        // 受注種別コード
                        g[colnm, j.Index].Value = t.売上金額.ToString("#,###");

                        // 月合計
                        tl += t.売上金額;
                        g[colnm, g.Rows.Count - 1].Value = tl.ToString("#,###");
                        break;
                    }
                }

                dYY = t.年;
                dMM = t.月;
            }            

            // 総計
            if (g.RowCount > 0)
            {
                g.CurrentCell = null;
                button2.Enabled = true;

                // カレント行
                g.CurrentCell = g[col, row];
                g.CurrentCell = null;
            }
            else
            {
                MessageBox.Show("対象となるデータがありませんでした", "月別受注内容別請求金額", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
            }
        }
        
        ///---------------------------------------------------------------------
        /// <summary>
        ///     受注種別のIDと名称をグリッドにロードする </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        ///---------------------------------------------------------------------
        private void gridloadOrderKbn(DataGridView g)
        {
            foreach (var t in dts.受注種別.OrderBy(a => a.ID))
            {
                g.Rows.Add();
                g[colID, g.Rows.Count - 1].Value = t.ID;
                g[colClient, g.Rows.Count - 1].Value = t.名称;
            }

            if (g.RowCount > 0)
            {
                g.Rows.Add();
                g[colID, g.Rows.Count - 1].Value = "";
                g[colClient, g.Rows.Count - 1].Value = "合　　計";

                g.CurrentCell = null;
            }
        }

        /// -------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// -------------------------------------------------------------
        private void dispClear()
        {
            fMode.Mode = 0;
            fMode.ID = 0;
            txtYear.Text = DateTime.Now.AddMonths(-11).Year.ToString();
            txtMonth.Text = DateTime.Now.AddMonths(-11).Month.ToString();
            txtYear2.Text = DateTime.Now.Year.ToString();
            txtMonth2.Text = DateTime.Now.Month.ToString();
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
            if (Utility.strToInt(txtYear.Text) == 0)
            {
                MessageBox.Show("年が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return;
            }

            if (Utility.strToInt(txtMonth.Text) == 0 || Utility.strToInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("月が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            // データグリッドビューデータ表示
            gridShow(dataGridView1, 0, 0);
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "月別受注内容別請求金額");
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

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
    public partial class frmUriageGaichu : Form
    {
        public frmUriageGaichu()
        {
            InitializeComponent();

            // データ読み込み
            adp.Fill(dts.外注先別粗利表);
            gAdp.Fill(dts.外注先);
        }

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.外注先別粗利表TableAdapter adp = new darwinDataSetTableAdapters.外注先別粗利表TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        #region グリッドビューカラム定義
        string colID = "col1";
        string colClient = "col2";
        string colUri = "col3";
        string colTotal = "colTl";
        #endregion

        public class uriGaichu
        {
            public int ID = 0;
            public int yymm = 0;
            public decimal uriage = 0;
            public decimal genka = 0;
            public decimal arari = 0;
        }

        uriGaichu[] ug = null;
        
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
                tempDGV.Rows.Clear();
                tempDGV.Columns.Clear();

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
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 各列幅指定
                //tempDGV.Columns.Add(colDt, "日付");

                tempDGV.Columns.Add(colID, "ID");
                tempDGV.Columns.Add(colClient, "外注先");
                tempDGV.Columns.Add(colUri, "");

                DateTime dt1 = DateTime.Parse(txtYear.Text + "/" + txtMonth.Text + "/01");
                DateTime dt2 = DateTime.Parse(txtYear2.Text + "/" + txtMonth2.Text + "/01");
                DateTime dt3 = dt1;

                for (int i = 0; dt3 <= dt2; i++)
                {
                    dt3 = dt1.AddMonths(i);
                    string colyymm = (dt3.Year * 100 + dt3.Month).ToString();
                    tempDGV.Columns.Add(colyymm, dt3.Year.ToString() + "年" + dt3.Month + "月");
                    tempDGV.Columns[colyymm].Width = 110;
                    tempDGV.Columns[colyymm].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                    // 合計エリア
                    Array.Resize(ref ug, i + 1);
                    ug[i] = new uriGaichu();
                    ug[i].yymm = dt3.Year * 100 + dt3.Month;
                }

                tempDGV.Columns.Add(colTotal, "合計");
                tempDGV.Columns[colTotal].Width = 110;

                tempDGV.Columns[colID].Width = 40;
                tempDGV.Columns[colClient].Width = 120;
                tempDGV.Columns[colUri].Width = 40;
                tempDGV.Columns[colTotal].Width = 110;

                tempDGV.Columns[colID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colUri].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colTotal].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                tempDGV.Columns[colUri].Frozen = true;

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
                //tempDGV.BorderStyle = BorderStyle.Fixed3D;
                //tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical;
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
            // カーソルを待機にする
            Cursor = Cursors.WaitCursor;

            // グリッドビュー
            g.Rows.Clear();
            
            int dt1 = Utility.strToInt(txtYear.Text) * 100 + Utility.strToInt(txtMonth.Text);  
            int dt2 = Utility.strToInt(txtYear2.Text) * 100 + Utility.strToInt(txtMonth2.Text);

            foreach (var it in dts.外注先.OrderBy(a => a.ID))
            {
                // 指定以外の外注先はネグる
                if (txtsClient.Text != string.Empty)
                {
                    if (!it.名称.Contains(txtsClient.Text))
                    {
                        continue;
                    }
                }

                g.Rows.Add(3);
                //g.Rows[g.RowCount - 1].DefaultCellStyle.BackColor = SystemColors.ControlLight;
                g[colID, g.RowCount - 3].Value = it.ID.ToString();
                g[colClient, g.RowCount - 3].Value = it.名称;

                decimal tUri = 0;
                decimal tGen = 0;
                decimal tArari = 0;

                foreach (var t in dts.外注先別粗利表.Where(a => a.外注先ID == it.ID && (a.年 * 100 + a.月) >= dt1 && (a.年 * 100 + a.月) <= dt2))
                {
                    g[colUri, g.RowCount - 3].Value = "売上";
                    g[colUri, g.RowCount - 2].Value = "原価";
                    g[colUri, g.RowCount - 1].Value = "粗利";

                    foreach (DataGridViewColumn c in g.Columns)
                    {
                        if (c.Name == (t.年 * 100 + t.月).ToString())
                        {
                            g[c.Name, g.RowCount - 3].Value = t.売上金額.ToString("#,0");
                            g[c.Name, g.RowCount - 2].Value = t.原価.ToString("#,0");
                            g[c.Name, g.RowCount - 1].Value = t.粗利.ToString("#,0");

                            tUri += t.売上金額;
                            tGen += t.原価;
                            tArari += t.粗利;

                            // 月別合計加算
                            foreach (var tu in ug)
                            {
                                if (c.Name == tu.yymm.ToString())
                                {
                                    tu.uriage += t.売上金額;
                                    tu.genka += t.原価;
                                    tu.arari += t.粗利;

                                    break;
                                }
                            }
                        }
                    }

                    // 合計
                    g[colTotal, g.RowCount - 3].Value = tUri.ToString("#,0");
                    g[colTotal, g.RowCount - 2].Value = tGen.ToString("#,0");
                    g[colTotal, g.RowCount - 1].Value = tArari.ToString("#,0");
                }
            }                

            // 総計
            if (g.RowCount > 0)
            {
                g.Rows.Add(3);
                //g.Rows[g.RowCount - 1].DefaultCellStyle.BackColor = SystemColors.ControlLight;
                g[colClient, g.RowCount - 3].Value = "合計";
                g[colUri, g.RowCount - 3].Value = "売上";
                g[colUri, g.RowCount - 2].Value = "原価";
                g[colUri, g.RowCount - 1].Value = "粗利";


                foreach (DataGridViewColumn c in g.Columns)
                {
                    // 月別合計
                    foreach (var tu in ug)
                    {
                        if (c.Name == tu.yymm.ToString())
                        {
                            g[c.Name, g.RowCount - 3].Value = tu.uriage.ToString("#,0");
                            g[c.Name, g.RowCount - 2].Value = tu.genka.ToString("#,0");
                            g[c.Name, g.RowCount - 1].Value = tu.arari.ToString("#,0");
                        }
                    }
                }

                // 総合計
                g[colTotal, g.RowCount - 3].Value = ug.Sum(a => a.uriage).ToString("#,0");
                g[colTotal, g.RowCount - 2].Value = ug.Sum(a => a.genka).ToString("#,0");
                g[colTotal, g.RowCount - 1].Value = ug.Sum(a => a.arari).ToString("#,0");

                g.CurrentCell = null;
                button2.Enabled = true;

                // カレント行
                g.CurrentCell = g[col, row];
                g.CurrentCell = null;
            }
            else
            {
                MessageBox.Show("対象となるデータがありませんでした", "外注先別粗利表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
            }

            // カーソルを戻す
            Cursor = Cursors.Default;
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

        private void txtKingaku_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime dt;
            if (!DateTime.TryParse(txtYear.Text + "/" + txtMonth.Text + "/01", out dt))
            {
                MessageBox.Show("開始年月が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return;
            }

            if (!DateTime.TryParse(txtYear2.Text + "/" + txtMonth2.Text + "/01", out dt))
            {
                MessageBox.Show("終了年月が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear2.Focus();
                return;
            }

            // データグリッドビューの定義
            gridSetting(dataGridView1);

            // 売掛残高表示
            dataShow();
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     売掛残高表示 </summary>
        ///--------------------------------------------------------
        private void dataShow()
        {
            // データグリッドビューデータ表示
            gridShow(dataGridView1, 0, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "外注先別粗利表");
        }

        private void txtMonth_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // 1行目や列ヘッダ、行ヘッダの場合は何もしない
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            // 粗利行を対象とする
            if (((e.RowIndex + 1) % 3) != 0)
            {
                if (e.ColumnIndex < 2)
                {
                    // セルの下側の境界線を「一重線の境界線」に設定
                    e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                }
            }
            else
            {
                // セルの上側の境界線を既定の境界線に設定
                e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                //e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.OutsetDouble;

                e.AdvancedBorderStyle.Bottom = dataGridView1.AdvancedCellBorderStyle.Bottom;
                //dataGridView1[e.ColumnIndex, e.RowIndex].Style.BackColor = SystemColors.ControlLight;                
            }
        }
    }
}

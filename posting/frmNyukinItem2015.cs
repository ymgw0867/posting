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
    public partial class frmNyukinItem2015 : Form
    {
        public frmNyukinItem2015(int sID)
        {
            InitializeComponent();

            _sID = sID;

            // データ読み込み
            jAdp.Fill(dts.新請求書);
            cAdp.Fill(dts.得意先);
            sAdp.Fill(dts.社員);
            nAdp.Fill(dts.新入金);

            fMode.Mode = global.FLGOFF;
        }

        // 請求書№
        int _sID;

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.新請求書TableAdapter jAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        darwinDataSetTableAdapters.新入金TableAdapter nAdp = new darwinDataSetTableAdapters.新入金TableAdapter();

        darwinDataSet.新請求書Row nr;

        // フォームモード
        Utility.formMode fMode = new Utility.formMode();

        #region グリッドビューカラム定義
        string colDate = "col8";        // 日付
        string colKingaku = "col4";     // 金額
        string colMemo = "col2";        // 備考
        string colBtn = "col1";         // 削除ボタン
        string colSel = "col3";         // 選択ボタン
        string colID = "col0";          // ID
        #endregion
        
        const string CHKON = "1";
        const string CHKOFF = "0";

        string[] kouzaArray = null;

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // データグリッドビューの定義
            gridSetting(dataGridView1);

            // 画面初期化
            dispClear();

            // 口座コンボボックスロード : 2017/08/08
            loadKouzaCsv();

            // データ表示
            gridShow(dataGridView1, _sID);

        }

        private void loadKouzaCsv()
        {
            int i = 0;

            string cpath = System.IO.Directory.GetCurrentDirectory();
            var s = System.IO.File.ReadAllLines(cpath + @"\口座.csv", Encoding.Default);
            foreach (var stBuffer in s)
            {
                // カンマ区切りで分割して配列に格納する
                string[] stCSV = stBuffer.Split(',');

                Array.Resize(ref kouzaArray, i + 1);
                kouzaArray[i] = stCSV[0];
                i++;
            }

            for (int iX = 0; iX < kouzaArray.Length; iX++)
            {
                cmbKouza.Items.Add(kouzaArray[iX]);
            }
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
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "選択";
                tempDGV.Columns.Add(btn);
                tempDGV.Columns[0].HeaderText = "";
                tempDGV.Columns[0].Name = colSel;

                tempDGV.Columns.Add(colDate, "日付");
                tempDGV.Columns.Add(colKingaku, "金額");
                tempDGV.Columns.Add(colMemo, "備考");

                DataGridViewButtonColumn btn2 = new DataGridViewButtonColumn();
                btn2.UseColumnTextForButtonValue = true;
                btn2.Text = "削除";
                tempDGV.Columns.Add(btn2);
                tempDGV.Columns[4].HeaderText = "";
                tempDGV.Columns[4].Name = colBtn;

                tempDGV.Columns.Add(colID, "SID");
                tempDGV.Columns[colID].Visible = false;

                // 各列幅指定
                tempDGV.Columns[colSel].Width = 50;
                tempDGV.Columns[colDate].Width = 100;
                tempDGV.Columns[colKingaku].Width = 90;
                tempDGV.Columns[colBtn].Width = 50;
                tempDGV.Columns[colMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colKingaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[colMemo].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
            lblZan.Text = nr.残金.ToString("#,##0");

            if (nr.入金完了 == global.FLGON)
            {
                checkBox1.Checked = true;
                label6.Visible = true;
            }
            else
            {
                checkBox1.Checked = false;
                label6.Visible = false;
            }

            // 入金済みメッセージ
            if (nr.入金完了 == global.FLGON && nr.残金 > 0)
            {
                label6.Text = "未収確定";
            }
            else if (nr.入金完了 == global.FLGON && nr.残金 == 0)
            {
                label6.Text = "入金完了";
            }
            else
            {
                label6.Text = string.Empty;
            }

            if (nr.Is備考Null())
            {
                txtSeikyuMemo.Text = string.Empty;
            }
            else
            {
                txtSeikyuMemo.Text = nr.備考;
            }

            // 入金日
            DateTime dt;
            if (DateTime.TryParse(lblSDt.Text, out dt))
            {
                dateTimePicker1.Value = dt;
            }
            else
            {
                dateTimePicker1.Value = DateTime.Today;
            }

            // 無効な請求書
            if (nr.無効 == global.FLGON)
            {
                lblMukou.Visible = true;
            }
            else
            {
                lblMukou.Visible = false;
            }

            // 精算日付
            if (nr.精算日付 == string.Empty)
            {
                dateTimeSai.Checked = false;
            }
            else
            {
                if (DateTime.TryParse(nr.精算日付, out dt))
                {
                    dateTimeSai.Checked = true;
                    dateTimeSai.Value =dt;
                }
                else
                {
                    dateTimeSai.Checked = false;
                }
            }

            // 精算額
            txtSai.Text = nr.精算額.ToString();

            // 精算備考
            txtSaiMemo.Text = nr.精算備考.ToString();

            // 口座 : 2017/08/15
            if (Utility.nullToStr(nr.口座) == string.Empty)
            {
                cmbKouza.SelectedIndex = -1;
            }
            else
            {
                cmbKouza.SelectedIndex = -1;

                for (int i = 0; i < kouzaArray.Length; i++)
                {
                    if (kouzaArray[i] == nr.口座)
                    {
                        cmbKouza.SelectedIndex = i;
                        break;
                    }
                }
            }

            g.Rows.Clear();
            int iX = 0;

            foreach (var t in nr.Get新入金Rows())
            {
                g.Rows.Add();

                g[colDate, iX].Value = t.入金年月日.ToShortDateString();
                g[colKingaku, iX].Value = t.金額.ToString("#,##0");
                g[colMemo, iX].Value = t.備考;
                g[colID, iX].Value = t.ID.ToString();

                iX++;
            }
            
            g.CurrentCell = null;
            button2.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // データ未更新：2017/08/15
            dataProperty = false;


            this.Close();
        }

        private void frmLoginType_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            //this.Dispose();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                label6.Visible = true;
            }
            else
            {
                label6.Visible = false;
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            int zan = Utility.strToInt(lblSeikyu.Text) - getNyukingaku(Utility.strToInt(lblNum.Text));

            // 編集のとき
            if (fMode.Mode == global.FLGON)
            {
                zan += fMode.kingaku;
            }
            
            lblZan.Text = (zan - Utility.strToInt(txtNyukin.Text)).ToString("#,##0");

            // 残金がゼロになったら入金済みとする
            if (Utility.strToInt(lblZan.Text) == 0)
            {
                checkBox1.Checked = true;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (!errCheck())
            {
                return;
            }

            int sMukou = global.FLGOFF;

            // 入金データ登録
            if (Utility.strToInt(txtNyukin.Text) > 0)
            {
                nyukinUpdate();
            }

            // 請求書データ更新（入金完了フラグ、無効フラグ）
            nyukinFinUpdate(sMukou);

            // データ更新：2017/08/15
            dataProperty = true;

            // 閉じる
            this.Close();
        }
        
        private void dispClear()
        {
            txtNyukin.Text = string.Empty;
            txtMemo.Text = string.Empty;
            checkBox1.Checked = false;
            txtSeikyuMemo.Text = string.Empty;
            button2.Enabled = false;
        }

        private bool errCheck()
        {
            //// 無効チェックがないとき
            //if (!checkBox2.Checked)
            //{
            //}

            // 日付：有効入金額があるとき    :   2016/01/21 撤廃
            //if (Utility.strToInt(txtNyukin.Text) > 0 && dateTimePicker1.Value < DateTime.Parse(lblHDt.Text))
            //{
            //    MessageBox.Show("入金日が請求書発行日以前になっています", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    dateTimePicker1.Focus();
            //    return false;
            //}

            // 入金額未入力を許容：請求書データ単独の書き換えができないため 2015/05/23
            //// 入金額
            //if (Utility.strToInt(txtNyukin.Text) == 0 && !checkBox1.Checked)
            //{
            //    MessageBox.Show("入金額を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    txtNyukin.Focus();
            //    return false;
            //}

            // 精算関連 2016/05/23
            if (dateTimeSai.Checked && Utility.strToInt(txtSai.Text) == 0)
            {
                MessageBox.Show("精算差異額を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSai.Focus();
                return false;
            }

            if (!dateTimeSai.Checked && Utility.strToInt(txtSai.Text) != 0)
            {
                MessageBox.Show("精算日付を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                dateTimeSai.Focus();
                return false;
            }
            
            // 合計入金額
            int kin = getNyukingaku(Utility.strToInt(lblNum.Text));

            // 編集のとき
            if (fMode.Mode == global.FLGON)
            {
                kin -= fMode.kingaku;
            }

            kin += Utility.strToInt(txtNyukin.Text);

            if (Utility.strToInt(lblSeikyu.Text) < kin)
            {
                if (MessageBox.Show("入金額が請求金額を超過していますがよろしいですか？", "エラー", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No)
                {
                    txtNyukin.Focus();
                    return false;
                }
            }

            // 口座 2017/08/10
            if (cmbKouza.SelectedIndex < 0)
            {
                MessageBox.Show("口座を選択してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbKouza.Focus();
                return false;
            }

            return true;
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     入金データ更新 </summary>
        /// ---------------------------------------------------------------
        private void nyukinUpdate()
        {
            if (fMode.Mode == global.FLGOFF)    // 追加モード
            {
                darwinDataSet.新入金Row r = dts.新入金.New新入金Row();
                r.入金年月日 = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                r.金額 = Utility.strToInt(txtNyukin.Text);
                r.登録年月日 = DateTime.Now;
                r.変更年月日 = DateTime.Now;
                r.ユーザーID = global.loginUserID;
                r.請求書ID = Utility.strToInt(lblNum.Text);
                r.備考 = txtMemo.Text;
                r.得意先ID = Utility.strToInt(lblClientCode.Text);
                dts.新入金.Add新入金Row(r);
            }
            else if (fMode.Mode == global.FLGON)    // 編集モード
            {
                var s = dts.新入金.Single(a => a.ID == fMode.ID);

                s.入金年月日 = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                s.金額 = Utility.strToInt(txtNyukin.Text);
                s.備考 = txtMemo.Text;
                s.変更年月日 = DateTime.Now;
                s.ユーザーID = global.loginUserID;
            }

            // データベース更新
            nAdp.Update(dts.新入金);
        }

        /// ------------------------------------------------------
        /// <summary>
        ///     新請求書データ更新 </summary>
        /// <param name="sMukou">
        ///     無効フラグ</param>
        /// ------------------------------------------------------
        private void nyukinFinUpdate(int sMukou)
        {
            var s = dts.新請求書.Single(a => a.ID == Utility.strToInt(lblNum.Text));

            // 無効でないとき
            if (sMukou == 0)
            {
                s.残金 = s.請求金額 - getNyukingaku(Utility.strToInt(lblNum.Text));

                if (checkBox1.Checked)
                {
                    s.入金完了 = global.FLGON;
                }
                else
                {
                    s.入金完了 = global.FLGOFF;
                }

                s.無効 = global.FLGOFF;
            }
            else
            {
                s.無効 = global.FLGON;
            }

            s.変更年月日 = DateTime.Now;
            s.備考 = txtSeikyuMemo.Text;

            // 精算関係 2016/05/23
            if (dateTimeSai.Checked)
            {
                s.精算日付 = dateTimeSai.Value.ToShortDateString();
            }
            else
            {
                s.精算日付 = string.Empty;
            }

            s.精算額 = Utility.strToInt(txtSai.Text);
            s.精算備考 = txtSaiMemo.Text;
            s.口座 = cmbKouza.Text;   // 2017/08/10

            // データベース更新
            jAdp.Update(dts.新請求書);
        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     任意の請求書データの入金済み金額合計を取得する </summary>
        /// <param name="sID">
        ///     新請求書ID</param>
        /// <returns>
        ///     入金済み金額</returns>
        /// -----------------------------------------------------------------
        private int getNyukingaku(int sID)
        {
            int kin = 0;

            var s = dts.新入金
               .Where(a => a.請求書ID == sID)
                   .GroupBy(a => a.請求書ID)
                   .Select(cg => new
                   {
                       cCode = cg.Key,
                       kingaku = cg.Sum(a => a.金額)
                   });

            foreach (var t in s)
            {
                kin = t.kingaku;
                break;
            }

            return kin;
        }

        private void txtNyukin_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void dataShow(int sID)
        {
            var s = dts.新入金.Single(a => a.ID == sID);

            // 編集対象の金額を取得
            fMode.kingaku = Utility.strToInt(s.金額.ToString());

            // データ表示
            dateTimePicker1.Value = s.入金年月日;
            txtNyukin.Text = s.金額.ToString();
            txtMemo.Text = s.備考;

            button2.Enabled = true;

            // ID
            fMode.ID = s.ID;


            // 口座 : 2017/08/15
            //if (Utility.nullToStr(nr.口座) == string.Empty)
            //{
            //    cmbKouza.SelectedIndex = -1;
            //}
            //else
            //{
            //    for (int i = 0; i < cmbKouza.Items.Count; i++)
            //    {
            //        cmbKouza.SelectedIndex = i;

            //        if (cmbKouza.SelectedValue.ToString() == nr.口座)
            //        {
            //            break;
            //        }
            //    }
            //}

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 選択ボタン
            if (e.ColumnIndex == 0)
            {
                // 編集モード
                fMode.Mode = global.FLGON;

                // データ表示
                dataShow(Utility.strToInt(dataGridView1[colID, e.RowIndex].Value.ToString()));            
            }

            // 削除ボタン
            if (e.ColumnIndex == 4)
            {
                if (MessageBox.Show(dataGridView1[colDate, e.RowIndex].Value.ToString() + "の入金情報を削除してよろしいですか？","削除確認",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
             
                // 新入金データ削除
                nyukinDataDelete(Utility.strToInt(dataGridView1[colID, e.RowIndex].Value.ToString()));
            }
        }

        /// -------------------------------------------------------------
        /// <summary>
        ///     新入金データ削除 </summary>
        /// <param name="nID">
        ///     新入金データID</param>
        /// -------------------------------------------------------------
        private void nyukinDataDelete(int nID)
        {
            int kin = 0;

            var s = dts.新入金.Single(a => a.ID == nID);
            kin = s.金額;
            s.Delete();

            // 新入金データ削除
            nAdp.Update(dts.新入金);
            nAdp.Fill(dts.新入金);

            // 新請求書データ更新（残金再計算）
            zanDataUpdate(_sID, kin);

            // データ表示
            gridShow(dataGridView1, _sID);
        }

        /// --------------------------------------------------------------
        /// <summary>
        ///     新請求書データ残金再計算 </summary>
        /// <param name="sID">
        ///     新請求書ID</param>
        /// <param name="sKingaku">
        ///     取り消す入金額</param>
        /// --------------------------------------------------------------
        private void zanDataUpdate(int sID, int sKingaku)
        {
            var s = dts.新請求書.Single(a => a.ID == sID);
            s.残金 += sKingaku;

            // 新請求書データ更新
            jAdp.Update(dts.新請求書);
            jAdp.Fill(dts.新請求書);
        }

        private void txtNyukin_MouseClick(object sender, MouseEventArgs e)
        {
            // 入金欄が空欄のときマウスクリックで残金を自動表示
            if (txtNyukin.Text == string.Empty)
            {
                txtNyukin.Text = Utility.strToInt(lblZan.Text).ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dispClear();

            // データ表示
            gridShow(dataGridView1, _sID);

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void txtSai_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        public bool dataProperty { get; set; }
    }
}

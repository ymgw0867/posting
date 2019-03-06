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
    public partial class frmTax : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.税率 cMaster = new Entity.税率();

        const string MESSAGE_CAPTION = "消費税率";

        public frmTax()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView1);

            // TODO: このコード行はデータを 'darwinDataSet.税率' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            this.税率TableAdapter.Fill(this.darwinDataSet.税率);

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
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // 全体の高さ
                    tempDGV.Height = 180;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    //tempDGV.Columns.Add("col1", "ｺｰﾄﾞ");
                    //tempDGV.Columns.Add("col2", "名称");
                    //tempDGV.Columns.Add("col3", "備考");

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].Width = 60;
                    tempDGV.Columns[2].Width = 120;
                    tempDGV.Columns[3].Width = 260;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.税率 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.税率 sTax = new Control.税率();
                OleDbDataReader dr;

                sqlStr = " where 税率.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = sTax.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.設定税率 = Convert.ToInt32(dr["税率"].ToString());
                        tempC.開始年月日 = Convert.ToDateTime(dr["開始年月日"].ToString());
                        tempC.備考 = dr["備考"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    sTax.Close();
                    return false;
                }

                dr.Close();
                sTax.Close();
                return true;
            }

            //////public static void ShowData(DataGridView tempDGV)
            //////{
            //////    string sqlSTRING = "";

            //////    try
            //////    {
            //////        tempDGV.RowCount = 0;

            //////        //原価名マスターのデータリーダーを取得する
            //////        Control.DataControl dCon = new Control.DataControl();

            //////        sqlSTRING = "select * from m_Costname " +
            //////                    "order by ID";

            //////        dR = dCon.FreeReader(sqlSTRING);

            //////        iX = 0;

            //////        while (dR.Read())
            //////        {
            //////            tempDGV.Rows.Add();

            //////            tempDGV[0, iX].Value = dR["ID"];
            //////            tempDGV[1, iX].Value = NullConvert.Noth(dR["原価名"]);
            //////            tempDGV[2, iX].Value = NullConvert.Noth(dR["備考"]);
            //////            //tempDGV[1, iX].Value = dR["原価名"];
            //////            //tempDGV[2, iX].Value = dR["備考"];
            //////            iX++;
            //////        }

            //////        dR.Close();

            //////        dCon.Close();

            //////    }
            //////    catch (Exception e)
            //////    {
            //////        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //////    }

            //////}

        }

        //グリッドからデータを選択
        private void GridEnter()
        {

            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[1, dataGridView1.SelectedRows[iX].Index].Value + "が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "税率選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                        MessageBox.Show("該当するデータがマスターに登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'データ値を取得
                    txtName1.Text = cMaster.設定税率.ToString();
                    tDate.Value = cMaster.開始年月日;
                    txtMemo.Text = cMaster.備考;

                    //IDテキストボックスは編集不可とする
                    //txtCode.Enabled = false;

                    //ボタン状態
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //フォームモードステータス:変更削除

                    txtName1.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "画面クリア", MessageBoxButtons.OK);
                }
            }

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
            GridEnter();
        }

        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {

            try
            {
                fMode.Mode = 0;

                //txtCode.Enabled = true;
                //txtCode.Text = "";
                txtName1.Text = "";
                tDate.Checked = false;
                txtMemo.Text = "";

                btnDel.Enabled = false;
                btnClr.Enabled = false;

                if (this.dataGridView1.RowCount > 0)
                {
                    btnCsv.Enabled = true;
                }
                else
                {
                    btnCsv.Enabled = false;
                }

                //txtCode.Focus();
                txtName1.Focus();
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

                if (fDataCheck() == true)
                {
                    Control.税率 sTax = new Control.税率();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                sTax.Close();
                                return;
                            }

                            if (sTax.DataInsert(cMaster) == true)
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
                            {
                                sTax.Close();
                                return;
                            }

                            if (sTax.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", "所属マスター", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    sTax.Close();

                    DispClear();

                    //データを 'darwinDataSet.税率' テーブルに読み込みます。
                    this.税率TableAdapter.Fill(this.darwinDataSet.税率);

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

                    // 数字か？
                    if (txtName1.Text == null)
                    {
                        this.txtName1.Focus();
                        throw new Exception("税率は数字で入力してください");
                    }

                    str = this.txtName1.Text;

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtName1.Focus();
                        throw new Exception("税率は数字で入力してください");
                    }

                    // 未入力またはスペースのみは不可
                    if ((this.txtName1.Text).Trim().Length < 1)
                    {
                        this.txtName1.Focus();
                        throw new Exception("税率入力してください");
                    }

                    ////ゼロは不可
                    //if (Convert.ToInt32(this.txtName1.Text.ToString()) == 0)
                    //{
                    //    this.txtName1.Focus();
                    //    throw new Exception("ゼロは登録できません");
                    //}

                //    //登録済みコードか調べる
                //    string sqlStr;
                //    Control.受注種別 sTax = new Control.受注種別();
                //    OleDbDataReader dr;

                //    sqlStr = " where ID = " + txtCode.Text.ToString();
                //    dr = sTax.FillBy(sqlStr);

                //    if (dr.HasRows == true)
                //    {
                //        txtCode.Focus();
                //        dr.Close();
                //        sTax.Close();
                //        throw new Exception("既に登録済みのコードです");
                //    }

                //    dr.Close();
                //    sTax.Close();

                }

                //開始年月日チェック
                if (tDate.Checked == false)
                {
                    tDate.Focus();
                    throw new Exception("摘要日を入力してください");
                }

                //クラスにデータをセット
                cMaster.設定税率 = Convert.ToInt32(txtName1.Text.ToString());
                cMaster.開始年月日 = Convert.ToDateTime(tDate.Text.ToString());
                cMaster.備考 = txtMemo.Text.ToString();

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

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            ////他に登録されているときは削除不可とする
            //string SqlStr;
            //SqlStr = " where ";
            //SqlStr += "(受注.受注種別ID = " + txtCode.Text.ToString() + ")  ";

            //OleDbDataReader dr;
            //Control.受注 Order = new Control.受注();
            //dr = Order.FillBy(SqlStr);

            ////該当種別の受注データが登録されているときは削除不可とする
            //if (dr.HasRows == true)
            //{
            //    MessageBox.Show(txtName1.Text.ToString() + "の受注データが登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    dr.Close();
            //    Order.Close();
            //    return;
            //}

            //dr.Close();
            //Order.Close();
            
            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.税率 sTax = new Control.税率();
            if (sTax.DataDelete(cMaster.ID)==true)
                MessageBox.Show("削除されました", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            sTax.Close();

            DispClear();

            //データを 'darwinDataSet.所属' テーブルに読み込みます。
            this.税率TableAdapter.Fill(this.darwinDataSet.税率);

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

    }
}
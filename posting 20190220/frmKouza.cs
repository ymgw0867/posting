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
    public partial class frmKouza : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.振込口座 cMaster = new Entity.振込口座();

        const string MESSAGE_CAPTION = "振込口座マスター";

        public frmKouza()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            // TODO: このコード行はデータを 'darwinDataSet.振込口座' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);
            this.振込口座TableAdapter.Fill(this.darwinDataSet.振込口座);

            //口座種別セット
            Utility.ComboKouza.load(comboBox1);

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

                    tempDGV.Columns[0].Width = 70;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 200;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.振込口座 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.振込口座 Kouza = new Control.振込口座();
                OleDbDataReader dr;

                sqlStr = " where 振込口座.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Kouza.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.金融機関名 = dr["金融機関名"].ToString() + "";
                        tempC.支店名 = dr["支店名"].ToString();
                        tempC.口座種別 = Convert.ToInt32(dr["口座種別"].ToString());
                        tempC.口座番号 = dr["口座番号"].ToString();
                        tempC.口座名義 = dr["口座名義"].ToString();
                    }
                }
                else
                {
                    dr.Close();
                    Kouza.Close();
                    return false;
                }

                dr.Close();
                Kouza.Close();
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

            if (MessageBox.Show(msgStr, "振込口座選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    //txtCode.Text = cMaster.ID.ToString();
                    txtName1.Text = cMaster.金融機関名;
                    txtName2.Text = cMaster.支店名;

                    Utility.ComboKouza.selectedIndex(comboBox1, Int32.Parse(cMaster.口座種別.ToString()));

                    txtName2.Text = cMaster.支店名;
                    txtNumber.Text = cMaster.口座番号;
                    txtMeigi.Text = cMaster.口座名義;

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

                //txtCode.Enabled = true;
                //txtCode.Text = "";
                txtName1.Text = "";
                txtName2.Text = "";
                comboBox1.SelectedIndex = -1;
                txtNumber.Text = ""; ;
                txtMeigi.Text = "";

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
                    Control.振込口座 Kouza = new Control.振込口座();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Kouza.Close();
                                return;
                            }

                            if (Kouza.DataInsert(cMaster) == true)
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
                                Kouza.Close();
                                return;
                            }

                            if (Kouza.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Kouza.Close();

                    DispClear();

                    //データを 'darwinDataSet.振込口座' テーブルに読み込みます。
                    this.振込口座TableAdapter.Fill(this.darwinDataSet.振込口座);

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
                    //Control.振込口座 Kouza = new Control.振込口座();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Kouza.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Kouza.Close();
                    //    throw new Exception("既に登録済みのコードです");
                    //}

                    //dr.Close();
                    //Kouza.Close();

                }

                //金融機関名チェック
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("金融機関名を入力してください");
                }

                //支店名チェック
                if (txtName2.Text.Trim().Length < 1)
                {
                    txtName2.Focus();
                    throw new Exception("支店名を入力してください");
                }

                //口座種別チェック
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("口座種別を選択してください");
                }

                //口座番号：数字か？
                if (txtNumber.Text == null)
                {
                    this.txtNumber.Focus();
                    throw new Exception("口座番号は数字で入力してください");
                }

                str = this.txtNumber.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtNumber.Focus();
                    throw new Exception("口座番号は数字で入力してください");
                }

                // 未入力またはスペースのみは不可
                if ((this.txtNumber.Text).Trim().Length < 1)
                {
                    this.txtNumber.Focus();
                    throw new Exception("口座番号を入力してください");
                }

                //口座名義チェック
                if (txtMeigi.Text.Trim().Length < 1)
                {
                    txtMeigi.Focus();
                    throw new Exception("口座名義を入力してください");
                }

                ////ゼロは不可
                //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                //{
                //    this.txtCode.Focus();
                //    throw new Exception("ゼロは登録できません");
                //}

                //クラスにデータセット
                //cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.金融機関名 = txtName1.Text.ToString();
                cMaster.支店名 = txtName2.Text.ToString();

                Utility.ComboKouza cmb1 = new Utility.ComboKouza();
                cmb1 = (Utility.ComboKouza)comboBox1.SelectedItem;
                cMaster.口座種別 = cmb1.ID;

                cMaster.口座番号 = txtNumber.Text.ToString();
                cMaster.口座名義 = txtMeigi.Text.ToString(); 

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

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;
            }

            if (sender == txtNumber)
            {
                objtxt = txtNumber;
            }

            if (sender == txtMeigi)
            {
                objtxt = txtMeigi;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {

            TextBox objtxt = new TextBox();

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;
            }

            if (sender == txtNumber)
            {
                objtxt = txtNumber;
            }

            if (sender == txtMeigi)
            {
                objtxt = txtMeigi;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //他に登録されているときは削除不可とする①
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(受注.振込口座ID = " + cMaster.ID.ToString() + ")  ";

            OleDbDataReader dr;
            Control.受注 Order = new Control.受注();
            dr = Order.FillBy(SqlStr);

            //該当振込口座の受注データが登録されているときは削除不可とする
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "の受注データが登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Order.Close();
                return;
            }

            dr.Close();
            Order.Close();

            //他に登録されているときは削除不可とする②
            SqlStr = " where ";
            SqlStr += "(請求書.振込口座ID1 = " + cMaster.ID.ToString() + ") or ";
            SqlStr += "(請求書.振込口座ID2 = " + cMaster.ID.ToString() + ") ";

            Control.請求書 Seikyu = new Control.請求書();
            dr = Seikyu.FillBy(SqlStr);

            //該当振込口座の受注データが登録されているときは削除不可とする
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "の請求データが登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Seikyu.Close();
                return;
            }

            dr.Close();
            Seikyu.Close();

            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.振込口座 Kouza = new Control.振込口座();
            if (Kouza.DataDelete(cMaster.ID)==true)
                MessageBox.Show("削除されました", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Kouza.Close();

            DispClear();

            //データを 'darwinDataSet.振込口座' テーブルに読み込みます。
            this.振込口座TableAdapter.Fill(this.darwinDataSet.振込口座);

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
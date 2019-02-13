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
    public partial class frmShain : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.社員 cMaster = new Entity.社員();

        const string MESSAGE_CAPTION = "社員マスター";

        public frmShain()
        {
            InitializeComponent();
        }

        private void Gengo_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: このコード行はデータを 'darwinDataSet.社員' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);
            this.社員TableAdapter.Fill(this.darwinDataSet.社員);

            //所属コンボボックスデータロード
            Utility.ComboShozoku.load(this.comboBox1);

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
                    tempDGV.Columns[1].Width = 160;
                    tempDGV.Columns[2].Width = 160;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 160;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 200;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.社員 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.社員 Shain = new Control.社員();
                OleDbDataReader dr;

                sqlStr = " where 社員.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Shain.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.氏名 = dr["氏名"].ToString() + "";
                        tempC.フリガナ = dr["フリガナ"].ToString() + "";
                        tempC.所属コード = Int32.Parse(dr["所属コード"].ToString());
                        tempC.役職 = dr["役職"].ToString() + "";
                        tempC.入社年月日 = dr["入社年月日"].ToString();
                        tempC.備考 = dr["備考"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    Shain.Close();
                    return false;
                }

                dr.Close();
                Shain.Close();
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

            if (MessageBox.Show(msgStr, "社員選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    //txtID.Text = cCostname.ID.ToString();
                    //txtName.Text = cCostname.Name;
                    //txtNote.Text = cCostname.Note;

                    txtCode.Text = cMaster.ID.ToString();
                    txtName.Text = cMaster.氏名;
                    txtFuri.Text = cMaster.フリガナ;

                    Utility.ComboShozoku.selectedIndex(comboBox1,cMaster.所属コード);

                    txtYaku.Text = cMaster.役職;

                    if (cMaster.入社年月日 == "")
                    {
                        iDate.Checked = false;
                    }
                    else
                    {
                        iDate.Checked = true;
                        iDate.Value = Convert.ToDateTime(cMaster.入社年月日);
                    }

                    txtMemo.Text = cMaster.備考;

                    //IDテキストボックスは編集不可とする
                    txtCode.Enabled = false;

                    //ボタン状態
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //フォームモードステータス:変更削除

                    txtName.Focus();
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

                txtCode.Enabled = true;
                txtCode.Text = "";
                txtName.Text = "";
                txtFuri.Text = "";
                comboBox1.SelectedIndex = -1;
                txtYaku.Text = "";
                iDate.Checked = false;
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

                txtCode.Focus();
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
                    Control.社員 Shain = new Control.社員();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Shain.Close();
                                return;
                            }

                            if (Shain.DataInsert(cMaster) == true)
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
                                Shain.Close();
                                return;
                            }

                            if (Shain.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Shain.Close();

                    DispClear();

                    //データを 'darwinDataSet.社員' テーブルに読み込みます。
                    this.社員TableAdapter.Fill(this.darwinDataSet.社員);

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
                    if (txtCode.Text == null)
                    {
                        this.txtCode.Focus();
                        throw new Exception("コードは数字で入力してください");
                    }

                    str = this.txtCode.Text;

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtCode.Focus();
                        throw new Exception("コードは数字で入力してください");
                    }

                    // 未入力またはスペースのみは不可
                    if ((this.txtCode.Text).Trim().Length < 1)
                    {
                        this.txtCode.Focus();
                        throw new Exception("コードを入力してください");
                    }

                    //ゼロは不可
                    if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                    {
                        this.txtCode.Focus();
                        throw new Exception("ゼロは登録できません");
                    }

                    //登録済みコードか調べる
                    string sqlStr;
                    Control.社員 Shain = new Control.社員();
                    OleDbDataReader dr;

                    sqlStr = " where ID = " + txtCode.Text.ToString();
                    dr = Shain.FillBy(sqlStr);

                    if (dr.HasRows == true)
                    {
                        txtCode.Focus();
                        dr.Close();
                        Shain.Close();
                        throw new Exception("既に登録済みのコードです");
                    }

                    dr.Close();
                    Shain.Close();

                }

                //名称チェック
                if (txtName.Text.Trim().Length < 1)
                {
                    txtName.Focus();
                    throw new Exception("氏名を入力してください");
                }

                // 所属
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("所属を選択してください");
                }

                //社員クラスにデータセット
                cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.氏名 = txtName.Text.ToString();
                cMaster.フリガナ = txtFuri.Text.ToString();

                Utility.ComboShozoku cmb1 = new Utility.ComboShozoku();
                cmb1 = (Utility.ComboShozoku)comboBox1.SelectedItem;
                cMaster.所属コード = cmb1.ID;
                
                cMaster.役職 = txtYaku.Text.ToString();

                if (iDate.Checked == false)
                {
                    cMaster.入社年月日 = "";
                }
                else
                {
                    cMaster.入社年月日 = iDate.Value.ToShortDateString();
                }

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

            if (sender == txtCode)
            {
                objtxt = txtCode;
            }

            if (sender == txtName)
            {
                objtxt = txtName;
            }

            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtYaku)
            {
                objtxt = txtYaku;
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

            if (sender == txtCode)
            {
                objtxt = txtCode;
            }

            if (sender == txtName)
            {
                objtxt = txtName;
            }

            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtYaku)
            {
                objtxt = txtYaku;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //他に社員登録されているときは削除不可とする
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(受注.社員ID = " + txtCode.Text.ToString() + ")  ";

            OleDbDataReader dr;
            Control.受注 Jyuchu = new Control.受注();
            dr = Jyuchu.FillBy(SqlStr);

            //該当社員の受注データが登録されているときは削除不可とする
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName.Text.ToString() + "の受注データ登録が存在します", txtName.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Jyuchu.Close();
                return;
            }

            dr.Close();
            Jyuchu.Close();

            //得意先に担当者登録されているときは削除不可とする
            SqlStr = " where ";
            SqlStr += "(得意先.担当社員コード = " + txtCode.Text.ToString() + ")  ";

            Control.得意先 tokui = new Control.得意先();

            dr =  tokui.FillBy(SqlStr);

            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName.Text.ToString() + "の担当得意先が存在します", txtName.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                tokui.Close();
                return;
            }

            dr.Close();
            tokui.Close();

            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.社員 Shain = new Control.社員();
            if (Shain.DataDelete(Convert.ToInt32(txtCode.Text.ToString()))==true)
                MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            Shain.Close();

            DispClear();

            //データを 'darwinDataSet.社員' テーブルに読み込みます。
            this.社員TableAdapter.Fill(this.darwinDataSet.社員);

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

        private void label5_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Form frm = new frmShozoku();

            frm.ShowDialog();

            //所属コンボボックスデータロード
            Utility.ComboShozoku.load(this.comboBox1);


        }

        private void txtName_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtName.Text) == false)
            {
                txtFuri.Text = "";
            }
        }

        private void txtName_KeyDown(object sender, KeyEventArgs e)
        {
            txtFuri.Text = MyTextKana.TextKana.textBox_KeyDown(txtName, sender, e);
        }

    }
}
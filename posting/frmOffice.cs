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
    public partial class frmOffice : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.事業所 cOffice = new Entity.事業所();

        public frmOffice()
        {
            InitializeComponent();
        }

        string[] zipArray = null;

        private void Gengo_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: このコード行はデータを 'darwinDataSet.事業所' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            this.事業所TableAdapter.Fill(this.darwinDataSet.事業所);
            
            // TODO: このコード行はデータを 'darwinDataSet.事業所' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);
            this.事業所TableAdapter.Fill(this.darwinDataSet.事業所);
            
            // 郵便番号CSV読み込み
            Utility.zipCsvLoad(ref zipArray);

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
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 300;
                    tempDGV.Columns[4].Width = 200;
                    tempDGV.Columns[5].Width = 110;
                    tempDGV.Columns[6].Width = 110;
                    tempDGV.Columns[7].Width = 200;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.事業所 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.事業所 Office = new Control.事業所();
                OleDbDataReader dr;

                sqlStr = " where 事業所.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Office.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.名称 = dr["名称"].ToString() + "";
                        tempC.郵便番号 = dr["郵便番号"].ToString() + "";
                        tempC.住所1 = dr["住所1"].ToString() + "";
                        tempC.住所2 = dr["住所2"].ToString() + "";
                        tempC.電話番号 = dr["電話番号"].ToString() + "";
                        tempC.FAX番号 = dr["FAX番号"].ToString() + "";
                        tempC.備考 = dr["備考"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    Office.Close();
                    return false;
                }

                dr.Close();
                Office.Close();
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

            if (MessageBox.Show(msgStr, "事業所選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {

                    //データを取得する
                    if (GridviewSet.GetData(dataGridView1,ref cOffice) == false)
                    {
                        MessageBox.Show("該当するデータがマスターに登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'データ値を取得
                    //txtID.Text = cCostname.ID.ToString();
                    //txtName.Text = cCostname.Name;
                    //txtNote.Text = cCostname.Note;

                    txtCode.Text = cOffice.ID.ToString();
                    txtName.Text = cOffice.名称;
                    txtZipCode.Text = cOffice.郵便番号;
                    txtAddress1.Text = cOffice.住所1;
                    txtAddress2.Text = cOffice.住所2;
                    txtTel.Text = cOffice.電話番号;
                    txtFax.Text = cOffice.FAX番号;
                    txtMemo.Text = cOffice.備考;

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
                txtZipCode.Text = "";
                txtAddress1.Text = "";
                txtAddress2.Text = "";
                txtTel.Text = "";
                txtFax.Text = "";
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
                    Control.事業所 dCon = new Control.事業所();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                dCon.Close();
                                return;
                            }

                            if (dCon.DataInsert(cOffice) == true)
                            {
                                MessageBox.Show("新規登録されました", "事業所マスター", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("新規登録に失敗しました", "事業所マスター", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                dCon.Close();
                                return;
                            }

                            if (dCon.DataUpdate(cOffice) == true)
                            {
                                MessageBox.Show("更新されました", "事業所マスター", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", "事業所マスター", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    dCon.Close();

                    DispClear();

                    //データを 'jFGMSTSQLDataSet.言語' テーブルに読み込みます。
                    this.事業所TableAdapter.Fill(this.darwinDataSet.事業所);

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
                    Control.事業所 Office = new Control.事業所();
                    OleDbDataReader dr;

                    sqlStr = " where ID = " + txtCode.Text.ToString();
                    dr = Office.FillBy(sqlStr);

                    if (dr.HasRows == true)
                    {
                        txtCode.Focus();
                        dr.Close();
                        Office.Close();
                        throw new Exception("既に登録済みのコードです");
                    }

                    dr.Close();
                    Office.Close();

                }

                //名称チェック
                if (txtName.Text.Trim().Length < 1)
                {
                    txtName.Focus();
                    throw new Exception("名称を入力してください");
                }

                //事業所クラスにデータセット
                cOffice.ID = Convert.ToInt32(txtCode.Text.ToString());
                cOffice.名称 = txtName.Text.ToString();
                cOffice.郵便番号 = txtZipCode.Text.ToString();
                cOffice.住所1 = txtAddress1.Text.ToString();
                cOffice.住所2 = txtAddress2.Text.ToString();
                cOffice.電話番号 = txtTel.Text.ToString();
                cOffice.FAX番号 = txtFax.Text.ToString();
                cOffice.備考 = txtMemo.Text.ToString();

                if (fMode.Mode ==0) cOffice.登録年月日 = DateTime.Today;
                cOffice.変更年月日 = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "事業所マスター保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

            if (sender == txtZipCode)
            {
                objtxt = txtZipCode;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
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

            if (sender == txtZipCode)
            {
                objtxt = txtZipCode;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //他に事業所登録されているときは削除不可とする
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(受注.事業所ID = " + txtCode.Text.ToString() + ")  ";

            OleDbDataReader dr;
            Control.受注 Jyuchu = new Control.受注();
            dr = Jyuchu.FillBy(SqlStr);

            //該当事業所の受注データが登録されているときは削除不可とする
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName.Text.ToString() + "の受注データ登録が存在します", txtName.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Jyuchu.Close();
                return;
            }

            dr.Close();
            Jyuchu.Close();
            
            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.事業所 Office = new Control.事業所();
            if (Office.DataDelete(Convert.ToInt32(txtCode.Text.ToString()))==true)
                MessageBox.Show("削除されました", "事業所マスター",  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Office.Close();

            DispClear();

            //データを 'darwinDataSet.事業所' テーブルに読み込みます。
            this.事業所TableAdapter.Fill(this.darwinDataSet.事業所);

        }

        private void btnCsv_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, "事業所マスター");
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

        private void txtZipCode_TextChanged(object sender, EventArgs e)
        {
            if (zipArray == null)
            {
                return;
            }

            TextBox mTxt = null;
            TextBox txtAdd = null;
            bool mc = false;

            mTxt = txtZipCode;
            txtAdd = txtAddress1;

            string zipText = mTxt.Text.Replace("-", "");

            if (zipText.Length == 7)
            {
                foreach (var t in zipArray)
                {
                    string[] r = t.Split(',');

                    if (zipText == r[2].Replace("\"", ""))
                    {
                        // 住所
                        string ad = r[6].Replace("\"", "") + r[7].Replace("\"", "") + r[8].Replace("\"", "");

                        // 同じ住所ならば再表示しない
                        if (!txtAdd.Text.Contains(ad))
                        {
                            txtAdd.Text = ad;
                        }

                        txtAdd.Focus();
                        mc = true;
                        break;
                    }
                }
            }
        }

    }
}
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
    public partial class frmSize : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.判型 cMaster = new Entity.判型();

        const string MESSAGE_CAPTION = "判型マスター";

        public frmSize()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);
            
            // TODO: このコード行はデータを 'darwinDataSet.判型' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            GridviewSet.Setting(dataGridView1);
            this.判型TableAdapter.Fill(this.darwinDataSet.判型);

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
                    tempDGV.Columns[1].Width = 120;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 120;
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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.判型 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.判型 Hangata = new Control.判型();
                OleDbDataReader dr;

                sqlStr = " where 判型.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Hangata.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.名称 = dr["名称"].ToString() + "";
                        tempC.卸単価1 = Convert.ToDouble(dr["卸単価1"].ToString());
                        tempC.卸単価2 = Convert.ToDouble(dr["卸単価2"].ToString());
                        tempC.卸単価3 = Convert.ToDouble(dr["卸単価3"].ToString());
                        tempC.備考 = dr["備考"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    Hangata.Close();
                    return false;
                }

                dr.Close();
                Hangata.Close();
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

            if (MessageBox.Show(msgStr, "判型選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    txtName1.Text = cMaster.名称;
                    txtTanka1.Text = cMaster.卸単価1.ToString("##0.00");
                    txtTanka2.Text = cMaster.卸単価2.ToString("##0.00");
                    txtTanka3.Text = cMaster.卸単価3.ToString("##0.00");
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
                txtTanka1.Text = "0";
                txtTanka2.Text = "0";
                txtTanka3.Text = "0";
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
                    Control.判型 Hangata = new Control.判型();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Hangata.Close();
                                return;
                            }

                            if (Hangata.DataInsert(cMaster) == true)
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
                                Hangata.Close();
                                return;
                            }

                            if (Hangata.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Hangata.Close();

                    DispClear();

                    //データを 'darwinDataSet.判型' テーブルに読み込みます。
                    this.判型TableAdapter.Fill(this.darwinDataSet.判型);
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
                    //Control.判型 Hangata = new Control.判型();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Hangata.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Hangata.Close();
                    //    throw new Exception("既に登録済みのコードです");
                    //}

                    //dr.Close();
                    //Hangata.Close();

                }

                //名称チェック
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("名称を入力してください");
                }

                // 卸単価１：数字か？
                if (txtTanka1.Text == null)
                {
                    this.txtTanka1.Focus();
                    throw new Exception("卸単価は数字で入力してください");
                }

                str = this.txtTanka1.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtTanka1.Focus();
                    throw new Exception("卸単価は数字で入力してください");
                }

                // 未入力またはスペースのみは不可
                if ((this.txtTanka1.Text).Trim().Length < 1)
                {
                    this.txtTanka1.Focus();
                    throw new Exception("卸単価を入力してください");
                }

                // 卸単価２：数字か？
                if (txtTanka2.Text == null)
                {
                    this.txtTanka2.Focus();
                    throw new Exception("卸単価は数字で入力してください");
                }

                str = this.txtTanka2.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtTanka2.Focus();
                    throw new Exception("卸単価は数字で入力してください");
                }

                // 未入力またはスペースのみは不可
                if ((this.txtTanka2.Text).Trim().Length < 1)
                {
                    this.txtTanka2.Focus();
                    throw new Exception("卸単価を入力してください");
                }

                // 卸単価３：数字か？
                if (txtTanka3.Text == null)
                {
                    this.txtTanka3.Focus();
                    throw new Exception("卸単価は数字で入力してください");
                }

                str = this.txtTanka3.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtTanka3.Focus();
                    throw new Exception("卸単価は数字で入力してください");
                }

                // 未入力またはスペースのみは不可
                if ((this.txtTanka3.Text).Trim().Length < 1)
                {
                    this.txtTanka3.Focus();
                    throw new Exception("卸単価を入力してください");
                }

                ////ゼロは不可
                //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                //{
                //    this.txtCode.Focus();
                //    throw new Exception("ゼロは登録できません");
                //}

                //クラスにデータセット
                //cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.名称 = txtName1.Text.ToString();
                cMaster.卸単価1 = Convert.ToDouble(txtTanka1.Text.ToString());
                cMaster.卸単価2 = Convert.ToDouble(txtTanka2.Text.ToString());
                cMaster.卸単価3 = Convert.ToDouble(txtTanka3.Text.ToString());
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

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtTanka1)
            {
                objtxt = txtTanka1;
            }

            if (sender == txtTanka2)
            {
                objtxt = txtTanka2;
            }

            if (sender == txtTanka3)
            {
                objtxt = txtTanka3;
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

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtTanka1)
            {
                objtxt = txtTanka1;
            }

            if (sender == txtTanka2)
            {
                objtxt = txtTanka2;
            }

            if (sender == txtTanka3)
            {
                objtxt = txtTanka3;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //他に登録されているときは削除不可とする
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(受注.判型 = " + cMaster.ID.ToString() + ")  ";

            OleDbDataReader dr;
            Control.受注 Order = new Control.受注();
            dr = Order.FillBy(SqlStr);

            //該当判型の受注データが登録されているときは削除不可とする
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "の受注データが登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Order.Close();
                return;
            }

            dr.Close();
            Order.Close();
            
            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.判型 Hangata = new Control.判型();
            if (Hangata.DataDelete(cMaster.ID)==true)
                MessageBox.Show("削除されました", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Hangata.Close();

            DispClear();

            //データを 'darwinDataSet.判型' テーブルに読み込みます。
            this.判型TableAdapter.Fill(this.darwinDataSet.判型);

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

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            GridEnter();
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }
    }
}
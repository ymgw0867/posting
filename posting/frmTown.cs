using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmTown : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.町名 cMaster = new Entity.町名();

        const string MESSAGE_CAPTION = "町名マスター";

        public frmTown()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //グリッド定義
            GridviewSet.Setting(dataGridView1);

            // TODO: このコード行はデータを 'darwinDataSet.町名' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            this.町名TableAdapter.Fill(this.darwinDataSet.町名);

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
                    tempDGV.Height = 181;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // 各列幅指定
                    //tempDGV.Columns.Add("col1", "ｺｰﾄﾞ");
                    //tempDGV.Columns.Add("col2", "名称");
                    //tempDGV.Columns.Add("col3", "備考");

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 100;
                    //tempDGV.Columns[3].Width = 160;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;

                    tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.町名 tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.町名 Town = new Control.町名();
                OleDbDataReader dr;

                sqlStr = " where 町名.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Town.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.名称 = dr["名称"].ToString() + "";
                        tempC.市区町村コード = int.Parse(dr["市区町村コード"].ToString(),System.Globalization.NumberStyles.Any);
                        tempC.備考 = dr["備考"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    Town.Close();
                    return false;
                }

                dr.Close();
                Town.Close();
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

            if (MessageBox.Show(msgStr, "町名選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    txtCode.Text = cMaster.ID.ToString();
                    txtName1.Text = cMaster.名称.ToString();
                    txtCityCode.Text = cMaster.市区町村コード.ToString();
                    txtMemo.Text = cMaster.備考.ToString();

                    //IDテキストボックスは編集不可とする
                    txtCode.Enabled = false;

                    //ボタン状態
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //フォームモードステータス:変更削除

                    txtName1.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "画面表示", MessageBoxButtons.OK);
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
                txtName1.Text = "";
                txtCityCode.Text = "0";
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
                    Control.町名 Town = new Control.町名();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Town.Close();
                                return;
                            }

                            if (Town.DataInsert(cMaster) == true)
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
                                Town.Close();
                                return;
                            }

                            if (Town.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", "所属マスター", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Town.Close();

                    DispClear();

                    //データを 'darwinDataSet.町名' テーブルに読み込みます。
                    this.町名TableAdapter.Fill(this.darwinDataSet.町名);
                    dataGridView1.DataSource = this.darwinDataSet.町名;

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
                    Control.町名 Town = new Control.町名();
                    OleDbDataReader dr;

                    sqlStr = " where ID = " + txtCode.Text.ToString();
                    dr = Town.FillBy(sqlStr);

                    if (dr.HasRows == true)
                    {
                        txtCode.Focus();
                        dr.Close();
                        Town.Close();
                        throw new Exception("既に登録済みのコードです");
                    }

                    dr.Close();
                    Town.Close();

                }

                //名称チェック
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("名称を入力してください");
                }

                //市区町村コード
                if (Utility.NumericCheck(txtCityCode.Text) == false)
                {
                    txtCityCode.Focus();
                    throw new Exception("市区町村コードは数字で入力してください");
                }

                //マスターチェック
                if (txtCityCode.Text != "0")
                {
                    string sqlSTR;
                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();

                    sqlSTR = "";
                    sqlSTR += "select * from 市区町村 where ID = " + txtCityCode.Text;
                    dR = fCon.free_dsReader(sqlSTR);

                    if (dR.HasRows == false)
                    {
                        txtCityCode.Focus();
                        dR.Close();
                        fCon.Close();
                        throw new Exception("該当する市区町村コードがありません");
                    }

                    dR.Close();
                    fCon.Close();
                }


                //クラスにデータセット
                cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.名称 = txtName1.Text.ToString();
                cMaster.市区町村コード = int.Parse(txtCityCode.Text.ToString(), System.Globalization.NumberStyles.Any);
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

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtCityCode)
            {
                objtxt = txtCityCode;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
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

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtCityCode)
            {
                objtxt = txtCityCode;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //他に登録されているときは削除不可とする
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(配布エリア.町名ID = " + txtCode.Text.ToString() + ")  ";

            OleDbDataReader dr;
            Control.配布エリア Area = new Control.配布エリア();
            dr = Area.FillBy(SqlStr);

            //該当町名が登録されているときは削除不可とする
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "が配布エリアデータに登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Area.Close();
                return;
            }

            dr.Close();
            Area.Close();
            
            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            Control.町名 Town = new Control.町名();
            if (Town.DataDelete(Convert.ToInt32(txtCode.Text.ToString()))==true)
                MessageBox.Show("削除されました", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Town.Close();

            DispClear();

            //データを 'darwinDataSet.町名' テーブルに読み込みます。
            this.町名TableAdapter.Fill(this.darwinDataSet.町名);
            dataGridView1.DataSource = this.darwinDataSet.町名;

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

        private void button1_Click(object sender, EventArgs e)
        {
            darwinDataSet ds = new darwinDataSet();
            this.町名TableAdapter.FillByName(ds.町名,"%" + textBox1.Text.ToString() + "%");
            dataGridView1.DataSource = ds.町名;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GetExcelPos();
        }

        private void GetExcelPos()
        {

            DialogResult ret;

            //ダイアログボックスの初期設定
            openFileDialog1.Title = "市区町村コード表の選択";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Microsoft Office Excelファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel) return;

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "市区町村コード表取り込み", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int S_GYO = 1;    //エクセルファイル見出し行（明細は1行目から）

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            //string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            Excel.Range dRng;
            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            int iX = S_GYO;
            int err;
            string cellID, cellItem1, cellItem2;
            string sqlSTR;

            try
            {

                while (true)
                {
                    err = 0;

                    //市区町村コード
                    dRng = (Excel.Range)oxlsSheet.Cells[iX, 2];

                    //空白なら処理終了
                    if ((dRng.Text.ToString().Trim() + "") == "")
                        break;

                    cellID = dRng.Text.ToString();

                    //都道府県
                    dRng = (Excel.Range)oxlsSheet.Cells[iX, 3];
                    cellItem1 = dRng.Text.ToString();

                    //市区町村名
                    dRng = (Excel.Range)oxlsSheet.Cells[iX, 4];
                    cellItem2 = dRng.Text.ToString();

                    //コードチェック
                    if (Utility.NumericCheck(cellID) == false)
                    {
                        err = 1;
                    }

                    //エラーのときは読み飛ばし
                    if (err == 0)
                    {
                        Control.FreeSql fCon = new Control.FreeSql();

                        sqlSTR = "";
                        sqlSTR += "insert into 市区町村 ";
                        sqlSTR += "(ID,都道府県,市区町村,区分1,区分2,備考,登録年月日,変更年月日) ";
                        sqlSTR += "values (" + cellID + ",";
                        sqlSTR += "'" + cellItem1 + "',";
                        sqlSTR += "'" + cellItem2 + "',";
                        sqlSTR += "0,0,'',";
                        sqlSTR += "'" + DateTime.Today.ToShortDateString() + "',";
                        sqlSTR += "'" + DateTime.Today.ToShortDateString() + "')";

                        if (fCon.Execute(sqlSTR) == false)
                        {
                            MessageBox.Show("市区町村マスターの登録に失敗しました。" + cellID, "市区町村コード表読み込み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            fCon.Close();
                            break;
                        }

                        fCon.Close();

                    }

                    iX++;
                }

                MessageBox.Show("終了しました", "市区町村コード表読み込み", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // 確認のためExcelのウィンドウを表示する
                //oXls.Visible = true;

                //印刷
                //oxlsSheet.PrintPreview(true);

                //保存処理
                oXls.DisplayAlerts = false;

                //Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                //Excelを終了
                oXls.Quit();
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            finally
            {
                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string cName;
            string sqlSTR;
            OleDbDataReader dR;
            Control.FreeSql fCon = new Control.FreeSql();

            dR = fCon.free_dsReader("select * from 市区町村 order by ID");

            while(dR.Read())
            {
                cName = dR["市区町村"].ToString();
                Control.FreeSql fCon2 = new Control.FreeSql();

                sqlSTR = "";
                sqlSTR += "update 町名 ";
                sqlSTR += "set 町名.市区町村コード = " + dR["ID"].ToString();
                sqlSTR += "where (町名.名称 like '" + cName + "%') and ";
                sqlSTR += "(町名.市区町村コード = 0)";

                fCon2.Execute(sqlSTR);

                fCon2.Close();

            }

            dR.Close();

            fCon.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmTownSub fSub = new frmTownSub();

            if (fSub.ShowDialog(this) == DialogResult.OK)
            {
                txtCityCode.Text = fSub.市区町村コード.ToString(); ;
            }
            else
            {
            }

            fSub.Dispose();
        }

    }
}
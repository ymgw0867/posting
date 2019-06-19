using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;
using System.Linq;

namespace posting
{
    public partial class frmClient : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        //Entity.得意先 cMaster = new Entity.得意先();

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.得意先TableAdapter cAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.請求書TableAdapter rAdp = new darwinDataSetTableAdapters.請求書TableAdapter();
        darwinDataSetTableAdapters.社員TableAdapter sAdp = new darwinDataSetTableAdapters.社員TableAdapter();
        //darwinDataSetTableAdapters.会社情報TableAdapter kAdp = new darwinDataSetTableAdapters.会社情報TableAdapter();

        const string MESSAGE_CAPTION = "得意先マスター";

        string[] zipArray = null;

        bool jStatus = false;
        bool sStatus = false;

        public frmClient()
        {
            InitializeComponent();

            // データ読み込み 2015/07/05
            cAdp.Fill(dts.得意先);
            //jAdp.Fill(dts.受注1);
            //rAdp.Fill(dts.請求書);
            sAdp.Fill(dts.社員);
            //kAdp.Fill(dts.会社情報);
        }

        private void form_Load(object sender, EventArgs e)
        {
            // 請求先名称セット
            setSeikyuName();

            // ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: このコード行はデータを 'darwinDataSet.得意先' テーブルに読み込みます。必要に応じて移動、または削除をしてください
            Setting(dataGridView1);
            gridSearch(dataGridView1);
            //this.得意先TableAdapter.Fill(this.darwinDataSet.得意先);

            // 敬称コンボ
            cmbKeishoSet();
            cmbKeishoSSet();    // 請求先敬称 2019/02/20

            // 税通知コンボ
            cmbTaxSet();

            // 都道府県コンボ
            cmbCitySet(cmbCity);
            cmbCitySet(cmbCityS);

            // 担当社員コンボ
            Utility.ComboShain.load(cmbShain);
            
            // データ初期化
            DispClear();

            // 郵便番号CSV読み込み
            Utility.zipCsvLoad(ref zipArray);

            // 請求先・敬称、部署名セットツールボタン表示 2019/04/02
            DateTime dt = new DateTime(2019, 04, 30);
            if (DateTime.Today > dt)
            {
                button2.Visible = false;
            }
            else
            {
                button2.Visible = true;
            }
        }

        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     請求先名称セット null、または空白のとき名称をセットする</summary>
        /// -----------------------------------------------------------------------------------
        private void setSeikyuName()
        {
            int cnt = 0;

            foreach (var t in dts.得意先)
            {
                //if (t.Is請求先名称Null())
                //{
                //    t.請求先名称 = t.名称;
                //    cnt++;
                //}
                //else if (t.請求先名称.Trim() == string.Empty)
                //{
                //    t.請求先名称 = t.名称;
                //    cnt++;
                //}

                if (t.請求先名称.Trim() == string.Empty)
                {
                    t.請求先名称 = t.名称;
                    cnt++;
                }
            }

            if (cnt > 0)
            {
                cAdp.Update(dts.得意先);
                cAdp.Fill(dts.得意先);
            }
        }


        #region グリッドビューカラム定義
        string colID = "col1";
        string colRyaku = "col2";
        string colFuri = "col3";
        string colName = "col4";
        string colKeisho = "col5";
        string colTantou = "col6";
        string colBusho = "col7";
        string colZip = "col8";
        string colCity = "col9";
        string colAdd1 = "col10";
        string colAdd2 = "col11";
        string colTel = "col12";
        string colFax = "col13";
        string colEmail = "col14";
        string colEigyo = "col15";
        string colShimebi = "col16";
        string colZei = "col17";
        string colSeName = "col18";
        string colSeZip = "col19";
        string colSeCity = "col20";
        string colSeAdd1 = "col21";
        string colSeAdd2 = "col22";
        string colBiko = "col23";
        string colAddDt = "col24";
        string colUpDt = "col25";
        string colTantouS = "col26";
        #endregion

        /// -------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// -------------------------------------------------------------------
        private void Setting(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 183;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(colID, "ｺｰﾄﾞ");
                tempDGV.Columns.Add(colRyaku, "略称");
                tempDGV.Columns.Add(colFuri, "フリガナ");
                tempDGV.Columns.Add(colName, "名称");
                tempDGV.Columns.Add(colKeisho, "敬称");
                tempDGV.Columns.Add(colTantou, "担当者名");
                tempDGV.Columns.Add(colBusho, "部署名");
                tempDGV.Columns.Add(colZip, "郵便番号");
                tempDGV.Columns.Add(colCity, "都道府県");
                tempDGV.Columns.Add(colAdd1, "住所１");
                tempDGV.Columns.Add(colAdd2, "住所２");
                tempDGV.Columns.Add(colTel, "電話番号");
                tempDGV.Columns.Add(colFax, "FAX番号");
                tempDGV.Columns.Add(colEmail, "eMail");
                tempDGV.Columns.Add(colEigyo, "担当営業");
                tempDGV.Columns.Add(colShimebi, "締日");
                tempDGV.Columns.Add(colZei, "税通知");
                tempDGV.Columns.Add(colSeName, "請求先名称");
                tempDGV.Columns.Add(colSeZip, "請求先〒");
                tempDGV.Columns.Add(colSeCity, "請求先都道府県");
                tempDGV.Columns.Add(colSeAdd1, "請求先１");
                tempDGV.Columns.Add(colSeAdd2, "請求先２");
                tempDGV.Columns.Add(colTantouS, "請求先担当者");
                tempDGV.Columns.Add(colBiko, "備考");
                tempDGV.Columns.Add(colAddDt, "登録年月日");
                tempDGV.Columns.Add(colUpDt, "更新年月日");

                tempDGV.Columns[0].Width = 70;
                tempDGV.Columns[1].Width = 200;
                tempDGV.Columns[2].Width = 200;
                tempDGV.Columns[3].Width = 200;
                tempDGV.Columns[4].Width = 60;
                tempDGV.Columns[5].Width = 100;
                tempDGV.Columns[6].Width = 160;
                tempDGV.Columns[7].Width = 100;
                tempDGV.Columns[8].Width = 120;
                tempDGV.Columns[9].Width = 200;
                tempDGV.Columns[10].Width = 200;

                tempDGV.Columns[11].Width = 120;
                tempDGV.Columns[12].Width = 120;
                tempDGV.Columns[13].Width = 160;
                tempDGV.Columns[14].Width = 100;
                tempDGV.Columns[15].Width = 60;
                tempDGV.Columns[16].Width = 80;
                tempDGV.Columns[17].Width = 120;
                tempDGV.Columns[18].Width = 120;
                tempDGV.Columns[19].Width = 200;
                tempDGV.Columns[20].Width = 200;
                tempDGV.Columns[21].Width = 200;
                tempDGV.Columns[22].Width = 110;
                tempDGV.Columns[23].Width = 120;
                tempDGV.Columns[24].Width = 140;
                tempDGV.Columns[25].Width = 140;

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

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの指定行のデータを取得する </summary>
        /// <param name="dgv">
        ///     対象とするデータグリッドビューオブジェクト</param>
        /// -----------------------------------------------------------------------------
        private Boolean GetData(DataGridView dgv,ref Entity.得意先 tempC, darwinDataSet dts, int sID)
        {
            //foreach (var t in dts.得意先.Where(a => a.ID == sID))
            //{

            //}

            int iX = 0;
            string sqlStr;
            Control.得意先 Client = new Control.得意先();
            OleDbDataReader dr;

            sqlStr = " where 得意先.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
            dr = Client.FillBy(sqlStr);

            if (dr.HasRows == true)
            {
                while (dr.Read() == true)
                {
                    tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                    tempC.略称 = dr["略称"].ToString() + "";
                    tempC.フリガナ = dr["フリガナ"].ToString();
                    tempC.名称 = dr["名称"].ToString();
                    tempC.敬称 = dr["敬称"].ToString();
                    tempC.担当者名 = dr["担当者名"].ToString();
                    tempC.部署名 = dr["部署名"].ToString();
                    tempC.担当者名 = dr["担当者名"].ToString();
                    tempC.郵便番号 = dr["郵便番号"].ToString();
                    tempC.都道府県 = dr["都道府県"].ToString();
                    tempC.住所1 = dr["住所1"].ToString();
                    tempC.住所2 = dr["住所2"].ToString();
                    tempC.電話番号 = dr["電話番号"].ToString();
                    tempC.FAX番号 = dr["FAX番号"].ToString();
                    tempC.メールアドレス = dr["メールアドレス"].ToString();
                    tempC.担当社員コード = Int32.Parse(dr["担当社員コード"].ToString());
                    tempC.締日 = Int32.Parse(dr["締日"].ToString());
                    tempC.税通知 = dr["税通知"].ToString();
                    tempC.請求先郵便番号 = dr["請求先郵便番号"].ToString();
                    tempC.請求先都道府県 = dr["請求先都道府県"].ToString();
                    tempC.請求先住所1 = dr["請求先住所1"].ToString();
                    tempC.請求先住所2 = dr["請求先住所2"].ToString();
                    tempC.備考 = dr["備考"].ToString();
                    tempC.請求先担当者名 = dr["請求先担当者名"].ToString();   // 2015/11/20
                }
            }
            else
            {
                dr.Close();
                Client.Close();
                return false;
            }

            dr.Close();
            Client.Close();
            return true;
        }

        //グリッドからデータを選択
        private void GridEnter()
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[1, dataGridView1.SelectedRows[iX].Index].Value + "が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "得意先選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    // データを取得する
                    int sID = int.Parse(dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value.ToString());
                    if (dts.得意先.Any(a => a.ID == sID))
                    {
                        darwinDataSet.得意先Row r = dts.得意先.Single(a => a.ID == sID);

                        //'データを表示します 2015/07/05
                        //txtCode.Text = cMaster.ID.ToString();
                        txtName1.Text = r.略称;
                        txtFuri.Text = r.フリガナ;
                        txtName2.Text = r.名称;
                        cmbKeisho.Text = r.敬称;
                        txtTantou.Text = r.担当者名;
                        txtBusho.Text = r.部署名;
                        mtxtZipCode.Text = r.郵便番号;
                        cmbCity.Text = r.都道府県;
                        txtAddress1.Text = r.住所1;
                        txtAddress2.Text = r.住所2;
                        txtTel.Text = r.電話番号;
                        txtFax.Text = r.FAX番号;
                        txtEmail.Text = r.メールアドレス;

                        Utility.ComboShain.selectedIndex(cmbShain, Int32.Parse(r.担当社員コード.ToString()));

                        txtShimebi.Text = r.締日.ToString();
                        cmbTax.Text = r.税通知.ToString();

                        //if (r.Is請求先名称Null())
                        //{
                        //    txtNameSeikyu.Text = string.Empty;
                        //}
                        //else
                        //{
                        //    txtNameSeikyu.Text = r.請求先名称;
                        //}

                        txtNameSeikyu.Text = r.請求先名称;

                        mtxtZipCodeS.Text = r.請求先郵便番号;
                        cmbCityS.Text = r.請求先都道府県;
                        txtAddress1S.Text = r.請求先住所1;
                        txtAddress2S.Text = r.請求先住所2;

                        // 2015/11/20
                        if (r.請求先担当者名 != null)
                        {
                            txtTantouS.Text = r.請求先担当者名;
                        }
                        else
                        {
                            txtTantouS.Text = string.Empty;
                        }

                        txtMemo.Text = r.備考;

                        txtBushoS.Text = r.請求先部署名;  // 2019/02/20
                        cmbKeishoS.Text = r.請求先敬称;  // 2019/02/20
                        

                        //IDテキストボックスは編集不可とする
                        //txtCode.Enabled = false;

                        //ボタン状態
                        btnDel.Enabled = true;
                        btnClr.Enabled = true;

                        fMode.Mode = 1;     // フォームモードステータス:変更削除
                        fMode.ID = r.ID;    // 得意先ID

                        txtName1.Focus();

                        button2.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("該当するデータがマスターに登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                }
                catch (Exception e)
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
                txtFuri.Text = "";
                txtName2.Text = "";
                cmbKeisho.Text = "";
                txtTantou.Text = "";
                txtBusho.Text = "";
                mtxtZipCode.Text = "";
                cmbCity.Text = "";
                txtAddress1.Text = "";
                txtAddress2.Text = "";
                txtTel.Text = "";
                txtFax.Text = "";
                txtEmail.Text = "";
                cmbShain.SelectedIndex = -1;
                txtShimebi.Text = "0";
                cmbTax.Text = "";

                txtNameSeikyu.Text = "";    // 2015/07/05
                mtxtZipCodeS.Text = "";
                cmbCityS.Text = "";
                txtAddress1S.Text = "";
                txtAddress2S.Text = "";
                txtTantouS.Text = "";   // 2015/11/20
                txtMemo.Text = "";

                txtBushoS.Text = "";    // 2019/02/20
                cmbKeishoS.Text = "";   // 2019/02/20
                cmbKeishoS.SelectedIndex = -1;  // 2019/02/20

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

                button2.Enabled = true; // 2019/04/02
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

                if (fDataCheck())
                {
                    switch (fMode.Mode)
                    {
                        case 0: // 新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                return;
                            }

                            dts.得意先.Add得意先Row(setClientRow(dts.得意先.New得意先Row()));
                            break;

                        case 1: // 更新
                            if (MessageBox.Show("更新します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                return;
                            }

                            darwinDataSet.得意先Row r = dts.得意先.Single(a => a.ID == fMode.ID);

                            if (!r.HasErrors)
                            {
                                setClientRow(r);
                            }
                            else
                            {
                                MessageBox.Show(fMode.ID + "がキー不在です：データの更新に失敗しました", "更新エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }

                            break;
                    }

                    // データベース更新
                    cAdp.Update(dts.得意先);

                    // 画面初期化
                    DispClear();

                    //データを 'darwinDataSet.得意先' テーブルに読み込みます。
                    //dataGridView1.DataSource = null;
                    cAdp.Fill(dts.得意先);
                    gridSearch(dataGridView1);

                    //this.得意先TableAdapter.Fill(this.darwinDataSet.得意先);
                    //dataGridView1.DataSource = this.darwinDataSet.得意先;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"更新処理",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     darwinDataSet.得意先Row行データ作成 </summary>
        /// <param name="r">
        ///     darwinDataSet.得意先Row オブジェクト</param>
        /// <returns>
        ///     darwinDataSet.得意先Row オブジェクト</returns>
        /// ------------------------------------------------------------------
        private darwinDataSet.得意先Row setClientRow(darwinDataSet.得意先Row r)
        {
            r.略称 = txtName1.Text;
            r.フリガナ = txtFuri.Text;
            r.名称 = txtName2.Text;
            r.敬称 = cmbKeisho.Text;
            r.担当者名 = txtTantou.Text;
            r.部署名 = txtBusho.Text;

            if (((mtxtZipCode.Text).Replace("-", "")).Trim() == "")
            {
                r.郵便番号 = "";
            }
            else
            {
                r.郵便番号 = mtxtZipCode.Text;
            }

            r.都道府県 = cmbCity.Text;
            r.住所1 = txtAddress1.Text;
            r.住所2 = txtAddress2.Text;
            r.電話番号 = txtTel.Text;
            r.FAX番号 = txtFax.Text;
            r.メールアドレス = txtEmail.Text;

            Utility.ComboShain cmb1 = new Utility.ComboShain();

            if (cmbShain.SelectedIndex == -1)
            {
                r.担当社員コード = 0;
            }
            else
            {
                cmb1 = (Utility.ComboShain)cmbShain.SelectedItem;
                r.担当社員コード = cmb1.ID;
            }

            r.締日 = Int32.Parse(txtShimebi.Text);
            r.税通知 = cmbTax.Text.ToString();

            if (((mtxtZipCodeS.Text).Replace("-", "")).Trim() == "")
            {
                r.請求先郵便番号 = "";
            }
            else
            {
                r.請求先郵便番号 = mtxtZipCodeS.Text;
            }

            r.請求先都道府県 = cmbCityS.Text;
            r.請求先住所1 = txtAddress1S.Text;
            r.請求先住所2 = txtAddress2S.Text;
            r.請求先担当者名 = txtTantouS.Text;

            r.備考 = txtMemo.Text;

            if (fMode.Mode == 0) r.登録年月日 = DateTime.Now;

            r.変更年月日 = DateTime.Now;
            r.請求先名称 = txtNameSeikyu.Text;   // 2015/07/05

            r.請求先部署名 = txtBushoS.Text;  // 2019/02/20
            r.請求先敬称 = cmbKeishoS.Text;  // 2019/02/20

            return r;
        }

        // 登録データチェック
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
                    //Control.得意先 Client = new Control.得意先();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Client.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Client.Close();
                    //    throw new Exception("既に登録済みのコードです");
                    //}

                    //dr.Close();
                    //Client.Close();

                }

                //略称チェック
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("略称を入力してください");
                }


                //締日：数字か？
                if (txtShimebi.Text == null)
                {
                    this.txtShimebi.Focus();
                    throw new Exception("締日は数字で入力してください");
                }

                str = this.txtShimebi.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtShimebi.Focus();
                    throw new Exception("締日は数字で入力してください");
                }

                ////ゼロは不可
                //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                //{
                //    this.txtCode.Focus();
                //    throw new Exception("ゼロは登録できません");
                //}

                //クラスにデータセット
                //cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());

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
            MaskedTextBox objMtxt = new MaskedTextBox();

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            if (sender == txtName1)
            {
                objtxt = txtName1;

                // 2019/03/16
                MyTextKana.TextKana.textEnter(txtName1);
            }

            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;
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

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtShimebi)
            {
                objtxt = txtShimebi;
            }

            if (sender == mtxtZipCodeS)
            {
                objMtxt = mtxtZipCodeS;
            }

            if (sender == txtAddress1S)
            {
                objtxt = txtAddress1S;
            }

            if (sender == txtAddress2S)
            {
                objtxt = txtAddress2S;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            // 請求先担当者名 2015/11/20
            if (sender == txtTantouS)
            {
                objtxt = txtTantouS;
            }

            // 請求先名称 2015/07/05
            if (sender == txtNameSeikyu)
            {
                objtxt = txtNameSeikyu;
            }

            // 検索電話番号 2015/07/05
            if (sender == txtsTel)
            {
                objtxt = txtsTel;
            }

            // 検索〒番号 2015/07/05
            if (sender == txtsZip)
            {
                objtxt = txtsZip;
            }

            // 検索請求先名称 2015/07/05
            if (sender == textBox2)
            {
                objtxt = textBox2;
            }

            // 検索請求先部署名 2019/02/20
            if (sender == txtBushoS)
            {
                objtxt = txtBushoS;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

            objMtxt.SelectAll();
            objMtxt.BackColor = Color.LightGray;
        }

        private void txtLeave(object sender, EventArgs e)
        {
           TextBox objtxt = new TextBox();
           MaskedTextBox objMtxt = new MaskedTextBox();

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            // 請求先担当者名 2015/11/20
            if (sender == txtTantouS)
            {
                objtxt = txtTantouS;
            }

            if (sender == txtName1)
            {
                objtxt = txtName1;

                // 2019/03/16
                MyTextKana.TextKana.textLeave(txtName1);
            }

            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;

                //請求先名称へコピー 2015/07/05
                if (txtNameSeikyu.Text == "")
                {
                    txtNameSeikyu.Text = txtName2.Text;
                }
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;

                // 請求先へコピー : 2019/02/20
                if (txtBushoS.Text == "")
                {
                    txtBushoS.Text = txtBusho.Text;
                }
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;

                //請求先へコピー
                if (mtxtZipCodeS.Text.Trim() == "-")
                {
                    mtxtZipCodeS.Text = mtxtZipCode.Text;
                }
            }

            if (sender == cmbCity)
            {
                //請求先へコピー
                if (cmbCityS.Text == "")
                {
                    cmbCityS.Text = cmbCity.Text;
                }
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;

                //請求先へコピー
                if (txtAddress1S.Text == "")
                {
                    txtAddress1S.Text = txtAddress1.Text;
                }
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;

                //請求先へコピー
                if (txtAddress2S.Text == "")
                {
                    txtAddress2S.Text = txtAddress2.Text;
                }
            }

            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtShimebi)
            {
                objtxt = txtShimebi;
            }

            if (sender == mtxtZipCodeS)
            {
                objMtxt = mtxtZipCodeS;
            }

            if (sender == txtAddress1S)
            {
                objtxt = txtAddress1S;
            }

            if (sender == txtAddress2S)
            {
                objtxt = txtAddress2S;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }
            
            // 請求先名称 2015/07/05
            if (sender == txtNameSeikyu)
            {
                objtxt = txtNameSeikyu;
            }

            // 検索電話番号 2015/07/05
            if (sender == txtsTel)
            {
                objtxt = txtsTel;
            }

            // 検索〒番号 2015/07/05
            if (sender == txtsZip)
            {
                objtxt = txtsZip;
            }

            // 検索請求先名称 2015/07/05
            if (sender == textBox2)
            {
                objtxt = textBox2;
            }

            // 検索請求先部署名 2019/02/20
            if (sender == txtBushoS)
            {
                objtxt = txtBushoS;
            }

            objtxt.BackColor = Color.White;
            objMtxt.BackColor = Color.White;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            // 削除確認：2018/01/03
            if (MessageBox.Show("削除処理が選択されました。続行してよろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // データ読み込み：2018/01/03
            if (!jStatus)
            {
                Cursor = Cursors.WaitCursor;
                jAdp.Fill(dts.受注1);
                jStatus = true;
                Cursor = Cursors.Default;
            }

            // 受注データに登録されているときは削除不可とする
            if (dts.受注1.Any(a => a.得意先ID == fMode.ID))
            {
                MessageBox.Show(txtName1.Text.ToString() + "の受注データが登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // データ読み込み：2018/01/03
            if (!sStatus)
            {
                Cursor = Cursors.WaitCursor;
                rAdp.Fill(dts.請求書);
                sStatus = true;
                Cursor = Cursors.Default;
            }

            // 請求データに登録されているときは削除不可とする
            if (dts.請求書.Any(a => a.得意先ID == fMode.ID))
            {
                MessageBox.Show(txtName1.Text.ToString() + "の請求データが登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //データ削除
            try
            {
                darwinDataSet.得意先Row r = dts.得意先.Single(a => a.ID == fMode.ID);
                if (!r.HasErrors)
                {
                    r.Delete();

                    // データベース更新
                    cAdp.Update(dts.得意先);

                    MessageBox.Show("削除されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // データを 'darwinDataSet.得意先' テーブルに読み込みます。
                    cAdp.Fill(dts.得意先);
                    gridSearch(dataGridView1);
                }
                else
                {
                    MessageBox.Show("削除に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show("削除に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 画面初期化
            DispClear();
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

        private void cmbKeishoSet()
        {
            cmbKeisho.Items.Clear();
            cmbKeisho.Items.Add("様");
            cmbKeisho.Items.Add("殿");
            cmbKeisho.Items.Add("御中");
        }
        private void cmbKeishoSSet()
        {
            cmbKeishoS.Items.Clear();
            cmbKeishoS.Items.Add("様");
            cmbKeishoS.Items.Add("殿");
            cmbKeishoS.Items.Add("御中");
        }

        private void cmbTaxSet()
        {
            cmbTax.Items.Clear();
            cmbTax.Items.Add("伝票計");
            cmbTax.Items.Add("請求時");
            cmbTax.Items.Add("非課税");
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     都道府県コンボ値セット </summary>
        /// ------------------------------------------------------------------
        private void cmbCitySet(ComboBox tempCmb)
        {
            tempCmb.Items.Add("北海道");
            tempCmb.Items.Add("青森県");
            tempCmb.Items.Add("岩手県");
            tempCmb.Items.Add("宮城県");
            tempCmb.Items.Add("秋田県");
            tempCmb.Items.Add("山形県");
            tempCmb.Items.Add("福島県");
            tempCmb.Items.Add("茨城県");
            tempCmb.Items.Add("栃木県");
            tempCmb.Items.Add("群馬県");
            tempCmb.Items.Add("埼玉県");
            tempCmb.Items.Add("千葉県");
            tempCmb.Items.Add("東京都");
            tempCmb.Items.Add("神奈川県");
            tempCmb.Items.Add("山梨県");
            tempCmb.Items.Add("長野県");
            tempCmb.Items.Add("新潟県");
            tempCmb.Items.Add("富山県");
            tempCmb.Items.Add("石川県");
            tempCmb.Items.Add("福井県");
            tempCmb.Items.Add("岐阜県");
            tempCmb.Items.Add("静岡県");
            tempCmb.Items.Add("愛知県");
            tempCmb.Items.Add("三重県");
            tempCmb.Items.Add("滋賀県");
            tempCmb.Items.Add("京都府");
            tempCmb.Items.Add("大阪府");
            tempCmb.Items.Add("兵庫県");
            tempCmb.Items.Add("奈良県");
            tempCmb.Items.Add("和歌山県");
            tempCmb.Items.Add("鳥取県");
            tempCmb.Items.Add("島根県");
            tempCmb.Items.Add("岡山県");
            tempCmb.Items.Add("広島県");
            tempCmb.Items.Add("山口県");
            tempCmb.Items.Add("徳島県");
            tempCmb.Items.Add("香川県");
            tempCmb.Items.Add("愛媛県");
            tempCmb.Items.Add("高知県");
            tempCmb.Items.Add("福岡県");
            tempCmb.Items.Add("佐賀県");
            tempCmb.Items.Add("長崎県");
            tempCmb.Items.Add("熊本県");
            tempCmb.Items.Add("大分県");
            tempCmb.Items.Add("宮崎県");
            tempCmb.Items.Add("鹿児島県");
            tempCmb.Items.Add("沖縄県");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // グリッド得意先データを表示する
            gridSearch(dataGridView1);
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     グリッド得意先データを表示する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------
        private void gridSearch(DataGridView g)
        {
            g.Rows.Clear();
            int iX = 0;

            // 略称検索
            var s = dts.得意先.Where(a => a.略称.Contains(textBox1.Text.Trim())).OrderBy(a => a.ID);

            // 電話番号検索
            s = s.Where(a => a.電話番号.Contains(txtsTel.Text.Trim())).OrderBy(a => a.ID);

            // 郵便番号検索
            s = s.Where(a => a.郵便番号.Contains(txtsZip.Text.Trim())).OrderBy(a => a.ID);

            // 請求先名称検索
            s = s.Where(a => a.請求先名称.Contains(textBox2.Text.Trim())).OrderBy(a => a.ID);

            // グリッドに表示
            foreach (var t in s)
            {
                g.Rows.Add();
                g[colID, iX].Value = t.ID.ToString();
                g[colRyaku, iX].Value = t.略称;
                g[colFuri, iX].Value = t.フリガナ;
                g[colName, iX].Value = t.名称;
                g[colKeisho, iX].Value = t.敬称;
                g[colTantou, iX].Value = t.担当者名;
                g[colBusho, iX].Value = t.部署名;
                g[colZip, iX].Value = t.郵便番号;
                g[colCity, iX].Value = t.都道府県;
                g[colAdd1, iX].Value = t.住所1;
                g[colAdd2, iX].Value = t.住所2;
                g[colTel, iX].Value = t.電話番号;
                g[colFax, iX].Value = t.FAX番号;
                g[colEmail, iX].Value = t.メールアドレス;

                if (t.社員Row == null)
                {
                    g[colEigyo, iX].Value = string.Empty;
                }
                else
                {
                    g[colEigyo, iX].Value = t.社員Row.氏名;
                }

                g[colShimebi, iX].Value = t.締日.ToString();
                g[colZei, iX].Value = t.税通知;

                //if (t.Is請求先名称Null())
                //{
                //    g[colSeName, iX].Value = string.Empty;
                //}
                //else
                //{
                //    g[colSeName, iX].Value = t.請求先名称;
                //}

                g[colSeName, iX].Value = t.請求先名称;

                g[colSeCity, iX].Value = t.請求先都道府県;
                g[colSeZip, iX].Value = t.請求先郵便番号;
                g[colSeAdd1, iX].Value = t.請求先住所1;
                g[colSeAdd2, iX].Value = t.請求先住所2;

                if (t.請求先担当者名 == null)
                {
                    g[colTantouS, iX].Value = string.Empty;
                }
                else
                {
                    g[colTantouS, iX].Value = t.請求先担当者名;
                }

                g[colBiko, iX].Value = t.備考;
                g[colAddDt, iX].Value = t.登録年月日;
                g[colUpDt, iX].Value = t.変更年月日;

                iX++;
            }

            g.CurrentCell = null;
        }
        
        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmShain();

            frm.ShowDialog();
            Utility.ComboShain.load(cmbShain);
        }

        private void txtName1_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtName1.Text) == false)
            {
                txtFuri.Text = "";
            }
        }

        private void txtName1_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/14
           string furi = MyTextKana.TextKana.textBox_KeyDown(txtName1, sender, e);

            if (furi != "")
            {
                //txtFuri.Text = furi;  // 2019/03/16 コメント化
                txtFuri.Text += furi;   // 2019/03/16
            }
        }

        private void mtxtZipCode_TextChanged(object sender, EventArgs e)
        {
            if (zipArray == null)
            {
                return;
            }

            MaskedTextBox mTxt = null;
            TextBox txtAdd = null;
            ComboBox cmbC = null;
            bool mc = false;

            if (sender == mtxtZipCode)
            {
                mTxt = mtxtZipCode;
                txtAdd = txtAddress1;
                cmbC = cmbCity;
            }
            else if (sender == mtxtZipCodeS)
            {
                mTxt = mtxtZipCodeS;
                txtAdd = txtAddress1S;
                cmbC = cmbCityS;
            }

            string zipText = mTxt.Text.Replace("-", "");

            if (zipText.Length == 7)
            {
                foreach (var t in zipArray)
                {
                    string[] r = t.Split(',');

                    if (zipText == r[2].Replace("\"", ""))
                    {
                        // 都道府県名を表示
                        for (int i = 0; i < cmbC.Items.Count; i++)
                        {
                            if (cmbC.Items[i].ToString() == r[6].Replace("\"", ""))
                            {
                                cmbC.SelectedIndex = i;
                                break;
                            }
                        }

                        // 住所
                        string ad = r[7].Replace("\"", "") + r[8].Replace("\"", "");
                            
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

                // 該当する郵便番号なしのとき
                if (!mc)
                {
                    cmbC.SelectedIndex = -1;
                    txtAdd.Text = string.Empty;
                }
            }
        }

        private void txtShimebi_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void cmbKeisho_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 請求先敬称にコピー 2019/02/20
            if (cmbKeishoS.SelectedIndex == -1)
            {
                cmbKeishoS.SelectedIndex = cmbKeisho.SelectedIndex;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmChargeName frm = new frmChargeName();
            frm.ShowDialog();

            this.Close();
        }
    }
}
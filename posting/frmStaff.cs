using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;
using MyLibrary;
using MyTextKana;


namespace posting
{
    public partial class frmStaff : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.配布員 cMaster = new Entity.配布員();

        const string MESSAGE_CAPTION = "配布員マスター";

        public frmStaff()
        {
            InitializeComponent();
        }

        string[] zipArray = null;

        private void form_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);
            
            this.darwinDataSet.Clear();
            this.darwinDataSet.EnforceConstraints = false;
            //this.配布員TableAdapter.Fill(this.darwinDataSet.配布員);
            this.ログインユーザーTableAdapter.Fill(this.darwinDataSet.ログインユーザー);

            // データグリッドビュー定義
            gridSetting(dataGridView1);

            // グリッドビューデータ表示
            //gridSerach(dataGridView1);

            //所属コンボデータロード
            Utility.ComboOffice.load(cmbShozoku);

            //支払形態コンボ
            cmbShiharaiSet();

            //口座種別コンボデータロード
            Utility.ComboKouza.load(cmbShubetsu);

            // 郵便番号CSV読み込み
            Utility.zipCsvLoad(ref zipArray);

            // 画面初期化
            DispClear();
        }

        #region グリッドビューカラム定義
        string colID = "col1";
        string colName = "col2";
        string colFuri = "col3";
        string colZip = "col4";
        string colAdd = "col5";
        string colMobile = "col6";
        string colTel = "col7";
        string colPcMail = "col8";
        string colMbMail = "col9";
        string colTourokuDt = "col10";
        string colKinmuKbn = "col11";
        string colHaihuKbn = "col12";
        string colHaifuMemo = "col13";
        string colShiharaiKbn = "col14";
        string colOfficeCode = "col15";
        string colBankCode1 = "col16";
        string colBankName = "col17";
        string colBankFuri = "col18";
        string colShitenCode = "col19";
        string colShitenName2 = "col20";
        string colShitenFuri = "col21";
        string colKouzaKbn = "col22";
        string colKouzaNum = "col23";
        string colKouzaName = "col24";
        string colMemo = "col25";
        string colAddDt = "col26";
        string colUpDt = "col27";
        string colMyNum = "col28";
        string colUser = "col29";
        #endregion

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
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
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 184;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(colID, "ｺｰﾄﾞ");
                tempDGV.Columns.Add(colName, "名称");
                tempDGV.Columns.Add(colFuri, "フリガナ");
                tempDGV.Columns.Add(colZip, "郵便番号");
                tempDGV.Columns.Add(colAdd, "住所");
                tempDGV.Columns.Add(colMobile, "携帯番号");
                tempDGV.Columns.Add(colTel, "自宅TEL");
                tempDGV.Columns.Add(colPcMail, "PCMail");
                tempDGV.Columns.Add(colMbMail, "携帯MAil");
                tempDGV.Columns.Add(colTourokuDt, "登録日");
                tempDGV.Columns.Add(colKinmuKbn, "勤務区分");
                tempDGV.Columns.Add(colHaihuKbn, "街頭配布");
                tempDGV.Columns.Add(colHaifuMemo, "街頭備考");
                tempDGV.Columns.Add(colShiharaiKbn, "支払区分");
                tempDGV.Columns.Add(colOfficeCode, "事務所");
                tempDGV.Columns.Add(colBankCode1, "銀行コード");
                tempDGV.Columns.Add(colBankName, "金融機関");
                tempDGV.Columns.Add(colBankFuri, "カナ");
                tempDGV.Columns.Add(colShitenCode, "支店コード");
                tempDGV.Columns.Add(colShitenName2, "支店名");
                tempDGV.Columns.Add(colShitenFuri, "カナ");
                tempDGV.Columns.Add(colKouzaKbn, "口座種別");
                tempDGV.Columns.Add(colKouzaNum, "口座番号");
                tempDGV.Columns.Add(colKouzaName, "名義");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colAddDt, "登録年月日");
                tempDGV.Columns.Add(colUpDt, "更新年月日");
                tempDGV.Columns.Add(colMyNum, "マイナンバー");
                tempDGV.Columns.Add(colUser, "更新ユーザー");

                tempDGV.Columns[0].Width = 60;
                tempDGV.Columns[1].Width = 180;
                tempDGV.Columns[2].Width = 160;
                tempDGV.Columns[3].Width = 80;
                tempDGV.Columns[4].Width = 300;
                tempDGV.Columns[5].Width = 120;
                tempDGV.Columns[6].Width = 120;
                tempDGV.Columns[7].Width = 140;
                tempDGV.Columns[8].Width = 140;
                tempDGV.Columns[9].Width = 100;
                tempDGV.Columns[10].Width = 80;
                tempDGV.Columns[11].Width = 100;
                tempDGV.Columns[12].Width = 160;
                tempDGV.Columns[13].Width = 80;
                tempDGV.Columns[14].Width = 80;
                tempDGV.Columns[15].Width = 100;
                tempDGV.Columns[16].Width = 160;
                tempDGV.Columns[17].Width = 80;
                tempDGV.Columns[25].Width = 130;    // 登録年月日 2015/07/16
                tempDGV.Columns[26].Width = 130;    // 更新年月日 2015/07/16
                tempDGV.Columns[27].Width = 130;    // マイナンバー
                tempDGV.Columns[28].Width = 130;    // 更新ユーザー

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

        /// --------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの指定行のデータを取得する </summary>
        /// <param name="dgv">
        ///     対象とするデータグリッドビューオブジェクト</param>
        /// --------------------------------------------------------------------------
        private Boolean GetData(DataGridView dgv,ref Entity.配布員 tempC)
        {
            int iX = 0;
            string sqlStr;
            Control.配布員 Staff = new Control.配布員();
            OleDbDataReader dr;

            sqlStr = " where 配布員.ID = " + Utility.strToInt(dgv[0, dgv.SelectedRows[iX].Index].Value.ToString());
            dr = Staff.FillBy(sqlStr);

            if (dr.HasRows == true)
            {
                while (dr.Read() == true)
                {
                    tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                    tempC.氏名 = dr["氏名"].ToString() + "";
                    tempC.フリガナ = dr["フリガナ"].ToString() + "";
                    tempC.郵便番号 = dr["郵便番号"].ToString() + "";
                    tempC.住所 = dr["住所"].ToString() + "";
                    tempC.携帯電話番号 = dr["携帯電話番号"].ToString() + "";
                    tempC.自宅電話番号 = dr["自宅電話番号"].ToString() + "";
                    tempC.PCアドレス = dr["PCアドレス"].ToString() + "";
                    tempC.携帯アドレス = dr["携帯アドレス"].ToString() + "";
                    tempC.登録日 = dr["登録日"].ToString() + "";
                    tempC.勤務区分 = Int32.Parse(dr["勤務区分"].ToString());
                    tempC.街頭配布区分 = Int32.Parse(dr["街頭配布区分"].ToString());
                    tempC.街頭配布備考 = dr["街頭配布備考"].ToString() + "";
                    tempC.事業所コード = Int32.Parse(dr["事業所コード"].ToString());
                    tempC.支払区分 = dr["支払区分"].ToString() + "";
                    tempC.金融機関コード = dr["金融機関コード"].ToString();
                    tempC.金融機関名 = dr["金融機関名"].ToString() + "";
                    tempC.金融機関名カナ = dr["金融機関名カナ"].ToString() + "";
                    tempC.支店コード = dr["支店コード"].ToString();
                    tempC.支店名 = dr["支店名"].ToString() + "";
                    tempC.支店名カナ = dr["支店名カナ"].ToString() + "";
                    tempC.口座種別 = Int32.Parse(dr["口座種別"].ToString());
                    tempC.口座番号 = dr["口座番号"].ToString() + "";
                    tempC.口座名義カナ = dr["口座名義カナ"].ToString() + "";
                    tempC.備考 = dr["備考"].ToString() + "";
                    tempC.マイナンバー = dr["マイナンバー"].ToString() + "";
                    tempC.ユーザーID = Utility.strToInt(dr["ユーザーID"].ToString());
                }
            }
            else
            {
                dr.Close();
                Staff.Close();
                return false;
            }

            dr.Close();
            Staff.Close();
            return true;
        }


        //グリッドからデータを選択
        private void GridEnter()
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[colName, dataGridView1.SelectedRows[iX].Index].Value + "が選択されました" + "\n";
            msgStr += "よろしいですか？";

            if (MessageBox.Show(msgStr, "配布員選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    //データを取得する
                    if (GetData(dataGridView1,ref cMaster) == false)
                    {
                        MessageBox.Show("該当するデータがマスターに登録されていません", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'データ値を取得
                    txtCode.Text = cMaster.ID.ToString();
                    txtName1.Text = cMaster.氏名;
                    txtFuri.Text = cMaster.フリガナ;
                    mtxtZipCode.Text = cMaster.郵便番号;
                    txtAddress.Text = cMaster.住所;
                    txtMobile.Text = cMaster.携帯電話番号;
                    txtTel.Text = cMaster.自宅電話番号;
                    txteMail.Text = cMaster.PCアドレス;
                    txteMailm.Text = cMaster.携帯アドレス;

                    if (cMaster.登録日 == "")
                    {
                        iDate.Checked = false;
                    }
                    else
                    {
                        iDate.Checked = true;
                        iDate.Value = Convert.ToDateTime(cMaster.登録日);
                    }

                    if (cMaster.勤務区分 == 1)
                    {
                        chkKinmu.Checked = true;
                    }
                    else
                    {
                        chkKinmu.Checked = false;
                    }

                    if (cMaster.街頭配布区分 == 1)
                    {
                        chkGaitou.Checked = true;
                    }
                    else
                    {
                        chkGaitou.Checked = false;
                    }

                    txtGaitouMemo.Text = cMaster.街頭配布備考;
                    cmbShiharai.Text = cMaster.支払区分;

                    Utility.ComboOffice.selectedIndex(cmbShozoku, cMaster.事業所コード);

                    txtBankCode.Text = cMaster.金融機関コード;
                    txtBank.Text = cMaster.金融機関名;
                    txtBankFuri.Text = cMaster.金融機関名カナ;

                    txtShitenCode.Text = cMaster.支店コード;
                    txtShiten.Text = cMaster.支店名;
                    txtShitenFuri.Text = cMaster.支店名カナ;

                    Utility.ComboKouza.selectedIndex(cmbShubetsu, cMaster.口座種別);
                    txtKouza.Text = cMaster.口座番号;
                    txtMeigi.Text = cMaster.口座名義カナ;
                    txtMyNumber.Text = cMaster.マイナンバー;      // 2015/07/16
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
                txtFuri.Text = "";
                mtxtZipCode.Text = "";
                txtAddress.Text = "";
                txtMobile.Text = "";
                txtTel.Text = "";
                txteMail.Text = "";
                txteMailm.Text = "";
                iDate.Checked = false;
                chkKinmu.Checked = false;
                chkGaitou.Checked = false;
                txtGaitouMemo.Text = "";
                cmbShiharai.SelectedIndex = -1;
                cmbShiharai.SelectedIndex = 0;  // 2018/02/22 週をデフォルトとする
                //cmbShozoku.SelectedIndex = -1;
                cmbShozoku.SelectedIndex = 1;   // 2018/02/22 新宿をデフォルトとする 
                txtBankCode.Text = "";
                txtBank.Text = "";
                txtBankFuri.Text = "";
                txtShitenCode.Text = "";
                txtShiten.Text = "";
                txtShitenFuri.Text = "";
                cmbShubetsu.SelectedIndex = -1;
                txtKouza.Text = "";
                txtMeigi.Text = "";
                txtMemo.Text = "";
                txtMyNumber.Text = "";  // 2015/07/16

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
                    Control.配布員 Staff = new Control.配布員();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Staff.Close();
                                return;
                            }

                            if (Staff.DataInsert(cMaster) == true)
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
                                Staff.Close();
                                return;
                            }

                            if (Staff.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", "所属マスター", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Staff.Close();

                    DispClear();

                    //データを 'darwinDataSet.配布員' テーブルに読み込みます。
                    this.配布員TableAdapter.Fill(this.darwinDataSet.配布員);
                    //this.配布員gridviewTableAdapter.Fill(this.darwinDataSet.配布員gridview);

                    // グリッド再表示
                    gridSerach(dataGridView1);
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
                    Control.配布員 Staff = new Control.配布員();
                    OleDbDataReader dr;

                    sqlStr = " where ID = " + txtCode.Text.ToString();
                    dr = Staff.FillBy(sqlStr);

                    if (dr.HasRows == true)
                    {
                        txtCode.Focus();
                        dr.Close();
                        Staff.Close();
                        throw new Exception("既に登録済みのコードです");
                    }

                    dr.Close();
                    Staff.Close();
                }

                //名称チェック
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("氏名を入力してください");
                }

                ////支払形態チェック
                //if (cmbShiharai.SelectedIndex == -1)
                //{
                //    cmbShiharai.Focus();
                //    throw new Exception("支払形態を選択してください");
                //}

                ////事業所チェック
                //if (cmbShozoku.SelectedIndex == -1)
                //{
                //    cmbShozoku.Focus();
                //    throw new Exception("事業所を選択してください");
                //}

                //金融機関コード
                if (txtBankCode.Text.ToString().Trim() != "")
                {
                    str = this.txtBankCode.Text.ToString();

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtBankCode.Focus();
                        throw new Exception("金融機関コードは数字で入力してください");
                    }
                }

                // 支店コード
                if (txtShitenCode.Text.ToString().Trim() != "")
                {
                    str = this.txtShitenCode.Text.ToString();

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtShitenCode.Focus();
                        throw new Exception("支店コードは数字で入力してください");
                    }
                }

                // マイナンバー
                if (txtMyNumber.Text != string.Empty)
                {
                    if (Utility.strToLong(txtMyNumber.Text) == 0)
                    {
                        this.txtShitenCode.Focus();
                        throw new Exception("マイナンバーは数字で入力してください");
                    }

                    if (txtMyNumber.Text.Length != 12)
                    {
                        this.txtShitenCode.Focus();
                        throw new Exception("マイナンバーは数字で12桁入力してください");
                    }
                }

                //クラスにデータセット
                cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.氏名 = txtName1.Text.ToString();
                cMaster.フリガナ = txtFuri.Text.ToString();

                if (mtxtZipCode.Text.ToString().Replace("-", "").Trim() == "")
                {
                    cMaster.郵便番号 = "";
                }
                else
                {
                    cMaster.郵便番号 = mtxtZipCode.Text.ToString();
                }

                cMaster.住所 = txtAddress.Text.ToString();
                cMaster.携帯電話番号 = txtMobile.Text.ToString();
                cMaster.自宅電話番号 = txtTel.Text.ToString();
                cMaster.PCアドレス = txteMail.Text.ToString();
                cMaster.携帯アドレス = txteMailm.Text.ToString();

                if (iDate.Checked == false)
                {
                    cMaster.登録日 = "";
                }
                else
                {
                    cMaster.登録日 = iDate.Value.ToShortDateString();
                }

                if (chkKinmu.Checked == true)
                {
                    cMaster.勤務区分 = 1;
                }
                else
                {
                    cMaster.勤務区分 = 0;
                }

                if (chkGaitou.Checked == true)
                {
                    cMaster.街頭配布区分 = 1;
                }
                else
                {
                    cMaster.街頭配布区分 = 0;
                }

                cMaster.街頭配布備考 = txtGaitouMemo.Text.ToString();
                cMaster.支払区分 = cmbShiharai.Text;

                if (cmbShozoku.SelectedIndex == -1)
                {
                    cMaster.事業所コード = 0;
                }
                else
                {
                    Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                    cmb1 = (Utility.ComboOffice)cmbShozoku.SelectedItem;
                    cMaster.事業所コード = cmb1.ID;
                }

                cMaster.金融機関コード = txtBankCode.Text.ToString().Trim();
                cMaster.金融機関名 = txtBank.Text.ToString();

                if (txtBankFuri.Text.ToString().Length > 15)
                {
                    cMaster.金融機関名カナ = txtBankFuri.Text.ToString().Substring(0, 15);
                }
                else
                {
                    cMaster.金融機関名カナ = txtBankFuri.Text.ToString();
                }

                cMaster.支店コード = txtShitenCode.Text.ToString().Trim();
                cMaster.支店名 = txtShiten.Text.ToString();

                if (txtShitenFuri.Text.ToString().Length > 15)
                {
                    cMaster.支店名カナ = txtShitenFuri.Text.ToString().Substring(0, 15);
                }
                else
                {
                    cMaster.支店名カナ = txtShitenFuri.Text.ToString();
                }

                if (cmbShubetsu.SelectedIndex == -1)
                {
                    cMaster.口座種別 = 0;
                }
                else
                {
                    Utility.ComboKouza cmb2 = new Utility.ComboKouza();
                    cmb2 = (Utility.ComboKouza)cmbShubetsu.SelectedItem;
                    cMaster.口座種別 = cmb2.ID;
                }

                cMaster.口座番号 = txtKouza.Text.ToString();
                cMaster.口座名義カナ = txtMeigi.Text.ToString();
                cMaster.備考 = txtMemo.Text.ToString();

                if (fMode.Mode == 0)
                {
                    cMaster.登録年月日 = DateTime.Now;
                    cMaster.変更年月日 = DateTime.Now;
                }
                else
                {
                    cMaster.変更年月日 = DateTime.Now;
                }

                cMaster.マイナンバー = txtMyNumber.Text;      // 2015/07/16
                cMaster.ユーザーID = global.loginUserID;      // 2015/07/16

                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        private void txtMEnter(object sender, EventArgs e)
        {
            MaskedTextBox objMtxt = (MaskedTextBox)sender;
            
            objMtxt.SelectAll();
            objMtxt.BackColor = Color.LightGray;
        }

        private void txtEnter(object sender, EventArgs e)
        {
            TextBox objtxt = (TextBox)sender;

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

            // 2019/03/16
            if (objtxt == txtName1 || objtxt == txtBank || objtxt == txtShiten)
            {
                MyTextKana.TextKana.textEnter(objtxt);
            }
        }

        private void txtMLeave(object sender, EventArgs e)
        {
            MaskedTextBox objMtxt = (MaskedTextBox)sender;
            objMtxt.BackColor = Color.White;
        }

        private void txtLeave(object sender, EventArgs e)
        {
            TextBox objtxt = (TextBox)sender;
            objtxt.BackColor = Color.White;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            // 他に登録されているときは削除不可とする
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(配布指示.配布員ID = " + txtCode.Text.ToString() + ")  ";

            OleDbDataReader dr;
            Control.配布指示 Shiji = new Control.配布指示();
            dr = Shiji.FillBy(SqlStr);

            // 該当配布員が登録されているときは削除不可とする
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "が配布指示データに登録されています", txtName1.Text.ToString() + "は削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Shiji.Close();
                return;
            }

            dr.Close();
            Shiji.Close();
            
            // 削除確認
            if (MessageBox.Show("削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // データ削除
            Control.配布員 Staff = new Control.配布員();
            if (Staff.DataDelete(Convert.ToInt32(txtCode.Text.ToString()))==true)
                MessageBox.Show("削除されました", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Staff.Close();

            DispClear();

            // データを 'darwinDataSet.配布員' テーブルに読み込みます。
            this.配布員TableAdapter.Fill(this.darwinDataSet.配布員);
            //this.配布員TableAdapter.Fill(this.darwinDataSet.配布員gridview);

            // グリッド再表示
            gridSerach(dataGridView1);
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

        private void cmbShiharaiSet()
        {
            cmbShiharai.Items.Clear();
            cmbShiharai.Items.Add("週");
            cmbShiharai.Items.Add("月");
        }

        private void label14_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Form frm = new frmOffice();
            frm.ShowDialog();

            //事業所コンボボックスデータロード
            Utility.ComboOffice.load(this.cmbShozoku);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // グリッドビューに絞り込み表示
            gridSerach(dataGridView1);
        }

        private void gridSerach(DataGridView d)
        {
            ////データを 'darwinDataSet.配布員' テーブルに読み込みます。
            //darwinDataSet ds = new darwinDataSet();
            //ds.Clear();
            //ds.EnforceConstraints = false;

            //const int MINID = 0;
            //const int MAXID = 999999;

            //int ID1, ID2;

            ////配布員IDの指定有無
            //if (textBoxID.Text.Length > 0)
            //{
            //    ID1 = int.Parse(textBoxID.Text);
            //    ID2 = int.Parse(textBoxID.Text);
            //}
            //else
            //{
            //    ID1 = MINID;
            //    ID2 = MAXID;
            //}

            //this.配布員TableAdapter.FillByName(ds.配布員, ID1, ID2, "%" + textBox1.Text.ToString() + "%");
            //dataGridView1.DataSource = ds.配布員;


            // 2015/07/17
            this.Cursor = Cursors.WaitCursor;   // 待機カーソルを表示
            this.配布員TableAdapter.Fill(this.darwinDataSet.配布員);
            //this.配布員TableAdapter.Fill(this.darwinDataSet.配布員gridview);

            // 2015/07/16
            var s = this.darwinDataSet.配布員.Where(a => a.ID > 0);

            if (textBoxID.Text != string.Empty)
            {
                s = s.Where(a => a.ID == Utility.strToInt(textBoxID.Text.Trim()));
            }

            if (textBox1.Text != string.Empty)
            {
                s = s.Where(a => a.氏名.Contains(textBox1.Text));
            }

            d.Rows.Clear();
            int r = 0;

            foreach (var t in s)
            {
                d.Rows.Add();

                d[colID, r].Value = t.ID.ToString();
                d[colName, r].Value = t.氏名;
                d[colFuri, r].Value = t.フリガナ;
                d[colZip, r].Value = t.郵便番号;
                d[colAdd, r].Value = t.住所;
                d[colMobile, r].Value = t.携帯電話番号;
                d[colTel, r].Value = t.自宅電話番号;
                d[colPcMail, r].Value = t.PCアドレス;
                d[colMbMail, r].Value = t.携帯アドレス;

                if (t.Is登録日Null())
                {
                    d[colTourokuDt, r].Value = string.Empty;
                }
                else
                {
                    d[colTourokuDt, r].Value = t.登録日.ToShortDateString();
                }

                d[colKinmuKbn, r].Value = t.勤務区分.ToString();
                d[colHaihuKbn, r].Value = t.街頭配布区分.ToString();
                d[colHaifuMemo, r].Value = t.街頭配布備考;
                d[colShiharaiKbn, r].Value = t.支払区分;
                d[colOfficeCode, r].Value = t.事業所コード.ToString();
                d[colBankCode1, r].Value = t.金融機関コード;
                d[colBankName, r].Value = t.金融機関名;
                d[colBankFuri, r].Value = t.金融機関名カナ;
                d[colShitenCode, r].Value = t.支店コード;
                d[colShitenName2, r].Value = t.支店名;
                d[colShitenFuri, r].Value = t.支店名カナ;
                d[colKouzaKbn, r].Value = t.口座種別.ToString();
                d[colKouzaNum, r].Value = t.口座番号;
                d[colKouzaName, r].Value = t.口座名義カナ;
                d[colMemo, r].Value = t.備考;
                d[colAddDt, r].Value = t.登録年月日;
                d[colUpDt, r].Value = t.変更年月日;
                d[colMyNum, r].Value = t.マイナンバー;

                if (t.ログインユーザーRow == null)
                {
                    d[colUser, r].Value = string.Empty;
                }
                else
                {
                    d[colUser, r].Value = t.ログインユーザーRow.ログインユーザー;
                }

                r++;
            }

            d.CurrentCell = null;
            this.Cursor = Cursors.Default;      // カーソルを戻す
        }
        
        private void txtName1_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtName1.Text) == false)
            {
                txtFuri.Text = "";
                txtMeigi.Text = "";
            }           
        }

        private void txtName1_KeyDown(object sender, KeyEventArgs e)
        {
            string furiName;
            furiName  = MyTextKana.TextKana.textBox_KeyDown(txtName1, sender, e);

            //txtFuri.Text = furiName;      // 2019/03/16 コメント化
            //txtMeigi.Text = furiName;     // 2019/03/16 コメント化

            // 2019/03/16
            if (furiName != "")
            {
                txtFuri.Text += furiName; 
                txtMeigi.Text += furiName;
            }
        }

        private void txtBank_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtBank.Text) == false)
            {
                txtBankFuri.Text = "";
            }
        }

        private void txtBank_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/16
            string furi = MyTextKana.TextKana.textBox_KeyDown(txtBank, sender, e);

            // 2019/03/16
            txtBankFuri.Text += furi;
        }

        private void txtShiten_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtShiten.Text) == false)
            {
                txtShiten.Text = "";
            }
        }

        private void txtShiten_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/16
            string furi = MyTextKana.TextKana.textBox_KeyDown(txtShiten, sender, e);

            // 2019/03/16
            txtShitenFuri.Text += furi;
        }

        private void textBoxID_Validated(object sender, EventArgs e)
        {
            if (textBoxID.Text.Length > 0)
            {
                if (Utility.NumericCheck(textBoxID.Text) == false)
                {
                    MessageBox.Show("配布員IDは数字で入力してください","検索ID");
                    textBoxID.Focus();
                }
            }
        }

        private void txtMyNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void txtMobile_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        private void fillByAllToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.配布員TableAdapter.FillByAll(this.darwinDataSet.配布員);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillByAllToolStripButton_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.配布員TableAdapter.FillByAll(this.darwinDataSet1.配布員);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillByAllToolStripButton_Click_2(object sender, EventArgs e)
        {
            try
            {
                this.配布員TableAdapter.FillByAll(this.darwinDataSet1.配布員);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
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
            bool mc = false;

            mTxt = mtxtZipCode;
            txtAdd = txtAddress;

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
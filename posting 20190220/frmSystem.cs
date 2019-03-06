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
    public partial class frmSystem : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.会社情報 cMaster = new Entity.会社情報();

        const string MESSAGE_CAPTION = "会社情報マスター";

        public frmSystem()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            // TODO: このコード行はデータを 'darwinDataSet.会社情報' テーブルに読み込みます。必要に応じて移動、または削除をしてください。
            //GridviewSet.Setting(dataGridView1);
            //this.会社情報TableAdapter.Fill(this.darwinDataSet.会社情報);

            //口座種別セット
            Utility.ComboKouza.load(comboBox1);

            DispClear();

            GridEnter();

        }

        private Boolean GetData(ref Entity.会社情報 tempC)
        {
            Control.会社情報 Kaisha = new Control.会社情報();
            OleDbDataReader dr;

            dr = Kaisha.Fill();

            if (dr.HasRows == true)
            {
                while (dr.Read() == true)
                {
                    tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                    tempC.会社名 = dr["会社名"].ToString() + "";
                    tempC.代表者氏名 = dr["代表者氏名"].ToString() + "";
                    tempC.役職名 = dr["役職名"].ToString() + "";
                    tempC.電話番号 = dr["電話番号"].ToString() + "";
                    tempC.FAX番号 = dr["FAX番号"].ToString() + "";
                    tempC.住所1 = dr["住所1"].ToString() + "";
                    tempC.住所2 = dr["住所2"].ToString() + "";
                    tempC.郵便番号 = dr["郵便番号"].ToString() + "";
                    tempC.メールアドレス = dr["メールアドレス"].ToString() + "";
                    tempC.部署名 = dr["部署名"].ToString() + "";
                    tempC.担当者名 = dr["担当者名"].ToString() + "";
                    tempC.特記事項1 = dr["特記事項1"].ToString() + "";
                    tempC.特記事項2 = dr["特記事項2"].ToString() + "";
                    tempC.依頼人コード = dr["依頼人コード"].ToString() + "";
                    tempC.依頼人名 = dr["依頼人名"].ToString() + "";
                    tempC.金融機関コード = dr["金融機関コード"].ToString() + "";
                    tempC.金融機関名 = dr["金融機関名"].ToString() + "";
                    tempC.支店コード = dr["支店コード"].ToString() + "";
                    tempC.支店名 = dr["支店名"].ToString() + "";
                    tempC.口座種別 = Int32.Parse(dr["口座種別"].ToString());
                    tempC.口座番号 = dr["口座番号"].ToString() + "";
                    tempC.配布フラグ = int.Parse(dr["配布フラグ"].ToString());
                    tempC.郵便番号CSVパス = dr["郵便番号CSVパス"].ToString();
                }
            }
            else
            {
                dr.Close();
                Kaisha.Close();
                return false;
            }

            dr.Close();
            Kaisha.Close();
            return true;
        }

        //グリッドからデータを選択
        private void GridEnter()
        {
            try
            {
                //データを取得する
                if (GetData(ref cMaster) == true)
                {
                    //'データ値を取得
                    //txtCode.Text = cMaster.ID.ToString();
                    txtName.Text = cMaster.会社名;
                    txtDaihyo.Text = cMaster.代表者氏名;
                    txtYaku.Text = cMaster.役職名;
                    txtTel.Text = cMaster.電話番号;
                    txtFax.Text = cMaster.FAX番号;
                    txtAddress1.Text = cMaster.住所1;
                    txtAddress2.Text = cMaster.住所2;
                    mtxtZipCode.Text = cMaster.郵便番号;
                    txtEmail.Text = cMaster.メールアドレス;
                    txtBusho.Text = cMaster.部署名;
                    txtTantou.Text = cMaster.担当者名;
                    txtMemo1.Text = cMaster.特記事項1;
                    txtMemo2.Text = cMaster.特記事項2;
                    txtIraiCode.Text = cMaster.依頼人コード;
                    txtIraiName.Text = cMaster.依頼人名;
                    txtBankCode.Text = cMaster.金融機関コード;
                    txtBankName.Text = cMaster.金融機関名;
                    txtShitenCode.Text = cMaster.支店コード;
                    txtShitenName.Text = cMaster.支店名;
                    Utility.ComboKouza.selectedIndex(comboBox1, Int32.Parse(cMaster.口座種別.ToString()));
                    txtNumber.Text = cMaster.口座番号;

                    txtFlg.Text = cMaster.配布フラグ.ToString();
                    txtZipPath.Text = cMaster.郵便番号CSVパス;   // 2015/10/04

                    //IDテキストボックスは編集不可とする
                    //txtCode.Enabled = false;

                    //ボタン状態

                    fMode.Mode = 1;     //フォームモードステータス:変更削除

                    txtName.Focus();
                }
                else
                {
                    fMode.Mode = 0;
                    txtName.Focus();
                }


            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "画面クリア", MessageBoxButtons.OK);
            }

        }

        /// <summary>
        /// 画面をクリアする
        /// </summary>
        private void DispClear()
        {

            try
            {
                fMode.Mode = 0;

                txtName.Text = "";
                txtDaihyo.Text = "";
                txtYaku.Text = "";
                txtTel.Text = "";
                txtFax.Text = "";
                txtAddress1.Text = "";
                txtAddress2.Text = "";
                mtxtZipCode.Text = "";
                txtEmail.Text = "";
                txtBusho.Text = "";
                txtTantou.Text = "";
                txtMemo1.Text = "";
                txtMemo2.Text = "";
                txtIraiCode.Text = "";
                txtIraiName.Text = "";
                txtBankCode.Text = "";
                txtBankName.Text = "";
                txtShitenCode.Text = "";
                txtShitenName.Text = "";
                comboBox1.SelectedIndex = -1;
                txtNumber.Text = "";

                //配布フラグ
                label22.Visible = false;
                txtFlg.Visible = false;

                txtZipPath.Text = string.Empty;     // 2015/10/04

                txtName.Focus();
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
                    Control.会社情報 Kaisha = new Control.会社情報();

                    switch (fMode.Mode)
                    {
                        case 0: //新規登録
                            if (MessageBox.Show("新規登録します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Kaisha.Close();
                                return;
                            }

                            if (Kaisha.DataInsert(cMaster) == true)
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
                                Kaisha.Close();
                                return;
                            }

                            if (Kaisha.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("更新されました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("更新に失敗しました", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Kaisha.Close();
                    this.Close();
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
                    //Control.会社情報 Kaisha = new Control.会社情報();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Kaisha.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Kaisha.Close();
                    //    throw new Exception("既に登録済みのコードです");
                    //}

                    //dr.Close();
                    //Kaisha.Close();

                }

                //会社名チェック
                if (txtName.Text.Trim().Length < 1)
                {
                    txtName.Focus();
                    throw new Exception("会社名を入力してください");
                }

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

                //支店コード
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

                //クラスにデータセット
                //cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.会社名 = txtName.Text.ToString();
                cMaster.代表者氏名 = txtDaihyo.Text.ToString();
                cMaster.役職名 = txtYaku.Text.ToString();
                cMaster.電話番号 = txtTel.Text.ToString();
                cMaster.FAX番号 = txtFax.Text.ToString();
                cMaster.住所1 = txtAddress1.Text.ToString();
                cMaster.住所2 = txtAddress2.Text.ToString();

                if (mtxtZipCode.Text.ToString().Replace("-", "").Trim() == "")
                {
                    cMaster.郵便番号 = "";
                }
                else
                {
                    cMaster.郵便番号 = mtxtZipCode.Text.ToString();
                }

                cMaster.メールアドレス = txtEmail.Text.ToString();
                cMaster.部署名 = txtBusho.Text.ToString();
                cMaster.担当者名 = txtTantou.Text.ToString();
                cMaster.特記事項1 = txtMemo1.Text.ToString();
                cMaster.特記事項2 = txtMemo2.Text.ToString();
                cMaster.依頼人コード = txtIraiCode.Text.ToString();
                cMaster.依頼人名 = txtIraiName.Text.ToString();
                cMaster.金融機関コード = txtBankCode.Text.ToString();
                cMaster.金融機関名 = txtBankName.Text.ToString();
                cMaster.支店コード = txtShitenCode.Text.ToString();
                cMaster.支店名 = txtShitenName.Text.ToString();

                Utility.ComboKouza cmb1 = new Utility.ComboKouza();
                cmb1 = (Utility.ComboKouza)comboBox1.SelectedItem;
                cMaster.口座種別 = cmb1.ID;
                cMaster.口座番号 = txtNumber.Text.ToString();

                if (txtFlg.Visible == true)
                {
                    cMaster.配布フラグ = int.Parse(txtFlg.Text);
                }

                if (fMode.Mode == 0) cMaster.登録年月日 = DateTime.Today;
                cMaster.変更年月日 = DateTime.Today;
                cMaster.郵便番号CSVパス = txtZipPath.Text;

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

            if (sender == txtName)
            {
                objtxt = txtName;
            }

            if (sender == txtDaihyo)
            {
                objtxt = txtDaihyo;
            }

            if (sender == txtYaku)
            {
                objtxt = txtYaku;
            }

            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;
            }
            
            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtMemo1)
            {
                objtxt = txtMemo1;
            }

            if (sender == txtMemo2)
            {
                objtxt = txtMemo2;
            }

            if (sender == txtIraiCode)
            {
                objtxt = txtIraiCode;
            }

            if (sender == txtIraiName)
            {
                objtxt = txtIraiName;
            }
            
            if (sender == txtBankCode)
            {
                objtxt = txtBankCode;
            }
            
            if (sender == txtBankName)
            {
                objtxt = txtBankName;
            }

            if (sender == txtShitenName)
            {
                objtxt = txtShitenName;
            }
            
            if (sender == txtShitenCode)
            {
                objtxt = txtShitenCode;
            }

            if (sender == txtNumber)
            {
                objtxt = txtNumber;
            }

            if (sender == txtFlg)
            {
                objtxt = txtFlg;
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

           if (sender == txtName)
           {
               objtxt = txtName;
           }

           if (sender == txtDaihyo)
           {
               objtxt = txtDaihyo;
           }

           if (sender == txtYaku)
           {
               objtxt = txtYaku;
           }

           if (sender == txtTel)
           {
               objtxt = txtTel;
           }

           if (sender == txtFax)
           {
               objtxt = txtFax;
           }

           if (sender == txtAddress1)
           {
               objtxt = txtAddress1;
           }

           if (sender == txtAddress2)
           {
               objtxt = txtAddress2;
           }

           if (sender == mtxtZipCode)
           {
               objMtxt = mtxtZipCode;
           }

           if (sender == txtEmail)
           {
               objtxt = txtEmail;
           }

           if (sender == txtBusho)
           {
               objtxt = txtBusho;
           }

           if (sender == txtTantou)
           {
               objtxt = txtTantou;
           }

           if (sender == txtMemo1)
           {
               objtxt = txtMemo1;
           }

           if (sender == txtMemo2)
           {
               objtxt = txtMemo2;
           }

           if (sender == txtIraiCode)
           {
               objtxt = txtIraiCode;
           }

           if (sender == txtIraiName)
           {
               objtxt = txtIraiName;
           }

           if (sender == txtBankCode)
           {
               objtxt = txtBankCode;
           }

           if (sender == txtBankName)
           {
               objtxt = txtBankName;
           }

           if (sender == txtShitenName)
           {
               objtxt = txtShitenName;
           }

           if (sender == txtShitenCode)
           {
               objtxt = txtShitenCode;
           }

           if (sender == txtNumber)
           {
               objtxt = txtNumber;
           }

           if (sender == txtFlg)
           {
               objtxt = txtFlg;
           }

           objtxt.BackColor = Color.White;
           objMtxt.BackColor = Color.White;

        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void txtFlg_Validating(object sender, CancelEventArgs e)
        {
            if (Utility.NumericCheck(txtFlg.Text) == false)
            {
                MessageBox.Show("数字の0,または1のみ有効です","配布フラグ");
                e.Cancel = true;
                return;
            }

            if ((txtFlg.Text != "0") && (txtFlg.Text != "1"))
            {
                MessageBox.Show("数字の0,または1のみ有効です", "配布フラグ");
                e.Cancel = true;
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("配布フラグはシステム管理者の確認のうえ変更してください。よろしいですか","配布フラグ設定",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            label22.Visible = true;
            txtFlg.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sPath = getZipPath();
            if (sPath != string.Empty)
            {
                txtZipPath.Text = sPath;
            }
        }

        private string getZipPath()
        {
            DialogResult ret;

            // ダイアログボックスの初期設定
            openFileDialog1.Title = "郵便番号CSVファイル選択";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "CSVファイル(*.csv)|*.csv|全てのファイル(*.*)|*.*";

            // ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

    }
}
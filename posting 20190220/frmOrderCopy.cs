using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmOrderCopy : Form
    {
        public frmOrderCopy(darwinDataSet _dts, darwinDataSetTableAdapters.受注1TableAdapter _adp, Int64 _jNum)
        {
            InitializeComponent();

            jNum = _jNum;
            dts = _dts;
            adp = _adp;
            
            Utility.comboGaichu.itemLoad(cmbGaichu2);   // 外注先２
            Utility.comboGaichu.itemLoad(cmbGaichu3);   // 外注先３
        }

        Int64 jNum = 0;
        darwinDataSet dts = new darwinDataSet();

        darwinDataSetTableAdapters.受注1TableAdapter adp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.受注種別TableAdapter jAdp = new darwinDataSetTableAdapters.受注種別TableAdapter();
        darwinDataSetTableAdapters.外注先TableAdapter gAdp = new darwinDataSetTableAdapters.外注先TableAdapter();

        private void frmOrderCopy_Load(object sender, EventArgs e)
        {
            jAdp.Fill(dts.受注種別);
            gAdp.Fill(dts.外注先);

            if (!dts.受注1.Any(a => a.ID == jNum))
            {
                btnAdd.Enabled = false;
                return;
            }

            btnAdd.Enabled = true;
            var s = dts.受注1.Single(a => a.ID == jNum);

            if (s.受注種別Row != null)
            {
                lblNaiyou.Text = s.受注種別Row.名称;
            }
            else
            {
                lblNaiyou.Text = string.Empty;
            }

            if (s.得意先Row != null)
            {
                lblClient.Text = s.得意先Row.略称;
                lblClientShimebi.Text = "※" + s.得意先Row.締日 + "日";
            }
            else
            {
                lblClient.Text = string.Empty;
                lblClientShimebi.Text = string.Empty;
            }

            txtChirashi.Text = s.チラシ名;

            StartDate.Checked = false;
            EndDate.Checked = false;
            NouhinDate.Checked = false;
            dtSeikyu.Checked = false;
            NyukinDate.Checked = false;
            dteGaichuPay.Checked = false;
            dtpGaichuPay.Checked = false;
            dtpGaichuPay2.Checked = false;
            dtpGaichuPay3.Checked = false;
            
            //if (!s.Is配布開始日Null())
            //{
            //    StartDate.Value = s.配布開始日;
            //}

            //if (!s.Is配布終了日Null())
            //{
            //    EndDate.Value = s.配布終了日;
            //}

            //if (!s.Is納品予定日Null())
            //{
            //    NouhinDate.Value = s.納品予定日;
            //}

            //if (!s.Is請求書発行日Null())
            //{
            //    dtSeikyu.Value = s.請求書発行日;
            //}

            //if (!s.Is入金予定日Null())
            //{
            //    NyukinDate.Value = s.入金予定日;
            //}

            //if (!s.Is外注支払日営業Null())
            //{
            //    dteGaichuPay.Value = s.外注支払日営業;
            //}

            //if (!s.Is外注支払日支払Null())
            //{
            //    dtpGaichuPay.Value = s.外注支払日支払;
            //}

            //if (!s.Is外注支払日支払2Null())
            //{
            //    dtpGaichuPay2.Value = s.外注支払日支払2;
            //}

            //if (!s.Is外注支払日支払3Null())
            //{
            //    dtpGaichuPay3.Value = s.外注支払日支払3;
            //}

            // 営業原価外注先
            if (s.外注先RowBy外注先_受注1 != null)
            {
                lblGaichu.BackColor = Color.White;
                lblGaichu.Text = s.外注先RowBy外注先_受注1.名称;
                dteGaichuPay.Enabled = true;
            }
            else
            {
                lblGaichu.BackColor = SystemColors.Control;
                lblGaichu.Text = string.Empty;
                dteGaichuPay.Enabled = false;
            }

            // 外注先１
            if (s.外注先RowBy外注先_受注11 != null)
            {
                lblGaichu1.BackColor = Color.White;
                lblGaichu1.Text = s.外注先RowBy外注先_受注11.名称;
                lblSD1.Text = "※" + s.外注先RowBy外注先_受注11.支払日 + "日";
                dtpGaichuPay.Enabled = true;

                label4.Enabled = true;
                label6.Enabled = true;
            }
            else
            {
                lblGaichu1.BackColor = SystemColors.Control;
                lblGaichu1.Text = string.Empty;
                lblSD1.Text = string.Empty;
                dtpGaichuPay.Enabled = false;

                //label4.Enabled = false;
                //label6.Enabled = false;
            }

            // 外注先２
            if (s.外注先RowBy外注先_受注12 != null)
            {
                //lblGaichu2.BackColor = Color.White;
                //lblGaichu2.Text = s.外注先RowBy外注先_受注12.名称;

                lblSD2.Text = "※" + s.外注先RowBy外注先_受注12.支払日 + "日";
                dtpGaichuPay2.Enabled = true;

                // 2018/03/02
                cmbGaichu2.Enabled = true;
                cmbGaichu2.SelectedValue = s.外注先ID支払2;

                label8.Enabled = true;
                label7.Enabled = true;
            }
            else
            {
                //lblGaichu2.BackColor = SystemColors.Control;
                //lblGaichu2.Text = string.Empty;

                lblSD2.Text = string.Empty;

                //dtpGaichuPay2.Enabled = false;

                // 2018/03/02 
                cmbGaichu2.Enabled = true;
                cmbGaichu2.SelectedIndex = -1;

                //label8.Enabled = false;
                //label7.Enabled = false;
            }

            // 外注先３
            if (s.外注先RowBy外注先_受注13 != null)
            {
                //lblGaichu3.BackColor = Color.White;
                //lblGaichu3.Text = s.外注先RowBy外注先_受注13.名称;

                lblSD3.Text = "※" + s.外注先RowBy外注先_受注13.支払日 + "日";
                dtpGaichuPay3.Enabled = true;

                // 2018/03/02
                cmbGaichu3.Enabled = true;
                cmbGaichu3.SelectedValue = s.外注先ID支払3;

                label12.Enabled = true;
                label9.Enabled = true;
            }
            else
            {
                //lblGaichu3.BackColor = SystemColors.Control;
                //lblGaichu3.Text = string.Empty;

                lblSD3.Text = string.Empty;

                //dtpGaichuPay3.Enabled = false;

                // 2018/03/02 
                cmbGaichu3.Enabled = true;
                cmbGaichu3.SelectedIndex = -1;

                //label12.Enabled = false;
                //label9.Enabled = false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            fStatus = false;
            Close();
        }

        private void frmOrderCopy_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            //Dispose();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!fDataCheck())
            {
                return;
            }

            // 確認メッセージ
            if (MessageBox.Show("受注確定書データのコピーを実施します。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // データ新規追加
            jDataCopy();

            // 確認メッセージ
            MessageBox.Show("データがコピーされ新たな受注確定書データが登録されました","処理完了",MessageBoxButtons.OK, MessageBoxIcon.Information);

            fStatus = true;

            // フォームを閉じる
            Close();
        }

        ///----------------------------------------------------
        /// <summary>
        ///     エラーチェック </summary>
        /// <returns>
        ///     true:エラーなし, false:エラー有り</returns>
        ///----------------------------------------------------
        private Boolean fDataCheck()
        {
            try
            {
                //ポスティングのときのみチェック対象とする項目
                if (lblNaiyou.Text == "ポスティング")
                {
                    //配布期間
                    if (StartDate.Value > EndDate.Value)
                    {
                        StartDate.Focus();
                        throw new Exception("配布期間が正しくありません");
                    }

                    // 納品日必須とするi
                    if (!NouhinDate.Checked)
                    {
                        NouhinDate.Focus();
                        throw new Exception("納品日が未入力です");
                    }
                }

                // 配布開始日：2016/08/03 受注内容がポスティングに限らず必須入力項目とする
                if (StartDate.Checked == false)
                {
                    StartDate.Focus();
                    throw new Exception("配布開始日を入力してください");
                }

                //配布終了日：2017/01/27 受注内容がポスティングに限らず必須入力項目とする
                if (EndDate.Checked == false)
                {
                    EndDate.Focus();
                    throw new Exception("配布終了日を入力してください");
                }
                
                // 請求書発行日必須とする
                if (!dtSeikyu.Checked)
                {
                    dtSeikyu.Focus();
                    throw new Exception("請求書発行日が未入力です");
                }

                // 支払期日必須とする
                if (!NyukinDate.Checked)
                {
                    NyukinDate.Focus();
                    throw new Exception("支払期日が未入力です");
                }

                // 営業原価
                // 外注先が選択済みで
                if (lblGaichu.Text != string.Empty)
                {
                    // 支払日が未入力のとき
                    if (!dteGaichuPay.Checked)
                    {
                        dteGaichuPay.Focus();
                        throw new Exception("営業原価の支払日が未入力です");
                    }
                }

                // 外注費支払用
                // 支払用外注先が選択済みで
                if (lblGaichu1.Text != string.Empty)
                {
                    // 支払日が未入力のとき
                    if (!dtpGaichuPay.Checked)
                    {
                        dtpGaichuPay.Focus();
                        throw new Exception("外注費支払用の支払日が未入力です");
                    }
                }

                // 外注費支払用２
                // 支払用外注先２が選択済みで
                //if (lblGaichu2.Text != string.Empty)
                //{
                //    // 支払日が未入力のとき
                //    if (!dtpGaichuPay2.Checked)
                //    {
                //        dtpGaichuPay2.Focus();
                //        throw new Exception("外注費支払用２の支払日が未入力です");
                //    }
                //}

                // 外注費支払用２
                // 支払用外注先２が選択済みで
                // 2018/03/02
                if (cmbGaichu2.SelectedIndex != -1)
                {
                    // 支払日が未入力のとき
                    if (!dtpGaichuPay2.Checked)
                    {
                        dtpGaichuPay2.Focus();
                        throw new Exception("外注費支払用２の支払日が未入力です");
                    }
                }

                // 外注費支払用３
                // 支払用外注先３が選択済みで
                //if (lblGaichu3.Text != string.Empty)
                //{
                //    // 支払日が未入力のとき
                //    if (!dtpGaichuPay3.Checked)
                //    {
                //        dtpGaichuPay3.Focus();
                //        throw new Exception("外注費支払用３の支払日が未入力です");
                //    }
                //}

                // 外注費支払用３
                // 支払用外注先３が選択済みで
                // 2018/03/02
                if (cmbGaichu3.SelectedIndex != -1)
                {
                    // 支払日が未入力のとき
                    if (!dtpGaichuPay3.Checked)
                    {
                        dtpGaichuPay3.Focus();
                        throw new Exception("外注費支払用３の支払日が未入力です");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "受注確定書コピー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }


        private void jDataCopy()
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                var s = dts.受注1.Single(a => a.ID == jNum);

                var r = dts.受注1.New受注1Row();

                r.ID = Utility.getOrderNum(jDate.Value);
                r.事業所ID = s.事業所ID;
                r.受注日 = DateTime.Parse(jDate.Value.ToShortDateString());
                r.受注区分 = s.受注区分;
                r.得意先ID = s.得意先ID;
                r.社員ID = s.社員ID;
                //r.チラシ名 = s.チラシ名;
                r.チラシ名 = txtChirashi.Text.Trim();   // 2018/02/15
                r.受注種別ID = s.受注種別ID;
                r.単価 = s.単価;
                r.枚数 = s.枚数;
                r.金額 = s.金額;
                r.消費税 = s.消費税;
                r.税込金額 = s.税込金額;
                r.値引額 = s.値引額;
                r.売上金額 = s.売上金額;
                r.税率 = s.税率;
                r.判型 = s.判型;
                r.配布単価 = s.配布単価;
                r.依頼先 = s.依頼先;
                r.原価 = s.原価;
                r.配布形態 = s.配布形態;
                r.配布条件 = s.配布条件;
                r.配布開始日 = DateTime.Parse(StartDate.Value.ToShortDateString());
                r.配布終了日 = DateTime.Parse(EndDate.Value.ToShortDateString());
                r.配布猶予 = s.配布猶予;

                if (NouhinDate.Checked)
                {
                    r.納品予定日 = DateTime.Parse(NouhinDate.Value.ToShortDateString());
                }
                else
                {
                    r.Set納品予定日Null();
                }

                r.納品形態 = s.納品形態;
                r.請求書 = s.請求書;
                //r.請求書ID = s.請求書ID;
                r.請求書ID = global.FLGOFF;    // 2018/02/15
                r.請求書発行日 = DateTime.Parse(dtSeikyu.Value.ToShortDateString());
                r.入金方法 = s.入金方法;
                r.入金予定日 = DateTime.Parse(NyukinDate.Value.ToShortDateString());
                r.報告時期 = s.報告時期;
                r.報告精度 = s.報告精度;
                r.報告方法 = s.報告方法;
                r.メールアドレス = s.メールアドレス;
                r.振込口座ID = s.振込口座ID;
                r.未配布情報有無 = s.未配布情報有無;
                r.枝番有無 = s.枝番有無;
                r.特記事項 = s.特記事項;
                r.エリア備考 = s.エリア備考;
                r.完了区分 = s.完了区分;
                r.併配除外 = s.併配除外;
                r.登録年月日 = DateTime.Now;
                r.変更年月日 = DateTime.Now;
                r.外注先ID営業 = s.外注先ID営業;

                if (dteGaichuPay.Checked)
                {
                    r.外注支払日営業 = DateTime.Parse(dteGaichuPay.Value.ToShortDateString());
                }
                else
                {
                    r.Set外注支払日営業Null();
                }

                r.外注原価営業 = s.外注原価営業;

                if (s.Is外注依頼日営業Null())
                {
                    r.Set外注依頼日営業Null();
                }
                else
                {
                    r.外注依頼日営業 = s.外注依頼日営業;
                }

                r.外注先ID支払 = s.外注先ID支払;

                if (dtpGaichuPay.Checked)
                {
                    r.外注支払日支払 = DateTime.Parse(dtpGaichuPay.Value.ToShortDateString());
                }
                else
                {
                    r.Set外注支払日支払Null();
                }

                r.外注原価支払 = s.外注原価支払;

                if (s.Is外注依頼日支払Null())
                {
                    r.Set外注依頼日支払Null();
                }
                else
                {
                    r.外注依頼日支払 = s.外注依頼日支払;
                }

                r.ユーザーID = global.loginUserID;
                r.案件種別 = s.案件種別;
                //r.外注支払ID = s.外注支払ID;
                r.外注支払ID = string.Empty;    // 2018/03/07
                //r.受注確定書発行 = s.受注確定書発行;
                r.受注確定書発行 = global.FLGOFF;  // 2018/02/15
                r.登録ユーザーID = global.loginUserID;

                if (s.Is外注渡し日Null())
                {
                    r.Set外注渡し日Null();
                }
                else
                {
                    r.外注渡し日 = s.外注渡し日;
                }

                r.外注受け渡し担当者 = s.外注受け渡し担当者;
                r.外注委託枚数 = s.外注委託枚数;
                r.業種 = s.業種;

                //r.外注先ID支払2 = s.外注先ID支払2; 2018/03/02

                // 2018/03/02
                if (cmbGaichu2.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbGaichu2.SelectedItem;
                    r.外注先ID支払2 = cmb.ID;
                }
                else
                {
                    r.外注先ID支払2 = 0;
                }

                if (dtpGaichuPay2.Checked)
                {
                    r.外注支払日支払2 = DateTime.Parse(dtpGaichuPay2.Value.ToShortDateString());
                }
                else
                {
                    r.Set外注支払日支払2Null();
                }

                r.外注原価支払2 = s.外注原価支払2;

                //r.外注先ID支払3 = s.外注先ID支払3; 2018/03/02

                // 2018/03/02
                if (cmbGaichu3.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbGaichu3.SelectedItem;
                    r.外注先ID支払3 = cmb.ID;
                }
                else
                {
                    r.外注先ID支払3 = 0;
                }

                if (dtpGaichuPay3.Checked)
                {
                    r.外注支払日支払3 = DateTime.Parse(dtpGaichuPay3.Value.ToShortDateString());
                }
                else
                {
                    r.Set外注支払日支払3Null();
                }
                
                r.外注原価支払3 = s.外注原価支払3;
                r.外注委託枚数2 = s.外注委託枚数2;
                r.外注委託枚数3 = s.外注委託枚数3;

                if (s.Is外注渡し日2Null())
                {
                    r.Set外注渡し日2Null();
                }
                else
                {
                    r.外注渡し日2 = s.外注渡し日2;
                }

                if (s.Is外注渡し日3Null())
                {
                    r.Set外注渡し日3Null();
                }
                else
                {
                    r.外注渡し日3 = s.外注渡し日3;
                }

                r.外注受け渡し担当者2 = s.外注受け渡し担当者2;
                r.外注受け渡し担当者3 = s.外注受け渡し担当者3;
                //r.外注支払ID2 = s.外注支払ID2;
                //r.外注支払ID3 = s.外注支払ID3;
                r.外注支払ID2 = string.Empty;   // 2018/03/07
                r.外注支払ID3 = string.Empty;   // 2018/03/07
                r.納品書発行 = global.FLGOFF;

                dts.受注1.Add受注1Row(r);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        // 終了ステータス
        public bool fStatus { get; set; }
    }
}

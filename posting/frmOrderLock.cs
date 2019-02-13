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
    public partial class frmOrderLock : Form
    {
        public frmOrderLock()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void frmOrderLock_Load(object sender, EventArgs e)
        {
            adp.Fill(dts.受注編集制限);

            // データ表示
            dataShow();
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注編集制限TableAdapter adp = new darwinDataSetTableAdapters.受注編集制限TableAdapter();

        private void dataShow()
        {
            // ログインタイプ
            Utility.comboLogintype.itemLoad(comboBox1);

            // 受注編集制限を登録済みか？
            if (dts.受注編集制限.Any(a => a.ID == global.lockKey))
            {
                var s = dts.受注編集制限.Single(a => a.ID == global.lockKey);
                dateTimePicker1.Checked = true;
                dateTimePicker1.Value = s.請求書発行日;
                comboBox1.SelectedValue = s.ログイングループ;

                // 解除ボタンをアクティブにする
                button3.Enabled = true;
            }
            else
            {
                dateTimePicker1.Checked = false;
                comboBox1.SelectedIndex = -1;

                // 解除ボタンを非アクティブにする
                button3.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 閉じる
            Close();
        }

        private void frmOrderLock_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataUpdateRun();
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     編集ロック登録 </summary>
        ///-------------------------------------------------------------------
        private void dataUpdateRun()
        {
            if (errCheck())
            {
                string msg = dateTimePicker1.Value.ToShortDateString() + " までに発行した請求書の受注確定書の編集をロックします。" + Environment.NewLine;
                msg += "ただし、ログインタイプ「" + comboBox1.Text + "」のみ編集可能です。" + Environment.NewLine;
                msg += "よろしいですか？";

                if (MessageBox.Show(msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                // データ登録
                dataUpdate();

                // 結果
                MessageBox.Show("受注確定書編集の編集ロックを行いました", "登録完了", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 閉じる
                Close();
            }
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     エラーチェック </summary>
        /// <returns>
        ///     true:エラーなし、false:エラーあり</returns>
        ///-----------------------------------------------------------------
        private bool errCheck()
        {
            if (!dateTimePicker1.Checked)
            {
                MessageBox.Show("請求書発行日を指定してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("編集可能なログイングループを指定してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     データ登録 </summary>
        ///--------------------------------------------------------
        private void dataUpdate()
        {
            // 受注編集制限を登録済みか？
            if (dts.受注編集制限.Any(a => a.ID == global.lockKey))
            {
                var s = dts.受注編集制限.Single(a => a.ID == global.lockKey);

                s.請求書発行日 = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());

                Utility.comboLogintype cmb = new Utility.comboLogintype();
                cmb = (Utility.comboLogintype)comboBox1.SelectedItem;
                s.ログイングループ = cmb.ID;
                s.登録年月日 = DateTime.Now;
            }
            else
            {
                darwinDataSet.受注編集制限Row r = dts.受注編集制限.New受注編集制限Row();
                r.ID = global.lockKey;
                r.請求書発行日 = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());

                Utility.comboLogintype cmb = new Utility.comboLogintype();
                cmb = (Utility.comboLogintype)comboBox1.SelectedItem;
                r.ログイングループ = cmb.ID;
                r.登録年月日 = DateTime.Now;

                dts.受注編集制限.Add受注編集制限Row(r);
            }

            // データベース更新
            adp.Update(dts.受注編集制限);

        }

        ///------------------------------------------------------------
        /// <summary>
        ///     レコード削除 </summary>
        ///------------------------------------------------------------
        private void dataUnLock()
        {
            string msg = "受注確定書の編集ロックを解除します。" + Environment.NewLine;
            msg += "よろしいですか？";

            if (MessageBox.Show(msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 受注編集制限を登録済みか？
            if (dts.受注編集制限.Any(a => a.ID == global.lockKey))
            {
                var s = dts.受注編集制限.Single(a => a.ID == global.lockKey);
                s.Delete();

                // データベース更新
                adp.Update(dts.受注編集制限);
            }

            // 結果
            MessageBox.Show("受注確定書の編集ロックを解除しました", "解除完了", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // 閉じる
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 編集ロック制限データ削除
            dataUnLock();
        }
    }
}

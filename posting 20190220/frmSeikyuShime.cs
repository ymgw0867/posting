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
    public partial class frmSeikyuShime : Form
    {
        public frmSeikyuShime()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 閉じる
            this.Close();
        }

        private void frmSeikyuShime_Load(object sender, EventArgs e)
        {

        }

        private void frmSeikyuShime_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!errCheck())
            {
                return;
            }

            if (MessageBox.Show(txtYear.Text + "年" + txtMonth.Text + "月の請求締め処理を実行します。よろしいですか", "請求締め処理",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            this.Cursor = Cursors.WaitCursor;

            // 請求書データ集計
            clsSeikyuData cls = new clsSeikyuData(this);
            cls.Summary(Utility.strToInt(txtYear.Text), Utility.strToInt(txtMonth.Text));

            // 終了メッセージ
            MessageBox.Show("終了しました", "請求締め処理",MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Cursor = Cursors.Default;

            // 閉じる
            this.Close();
        }

        private bool errCheck()
        {
            bool res = true;

            if (Utility.strToInt(txtYear.Text) < 2010)
            {
                MessageBox.Show("対象年が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.strToInt(txtMonth.Text) < 1 || Utility.strToInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("対象月が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }
            
            return res;
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
    }
}

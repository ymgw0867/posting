using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace posting
{
    public partial class frmMenuMST : Form
    {
        public frmMenuMST()
        {
            InitializeComponent();
        }

        clsMenu _cm;

        private void button1_Click(object sender, EventArgs e)
        {
            Form frm = new frmOffice();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form frm = new frmShozoku();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form frm = new frmJShubetsu();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form frm = new frmTax();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form frm = new frmTown();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form frm = new frmShimebi();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form frm = new frmSize();

            this.Hide();
            frm.ShowDialog();
            this.Show();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form frm = new frmIssueMode();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Form frm = new frmShain();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Form frm = new frmKouza();

            this.Hide();
            frm.ShowDialog();
            this.Show();

        }

        private void button11_Click(object sender, EventArgs e)
        {
            Form frm = new frmClient();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Form frm = new frmStaff();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            frmSystem frm = new frmSystem();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmMenuMST_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("マスター管理画面を閉じます。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }

            this.Dispose();
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            frmMihaifuMST frm = new frmMihaifuMST();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            frmShowTable frm = new frmShowTable();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmGaichu frm = new frmGaichu();
            frm.ShowDialog();
            this.Show();
        }

        private void frmMenuMST_Load(object sender, EventArgs e)
        {
            // メニュータイトルクラス 2015/07/07
            clsMenu _cm = new clsMenu();

            // メニュータイトルCSVの読込 2015/07/07
            _cm.loadMenu();

            // メニュータイトルをセット 2015/07/07
            Utility.getMenuTittle(button11, _cm);
            Utility.getMenuTittle(button12, _cm);
            Utility.getMenuTittle(button9, _cm);
            Utility.getMenuTittle(button1, _cm);
            Utility.getMenuTittle(button2, _cm);
            Utility.getMenuTittle(button3, _cm);
            Utility.getMenuTittle(button7, _cm);
            Utility.getMenuTittle(button8, _cm);
            Utility.getMenuTittle(button6, _cm);
            Utility.getMenuTittle(button5, _cm);
            Utility.getMenuTittle(button13, _cm);
            Utility.getMenuTittle(button4, _cm);
            Utility.getMenuTittle(button10, _cm);
            Utility.getMenuTittle(button15, _cm);
            Utility.getMenuTittle(button14, _cm);
            Utility.getMenuTittle(button16, _cm);
            Utility.getMenuTittle(button17, _cm);
            Utility.getMenuTittle(button18, _cm);
            Utility.getMenuTittle(button19, _cm);

            // メニューボタン表示状態初期化
            button11.Enabled = false;
            button12.Enabled = false;
            button9.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            button6.Enabled = false;
            button5.Enabled = false;
            button13.Enabled = false;
            button4.Enabled = false;
            button10.Enabled = false;
            button15.Enabled = false;
            button16.Enabled = false;
            button17.Enabled = false;
            button18.Enabled = false;
            button19.Enabled = false;

            // ログインユーザーごとのメニュー制御
            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter hAdp = new darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter();
            darwinDataSetTableAdapters.ログインタイプタグTableAdapter tAdp = new darwinDataSetTableAdapters.ログインタイプタグTableAdapter();
            hAdp.Fill(dts.ログインタイプヘッダ);
            tAdp.Fill(dts.ログインタイプタグ);

            foreach (var h in dts.ログインタイプヘッダ.Where(a => a.Id == global.loginType))
            {
                foreach (var item in h.GetログインタイプタグRows())
                {
                    if (menuButtonStatus(button11, item.tag)) continue;
                    if (menuButtonStatus(button12, item.tag)) continue;
                    if (menuButtonStatus(button9, item.tag)) continue;
                    if (menuButtonStatus(button1, item.tag)) continue;
                    if (menuButtonStatus(button2, item.tag)) continue;
                    if (menuButtonStatus(button3, item.tag)) continue;
                    if (menuButtonStatus(button7, item.tag)) continue;
                    if (menuButtonStatus(button8, item.tag)) continue;
                    if (menuButtonStatus(button6, item.tag)) continue;
                    if (menuButtonStatus(button5, item.tag)) continue;
                    if (menuButtonStatus(button13, item.tag)) continue;
                    if (menuButtonStatus(button4, item.tag)) continue;
                    if (menuButtonStatus(button10, item.tag)) continue;
                    if (menuButtonStatus(button15, item.tag)) continue;
                    if (menuButtonStatus(button16, item.tag)) continue;
                    if (menuButtonStatus(button17, item.tag)) continue;
                    if (menuButtonStatus(button18, item.tag)) continue;
                    if (menuButtonStatus(button19, item.tag)) continue;
                }
            }
        }

        private bool menuButtonStatus(Button btn, int tag)
        {
            if (Utility.strToInt(btn.Tag.ToString()) == tag)
            {
                btn.Enabled = true;
                return true;
            }

            return false;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmLoginType frm = new frmLoginType();
            frm.ShowDialog();
            this.Show();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmLoginUser frm = new frmLoginUser();
            frm.ShowDialog();
            this.Show();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmOrderLock frm = new frmOrderLock();
            frm.ShowDialog();
            this.Show();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmEditLock frm = new frmEditLock();
            frm.ShowDialog();
            this.Show();
        }
    }
}
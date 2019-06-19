using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace posting
{
    public partial class frmChargeName : Form
    {
        public frmChargeName()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("得意先毎に請求先・敬称、部署名をセットします。よろしいですか？","確認",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            setClientdata();
        }

        private void setClientdata()
        {
            toolStripProgressBar1.Visible = true;
            label1.Visible = true;

            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.得意先TableAdapter adp = new darwinDataSetTableAdapters.得意先TableAdapter();

            adp.Fill(dts.得意先);

            toolStripProgressBar1.Minimum = 1;
            toolStripProgressBar1.Maximum = dts.得意先.Count();

            try
            {
                int dCnt = 0;

                Cursor = Cursors.WaitCursor;

                foreach (var t in dts.得意先.OrderBy(a => a.ID))
                {
                    label1.Text = t.ID.ToString("D4") + " " + t.略称;

                    // リストビューへ表示
                    listBox1.Items.Add(label1.Text);
                    listBox1.TopIndex = listBox1.Items.Count - 1;

                    dCnt++;
                    toolStripProgressBar1.Value = dCnt;
                    System.Threading.Thread.Sleep(50);
                    Application.DoEvents();

                    label2.Text = dCnt + "/" + toolStripProgressBar1.Maximum;

                    // 請求先敬称セット
                    if (t.請求先担当者名 == null || t.請求先担当者名 == string.Empty)
                    {
                        t.請求先敬称 = "御中";
                    }
                    else
                    {
                        t.請求先敬称 = "様";
                    }

                    // 請求先部署名セット
                    t.請求先部署名 = t.部署名;

                    t.変更年月日 = DateTime.Now;
                }

                // データベース更新
                adp.Update(dts.得意先);

                label1.Text = string.Empty;

                Cursor = Cursors.Default;
                MessageBox.Show(dCnt.ToString("#,##0") + "件、処理しました", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 画面を閉じる
                Close();
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

        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}

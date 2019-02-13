using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace posting
{
    public partial class frmPostingTree : Form
    {
        const string MESSAGE_CAPTION = "ポスティングツリー";

        public frmPostingTree()
        {
            InitializeComponent();
        }

        private void frmPostingTree_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            button2.Enabled = false;
            button3.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            NodesShow(int.Parse(txtYear.Text), int.Parse(txtMonth.Text));
        }

        private void NodesShow(int tempYear,int tempMonth)
        {
            string sqlSTR;
            string NodeName1, nDate, nKanryo,nName;
            long wID;
            int nIndex=0;
            int nCnt=0;

            //ツリービュークリア
            treeView1.Nodes.Clear();

            OleDbDataReader dR;
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();

            cn = Con.GetConnection();

            sqlSTR = "";
            sqlSTR += "select 受注.ID, 受注.チラシ名, 受注.枚数, 配布エリア.町名ID, 町名.名称 AS 町名,";
            sqlSTR += "配布指示.配布日, 配布員.氏名,配布エリア.予定枚数, 配布エリア.報告枚数, ";
            sqlSTR += "配布エリア.報告残数, 配布エリア.完了区分 ";
            sqlSTR += "from 受注 left join 配布エリア ";
            sqlSTR += "on 受注.ID = 配布エリア.受注ID left join 町名 ";
            sqlSTR += "on 配布エリア.町名ID = 町名.ID left join 配布指示 ";
            sqlSTR += "on 配布エリア.配布指示ID = 配布指示.ID left join 配布員 ";
            sqlSTR += "on 配布指示.配布員ID = 配布員.ID ";
            sqlSTR += "where (受注.受注種別ID = 1) and ";
            sqlSTR += "(year(受注日) = ?) AND (month(受注日) = ?) ";
            sqlSTR += "order by 受注.ID desc,町名ID";

            OleDbCommand SCom = new OleDbCommand();

            SCom.CommandText = sqlSTR;
            SCom.Parameters.AddWithValue("@year", tempYear);
            SCom.Parameters.AddWithValue("@month", tempMonth);

            SCom.Connection = cn;
            dR = SCom.ExecuteReader();

            if (dR.HasRows == true)
            {
                treeView1.Nodes.Add("受注確定書");
                treeView1.Nodes[0].ImageIndex = 4;  //アイコンの設定

                wID = 0;
                nIndex = 0;

                while (dR.Read())
                {
                    if (wID != long.Parse(dR["ID"].ToString()))
                    {
                        //受注確定書毎のエリア数
                        if (wID != 0)
                        {
                            treeView1.Nodes[0].Nodes[nIndex - 1].Text += "(" + nCnt.ToString() + ")";
                        }

                        //受注確定書情報
                        NodeName1 = dR["ID"].ToString() + " " + dR["チラシ名"].ToString() + " " + dR["枚数"].ToString();
                        treeView1.Nodes[0].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex].NodeFont = new Font("ＭＳ Ｐゴシック", 10, FontStyle.Regular);
                        treeView1.Nodes[0].Nodes[nIndex].ImageIndex = 2;    //アイコンの設定
                        nIndex++;
                        nCnt = 0;
                    }

                    //配布エリア情報
                    if (dR["町名"] != DBNull.Value)
                    {

                        if (dR["配布日"] == DBNull.Value)
                        {
                            nDate = "----/--/--";
                        }
                        else
                        {
                            nDate = DateTime.Parse(dR["配布日"].ToString()).ToShortDateString();
                        }

                        if (dR["氏名"] == DBNull.Value)
                        {
                            nName = "********";
                        }
                        else
                        {
                            nName = dR["氏名"].ToString();
                        }

                        if (dR["完了区分"].ToString() == "0")
                        {
                            nKanryo = "";
                        }
                        else
                        {
                            nKanryo = "完了";
                        }

                        NodeName1 = int.Parse(dR["町名ID"].ToString()).ToString("d4") + " " + dR["町名"].ToString() + " " + nDate + " " + nName + " " + dR["予定枚数"].ToString() + " " + nKanryo;
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes[nCnt].ImageIndex = 1;    //アイコンの設定

                        //未完了は赤表示
                        if (nKanryo == "")
                        {
                            treeView1.Nodes[0].Nodes[nIndex - 1].Nodes[nCnt].ForeColor = Color.Red;
                            treeView1.Nodes[0].Nodes[nIndex - 1].ForeColor = Color.Red;
                        }

                        nCnt++;
                    }
                    else
                    {
                        treeView1.Nodes[0].Nodes[nIndex - 1].ForeColor = Color.Red;
                    }

                    wID = long.Parse(dR["ID"].ToString());

                }

                treeView1.Nodes[0].Nodes[nIndex - 1].Text += "(" + nCnt.ToString() + ")";
            }
            else
            {
                MessageBox.Show("該当する受注確定書がありません",MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Information);
            }

            dR.Close();
            Con.Close();
            
            if (nIndex > 0)
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            treeView1.CollapseAll();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            treeView1.ExpandAll();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmPostingTree_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void txtYear_Validating(object sender, CancelEventArgs e)
        {

            string str;
            int d;

            if (txtYear.Text == null)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtYear.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

        }

        private void txtMonth_Validating(object sender, CancelEventArgs e)
        {

            string str;
            int d;

            if (txtMonth.Text == null)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtMonth.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("数字で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("1～12で入力してください", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {

            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {

            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.BackColor = Color.White;
        }
    }
}
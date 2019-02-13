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
    public partial class frmHaifuTree : Form
    {
        const string MESSAGE_CAPTION = "配布指示ツリー";

        public frmHaifuTree()
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

            progressBar1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            NodesShow(int.Parse(txtYear.Text), int.Parse(txtMonth.Text));
        }

        private void NodesShow(int tempYear,int tempMonth)
        {
            string sqlSTR;
            string NodeName1, nKanryo;
            string nDate = "";

            long wID;
            int nIndex=0;
            int nCnt=0;

            int PrgMax;
            int PrgVal = 0;

            //ツリービュークリア
            treeView1.Nodes.Clear();

            OleDbDataReader dR;
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();

            cn = Con.GetConnection();

            //データ件数取得
            sqlSTR = "";
            sqlSTR += "select count(*) as Cnt ";
            sqlSTR += "from 配布指示 LEFT OUTER JOIN 配布エリア ";
            sqlSTR += "ON 配布指示.ID = 配布エリア.配布指示ID ";
            sqlSTR += "where (year(配布指示.配布日) = " + tempYear.ToString() + ") and (month(配布指示.配布日) = " + tempMonth.ToString() + ")";

            OleDbCommand SCom = new OleDbCommand();

            SCom.CommandText = sqlSTR;
            SCom.Connection = cn;
            dR = SCom.ExecuteReader();
            dR.Read();
            PrgMax = int.Parse(dR["Cnt"].ToString());
            dR.Close();

            //プログレスバーMAX件数,MIN件数設定
            progressBar1.Maximum = PrgMax;
            progressBar1.Minimum = 0;

            //データ取得
            sqlSTR = "";
            sqlSTR += "select 配布指示.ID as 配布指示ID, 配布指示.配布日 as 配布指示配布日, ";
            sqlSTR += "配布員.氏名, 受注.チラシ名, 町名.ID AS 町名コード,";
            sqlSTR += "町名.名称 as 町名名称, 配布エリア.予定枚数, 配布エリア.配布単価, ";
            sqlSTR += "配布エリア.報告枚数, 配布エリア.報告残数,配布エリア.完了区分 ";
            sqlSTR += "from 配布指示 LEFT OUTER JOIN 配布エリア ";
            sqlSTR += "ON 配布指示.ID = 配布エリア.配布指示ID LEFT OUTER JOIN ";
            sqlSTR += "受注 ON 配布エリア.受注ID = 受注.ID LEFT OUTER JOIN ";
            sqlSTR += "配布員 ON 配布指示.配布員ID = 配布員.ID LEFT OUTER JOIN ";
            sqlSTR += "町名 ON 配布エリア.町名ID = 町名.ID ";
            sqlSTR += "where (year(配布指示.配布日) = ?) and (month(配布指示.配布日) = ?) ";
            sqlSTR += "ORDER BY 配布指示.ID DESC,配布エリア.受注ID";

            //OleDbCommand SCom = new OleDbCommand();

            SCom.CommandText = sqlSTR;
            SCom.Parameters.AddWithValue("@year", tempYear);
            SCom.Parameters.AddWithValue("@month", tempMonth);

            SCom.Connection = cn;
            dR = SCom.ExecuteReader();

            if (dR.HasRows == true)
            {
                treeView1.Nodes.Add("配布指示書");
                treeView1.Nodes[0].ImageIndex = 4;  //アイコンの設定

                wID = 0;
                nIndex = 0;

                //プログレスバー表示
                progressBar1.Visible = true;

                //データ読み込み
                while (dR.Read())
                {
                    //プログレスバー進行状況表示
                    PrgVal++;
                    progressBar1.Value = PrgVal;

                    if (wID != long.Parse(dR["配布指示ID"].ToString()))
                    {
                        //配布指示書毎のエリア数
                        if (wID != 0)
                        {
                            treeView1.Nodes[0].Nodes[nIndex - 1].Text += "(" + nCnt.ToString() + ")";
                        }

                        //配布指示書情報
                        if (dR["配布指示配布日"] == DBNull.Value) 
                        {
                            nDate = "";
                        }
                        else
                        {
                            nDate = DateTime.Parse(dR["配布指示配布日"].ToString()).ToShortDateString();
                        }
 
                        NodeName1 = int.Parse(dR["配布指示ID"].ToString()).ToString("d6") + " " + nDate + " " + dR["氏名"].ToString() + "";
                        treeView1.Nodes[0].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex].NodeFont = new Font("ＭＳ Ｐゴシック", 10, FontStyle.Regular);
                        treeView1.Nodes[0].Nodes[nIndex].ImageIndex = 1;    //アイコンの設定
                        nIndex++;
                        nCnt = 0;
                    }

                    //配布エリア情報
                    if (dR["チラシ名"] != DBNull.Value)
                    {
                        if (dR["完了区分"].ToString() == "0")
                        {
                            nKanryo = "";
                        }
                        else
                        {
                            nKanryo = "完了";
                        }

                        NodeName1 = dR["チラシ名"].ToString() + " " + int.Parse(dR["町名コード"].ToString()).ToString("d4") + " " + dR["町名名称"].ToString() + " " + double.Parse(dR["配布単価"].ToString(), System.Globalization.NumberStyles.Any).ToString("##0.0") + " " + dR["予定枚数"].ToString() + " " + nKanryo; 
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes[nCnt].ImageIndex = 2;    //アイコンの設定

                        //未完了は赤表示
                        if (nKanryo == "")
                        {
                            treeView1.Nodes[0].Nodes[nIndex - 1].Nodes[nCnt].ForeColor = Color.Red;
                            treeView1.Nodes[0].Nodes[nIndex - 1].ForeColor = Color.Red;
                        }
                        
                        nCnt++;
                    }

                    wID = long.Parse(dR["配布指示ID"].ToString());

                }

                treeView1.Nodes[0].Nodes[nIndex - 1].Text += "(" + nCnt.ToString() + ")";
            }
            else
            {
                MessageBox.Show("該当する配布指示書がありません", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
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
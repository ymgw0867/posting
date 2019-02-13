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
        const string MESSAGE_CAPTION = "�|�X�e�B���O�c���[";

        public frmPostingTree()
        {
            InitializeComponent();
        }

        private void frmPostingTree_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
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

            //�c���[�r���[�N���A
            treeView1.Nodes.Clear();

            OleDbDataReader dR;
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();

            cn = Con.GetConnection();

            sqlSTR = "";
            sqlSTR += "select ��.ID, ��.�`���V��, ��.����, �z�z�G���A.����ID, ����.���� AS ����,";
            sqlSTR += "�z�z�w��.�z�z��, �z�z��.����,�z�z�G���A.�\�薇��, �z�z�G���A.�񍐖���, ";
            sqlSTR += "�z�z�G���A.�񍐎c��, �z�z�G���A.�����敪 ";
            sqlSTR += "from �� left join �z�z�G���A ";
            sqlSTR += "on ��.ID = �z�z�G���A.��ID left join ���� ";
            sqlSTR += "on �z�z�G���A.����ID = ����.ID left join �z�z�w�� ";
            sqlSTR += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID left join �z�z�� ";
            sqlSTR += "on �z�z�w��.�z�z��ID = �z�z��.ID ";
            sqlSTR += "where (��.�󒍎��ID = 1) and ";
            sqlSTR += "(year(�󒍓�) = ?) AND (month(�󒍓�) = ?) ";
            sqlSTR += "order by ��.ID desc,����ID";

            OleDbCommand SCom = new OleDbCommand();

            SCom.CommandText = sqlSTR;
            SCom.Parameters.AddWithValue("@year", tempYear);
            SCom.Parameters.AddWithValue("@month", tempMonth);

            SCom.Connection = cn;
            dR = SCom.ExecuteReader();

            if (dR.HasRows == true)
            {
                treeView1.Nodes.Add("�󒍊m�菑");
                treeView1.Nodes[0].ImageIndex = 4;  //�A�C�R���̐ݒ�

                wID = 0;
                nIndex = 0;

                while (dR.Read())
                {
                    if (wID != long.Parse(dR["ID"].ToString()))
                    {
                        //�󒍊m�菑���̃G���A��
                        if (wID != 0)
                        {
                            treeView1.Nodes[0].Nodes[nIndex - 1].Text += "(" + nCnt.ToString() + ")";
                        }

                        //�󒍊m�菑���
                        NodeName1 = dR["ID"].ToString() + " " + dR["�`���V��"].ToString() + " " + dR["����"].ToString();
                        treeView1.Nodes[0].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex].NodeFont = new Font("�l�r �o�S�V�b�N", 10, FontStyle.Regular);
                        treeView1.Nodes[0].Nodes[nIndex].ImageIndex = 2;    //�A�C�R���̐ݒ�
                        nIndex++;
                        nCnt = 0;
                    }

                    //�z�z�G���A���
                    if (dR["����"] != DBNull.Value)
                    {

                        if (dR["�z�z��"] == DBNull.Value)
                        {
                            nDate = "----/--/--";
                        }
                        else
                        {
                            nDate = DateTime.Parse(dR["�z�z��"].ToString()).ToShortDateString();
                        }

                        if (dR["����"] == DBNull.Value)
                        {
                            nName = "********";
                        }
                        else
                        {
                            nName = dR["����"].ToString();
                        }

                        if (dR["�����敪"].ToString() == "0")
                        {
                            nKanryo = "";
                        }
                        else
                        {
                            nKanryo = "����";
                        }

                        NodeName1 = int.Parse(dR["����ID"].ToString()).ToString("d4") + " " + dR["����"].ToString() + " " + nDate + " " + nName + " " + dR["�\�薇��"].ToString() + " " + nKanryo;
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes[nCnt].ImageIndex = 1;    //�A�C�R���̐ݒ�

                        //�������͐ԕ\��
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
                MessageBox.Show("�Y������󒍊m�菑������܂���",MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Information);
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
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtYear.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtMonth.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("1�`12�œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
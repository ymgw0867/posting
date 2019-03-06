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
        const string MESSAGE_CAPTION = "�z�z�w���c���[";

        public frmHaifuTree()
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

            //�c���[�r���[�N���A
            treeView1.Nodes.Clear();

            OleDbDataReader dR;
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();

            cn = Con.GetConnection();

            //�f�[�^�����擾
            sqlSTR = "";
            sqlSTR += "select count(*) as Cnt ";
            sqlSTR += "from �z�z�w�� LEFT OUTER JOIN �z�z�G���A ";
            sqlSTR += "ON �z�z�w��.ID = �z�z�G���A.�z�z�w��ID ";
            sqlSTR += "where (year(�z�z�w��.�z�z��) = " + tempYear.ToString() + ") and (month(�z�z�w��.�z�z��) = " + tempMonth.ToString() + ")";

            OleDbCommand SCom = new OleDbCommand();

            SCom.CommandText = sqlSTR;
            SCom.Connection = cn;
            dR = SCom.ExecuteReader();
            dR.Read();
            PrgMax = int.Parse(dR["Cnt"].ToString());
            dR.Close();

            //�v���O���X�o�[MAX����,MIN�����ݒ�
            progressBar1.Maximum = PrgMax;
            progressBar1.Minimum = 0;

            //�f�[�^�擾
            sqlSTR = "";
            sqlSTR += "select �z�z�w��.ID as �z�z�w��ID, �z�z�w��.�z�z�� as �z�z�w���z�z��, ";
            sqlSTR += "�z�z��.����, ��.�`���V��, ����.ID AS �����R�[�h,";
            sqlSTR += "����.���� as ��������, �z�z�G���A.�\�薇��, �z�z�G���A.�z�z�P��, ";
            sqlSTR += "�z�z�G���A.�񍐖���, �z�z�G���A.�񍐎c��,�z�z�G���A.�����敪 ";
            sqlSTR += "from �z�z�w�� LEFT OUTER JOIN �z�z�G���A ";
            sqlSTR += "ON �z�z�w��.ID = �z�z�G���A.�z�z�w��ID LEFT OUTER JOIN ";
            sqlSTR += "�� ON �z�z�G���A.��ID = ��.ID LEFT OUTER JOIN ";
            sqlSTR += "�z�z�� ON �z�z�w��.�z�z��ID = �z�z��.ID LEFT OUTER JOIN ";
            sqlSTR += "���� ON �z�z�G���A.����ID = ����.ID ";
            sqlSTR += "where (year(�z�z�w��.�z�z��) = ?) and (month(�z�z�w��.�z�z��) = ?) ";
            sqlSTR += "ORDER BY �z�z�w��.ID DESC,�z�z�G���A.��ID";

            //OleDbCommand SCom = new OleDbCommand();

            SCom.CommandText = sqlSTR;
            SCom.Parameters.AddWithValue("@year", tempYear);
            SCom.Parameters.AddWithValue("@month", tempMonth);

            SCom.Connection = cn;
            dR = SCom.ExecuteReader();

            if (dR.HasRows == true)
            {
                treeView1.Nodes.Add("�z�z�w����");
                treeView1.Nodes[0].ImageIndex = 4;  //�A�C�R���̐ݒ�

                wID = 0;
                nIndex = 0;

                //�v���O���X�o�[�\��
                progressBar1.Visible = true;

                //�f�[�^�ǂݍ���
                while (dR.Read())
                {
                    //�v���O���X�o�[�i�s�󋵕\��
                    PrgVal++;
                    progressBar1.Value = PrgVal;

                    if (wID != long.Parse(dR["�z�z�w��ID"].ToString()))
                    {
                        //�z�z�w�������̃G���A��
                        if (wID != 0)
                        {
                            treeView1.Nodes[0].Nodes[nIndex - 1].Text += "(" + nCnt.ToString() + ")";
                        }

                        //�z�z�w�������
                        if (dR["�z�z�w���z�z��"] == DBNull.Value) 
                        {
                            nDate = "";
                        }
                        else
                        {
                            nDate = DateTime.Parse(dR["�z�z�w���z�z��"].ToString()).ToShortDateString();
                        }
 
                        NodeName1 = int.Parse(dR["�z�z�w��ID"].ToString()).ToString("d6") + " " + nDate + " " + dR["����"].ToString() + "";
                        treeView1.Nodes[0].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex].NodeFont = new Font("�l�r �o�S�V�b�N", 10, FontStyle.Regular);
                        treeView1.Nodes[0].Nodes[nIndex].ImageIndex = 1;    //�A�C�R���̐ݒ�
                        nIndex++;
                        nCnt = 0;
                    }

                    //�z�z�G���A���
                    if (dR["�`���V��"] != DBNull.Value)
                    {
                        if (dR["�����敪"].ToString() == "0")
                        {
                            nKanryo = "";
                        }
                        else
                        {
                            nKanryo = "����";
                        }

                        NodeName1 = dR["�`���V��"].ToString() + " " + int.Parse(dR["�����R�[�h"].ToString()).ToString("d4") + " " + dR["��������"].ToString() + " " + double.Parse(dR["�z�z�P��"].ToString(), System.Globalization.NumberStyles.Any).ToString("##0.0") + " " + dR["�\�薇��"].ToString() + " " + nKanryo; 
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes.Add(NodeName1);
                        treeView1.Nodes[0].Nodes[nIndex - 1].Nodes[nCnt].ImageIndex = 2;    //�A�C�R���̐ݒ�

                        //�������͐ԕ\��
                        if (nKanryo == "")
                        {
                            treeView1.Nodes[0].Nodes[nIndex - 1].Nodes[nCnt].ForeColor = Color.Red;
                            treeView1.Nodes[0].Nodes[nIndex - 1].ForeColor = Color.Red;
                        }
                        
                        nCnt++;
                    }

                    wID = long.Parse(dR["�z�z�w��ID"].ToString());

                }

                treeView1.Nodes[0].Nodes[nIndex - 1].Text += "(" + nCnt.ToString() + ")";
            }
            else
            {
                MessageBox.Show("�Y������z�z�w����������܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
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
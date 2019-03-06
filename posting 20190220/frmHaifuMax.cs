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
    public partial class frmHaifuMax : Form
    {
        const string MESSAGE_CAPTION = "�s�撬���ʌ��ԍő��z�z����";

        public frmHaifuMax()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            GridviewSet.Setting(dataGridView2);
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();
        }

        // �f�[�^�O���b�h�r���[�N���X
        private class GridviewSet
        {

            /// <summary>
            /// �f�[�^�O���b�h�r���[�̒�`���s���܂�
            /// </summary>
            /// <param name="tempDGV">�f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            public static void Setting(DataGridView tempDGV)
            {
                try
                {
                    //�t�H�[���T�C�Y��`

                    // ��X�^�C����ύX����

                    tempDGV.EnableHeadersVisualStyles = false;

                    // ��w�b�_�[�\���ʒu�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // ��w�b�_�[�t�H���g�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", 9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", (float)11, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 541;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�R�[�h");
                    tempDGV.Columns.Add("col2", "�s�撬��");
                    tempDGV.Columns.Add("col5", "�z�z����");

                    tempDGV.Columns[0].Width = 70;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 80;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

                    // �ҏW�s�Ƃ���
                    tempDGV.ReadOnly = false;

                    // �ǉ��s�\�����Ȃ�
                    tempDGV.AllowUserToAddRows = false;

                    // �f�[�^�O���b�h�r���[����s�폜���֎~����
                    tempDGV.AllowUserToDeleteRows = false;

                    // �蓮�ɂ���ړ��̋֎~
                    tempDGV.AllowUserToOrderColumns = false;

                    // ��T�C�Y�ύX�֎~
                    tempDGV.AllowUserToResizeColumns = false;

                    // �s�T�C�Y�ύX�֎~
                    tempDGV.AllowUserToResizeRows = false;

                    // �s�w�b�_�[�̎�������
                    //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void ShowData(DataGridView tempDGV,int tempYear,int tempMonth)
            {

                string sqlSTRING = "";
                int iX;

                try
                {
                    tempDGV.RowCount = 0;
                    
                    //�f�[�^���[�_�[���擾����        
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "SELECT m_tbl.ID,�s�撬��_1.�s���{��,m_tbl.�s�撬��,SUM(m_tbl.����) AS ���� ";
                    sqlSTRING += "from ";
                    sqlSTRING += "(SELECT �s�撬��.ID,�s�撬��.�s�撬��,�z�z�G���A.����ID,����.����,MAX(�z�z�G���A.�\�薇��) AS ���� ";
                    sqlSTRING += "FROM �z�z�G���A INNER JOIN ";
                    sqlSTRING += "�z�z�w�� ON �z�z�G���A.�z�z�w��ID = �z�z�w��.ID INNER JOIN ";
                    sqlSTRING += "���� ON �z�z�G���A.����ID = ����.ID INNER JOIN ";
                    sqlSTRING += "�s�撬�� ON ����.�s�撬���R�[�h = �s�撬��.ID ";
                    sqlSTRING += "WHERE (YEAR(�z�z�w��.�z�z��) = ?) AND (MONTH(�z�z�w��.�z�z��) = ?) ";
                    sqlSTRING += "GROUP BY �z�z�G���A.����ID,����.����,�s�撬��.ID,�s�撬��.�s�撬��) ";
                    sqlSTRING += "AS m_tbl LEFT OUTER JOIN ";
                    sqlSTRING += "�s�撬�� AS �s�撬��_1 ON m_tbl.ID = �s�撬��_1.ID ";
                    sqlSTRING += "GROUP BY m_tbl.ID,m_tbl.�s�撬��,�s�撬��_1.�s���{�� ";
                    sqlSTRING += "ORDER BY m_tbl.ID";

                    //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@sYear", tempYear);
                    SCom.Parameters.AddWithValue("@sMonth",tempMonth);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = int.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = dR["�s���{��"].ToString() + " " + dR["�s�撬��"].ToString() + "";
                            tempDGV[2, iX].Value = int.Parse(dR["����"].ToString());

                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();
                    cn.Close();
                    Con.Close();

                    //if (tempDGV.Rows.Count > 29)
                    //{
                    //    tempDGV.Columns[1].Width = 333;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 350;
                    //}

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�I�����܂�" + Environment.NewLine + "��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, MESSAGE_CAPTION);
        }

        private void txtYear_Validating(object sender, CancelEventArgs e)
        {
            if (Utility.NumericCheck(txtYear.Text) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        private void txtMonth_Validating(object sender, CancelEventArgs e)
        {
            if (Utility.NumericCheck(txtMonth.Text) == false)
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

        private void button2_Click(object sender, EventArgs e)
        {
            GridviewSet.ShowData(dataGridView2,int.Parse(txtYear.Text.ToString()),int.Parse(txtMonth.Text.ToString()));
            dataGridView2.CurrentCell = null; //�I����Ԃ��������
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
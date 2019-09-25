using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using MyLibrary;

namespace posting
{
    public partial class frmHaifuShijiRep : Form
    {
        const string MESSAGE_CAPTION = "�z�z�w���ꗗ";

        public frmHaifuShijiRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            viewSetting(dataGridView1);
            DispClear();
        }
        
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̒�`���s���܂� </summary>
        /// <param name="tempDGV">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        private void viewSetting(DataGridView tempDGV)
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
                tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", (float)9.5, FontStyle.Regular);
                    
                // �s�̍���
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // �S�̂̍���
                tempDGV.Height = 631;

                // ��s�̐F
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add("col1", "�w������");
                tempDGV.Columns.Add("col2", "�z�z��");
                tempDGV.Columns.Add("col3", "�z�z��ID");
                tempDGV.Columns.Add("col4", "�z�z����");
                tempDGV.Columns.Add("col5", "�󒍔ԍ�");
                tempDGV.Columns.Add("col6", "�`���V��");
                tempDGV.Columns.Add("col7", "�R�[�h");
                tempDGV.Columns.Add("col8", "�G���A");
                tempDGV.Columns.Add("col9", "�z�z�P��");
                tempDGV.Columns.Add("col10", "�\�薇��");
                tempDGV.Columns.Add("col11", "�z�z����");
                tempDGV.Columns.Add("col12", "��ʔ�");

                tempDGV.Columns[0].Width = 80;
                tempDGV.Columns[1].Width = 80;
                tempDGV.Columns[2].Width = 80;
                tempDGV.Columns[3].Width = 140;
                tempDGV.Columns[4].Width = 100;
                tempDGV.Columns[5].Width = 300;
                tempDGV.Columns[6].Width = 70;
                tempDGV.Columns[7].Width = 260;
                tempDGV.Columns[8].Width = 80;
                tempDGV.Columns[9].Width = 80;
                tempDGV.Columns[10].Width = 80;
                tempDGV.Columns[11].Width = 80;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                tempDGV.Columns[1].DefaultCellStyle.Format = "yyyy/MM/dd";
                tempDGV.Columns[6].DefaultCellStyle.Format = "0000";
                tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";
                tempDGV.Columns[11].DefaultCellStyle.Format = "#,##0";

                // �s�w�b�_��\�����Ȃ�
                tempDGV.RowHeadersVisible = false;

                // �I�����[�h
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // �ҏW�s�Ƃ���
                tempDGV.ReadOnly = true;

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

        private void showData(DataGridView tempDGV)
        {
            string sqlSTRING = "";

            try
            {
                //�f�[�^���[�_�[���擾����
                OleDbDataReader dR;

                sqlSTRING += "SELECT �z�z�w��.ID AS �z�z�w��ID, �z�z�w��.�z�z�� AS �z�z�w���z�z��,";
                sqlSTRING += "�z�z�w��.�z�z��ID,�z�z��.����, �z�z�G���A.��ID,";
                sqlSTRING += "��.�`���V��, �z�z�G���A.����ID AS �����R�[�h, ����.���� AS ��������,";
                sqlSTRING += "�z�z�G���A.�z�z�P��, �z�z�G���A.�\�薇��,"; 
                sqlSTRING += "�z�z�G���A.���z�z����, �z�z�w��.��ʔ� ";
                sqlSTRING += "FROM �z�z�w�� inner JOIN �z�z�G���A ";
                sqlSTRING += "ON �z�z�w��.ID = �z�z�G���A.�z�z�w��ID inner JOIN �� ";
                sqlSTRING += "ON �z�z�G���A.��ID = ��.ID LEFT OUTER JOIN �z�z�� ";
                sqlSTRING += "ON �z�z�w��.�z�z��ID = �z�z��.ID LEFT OUTER JOIN ���� ";
                sqlSTRING += "ON �z�z�G���A.����ID = ����.ID ";

                sqlSTRING += "where (�z�z�w��.ID >= ? and �z�z�w��.ID <= ?) and ";
                sqlSTRING += "(�z�z�w��.�z�z�� >= ? and �z�z�w��.�z�z�� <= ?) ";

                // ���͓��ݒ� 2016/09/19
                if (sInputDt.Checked)
                {
                    sqlSTRING += "and �z�z�w��.���͓� = ? ";
                }

                //// �����ݒ�
                //if (txtsChome.Text.Trim() != string.Empty)
                //{
                //    sqlSTRING += "and ����.���� like '*?*' ";
                //}

                sqlSTRING += "ORDER BY �z�z�w��ID DESC, �z�z�G���A.��ID ";

                Control.DataControl Con = new Control.DataControl();
                OleDbConnection Cn = new OleDbConnection();
                Cn = Con.GetConnection();

                OleDbCommand sCom = new OleDbCommand();

                sCom.CommandText = sqlSTRING;

                // �w�����͈͐ݒ� 2016/09/19
                sCom.Parameters.AddWithValue("@sNoS", Utility.strToInt(txtsShijiS.Text));

                if (Utility.strToInt(txtsShijiE.Text) == 0)
                {
                    sCom.Parameters.AddWithValue("@sNoE", 2000000000);
                }
                else
                {
                    sCom.Parameters.AddWithValue("@sNoE", Utility.strToInt(txtsShijiE.Text));
                }

                // �z�z�J�n���ݒ� 2016/09/19
                if (sHaifuDtS.Checked)
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtS", sHaifuDtS.Value.ToShortDateString());
                }
                else
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtS", "1900/01/01");
                }

                // �z�z�I�����ݒ� 2016/09/19
                if (sHaifuDtE.Checked)
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtE", sHaifuDtE.Value.ToShortDateString());
                }
                else
                {
                    sCom.Parameters.AddWithValue("@sHaifuDtE", "2900/01/01");
                }

                // ���͓��ݒ� 2016/09/19
                if (sInputDt.Checked)
                {
                    sCom.Parameters.AddWithValue("@sInputDt", sInputDt.Value.ToShortDateString());
                }
                
                sCom.Connection = Cn;
                dR = sCom.ExecuteReader();

                //�O���b�h�r���[�ɕ\������
                int iX = 0;

                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    // ���ڌ����ݒ�̂Ƃ� 2016/09/19
                    if (txtsChome.Text.Trim() != string.Empty)
                    {
                        if (!dR["��������"].ToString().Contains(txtsChome.Text.Trim()))
                        {
                            continue;
                        }
                    }

                    // �O���b�h�r���[�\��
                    tempDGV.Rows.Add();

                    tempDGV[0, iX].Value = int.Parse(dR["�z�z�w��ID"].ToString());
                    tempDGV[1, iX].Value = DateTime.Parse(dR["�z�z�w���z�z��"].ToString());
                    tempDGV[2, iX].Value = dR["�z�z��ID"].ToString();
                    tempDGV[3, iX].Value = dR["����"].ToString();
                    tempDGV[4, iX].Value = long.Parse(dR["��ID"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[5, iX].Value = dR["�`���V��"].ToString();
                    tempDGV[6, iX].Value = int.Parse(dR["�����R�[�h"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[7, iX].Value = dR["��������"].ToString();
                    tempDGV[8, iX].Value = double.Parse(dR["�z�z�P��"].ToString(), System.Globalization.NumberStyles.Any).ToString("#,##0.00");
                    tempDGV[9, iX].Value = int.Parse(dR["�\�薇��"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[10, iX].Value = int.Parse(dR["���z�z����"].ToString(), System.Globalization.NumberStyles.Any);
                    tempDGV[11, iX].Value = int.Parse(dR["��ʔ�"].ToString(), System.Globalization.NumberStyles.Any);
                        
                    iX++;
                }

                dR.Close();
                Con.Close();
                    
                //if (tempDGV.RowCount > 25)
                //{
                //    tempDGV.Columns[1].Width = 198;
                //}
                //else
                //{
                //    tempDGV.Columns[1].Width = 215;
                //}

                tempDGV.CurrentCell = null;

                // �\�������\�� 2016/09/19
                if (tempDGV.Rows.Count > 0)
                {
                    this.Text = MESSAGE_CAPTION + " �y" + tempDGV.Rows.Count.ToString("#,##0") + "���z";
                }
                else
                {
                    this.Text = MESSAGE_CAPTION;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
            }
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�z�z�w���ꗗ��\�����܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            Cursor.Current = Cursors.WaitCursor;
            showData(dataGridView1);
            dataGridView1.CurrentCell = null;
            Cursor.Current = Cursors.Default;

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("�Y���f�[�^������܂���ł���", MESSAGE_CAPTION);
                btnPrn.Enabled = false;
            }
            else
            {
                btnPrn.Enabled = true;
            }
        }

        private void DispClear()
        {
            // 2016/09/19
            txtsShijiS.Text = string.Empty;
            txtsShijiE.Text = string.Empty;
            sHaifuDtS.Checked = false;
            sHaifuDtE.Checked = false;
            sInputDt.Checked = false;
            txtsChome.Text = string.Empty;
        }

        private void txtsShijiS_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
    }
}
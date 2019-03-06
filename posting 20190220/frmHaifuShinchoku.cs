using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmHaifuShinchoku : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "�z�z�i����";

        public frmHaifuShinchoku()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShinchoku_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView2);
            GridviewSet.Setting2(dataGridView1);
            button1.Enabled = false;

            comboBox1.Items.Add("�S��");
            comboBox1.Items.Add("������");
            comboBox1.Items.Add("����");
            comboBox1.SelectedIndex = 0;
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
                    tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", 9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 235;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col2", "���Ӑ�");
                    tempDGV.Columns.Add("col3", "�`���V��");
                    tempDGV.Columns.Add("col4", "�S����");
                    tempDGV.Columns.Add("col5", "�ō�����");
                    tempDGV.Columns.Add("col6", "�\�薇��");
                    tempDGV.Columns.Add("col7", "�z�z����");
                    tempDGV.Columns.Add("col8", "�c����");
                    tempDGV.Columns.Add("col9", "������");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 78;
                    tempDGV.Columns[5].Width = 78;
                    tempDGV.Columns[6].Width = 78;
                    tempDGV.Columns[7].Width = 75;
                    tempDGV.Columns[8].Width = 80;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

            public static void Setting2(DataGridView tempDGV)
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
                    tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", 9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 253;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col2", "�w���ԍ�");
                    tempDGV.Columns.Add("col3", "�z�z��");
                    tempDGV.Columns.Add("col4", "ID");
                    tempDGV.Columns.Add("col5", "�z�z��");
                    tempDGV.Columns.Add("col6", "ID");
                    tempDGV.Columns.Add("col7", "����");
                    tempDGV.Columns.Add("col8", "�\�薇��");
                    tempDGV.Columns.Add("col9", "�c����");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 90;
                    tempDGV.Columns[3].Width = 60;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 60;
                    tempDGV.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[7].Width = 78;
                    tempDGV.Columns[8].Width = 78;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV,int tempSel,string tempCName)
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
                    sqlSTRING += "select ��.ID,�Ј�.����,���Ӑ�.����,��.�`���V��,��.�ō����z,";
                    sqlSTRING += "��.���� as �[�i��,h_tbl.�z�z����,";
                    sqlSTRING += "��.���� - h_tbl.�z�z���� as �c���� ";
                    sqlSTRING += "from �� left join  ";

                    sqlSTRING += "(select ��ID,sum(�\�薇��) as �z�z���� ";
                    sqlSTRING += "from �z�z�G���A ";
                    sqlSTRING += "where �����敪 = 1 ";
                    sqlSTRING += "group by ��ID ) as h_tbl ";

                    sqlSTRING += "on ��.ID = h_tbl.��ID left join ���Ӑ� ";
                    sqlSTRING += "on ��.���Ӑ�ID = ���Ӑ�.ID left join �Ј� ";
                    sqlSTRING += "on ���Ӑ�.�S���Ј��R�[�h = �Ј�.ID ";
                    sqlSTRING += "where (��.�󒍎��ID = 1) ";

                    switch (tempSel)
                    {
                        case 1:
                            sqlSTRING += "and (((��.���� - h_tbl.�z�z����) > 0) or h_tbl.�z�z���� is null) ";
                            break;

                        case 2:
                            sqlSTRING += "and (��.���� - h_tbl.�z�z���� = 0) ";
                            break;
                    }

                    if (tempCName.Trim().Length > 0)
                    {
                        sqlSTRING += "and (��.�`���V�� like ?)";
                    }

                    sqlSTRING += "order by ��.ID desc";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    if (tempCName.Trim().Length > 0)
                    {
                        SCom.Parameters.AddWithValue("@CName", "%" + tempCName + "%");
                    }

                    SCom.Connection = cn;
                
                    //�z�z�i���󋵂̃f�[�^���[�_�[���擾����
                    //Control.FreeSql fCon = new Control.FreeSql();
                    //dR = fCon.free_dsReader(sqlSTRING);

                    dR = SCom.ExecuteReader();
                    
                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = long.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = dR["����"].ToString() + "";
                            tempDGV[2, iX].Value = dR["�`���V��"].ToString();
                            tempDGV[3, iX].Value = dR["����"].ToString() + "";
                            tempDGV[4, iX].Value = int.Parse(dR["�ō����z"].ToString(),System.Globalization.NumberStyles.Any);
                            tempDGV[5, iX].Value = int.Parse(dR["�[�i��"].ToString(), System.Globalization.NumberStyles.Any);

                            if (dR["�z�z����"] == DBNull.Value)
                            {
                                tempDGV[6, iX].Value = 0;
                            }
                            else
                            {
                                tempDGV[6, iX].Value = int.Parse(dR["�z�z����"].ToString(), System.Globalization.NumberStyles.Any);
                            }

                            if (dR["�c����"] == DBNull.Value)
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["�[�i��"].ToString(), System.Globalization.NumberStyles.Any);
                            }
                            else
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["�c����"].ToString(), System.Globalization.NumberStyles.Any);
                            }

                            //�z�z������
                            if (int.Parse(tempDGV[7, iX].Value.ToString(), System.Globalization.NumberStyles.Any) == 0)
                            {
                                string sqlSTR;
                                OleDbDataReader r;
                                Control.FreeSql rCon = new Control.FreeSql();
                                sqlSTR = "";
                                sqlSTR += "select max(�z�z�w��.�z�z��) as ������ ";
                                sqlSTR += "from �z�z�G���A inner join �z�z�w�� ";
                                sqlSTR += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                                sqlSTR += "where ";
                                sqlSTR += "(�z�z�G���A.��ID = " + long.Parse(dR["ID"].ToString()) + ") and ";
                                sqlSTR += "(�z�z�G���A.�����敪 = 1)";

                                r = rCon.free_dsReader(sqlSTR);

                                while (r.Read())
                                {
                                    if (r["������"] != DBNull.Value)
                                    {
                                        tempDGV[8, iX].Value = DateTime.Parse(r["������"].ToString()).ToShortDateString();
                                    }
                                    else
                                    {
                                        tempDGV[8, iX].Value = "";
                                    }
                                }

                                r.Close();
                                rCon.Close();
                            }
                            else
                            {
                                tempDGV[8, iX].Value = "";
                            }


                            iX++;

                            //frmP.valueCount = iX;
                            //frmP.ShowProgress();
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();
                    Con.Close();
                    cn.Close();

                    //fCon.Close();

                    //frmP.Close();

                    //frmP.Dispose();

                    //if (tempDGV.RowCount <= 12)
                    //{
                    //    tempDGV.Columns[3].Width = 97;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[3].Width = 80;
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
            if (MessageBox.Show("�I�����܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("�Ώۃ`���V��I�����Ă�������");
                comboBox1.Focus();
                return;
            }

            try
            {

                Cursor.Current = Cursors.WaitCursor;    //�J�[�\����ҋ@�\��

                GridviewSet.ShowData(dataGridView2, comboBox1.SelectedIndex,txtCName.Text);
                dataGridView2.CurrentCell = null;
                dataGridView1.RowCount = 0;

                button1.Enabled = false;
                label1.Text = "�y�|�X�e�B���O�ڍׁz";
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ShowPosting(dataGridView1,long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
        }

        private void ShowPosting(DataGridView tempDGV,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select ��.ID, �z�z�w��.ID as �w���ԍ�,�z�z�w��.�z�z��, �z�z��.ID AS �z�z��ID, �z�z��.���� AS �z�z������,";
            mySql += "����.ID AS ����ID, ����.���� AS ����,";
            mySql += "�z�z�G���A.�\�薇��,�z�z�G���A.�����敪 ";
            mySql += "from �� inner join �z�z�G���A ";
            mySql += "on ��.ID = �z�z�G���A.��ID left join �z�z�w�� ";
            mySql += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID left join ���� ";
            mySql += "on �z�z�G���A.����ID = ����.ID left join �z�z�� ";
            mySql += "on �z�z�w��.�z�z��ID = �z�z��.ID ";
            mySql += "where ��.ID = " + tempID.ToString() + " ";
            mySql += "order by �z�z�w��.�z�z�� desc";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;
            button1.Enabled = false;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["ID"].ToString();

                if (dR["�w���ԍ�"] == DBNull.Value)
                {
                    tempDGV[1, iX].Value = "** ���ݒ� **";
                }
                else
                {
                    tempDGV[1, iX].Value = dR["�w���ԍ�"].ToString();
                }

                if (dR["�z�z��"] == DBNull.Value)
                {
                    tempDGV[2, iX].Value = "";
                }
                else
                {
                    tempDGV[2, iX].Value =  DateTime.Parse(dR["�z�z��"].ToString()).ToShortDateString();
                }

                if (dR["�z�z��ID"] == DBNull.Value)
                {
                    tempDGV[3, iX].Value = "";
                }
                else
                {
                    tempDGV[3, iX].Value = dR["�z�z��ID"].ToString();
                }

                if (dR["�z�z������"] == DBNull.Value)
                {
                    tempDGV[4, iX].Value = "";
                }
                else
                {
                    tempDGV[4, iX].Value = dR["�z�z������"].ToString();
                }

                if (dR["����ID"] == DBNull.Value)
                {
                    tempDGV[5, iX].Value = "";
                }
                else
                {
                    tempDGV[5, iX].Value = dR["����ID"].ToString();
                }

                if (dR["����"] == DBNull.Value)
                {
                    tempDGV[6, iX].Value = "";
                }
                else
                {
                    tempDGV[6, iX].Value = dR["����"].ToString();
                }

                if (dR["�\�薇��"] == DBNull.Value)
                {
                    tempDGV[7, iX].Value = 0;
                }
                else
                {
                    tempDGV[7, iX].Value = int.Parse(dR["�\�薇��"].ToString());
                }

                //�c����
                if (dR["�\�薇��"] == DBNull.Value)
                {
                    tempDGV[8, iX].Value = 0;
                }
                else
                {
                    if (dR["�����敪"].ToString() == "1")
                    {
                        tempDGV[8, iX].Value = 0;
                    }
                    else
                    {
                        tempDGV[8, iX].Value = int.Parse(dR["�\�薇��"].ToString());
                    }
                }

                iX++;

                button1.Enabled = true;
            }

            //if (tempDGV.RowCount <= 13)
            //{
            //    tempDGV.Columns[6].Width = 330;
            //}
            //else
            //{
            //    tempDGV.Columns[6].Width = 313;
            //}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();

            label1.Text = "�y" + dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString() +" �z�z�󋵁z";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�z�z�i���󋵂𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 36; //�z�z�����񍐏��̖��׍s��

            int pCnt;

            //�y�[�W�J�E���g
            pCnt = dataGridView1.Rows.Count / G_COUNT + 1;

            for (int i = 1; i <= pCnt; i++)
            {
                KanryoReport(pCnt, i, G_COUNT);
            }

        }

        private void KanryoReport(int tempPage, int tempCurrentPage, int tempFixRows)
        {

            const int S_GYO = 8;    //�G�N�Z���t�@�C�����ׂ�8�s�ڂ����
            int dgvIndex;
            int i;
            int yMai, zMai;

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���z�z�i����, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    //���t�E�y�[�W��
                    oxlsSheet.Cells[1, 44] = "DATE : " + DateTime.Today.ToShortDateString() + "  P." + tempCurrentPage.ToString() + "/" + tempPage.ToString();

                    //�󒍔ԍ�
                    oxlsSheet.Cells[3, 5] = long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);
                    
                    //���Ӑ於
                    oxlsSheet.Cells[4, 5] = dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString();
                    
                    //�`���V��
                    oxlsSheet.Cells[5, 5] = dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //�S���Җ�
                    oxlsSheet.Cells[3, 18] = dataGridView2[3, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //�ō�����
                    oxlsSheet.Cells[4, 29] = int.Parse(dataGridView2[4, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                    //�\�薇��
                    oxlsSheet.Cells[4, 34] = int.Parse(dataGridView2[5, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                    
                    //�z�z����
                    oxlsSheet.Cells[4, 39] = int.Parse(dataGridView2[6, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                    //�c����
                    oxlsSheet.Cells[4, 44] = int.Parse(dataGridView2[7, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                    //������
                    oxlsSheet.Cells[4, 49] = dataGridView2[8, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //�z�z�G���A����
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //�f�[�^�O���b�h�r���[�̍s�C���f�b�N�X�����߂�

                        //�w���ԍ�
                        oxlsSheet.Cells[i + S_GYO, 1] = dataGridView1[1, dgvIndex].Value.ToString();
                        
                        //�z�z��
                        oxlsSheet.Cells[i + S_GYO, 5] = dataGridView1[2, dgvIndex].Value.ToString();
                        
                        //�z�z��ID
                        oxlsSheet.Cells[i + S_GYO, 10] = dataGridView1[3, dgvIndex].Value.ToString();
                        
                        //�z�z����
                        oxlsSheet.Cells[i + S_GYO, 13] = dataGridView1[4, dgvIndex].Value.ToString();
                        
                        //����ID
                        oxlsSheet.Cells[i + S_GYO, 19] = dataGridView1[5, dgvIndex].Value.ToString();
                        
                        //����
                        oxlsSheet.Cells[i + S_GYO, 22] = dataGridView1[6, dgvIndex].Value.ToString();
                        
                        //�\�薇��
                        oxlsSheet.Cells[i + S_GYO, 39] = dataGridView1[7, dgvIndex].Value.ToString();
                        
                        //�z�z����
                        yMai = int.Parse(dataGridView1[7, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);
                        zMai = int.Parse(dataGridView1[8, dgvIndex].Value.ToString(), System.Globalization.NumberStyles.Any);
                        oxlsSheet.Cells[i + S_GYO, 44] = yMai - zMai;
                        
                        //�c����
                        oxlsSheet.Cells[i + S_GYO, 49] = dataGridView1[8, dgvIndex].Value.ToString();

                        //�O���b�h�ŏI�s�̂Ƃ��I��
                        if (dgvIndex == (dataGridView1.Rows.Count - 1)) break;

                        //������׍ő�s�̂Ƃ��I��
                        if (i == (tempFixRows - 1)) break;

                        i++;
                    }

                    //�}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;

                    // �m�F�̂���Excel�̃E�B���h�E��\������
                    oXls.Visible = true;

                    //���
                    oxlsSheet.PrintPreview(true);

                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    //�ۑ�����
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //�_�C�A���O�{�b�N�X�̏����ݒ�
                    saveFileDialog1.Title = MESSAGE_CAPTION + "�ۑ�";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //�����y�[�W�̂Ƃ��A�y�[�W�����t�^
                    if (tempPage > 1)
                    {
                        saveFileDialog1.FileName += "_" + tempCurrentPage.ToString();
                    }

                    saveFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xls)|*.xls|�S�Ẵt�@�C��(*.*)|*.*";

                    //�_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();
                }

                finally
                {
                    // COM �I�u�W�F�N�g�̎Q�ƃJ�E���g��������� 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //�}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

    }
}
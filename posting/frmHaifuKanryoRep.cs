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
    public partial class frmHaifuKanryoRep : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "�z�z�����񍐏�";
        const string FILE_TYPE_XLS = ".xls"; 

        public frmHaifuKanryoRep()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            GridviewSet.Setting(dataGridView2);
            GridviewSet.Setting2(dataGridView1);
            button1.Enabled = false;
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
                    tempDGV.Height = 505;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col2", "�`���V��");
                    tempDGV.Columns.Add("col3", "�[�i��");
                    tempDGV.Columns.Add("col4", "�O��܂�");
                    tempDGV.Columns.Add("col5", "�c����");
                    tempDGV.Columns.Add("col6", "�z�z����");
                    tempDGV.Columns.Add("col7", "�����c");

                    tempDGV.Columns[0].Width = 80;
                    //tempDGV.Columns[1].Width = 210;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;

                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

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
                    tempDGV.Height = 505;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col0", "�z�z��");
                    tempDGV.Columns.Add("col1", "�z�z�ꏊ");
                    tempDGV.Columns.Add("col2", "����");

                    tempDGV.Columns[0].Width = 80;
                    //tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 57;

                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    tempDGV.Columns[0].DefaultCellStyle.Format = "yyyy/MM/dd";
                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

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

            public static void ShowData(DataGridView tempDGV, DateTime tempDate, DateTime tempDateE, string chirashiName)
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

                    //sqlSTRING = "";
                    //sqlSTRING += "select ��.ID, ��.�`���V��, sum(�z�z�G���A.�\�薇��) as �z�z����,";
                    //sqlSTRING += "��.���� as �[�i�� ";
                    //sqlSTRING += "from �z�z�G���A inner join �� ";
                    //sqlSTRING += "on �z�z�G���A.��ID = ��.ID inner join �z�z�w�� ";
                    //sqlSTRING += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                    //sqlSTRING += "where (�z�z�w��.�z�z�� = ?) and (�z�z�G���A.�����敪 = 1) ";
                    //sqlSTRING += "group by ��.ID, ��.�`���V��, ��.����";
                       
                    sqlSTRING = "";
                    sqlSTRING += "select ��.ID,��.�`���V��,SUM(�z�z�G���A.�\�薇��) AS �z�z����, ";
                    sqlSTRING += "��.���� AS �[�i��, x.�O��z�z���� ";
                    sqlSTRING += "from �z�z�G���A inner join �� ";
                    sqlSTRING += "on �z�z�G���A.��ID = ��.ID inner join �z�z�w�� ";
                    sqlSTRING += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID left join ";

                    sqlSTRING += "(";
                    sqlSTRING += "select �z�z�G���A.��ID, sum(�z�z�G���A.�\�薇��) AS �O��z�z���� ";
                    sqlSTRING += "from �z�z�G���A inner join �z�z�w�� ";
                    sqlSTRING += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(�z�z�G���A.�����敪 = 1) and ";
                    sqlSTRING += "(�z�z�w��.�z�z�� < ?) ";
                    sqlSTRING += "group by �z�z�G���A.��ID";
                    sqlSTRING += ") AS x ";

                    sqlSTRING += "on ��.ID = x.��ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(�z�z�w��.�z�z�� >= ?) and ";
                    sqlSTRING += "(�z�z�w��.�z�z�� <= ?) and "; 
                    sqlSTRING += "(�z�z�G���A.�����敪 = 1) and ";
                    sqlSTRING += "(��.�`���V�� like ?) ";
                    sqlSTRING += "group by ��.ID, ��.�`���V��, ��.����, x.�O��z�z����";
                 
                    //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@sDate", tempDate);
                    SCom.Parameters.AddWithValue("@sDate2", tempDate);
                    SCom.Parameters.AddWithValue("@DateE", tempDateE);
                    SCom.Parameters.AddWithValue("@cName",  "" + "%" + "" + chirashiName + "" + "%" + "");

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();
                    
                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = long.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = dR["�`���V��"].ToString();
                            tempDGV[2, iX].Value = int.Parse(dR["�[�i��"].ToString());

                            //tempDGV[3, iX].Value = 0;
                            //tempDGV[4, iX].Value = int.Parse(dR["�[�i��"].ToString());
                            
                            tempDGV[5, iX].Value = int.Parse(dR["�z�z����"].ToString());
                            
                            //tempDGV[6, iX].Value = int.Parse(dR["�[�i��"].ToString()) - 0 - int.Parse(dR["�z�z����"].ToString());

                            if (dR["�O��z�z����"] == DBNull.Value)
                            {
                                tempDGV[3, iX].Value = 0;
                                tempDGV[4, iX].Value = int.Parse(dR["�[�i��"].ToString());
                                tempDGV[6, iX].Value = int.Parse(dR["�[�i��"].ToString()) - 0 - int.Parse(dR["�z�z����"].ToString());
                            }
                            else
                            {
                                tempDGV[3, iX].Value = int.Parse(dR["�O��z�z����"].ToString()); 
                                tempDGV[4, iX].Value = int.Parse(dR["�[�i��"].ToString()) - int.Parse(dR["�O��z�z����"].ToString());
                                tempDGV[6, iX].Value = int.Parse(dR["�[�i��"].ToString()) - int.Parse(dR["�O��z�z����"].ToString()) - int.Parse(dR["�z�z����"].ToString());
                            }

                            ////�O��܂ł̔z�z����
                            //string mySql = "";
                            //OleDbDataReader dR2;

                            //mySql += "select sum(�z�z�G���A.�\�薇��) as �O��z�z���� ";
                            //mySql += "from �z�z�G���A inner join �z�z�w�� ";
                            //mySql += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                            //mySql += "where ";
                            //mySql += "(�z�z�G���A.��ID = " + dR["ID"].ToString() + ") and ";
                            //mySql += "(�z�z�G���A.�����敪 = 1) and ";
                            //mySql += "(�z�z�w��.�z�z�� < '" + tempDate.ToShortDateString() + "') ";
                            //mySql += "group by �z�z�G���A.��ID";

                            //Control.FreeSql fCon = new Control.FreeSql();
                            //dR2 = fCon.free_dsReader(mySql);

                            //while (dR2.Read())
                            //{
                            //    tempDGV[3, iX].Value = int.Parse(dR2["�O��z�z����"].ToString());
                            //    tempDGV[4, iX].Value = int.Parse(dR["�[�i��"].ToString()) - int.Parse(dR2["�O��z�z����"].ToString());
                            //    tempDGV[6, iX].Value = int.Parse(dR["�[�i��"].ToString()) - int.Parse(dR2["�O��z�z����"].ToString()) - int.Parse(dR["�z�z����"].ToString());
                            //}

                            //dR2.Close();
                            //fCon.Close();


                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    //�Y���Ȃ��̂Ƃ�
                    if (tempDGV.RowCount == 0) MessageBox.Show("�Y�����Ԃɔz�z���т͂���܂���ł���","�Y���Ȃ�");

                    dR.Close();

                    cn.Close();

                    Con.Close();

                    //tempDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    //if (tempDGV.RowCount <= 27)
                    //{
                    //    tempDGV.Columns[1].Width = 217;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 200;
                    //}

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�\�����̃`���V�f�[�^�S�Ă�I����Ԃɂ��܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.Rows[i].Selected = true;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("���݂̑I����Ԃ��������܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                dataGridView2.Rows[r.Index].Selected = false;
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
            this.Cursor = Cursors.WaitCursor;   // 2019/02/13

            GridviewSet.ShowData(dataGridView2, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), Convert.ToDateTime(dateTimePicker2.Value.ToShortDateString()), txtsName.Text);
            dataGridView2.CurrentCell = null;
            dataGridView1.RowCount = 0;

            this.Cursor = Cursors.Default;  // 2019/02/13

            button1.Enabled = false;
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ShowPosting(dataGridView1, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), Convert.ToDateTime(dateTimePicker2.Value.ToShortDateString()), long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
            
            button1.Enabled = true;
        }

        private void ShowPosting(DataGridView tempDGV, DateTime tempDate, DateTime tempDateE, long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select �z�z�w��.�z�z��,����.����,�z�z�G���A.�}�ԋL��,�z�z�G���A.�\�薇�� ";
            mySql += "from (�z�z�G���A inner join �z�z�w�� ";
            mySql += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID) left join ���� ";
            mySql += "on �z�z�G���A.����ID = ����.ID ";
            mySql += "where ";
            mySql += "(�z�z�G���A.��ID = " + tempID.ToString() + ") and ";
            mySql += "(�z�z�G���A.�����敪 = 1) and ";
            mySql += "(�z�z�w��.�z�z�� >= '" + tempDate.ToShortDateString() + "') and ";
            mySql += "(�z�z�w��.�z�z�� <= '" + tempDateE.ToShortDateString() + "') ";
            mySql += "order by �z�z�w��.�z�z��,����ID";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["�z�z��"];
                tempDGV[1, iX].Value = dR["����"].ToString() + " " + dR["�}�ԋL��"].ToString() + "";
                tempDGV[2, iX].Value = int.Parse(dR["�\�薇��"].ToString());

                iX++;
            }

            //if (tempDGV.RowCount <= 27)
            //{
            //    tempDGV.Columns[1].Width = 200;
            //}
            //else
            //{
            //    tempDGV.Columns[1].Width = 183;
            //}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�z�z�����񍐏��𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 32; //�z�z�����񍐏��̖��׍s��

            //const int G_COUNT = 10; //�z�z�����񍐏��̖��׍s��
            //int pCnt;

            //�y�[�W�J�E���g
            //pCnt = dataGridView1.Rows.Count / G_COUNT + 1;

            //for (int i = 1; i <= pCnt; i++)
            //{
            //    KanryoReport(pCnt, i, G_COUNT);
            //}

            KanryoReport(G_COUNT);
        }

        private void KanryoReport(int tempFixRows)
        {

            const int S_GYO = 13;    //�G�N�Z���t�@�C�����ׂ�13�s�ڂ����
            int i;

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���z�z�����񍐏�, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range rng;

                try
                {
                    //���Ӑ���
                    long sID;
                    string sqlSTR;

                    sID = long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString());

                    sqlSTR = "";
                    sqlSTR += "select ���Ӑ�.����,���Ӑ�.�S���Җ�,���Ӑ�.FAX�ԍ� ";
                    sqlSTR += "from �� inner join ���Ӑ� ";
                    sqlSTR += "on ��.���Ӑ�ID = ���Ӑ�.ID ";
                    sqlSTR += "where ��.ID = " + sID.ToString();

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTR);

                    while (dR.Read())
                    {
                        oxlsSheet.Cells[1, 3] = dR["����"].ToString() + " " + dR["�S���Җ�"].ToString() + "�l";
                        oxlsSheet.Cells[2, 3] = dR["FAX�ԍ�"].ToString();                        
                    }

                    dR.Close();
                    fCon.Close();

                    //�[�i��
                    oxlsSheet.Cells[8, 2] = int.Parse(dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                    
                    //�O��܂�
                    oxlsSheet.Cells[9, 2] = int.Parse(dataGridView2[3, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    //�c����
                    oxlsSheet.Cells[10, 2] = int.Parse(dataGridView2[4, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    //�`���V��
                    oxlsSheet.Cells[10, 3] = dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString();

                    //�z�z�G���A���׏������F2019/03/18
                    for (int iX = S_GYO; iX < (tempFixRows + S_GYO); iX++)
                    {
                        oxlsSheet.Cells[iX, 2] = string.Empty;
                        oxlsSheet.Cells[iX, 3] = string.Empty;
                        oxlsSheet.Cells[iX, 4] = string.Empty;
                    }
                    
                    //�z�z�G���A����
                    i = 0;
                    while (true)
                    {

                        //������׍ő�s�����̂Ƃ��s�}��
                        if (i >= (tempFixRows - 1))
                        {
                            rng = (Excel.Range)oxlsSheet.Cells[i + S_GYO,1];
                            rng.EntireRow.Insert(Type.Missing, Microsoft.Office.Interop.Excel.XlDirection.xlDown);
                        }

                        oxlsSheet.Cells[i + S_GYO, 2] = dataGridView1[0, i].Value.ToString();
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView1[1, i].Value.ToString();
                        oxlsSheet.Cells[i + S_GYO, 4] = int.Parse(dataGridView1[2, i].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //�O���b�h�ŏI�s�̂Ƃ��I��
                        if (i == (dataGridView1.Rows.Count - 1)) break;

                        i++;
                    }

                    ////�z�z�������v
                    //oxlsSheet.Cells[45, 4] = int.Parse(dataGridView2[5, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    ////�c����
                    //oxlsSheet.Cells[46, 4] = int.Parse(dataGridView2[6, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

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
                    //saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString() + "_" + DateTime.Today.ToLongDateString();
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString() + "_" + this.dateTimePicker1.Text + "-" + this.dateTimePicker2.Text + FILE_TYPE_XLS;


                    ////�����y�[�W�̂Ƃ��A�y�[�W�����t�^
                    //if (tempPage > 1)
                    //{
                    //    saveFileDialog1.FileName += "_" + tempCurrentPage.ToString();
                    //}

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

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


    }
}
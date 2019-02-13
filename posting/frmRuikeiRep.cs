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
    public partial class frmRuikeiRep : Form
    {
        const string MESSAGE_CAPTION = "�݌v�\";
        const int SHINJUKU = 2;
        const int JYOUTOU = 3;

        public frmRuikeiRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            GridviewSet.Setting(dataGridView1);

            //���Ə��R���{
            //Utility.ComboOffice.load(comboBox1);

            DispClear();

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
                    tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", (float)9.5, FontStyle.Regular);
                    
                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 595;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "���t");
                    tempDGV.Columns.Add("col2", "�j��");
                    tempDGV.Columns.Add("col3", "�V��");
                    tempDGV.Columns.Add("col4", "�V�h����");
                    tempDGV.Columns.Add("col5", "�铌����");
                    tempDGV.Columns.Add("col6", "�v");
                    tempDGV.Columns.Add("col7", "�V�h���x");
                    tempDGV.Columns.Add("col8", "�铌���x");
                    tempDGV.Columns.Add("col9", "�v");
                    tempDGV.Columns.Add("col10", "�V�h�z�z��");
                    tempDGV.Columns.Add("col11", "�铌�z�z��");
                    tempDGV.Columns.Add("col12", "�v");

                    tempDGV.Columns[0].Width = 55;
                    tempDGV.Columns[1].Width = 55;
                    tempDGV.Columns[2].Width = 60;
                    tempDGV.Columns[3].Width = 90;
                    tempDGV.Columns[4].Width = 90;
                    tempDGV.Columns[5].Width = 90;
                    tempDGV.Columns[6].Width = 90;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;
                    tempDGV.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
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

            public static void ShowData(DataGridView tempDGV,int temptYear,int temptMonth)
            {
                string sqlSTRING = "";

                const int GYOSU = 32;
                DateTime sDate;

                try
                {
                    Control.DataControl sdcon = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = sdcon.GetConnection();

                    //�f�[�^���[�_�[���擾����
                    OleDbDataReader dR;

                    sqlSTRING += "select �z�z��,���Ə�ID,count(distinct �z�z��ID) AS �z�z����,SUM(����) AS ����, ";
                    sqlSTRING += "SUM(����) AS ����, SUM(����) - SUM(����) AS ���x ";
                    sqlSTRING += "from ";
                    sqlSTRING += "(SELECT TOP (100) PERCENT �z�z�w��.�z�z��, �z�z�w��.�z�z��ID,";
                    sqlSTRING += "���Ə�.ID AS ���Ə�ID,��.�P��,�z�z�G���A.�z�z�P��,";
                    sqlSTRING += "�z�z�G���A.���z�z����,�z�z�G���A.�\�薇��,";
                    sqlSTRING += "��.�P�� * �z�z�G���A.�\�薇�� AS ����,";
                    sqlSTRING += "�z�z�G���A.�z�z�P�� * �z�z�G���A.���z�z���� AS ���� ";
                    sqlSTRING += "from �z�z�w�� INNER JOIN �z�z�G���A ";
                    sqlSTRING += "ON �z�z�w��.ID = �z�z�G���A.�z�z�w��ID INNER JOIN �� ";
                    sqlSTRING += "ON �z�z�G���A.��ID = ��.ID INNER JOIN �z�z�� ";
                    sqlSTRING += "ON �z�z�w��.�z�z��ID = �z�z��.ID LEFT OUTER JOIN ���Ə� ";
                    sqlSTRING += "ON �z�z��.���Ə��R�[�h = ���Ə�.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(YEAR(�z�z�w��.�z�z��) = ?) AND (MONTH(�z�z�w��.�z�z��) = ?) AND ";
                    sqlSTRING += "(���Ə�.ID = " + SHINJUKU.ToString() + ") ";
                    sqlSTRING += "order by �z�z�w��.�z�z��, �z�z�w��.�z�z��ID) AS sel_tbl ";
                    sqlSTRING += "group by �z�z��, ���Ə�ID ";

                    sqlSTRING += "union ";

                    sqlSTRING += "select �z�z��,���Ə�ID,count(distinct �z�z��ID) AS �z�z����,SUM(����) AS ����, ";
                    sqlSTRING += "SUM(����) AS ����, SUM(����) - SUM(����) AS ���x ";
                    sqlSTRING += "from ";
                    sqlSTRING += "(SELECT TOP (100) PERCENT �z�z�w��.�z�z��, �z�z�w��.�z�z��ID,";
                    sqlSTRING += "���Ə�.ID AS ���Ə�ID,��.�P��,�z�z�G���A.�z�z�P��,";
                    sqlSTRING += "�z�z�G���A.���z�z����,�z�z�G���A.�\�薇��,";
                    sqlSTRING += "��.�P�� * �z�z�G���A.�\�薇�� AS ����,";
                    sqlSTRING += "�z�z�G���A.�z�z�P�� * �z�z�G���A.���z�z���� AS ���� ";
                    sqlSTRING += "from �z�z�w�� INNER JOIN �z�z�G���A ";
                    sqlSTRING += "ON �z�z�w��.ID = �z�z�G���A.�z�z�w��ID INNER JOIN �� ";
                    sqlSTRING += "ON �z�z�G���A.��ID = ��.ID INNER JOIN �z�z�� ";
                    sqlSTRING += "ON �z�z�w��.�z�z��ID = �z�z��.ID LEFT OUTER JOIN ���Ə� ";
                    sqlSTRING += "ON �z�z��.���Ə��R�[�h = ���Ə�.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(YEAR(�z�z�w��.�z�z��) = ?) AND (MONTH(�z�z�w��.�z�z��) = ?) AND ";
                    sqlSTRING += "(���Ə�.ID = " + JYOUTOU.ToString() + ") ";
                    sqlSTRING += "order by �z�z�w��.�z�z��, �z�z�w��.�z�z��ID) AS sel_tbl ";
                    sqlSTRING += "group by �z�z��, ���Ə�ID ";
                    sqlSTRING += "order by �z�z��, ���Ə�ID ";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@year1", temptYear);
                    SCom.Parameters.AddWithValue("@month1", temptMonth);

                    SCom.Parameters.AddWithValue("@year2", temptYear);
                    SCom.Parameters.AddWithValue("@month2", temptMonth);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //�O���b�h�r���[�ɕ\������
                    int iX = 0;
                    double gzUri = 0;
                    double gtUri = 0;
                    double gzShushi = 0;
                    double gtShushi = 0;
                    double gzNin = 0;
                    double gtNin = 0;

                    //�O���b�h�쐬
                    tempDGV.RowCount = GYOSU;

                    //������(�[���Z�b�g�j
                    foreach (DataGridViewRow iRow in tempDGV.Rows)
                    {
                        foreach (DataGridViewColumn iColumn in tempDGV.Columns)
                        {
                            if ((iColumn.Index == 1) || (iColumn.Index == 2))
                            {
                                tempDGV[iColumn.Index, iRow.Index].Value = "";
                            }
                            else
                            {
                                tempDGV[iColumn.Index, iRow.Index].Value = (double)(0);
                            }
                        }
                    }

                    //���t�A�j���A�V����Z�b�g
                    for (int i = 0; i < GYOSU; i++)
                    {
                        int rDay;
                        string rDate;

                        rDay = i + 1;

                        //�Ώۓ��t
                        rDate = temptYear.ToString() + "/" + temptMonth.ToString() + "/" + rDay.ToString();
                        
                        if (DateTime.TryParse(rDate, out sDate) == true)
                        {
                            //���t
                            tempDGV[0, i].Value = rDay.ToString("00");

                            //�j��
                            tempDGV[1, i].Value = ("�����ΐ��؋��y").Substring(int.Parse(sDate.DayOfWeek.ToString("d")), 1);

                            //�V��
                            OleDbDataReader dRt;
                            Control.�V�� sTenkou = new Control.�V��();
                            dRt = sTenkou.FillBy("where ���t = '" + sDate.ToShortDateString() + "'");

                            while (dRt.Read())
                            {
                                tempDGV[2, i].Value = dRt["�V��"].ToString() + "";
                            }

                            dRt.Close();
                            sTenkou.Close();
                        }
                    }

                    //�f�[�^�\��
                    int sOffice = 0;

                    while (dR.Read())
                    {
                        iX = DateTime.Parse(dR["�z�z��"].ToString()).Day - 1;

                        //�V�hor�铌�̔��f
                        if (dR["���Ə�ID"] == DBNull.Value)
                        {
                            sOffice = 0;
                        }
                        else
                        {
                            sOffice = int.Parse(dR["���Ə�ID"].ToString(), System.Globalization.NumberStyles.Any);
                        }

                        switch (sOffice)
                        {
                            case SHINJUKU: //�V�h
                                tempDGV[3, iX].Value = double.Parse(dR["����"].ToString(), System.Globalization.NumberStyles.Any); //����
                                tempDGV[6, iX].Value = double.Parse(dR["���x"].ToString(), System.Globalization.NumberStyles.Any); //���x
                                tempDGV[9, iX].Value = double.Parse(dR["�z�z����"].ToString(), System.Globalization.NumberStyles.Any); //�z�z����

                                //���v
                                gtUri += double.Parse(dR["����"].ToString(), System.Globalization.NumberStyles.Any); //����
                                gtShushi += double.Parse(dR["���x"].ToString(), System.Globalization.NumberStyles.Any); //���x
                                gtNin += double.Parse(dR["�z�z����"].ToString(), System.Globalization.NumberStyles.Any); //�z�z����

                                break;

                            case JYOUTOU:
                                tempDGV[4, iX].Value = double.Parse(dR["����"].ToString(), System.Globalization.NumberStyles.Any); //����
                                tempDGV[7, iX].Value = double.Parse(dR["���x"].ToString(), System.Globalization.NumberStyles.Any); //���x
                                tempDGV[10, iX].Value = double.Parse(dR["�z�z����"].ToString(), System.Globalization.NumberStyles.Any); //���x

                                //���v
                                gzUri += double.Parse(dR["����"].ToString(), System.Globalization.NumberStyles.Any); //����
                                gzShushi += double.Parse(dR["���x"].ToString(), System.Globalization.NumberStyles.Any); //���x
                                gzNin += double.Parse(dR["�z�z����"].ToString(), System.Globalization.NumberStyles.Any); //�z�z����
 
                                break;
                        }
                    }

                    //���v�s
                    if (tempDGV.RowCount == 0)
                    {
                        MessageBox.Show("�Y������f�[�^������܂���ł���", MESSAGE_CAPTION);
                    }
                    else
                    {
                        tempDGV[0, GYOSU -1].Value = "�v";
                        tempDGV[3, GYOSU - 1].Value = gtUri;
                        tempDGV[4, GYOSU - 1].Value = gzUri;
                        tempDGV[6, GYOSU - 1].Value = gtShushi;
                        tempDGV[7, GYOSU - 1].Value = gzShushi;
                        tempDGV[9, GYOSU - 1].Value = gtNin;
                        tempDGV[10, GYOSU - 1].Value = gzNin;
                    }

                    //if (tempDGV.RowCount <= 25)
                    //{
                    //    tempDGV.Columns[2].Width = 318;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[2].Width = 301;
                    //}

                    dR.Close();
                    sdcon.Close();
                    cn.Close();

                    tempDGV.CurrentCell = null;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }

            }
        }

        /// <summary>
        /// ��ʂ��N���A����
        /// </summary>
        private void DispClear()
        {

            try
            {

                txtYear.Text = "";
                txtMonth.Text = "";

                btnPrn.Enabled = false;
                button1.Enabled = false;
                txtYear.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            int tYear, tMonth;

            try
            {

                if (Utility.NumericCheck(txtYear.Text) == false)
                {
                    MessageBox.Show("�Ώ۔N�͐����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtYear.Focus();
                    return;
                }

                if (Utility.NumericCheck(txtMonth.Text) == false)
                {
                    MessageBox.Show("�Ώی��͐����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMonth.Focus();
                    return;
                }

                if (int.Parse(txtMonth.Text, System.Globalization.NumberStyles.Any) < 1 || int.Parse(txtMonth.Text, System.Globalization.NumberStyles.Any) > 12)
                {
                    MessageBox.Show("�Ώی�������������܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMonth.Focus();
                    return;
                }

                //����
                tYear = int.Parse(txtYear.Text, System.Globalization.NumberStyles.Any);
                tMonth = int.Parse(txtMonth.Text, System.Globalization.NumberStyles.Any);

                //�f�[�^�\��
                GridviewSet.ShowData(dataGridView1, tYear,tMonth);

                //���o����������
                //dataGridView1.Columns[1].HeaderText = zMonth.ToString() + "���x����";
                //dataGridView1.Columns[2].HeaderText = tMonth.ToString() + "���x����";
                //dataGridView1.Columns[5].HeaderText = zMonth.ToString() + "���x���x";
                //dataGridView1.Columns[6].HeaderText = tMonth.ToString() + "���x���x";
                //dataGridView1.Columns[9].HeaderText = zMonth.ToString() + "���z�z��";
                //dataGridView1.Columns[10].HeaderText = tMonth.ToString() + "���z�z��";

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                    button1.Enabled = true;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"�I��",MessageBoxButtons.OK,MessageBoxIcon.Stop);
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
            EigyoReport(dataGridView1);
        }


        private void EigyoReport(DataGridView tempDGV)
        {
            const int S_GYO = 4;        //�G�N�Z���t�@�C�����o���s
            //const int S_ROWSMAX = 13;   //�G�N�Z���t�@�C����ő�l
            string exlHead = "";        //���o��

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���݌v�\, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX < tempDGV.RowCount; iX++)
                    {
                        //���o��
                        int tYear,tMonth;

                        tYear = int.Parse(txtYear.Text, System.Globalization.NumberStyles.Any);
                        tMonth = int.Parse(txtMonth.Text,System.Globalization.NumberStyles.Any);

                        exlHead = tYear.ToString() + "�N " + tMonth.ToString() + "���x �V�h�E�铌�݌v����";
                        oxlsSheet.Cells[S_GYO - 3, 3] = exlHead;

                        //oxlsSheet.Cells[S_GYO - 1, 2] = tempDGV.Columns[1].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 3] = tempDGV.Columns[2].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 6] = tempDGV.Columns[5].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 7] = tempDGV.Columns[6].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 10] = tempDGV.Columns[9].HeaderText;
                        //oxlsSheet.Cells[S_GYO - 1, 11] = tempDGV.Columns[10].HeaderText;

                        //����
                        if (tempDGV[0, iX].Value.ToString() == "0")
                        {
                            oxlsSheet.Cells[iX + S_GYO, 1] = "";
                            oxlsSheet.Cells[iX + S_GYO, 2] = "";
                            oxlsSheet.Cells[iX + S_GYO, 3] = "";
                            oxlsSheet.Cells[iX + S_GYO, 4] = "";
                            oxlsSheet.Cells[iX + S_GYO, 5] = "";
                            oxlsSheet.Cells[iX + S_GYO, 6] = "";
                            oxlsSheet.Cells[iX + S_GYO, 7] = "";
                            oxlsSheet.Cells[iX + S_GYO, 8] = "";
                            oxlsSheet.Cells[iX + S_GYO, 9] = "";
                            oxlsSheet.Cells[iX + S_GYO, 10] = "";
                            oxlsSheet.Cells[iX + S_GYO, 11] = "";
                            oxlsSheet.Cells[iX + S_GYO, 12] = "";
                        }
                        else
                        {
                            oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                            oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                            oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                            oxlsSheet.Cells[iX + S_GYO, 4] = double.Parse(tempDGV[3, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 5] = double.Parse(tempDGV[4, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 6] = double.Parse(tempDGV[5, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 7] = double.Parse(tempDGV[6, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 8] = double.Parse(tempDGV[7, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 9] = double.Parse(tempDGV[8, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 10] = double.Parse(tempDGV[9, iX].Value.ToString(),System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 11] = double.Parse(tempDGV[10, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                            oxlsSheet.Cells[iX + S_GYO, 12] = double.Parse(tempDGV[11, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                        }
                    }

                    ////////�Z���㕔�֎������R�r��������
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////�Z�������֎������R�r��������
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////�\�S�̂Ɏ����c�r��������
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////�\�S�̂̍��[�c�r��
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////////�\�S�̂̉E�[�c�r��
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    
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
                    saveFileDialog1.Title = MESSAGE_CAPTION;
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = exlHead;
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

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            switch (e.ColumnIndex)
	        {
		        case 3: //�V�h����  
                    dgvUri_Sai(3, 4, e.RowIndex);
                    break;

                case 4: //�铌����
                    dgvUri_Sai(3, 4, e.RowIndex);
                    break;

                case 6: //�V�h���x
                    dgvUri_Sai(6, 7, e.RowIndex);
                    break;

                case 7: //�铌���x
                    dgvUri_Sai(6, 7, e.RowIndex);
                    break;

                case 9: //�V�h�l��
                    dgvUri_Sai(9, 10, e.RowIndex);
                    break;

                case 10: //�铌�l��
                    dgvUri_Sai(9, 10, e.RowIndex);
                    break;
	        }

        }

        /// <summary>
        /// ���v�v�Z
        /// </summary>
        /// <param name="tempCol1">�O���J����</param>
        /// <param name="tempCol2">�����J����</param>
        /// <param name="tempRow">�s</param>
        private void dgvUri_Sai(int tempCol1, int tempCol2, int tempRow)
        {
            double z,t,r;

            //���v�v�Z
            if (dataGridView1[tempCol1, tempRow].Value == null) dataGridView1[tempCol1, tempRow].Value = (double)(0);
            if (dataGridView1[tempCol2, tempRow].Value == null) dataGridView1[tempCol2, tempRow].Value = (double)(0);

            z = double.Parse(dataGridView1[tempCol1,tempRow].Value.ToString(),System.Globalization.NumberStyles.Any);
            t = double.Parse(dataGridView1[tempCol2,tempRow].Value.ToString(),System.Globalization.NumberStyles.Any);
            r = t + z;

            dataGridView1[tempCol2 + 1, tempRow].Value = r;
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.BackColor = Color.LightGray;
            txtObj.SelectAll();
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.BackColor = Color.White;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }
    }
}
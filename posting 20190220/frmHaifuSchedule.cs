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
    public partial class frmSchedule : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "�z�z�X�P�W���[��";

        public frmSchedule()
        {            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView2,DateTime.Parse(dateTimePicker1.Value.ToShortDateString()));

            //���Ə��R���{
            Utility.ComboOffice.load(comboBox1);
            comboBox1.SelectedIndex = -1;
            comboBox1.Enabled = false;

            //�`�F�b�N�{�b�N�X
            checkBox1.Checked = false;

            button1.Enabled = false;

        }

        // �f�[�^�O���b�h�r���[�N���X
        private class GridviewSet
        {

            /// <summary>
            /// �f�[�^�O���b�h�r���[�̒�`���s���܂�
            /// </summary>
            /// <param name="tempDGV">�f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            public static void Setting(DataGridView tempDGV,DateTime tempDate)
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
                    tempDGV.Columns.Add("col1", "ID");
                    tempDGV.Columns.Add("col2", "�`���V��");
                    tempDGV.Columns.Add("col3", "���Ə�");
                    tempDGV.Columns.Add("col4", "�S����");
                    tempDGV.Columns.Add("col5", "�J�n��");
                    tempDGV.Columns.Add("col6", "�I����");
                    tempDGV.Columns.Add("col7", "�z�z�`��");
                    tempDGV.Columns.Add("col8", "�c����");
                    tempDGV.Columns.Add("col9", "�l��");
                    tempDGV.Columns.Add("col10", "������");

                    tempDGV.Columns[1].Frozen = true;

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 217;
                    tempDGV.Columns[2].Width = 70;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 90;
                    tempDGV.Columns[5].Width = 90;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 70;
                    tempDGV.Columns[8].Width = 60;
                    tempDGV.Columns[9].Width = 66;

                    //tempDGV.Columns.Add("col10", "d1");
                    //tempDGV.Columns.Add("col11", "d2");
                    //tempDGV.Columns.Add("col12", "d3");
                    //tempDGV.Columns.Add("col13", "d4");
                    //tempDGV.Columns.Add("col14", "d5");
                    //tempDGV.Columns.Add("col15", "d6");
                    //tempDGV.Columns.Add("col16", "d7");
                    //tempDGV.Columns.Add("col17", "d8");
                    //tempDGV.Columns.Add("col18", "d9");
                    //tempDGV.Columns.Add("col19", "d10");
                    //tempDGV.Columns.Add("col20", "d11");
                    //tempDGV.Columns.Add("col21", "d12");
                    //tempDGV.Columns.Add("col22", "d13");
                    //tempDGV.Columns.Add("col23", "d14");


                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
                    tempDGV.Columns.Add("col1", "�z�z�ꏊ");
                    tempDGV.Columns.Add("col2", "����");

                    tempDGV.Columns[0].Width = 200;
                    tempDGV.Columns[1].Width = 57;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[1].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV,DateTime tempDate,int tempOfficeID,ComboBox tempCmb)
            {
                try
                {

                    int iX;
                    int colno;
                    int  Nin;
                    string htxt;
                    string sqlSTRING = "";
                    DateTime sDate;
                    DateTime eDate;

                    const int DKIKAN = 14; 
                    DateTime [] mDate = new DateTime [DKIKAN];
                    int[] mTotal = new int[DKIKAN];

                    eDate = tempDate.AddDays(DKIKAN - 1);

                    //�J�[�\���ҋ@
                    Cursor.Current = Cursors.WaitCursor;

                    //���ԗ�폜
                    if (tempDGV.Columns.Count > 10)
                    {
                        for (int i = 0; i < DKIKAN; i++)
	                    {
	                        tempDGV.Columns.RemoveAt(10);
	                    }
                    }

                    //���ԗ�ǉ�
                    for (int i = 0; i < DKIKAN; i++)
                    {
                        colno = i + 10;
                        sDate = tempDate.AddDays(i);

                        htxt = sDate.Day.ToString() + Environment.NewLine + ("�����ΐ��؋��y").Substring(int.Parse(sDate.DayOfWeek.ToString("d")), 1);

                        tempDGV.Columns.Add("col" + colno.ToString(), htxt);
                        tempDGV.Columns[colno].Width = 40;
                        tempDGV.Columns[colno].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                        mDate[i] = sDate;

                        //�y���Ȃ�w�i�F�\��
                        if (sDate.DayOfWeek.ToString("d") == "0" || sDate.DayOfWeek.ToString("d") == "6")
                        {
                            tempDGV.Columns[colno].DefaultCellStyle.BackColor = Color.LightPink;
                        }
                    }

                    tempDGV.RowCount = 0;
                    
                    //�f�[�^���[�_�[���擾����
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select ��.ID,��.�`���V��,�Ј�.����,��.�z�z�J�n��,��.�z�z�I����,";
                    sqlSTRING += "�z�z�`��.����,�z�z�`��.��l�����薇��,��.����,sum_tbl.�z�z����,���Ə�.���� as ���Ə��� ";
                    sqlSTRING += "from �� LEFT OUTER JOIN ";
                    sqlSTRING += "���Ӑ� ON ��.���Ӑ�ID = ���Ӑ�.ID LEFT OUTER JOIN ";
                    sqlSTRING += "�Ј� ON ���Ӑ�.�S���Ј��R�[�h = �Ј�.ID LEFT OUTER JOIN ";
                    sqlSTRING += "�z�z�`�� ON ��.�z�z�`�� = �z�z�`��.ID LEFT OUTER JOIN ";

                    sqlSTRING += "(select �z�z�G���A.��ID, SUM(�z�z�G���A.�\�薇��) AS �z�z���� ";
                    sqlSTRING += "from �z�z�G���A INNER JOIN ";
                    sqlSTRING += "�z�z�w�� ON �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                    sqlSTRING += "where (�z�z�w��.�z�z�� < ?) ";
                    sqlSTRING += "group by �z�z�G���A.��ID) AS sum_tbl ON ��.ID = sum_tbl.��ID ";
                    sqlSTRING += "left join ";
                    sqlSTRING += "���Ə� on ��.���Ə�ID = ���Ə�.ID ";

                    sqlSTRING += "where ";

                    //���Ə��w��
                    if (tempCmb.SelectedIndex != -1)
                    {
                        sqlSTRING += "(��.���Ə�ID = ?) and ";
                    }

                    sqlSTRING += "(((��.�z�z�J�n�� >= ?) and (��.�z�z�J�n�� <= ?)) or ";
                    sqlSTRING += "((��.�z�z�I���� >= ?) and (��.�z�z�I���� <= ?)) or ";
                    sqlSTRING += "((��.�z�z�J�n�� <= ?) and (��.�z�z�I���� >= ?))) ";
                    sqlSTRING += "order by ��.ID";
                                        
                    //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@sDate", tempDate);

                    //���Ə��w��
                    if (tempCmb.SelectedIndex != -1)
                    {
                        SCom.Parameters.AddWithValue("@office", tempOfficeID);
                    }

                    SCom.Parameters.AddWithValue("@d1", tempDate);
                    SCom.Parameters.AddWithValue("@d2", eDate);
                    SCom.Parameters.AddWithValue("@d3", tempDate);
                    SCom.Parameters.AddWithValue("@d4", eDate);
                    SCom.Parameters.AddWithValue("@d5", tempDate);
                    SCom.Parameters.AddWithValue("@d6", eDate);

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
                            tempDGV[2, iX].Value = dR["���Ə���"].ToString() + "";
                            tempDGV[3, iX].Value = dR["����"].ToString();
                            tempDGV[4, iX].Value = DateTime.Parse(dR["�z�z�J�n��"].ToString()).ToShortDateString();
                            tempDGV[5, iX].Value = DateTime.Parse(dR["�z�z�I����"].ToString()).ToShortDateString();
                            tempDGV[6, iX].Value = dR["����"].ToString() + "";

                            if (dR["�z�z����"] == DBNull.Value)
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["����"].ToString());
                            }
                            else
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["����"].ToString()) - int.Parse(dR["�z�z����"].ToString());
                            }

                            if (dR["��l�����薇��"] == DBNull.Value)
                            {
                                Nin = (int)0;
                            }
                            else
                            {
                                Nin = int.Parse(dR["��l�����薇��"].ToString(), System.Globalization.NumberStyles.Any);
                            }

                            if (Nin == 0)
                            {
                                tempDGV[8, iX].Value = (double)(0);
                            }
                            else
                            {
                                tempDGV[8, iX].Value = System.Math.Floor(double.Parse(tempDGV[7,iX].Value.ToString(),System.Globalization.NumberStyles.Any) / Nin + 0.9);
                            }

                            TimeSpan tSpan;
                            tSpan = DateTime.Parse(dR["�z�z�I����"].ToString()) - DateTime.Parse(dR["�z�z�J�n��"].ToString());

                            tempDGV[9, iX].Value = tSpan.Days + 1;

                            //�X�P�W���[�����ɐl���\��
                            //�z�z�J�n�����N�_���ȑO�̂Ƃ��͋N�_���ɕ\��
                            if (DateTime.Parse(dR["�z�z�J�n��"].ToString()) < mDate[0])
                            {
                                tempDGV[10, iX].Value = tempDGV[8, iX].Value.ToString();
                                mTotal[0] += int.Parse(tempDGV[8, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                            }
                            else�@//�z�z�J�n���ɕ\��
                            {
                                for (int i = 0; i < mDate.Length; i++)
                                {
                                    if (DateTime.Parse(dR["�z�z�J�n��"].ToString()) == mDate[i])
                                    {
                                        tempDGV[i + 10, iX].Value = tempDGV[8, iX].Value.ToString();
                                        mTotal[i] += int.Parse(tempDGV[8, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                                    }
                                }
                            }

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

                    //if (tempDGV.RowCount <= 27)
                    //{
                    //    tempDGV.Columns[1].Width = 217;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 200;
                    //}

                    //�X�P�W���[���w�b�_�ɍ��v�l���\��
                    for (int i = 0; i < mTotal.Length; i++)
                    {
                        tempDGV.Columns[i + 10].HeaderText += Environment.NewLine + mTotal[i].ToString();
                    }

                    //�J�[�\���\���߂�
                    Cursor.Current = Cursors.Default;

                    if (iX == 0)
                    {
                        MessageBox.Show("�Y������f�[�^������܂���",MESSAGE_CAPTION);
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);

                    //�J�[�\���\���߂�
                    Cursor.Current = Cursors.Default;
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

            int officeID = 0;

            if (checkBox1.Checked == true)
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    MessageBox.Show("���Ə���I�����Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    comboBox1.Focus();
                    return;
                }
                else
                {
                    //���Ə��R�[�h
                    Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                    cmb1 = (Utility.ComboOffice)comboBox1.SelectedItem;
                    officeID = cmb1.ID;
                }
            }

            //��ʕ\��
            GridviewSet.ShowData(dataGridView2, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()),officeID,comboBox1);
            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount > 0)
            {
                button1.Enabled = true;
            }
        }

        private void ShowPosting(DataGridView tempDGV,DateTime tempDate,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select ����.����,�z�z�G���A.�}�ԋL��,�z�z�G���A.�񍐖��� ";
            mySql += "from (�z�z�G���A inner join �z�z�w�� ";
            mySql += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID) left join ���� ";
            mySql += "on �z�z�G���A.����ID = ����.ID ";
            mySql += "where (�z�z�G���A.��ID = " + tempID.ToString() + ") and ";
            mySql += "(�z�z�w��.�z�z�� = '" + tempDate.ToShortDateString() + "') ";
            mySql += "order by ����ID";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["����"].ToString() + " " + dR["�}�ԋL��"].ToString() + "";
                tempDGV[1, iX].Value = int.Parse(dR["�񍐖���"].ToString());

                iX++;
            }

            if (tempDGV.RowCount <= 27)
            {
                tempDGV.Columns[0].Width = 200;
            }
            else
            {
                tempDGV.Columns[0].Width = 183;
            }

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�z�z�X�P�W���[���\��������܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            PrintReport();

        }

        private void PrintReport()
        {

            const int S_GYO = 5;        //�G�N�Z���t�@�C�����ׂ�5�s�ڂ����
            const int S_ROWSMAX = 24;   //�G�N�Z���t�@�C����ő�l
            const int DKIKAN = 14;      //���ԗ�

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���z�z�X�P�W���[��, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];
                Excel.Range srng;

                try
                {
                    //���o��
                    int sPos,sIdx,sCnt;
                    string sHead;
                    oxlsSheet.Cells[1, 1] = "�z�z�X�P�W���[��  [" + dateTimePicker1.Value.ToShortDateString() + " �` " + dateTimePicker1.Value.AddDays(13).ToShortDateString() + "]";

                    //���E�j���E�l��
                    for (int i = 10; i < S_ROWSMAX; i++)
                    {
                        sHead = dataGridView2.Columns[i].HeaderText.Replace(Environment.NewLine, "*");
                        sPos = -1;
                        sIdx = 0;
                        sCnt = -3;

                        while (true)
	                    {
                            sPos =sHead.IndexOf("*", sPos + 1);

                            if (sPos == -1)
                            {
                                oxlsSheet.Cells[S_GYO + sCnt, i + 1] = sHead.Substring(sIdx, sHead.Length - sIdx);
                                break;
                            }

                            oxlsSheet.Cells[S_GYO + sCnt, i + 1] = sHead.Substring(sIdx, sPos -sIdx);
                            sIdx = sPos + 1;
                            sCnt++;
	                    }
                        
                    }

                    //�z�z�G���A����
                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        oxlsSheet.Cells[i + S_GYO, 1] = dataGridView2[0, i].Value.ToString();   //�󒍔ԍ�
                        oxlsSheet.Cells[i + S_GYO, 2] = dataGridView2[1, i].Value.ToString();   //�`���V��
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView2[2, i].Value.ToString();   //���Ə���                        
                        oxlsSheet.Cells[i + S_GYO, 4] = dataGridView2[3, i].Value.ToString();   //�S���Җ�
                        oxlsSheet.Cells[i + S_GYO, 5] = dataGridView2[4, i].Value.ToString();   //�z�z�J�n��
                        oxlsSheet.Cells[i + S_GYO, 6] = dataGridView2[5, i].Value.ToString();   //�z�z�I����
                        oxlsSheet.Cells[i + S_GYO, 7] = dataGridView2[6, i].Value.ToString();   //�z�z�`��
                        oxlsSheet.Cells[i + S_GYO, 8] = int.Parse(dataGridView2[7, i].Value.ToString(), System.Globalization.NumberStyles.Any);   //�c����
                        oxlsSheet.Cells[i + S_GYO, 9] = int.Parse(dataGridView2[8, i].Value.ToString(), System.Globalization.NumberStyles.Any);   //�l��
                        oxlsSheet.Cells[i + S_GYO, 10] = int.Parse(dataGridView2[9, i].Value.ToString(), System.Globalization.NumberStyles.Any);   //������

                        for (int n = 1; n <= DKIKAN; n++)
                        {
                            if (dataGridView2[n + 9, i].Value != null)
                            {
                                oxlsSheet.Cells[i + S_GYO, n + 10] = int.Parse(dataGridView2[n + 9, i].Value.ToString() + "", System.Globalization.NumberStyles.Any);   //�l��
                            }
                        }

                        ////�Z���㕔�֎������R�r��������
                        //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //�Z�������֎������R�r��������
                        rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }

                    //�\�S�̂Ɏ����c�r��������
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̍��[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̉E�[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�y���̗�͔w�i�F��ς���
                    for (int i = S_ROWSMAX - DKIKAN + 1; i <= S_ROWSMAX; i++)
                    {
                        srng = (Excel.Range)oxlsSheet.Cells[3, i];
                        
                        if (srng.Text.ToString() == "�y" || srng.Text.ToString() == "��")
                        {
                            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, i];
                            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, i];
                            oxlsSheet.get_Range(rng[0], rng[1]).Interior.ColorIndex = 15;
                        }
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
                    saveFileDialog1.Title = MESSAGE_CAPTION;
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = MESSAGE_CAPTION;
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                comboBox1.Enabled = true;
            }
            else
            {
                comboBox1.Enabled = false;
                comboBox1.SelectedIndex = -1;
            }
        }


    }
}
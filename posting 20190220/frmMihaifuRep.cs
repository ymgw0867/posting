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
    public partial class frmMihaifuRep : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "���z�z���X�g";

        public frmMihaifuRep()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView2);
            GridviewSet.Setting2(dataGridView1);

            //////GridviewSet.ShowData(dataGridView2,"");
            //////dataGridView2.CurrentCell = null;
            //////dataGridView1.RowCount = 0;

            button2.Enabled = false;
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
                    tempDGV.Height = 595;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col2", "�`���V��");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 217;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
                    tempDGV.Height = 595;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�z�z��");
                    tempDGV.Columns.Add("col2", "����ID");
                    tempDGV.Columns.Add("col3", "����");
                    tempDGV.Columns.Add("col4", "�Ԓn���");
                    tempDGV.Columns.Add("col5", "�}���V������");
                    tempDGV.Columns.Add("col6", "���R");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 66;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 200;
                    tempDGV.Columns[5].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    //tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    //tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    //tempDGV.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static void ShowData(DataGridView tempDGV,string tempCName)
            {

                string sqlSTRING = "";
                int iX;

                try
                {
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection Cn = new OleDbConnection();

                    Cn = Con.GetConnection();

                    tempDGV.RowCount = 0;
                    
                    //�f�[�^���[�_�[���擾����
                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "SELECT ��.ID,��.�`���V�� ";
                    sqlSTRING += "from �� INNER JOIN �z�z�G���A ON ��.ID = �z�z�G���A.��ID ";
                    sqlSTRING += "INNER JOIN ���z�z��� ON �z�z�G���A.ID = ���z�z���.�z�z�G���AID ";

                    sqlSTRING += "where ��.�`���V�� like ? ";

                    sqlSTRING += "group by ��.ID,��.�`���V�� ";
                    sqlSTRING += "ORDER BY ��.ID DESC";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@CName", "%" + tempCName + "%");

                    SCom.Connection = Cn;


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

                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();
                    Con.Close();
                    Cn.Close();

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
            dataGridView2.CurrentCell = null;
            dataGridView1.RowCount = 0;

            button2.Enabled = false;
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ShowPosting(dataGridView1, long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
            button2.Enabled = true;
        }

        private void ShowPosting(DataGridView tempDGV,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "SELECT ��.ID,��.�`���V��,�z�z�G���A.����ID,����.����,���z�z���.�Ԓn��,";
            mySql += "���z�z���.�}���V������,���z�z���.���R,���z�z���R.�E�v,���z�z���.���̑����e,";
            mySql += "�z�z�w��.�z�z�� ";
            mySql += "from �� ";
            mySql += "inner join �z�z�G���A ON ��.ID = �z�z�G���A.��ID ";
            mySql += "inner join ���z�z��� on �z�z�G���A.ID = ���z�z���.�z�z�G���AID ";
            mySql += "inner join ���� ON �z�z�G���A.����ID = ����.ID ";
            mySql += "inner join ���z�z���R ON ���z�z���.���R = ���z�z���R.ID ";
            mySql += "inner join �z�z�w�� ON �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
            mySql += "where ��.ID = " + tempID.ToString() + " ";
            mySql += " ORDER BY �z�z��,�z�z�G���A.����ID,�Ԓn��";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;

            while (dR.Read())
            {
                tempDGV.Rows.Add();

                if (dR["�z�z��"] != DBNull.Value)
                {
                    tempDGV[0, iX].Value = DateTime.Parse(dR["�z�z��"].ToString()).ToShortDateString() + "";
                }
                else
                {
                    tempDGV[0, iX].Value = "";
                }

                tempDGV[1, iX].Value = dR["����ID"].ToString() + "";
                tempDGV[2, iX].Value = dR["����"].ToString() + "";
                tempDGV[3, iX].Value = dR["�Ԓn��"].ToString() + "";
                tempDGV[4, iX].Value = dR["�}���V������"].ToString() + "";

                if (dR["�E�v"].ToString() == "���̑�")
                {
                    tempDGV[5, iX].Value = dR["���̑����e"].ToString() + "";
                }
                else
                {
                    tempDGV[5, iX].Value = dR["�E�v"].ToString() + "";
                }

                iX++;
            }

            //////if (tempDGV.RowCount <= 27)
            //////{
            //////    tempDGV.Columns[3].Width = 200;
            //////}
            //////else
            //////{
            //////    tempDGV.Columns[3].Width = 183;
            //////}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�z�z�����񍐏��𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 32; //�z�z�����񍐏��̖��׍s��
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

            const int S_GYO = 13;    //�G�N�Z���t�@�C�����ׂ�13�s�ڂ����
            int dgvIndex;
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

                try
                {
                    //���Ӑ���
                    long sID;
                    string sqlSTR;

                    sID = long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString());

                    sqlSTR = "";
                    sqlSTR += "select ���Ӑ�.����,���Ӑ�.�S���Җ�,���Ӑ�.�d�b�ԍ� ";
                    sqlSTR += "from �� inner join ���Ӑ� ";
                    sqlSTR += "on ��.���Ӑ�ID = ���Ӑ�.ID ";
                    sqlSTR += "where ��.ID = " + sID.ToString();

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTR);

                    while (dR.Read())
                    {
                        oxlsSheet.Cells[1, 3] = dR["����"].ToString() + " " + dR["�S���Җ�"].ToString() + "�l";
                        oxlsSheet.Cells[2, 3] = dR["�d�b�ԍ�"].ToString();                        
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
                    
                    //�z�z�G���A����
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //�f�[�^�O���b�h�r���[�̍s�C���f�b�N�X�����߂�

                        //oxlsSheet.Cells[i + S_GYO, 2] = dateTimePicker1.Value.ToShortDateString();   //�`���V��
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView1[0, dgvIndex].Value.ToString();   //�z�z�敪
                        oxlsSheet.Cells[i + S_GYO, 4] = int.Parse(dataGridView1[1, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //�O���b�h�ŏI�s�̂Ƃ��I��
                        if (dgvIndex == (dataGridView1.Rows.Count - 1)) break;

                        //������׍ő�s�̂Ƃ��I��
                        if (i == (tempFixRows - 1)) break;

                        i++;
                    }

                    //�z�z�������v
                    oxlsSheet.Cells[45, 4] = int.Parse(dataGridView2[5, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                    //�c����
                    oxlsSheet.Cells[46, 4] = int.Parse(dataGridView2[6, dataGridView2.SelectedRows[0].Index].Value.ToString(), System.Globalization.NumberStyles.Any);

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
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString() + "_" + DateTime.Today.ToLongDateString();

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

        private void button2_Click_1(object sender, EventArgs e)
        {
            string csvTittle;

            csvTittle = dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString();
            MyLibrary.CsvOut.GridView(dataGridView1, csvTittle);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            GridviewSet.ShowData(dataGridView2, txtCName.Text);
            dataGridView2.CurrentCell = null;
            dataGridView1.RowCount = 0;
        }


    }
}
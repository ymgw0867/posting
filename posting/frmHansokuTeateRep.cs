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
    public partial class frmHansokuTeateRep : Form
    {
        const string MESSAGE_CAPTION = "�̔����i�蓖";

        public frmHansokuTeateRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView1);

            //�Ј��R���{
            Utility.ComboShain.load(comboBox1);

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
                    tempDGV.Height = 469;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "������");
                    tempDGV.Columns.Add("col2", "���Ӑ旪��");
                    tempDGV.Columns.Add("col3", "�����z");
                    tempDGV.Columns.Add("col4", "�蓖");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    tempDGV.Columns[0].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";

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

            public static void ShowData(DataGridView tempDGV,DateTime tempDate1,DateTime tempDate2,int tempID,Label t1,Label t2)
            {
                string sqlSTRING = "";
                int nTotal = 0;
                int Total = 0;

                try
                {

                    Control.DataControl sdcon = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = sdcon.GetConnection();

                    //�f�[�^���[�_�[���擾����
                    OleDbDataReader dR;

                    sqlSTRING += "select ����.*,������.�ŗ�,���Ӑ�.���� ";
                    sqlSTRING += "from ���� inner join ������ ";
                    sqlSTRING += "on ����.������ID = ������.ID left join ���Ӑ� ";
                    sqlSTRING += "on ������.���Ӑ�ID = ���Ӑ�.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(����.�����N���� >= ?) and ";
                    sqlSTRING += "(����.�����N���� <= ?) and ";
                    sqlSTRING += "(���Ӑ�.�S���Ј��R�[�h = ?) ";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@Date1", tempDate1);
                    SCom.Parameters.AddWithValue("@Date2", tempDate2);
                    SCom.Parameters.AddWithValue("@SID", tempID);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //�O���b�h�r���[�ɕ\������
                    int iX = 0;
                    int RT;
                    double sKin;

                    RT = Properties.Settings.Default.�̑��蓖��;
                    tempDGV.RowCount = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = Convert.ToDateTime(dR["�����N����"].ToString());
                        tempDGV[1, iX].Value = dR["����"].ToString();
                        tempDGV[2, iX].Value = int.Parse(dR["���z"].ToString(), System.Globalization.NumberStyles.Any);
                        
                        sKin = double.Parse(dR["���z"].ToString(), System.Globalization.NumberStyles.Any) / (1 + double.Parse(dR["�ŗ�"].ToString(), System.Globalization.NumberStyles.Any) / 100);
                        sKin = Math.Floor(sKin * RT / 100 + 0.5);
                        tempDGV[3, iX].Value = int.Parse(sKin.ToString(), System.Globalization.NumberStyles.Any);

                        //�����z���v
                        nTotal += int.Parse(dR["���z"].ToString(), System.Globalization.NumberStyles.Any);

                        //�蓖���v
                        Total += int.Parse(sKin.ToString(),System.Globalization.NumberStyles.Any);

                        iX++;
                    }

                    if (tempDGV.RowCount == 0)
                    {
                        MessageBox.Show("�Y����������f�[�^������܂���ł���",MESSAGE_CAPTION);
                    }

                    //if (tempDGV.RowCount <= 25)
                    //{
                    //    tempDGV.Columns[1].Width = 300;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 293;
                    //}

                    dR.Close();
                    sdcon.Close();
                    cn.Close();

                    //���v�\��
                    t1.Text = nTotal.ToString("#,##0");
                    t2.Text = Total.ToString("#,##0");

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
                label5.Text = "";
                label6.Text = "";

                tDate.Checked = false;
                tDate2.Checked = false;
                comboBox1.SelectedIndex = -1;

                btnPrn.Enabled = false;
                tDate.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            DateTime StartDate;
            DateTime EndDate;
            int ShainID;

            try
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    MessageBox.Show("�S���҂�I�����Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (tDate.Checked == false)
                {
                    StartDate = Convert.ToDateTime("1900/01/01");
                }
                else
                {
                    StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
                }
                
                if (tDate2.Checked == false)
                {
                    EndDate = Convert.ToDateTime("9999/12/31");
                }
                else
                {
                    EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
                }

                if (comboBox1.SelectedIndex == -1)
                {
                    ShainID = 0;
                }
                else
                {
                    Utility.ComboShain cmb1 = new Utility.ComboShain();
                    cmb1 = (Utility.ComboShain)comboBox1.SelectedItem;
                    ShainID = cmb1.ID;
                }

                GridviewSet.ShowData(dataGridView1, StartDate, EndDate, ShainID,label5,label6);

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
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
            const int S_GYO = 4;    //�G�N�Z���t�@�C�����o���s�i���ׂ�3�s�ڂ���󎚁j
            const int S_ROWSMAX = 4; //�G�N�Z���t�@�C����ő�l
            string sMidashi;

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���̔����i�蓖, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        sMidashi = this.comboBox1.Text + "  ";

                        if (tDate.Checked == true)
                        {
                            sMidashi += this.tDate.Value.ToShortDateString() + " �` ";
                        }

                        if (tDate2.Checked == true)
                        {
                            if (tDate.Checked == false) sMidashi += " �` ";

                            sMidashi += this.tDate2.Value.ToShortDateString();
                        }

                        oxlsSheet.Cells[S_GYO - 2, 1] = sMidashi;
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[3, iX].Value.ToString();

                        //�Z�������֎������R�r��������
                        rng[0] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    }

                    ////�Z���㕔�֎������R�r��������
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////�Z�������֎������R�r��������
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

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

                    //���v
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count + 1, 3] = this.label5.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 4] = this.label6.Text;

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

                    string msgHD = "";

                    if (tDate.Checked == true)
                    {
                        msgHD += tDate.Value.ToLongDateString() + "����";
                    }

                    if (tDate2.Checked == true)
                    {
                        msgHD += tDate2.Value.ToLongDateString() + "�܂�";
                    }

                    //�_�C�A���O�{�b�N�X�̏����ݒ�
                    saveFileDialog1.Title = MESSAGE_CAPTION;
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = MESSAGE_CAPTION + "_" + comboBox1.Text + "_" + msgHD; 
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
                MessageBox.Show(e.Message, "�̔����i�蓖", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
            label5.Text = "";
            label6.Text = "";
        }

    }
}
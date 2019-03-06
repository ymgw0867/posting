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
    public partial class frmChirashiZaiko : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "�`���V�݌ɊǗ��\";

        public frmChirashiZaiko()
        {

            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView2);
            button1.Enabled = false;

            radioButton2.Checked = true;
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();
            txtChirashiName.Text = string.Empty;
            txtsOrderNum.Text = string.Empty;

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
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    // ��w�b�_�[�\���ʒu�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // ��w�b�_�[�t�H���g�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 508;

                    // ��s�̐F
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col9", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col1", "���i��");
                    tempDGV.Columns.Add("col2", "������");
                    tempDGV.Columns.Add("col3", "�z�z��");
                    tempDGV.Columns.Add("col4", "������");
                    tempDGV.Columns.Add("col5", "�z�z�w��");
                    tempDGV.Columns.Add("col6", "�w���c��");
                    tempDGV.Columns.Add("col7", "�z�z����");
                    tempDGV.Columns.Add("col8", "�����c��");

                    tempDGV.Columns[0].Width = 120;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 190;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 80;
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
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
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
                    
                    // �r��
                    tempDGV.BorderStyle = BorderStyle.Fixed3D;
                    tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void ShowData(DataGridView tempDGV, int tempYear, int tempMonth, int temprb, string cName, Int64 orderNum)
            {
                string sqlSTRING = "";
                int iX;

                try
                {
                    //�J�[�\���\����ҋ@���
                    Cursor.Current = Cursors.WaitCursor;

                    tempDGV.RowCount = 0;
                    
                    //�f�[�^���[�_�[���擾����
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select ��.ID,��.�`���V��,��.�[�i�\���,��.�z�z�J�n��,";
                    sqlSTRING += "��.�z�z�I����,��.����,t_tbl.�z�z��,";
                    sqlSTRING += "��.���� - t_tbl.�z�z�� AS �c��,";
                    sqlSTRING += "k_tbl.�����z�z��,";
                    sqlSTRING += "��.���� - k_tbl.�����z�z�� AS �����c��,";
                    sqlSTRING += "��.�O���˗����x��, ��.�O���ϑ����� ";
                    
                    sqlSTRING += "from �� inner join ";

                    sqlSTRING += "(SELECT ��_1.ID,SUM(�z�z�G���A.�\�薇��) AS �z�z�� ";
                    sqlSTRING += "from �� AS ��_1 INNER JOIN �z�z�G���A ";
                    sqlSTRING += "ON ��_1.ID = �z�z�G���A.��ID ";
                    sqlSTRING += "where (�z�z�G���A.�z�z�w��ID <> 0) ";
                    sqlSTRING += "GROUP BY ��_1.ID) AS t_tbl ";

                    sqlSTRING += "ON ��.ID = t_tbl.ID ";
                    sqlSTRING += "left join ";

                    sqlSTRING += "(SELECT ��_2.ID,SUM(�z�z�G���A.�\�薇��) AS �����z�z�� ";
                    sqlSTRING += "from �� AS ��_2 INNER JOIN �z�z�G���A ";
                    sqlSTRING += "ON ��_2.ID = �z�z�G���A.��ID ";
                    sqlSTRING += "where (�z�z�G���A.�z�z�w��ID <> 0) and (�z�z�G���A.�����敪 = 1)";
                    sqlSTRING += "GROUP BY ��_2.ID) AS k_tbl ";

                    sqlSTRING += "ON ��.ID = k_tbl.ID ";

                    sqlSTRING += "where (�󒍎��ID = 1) ";
                    sqlSTRING += "and (��.�O���˗����x�� is Null) "; // 2015/11/19

                    if (temprb == 1)
                    {
                        sqlSTRING += "and (year(��.�[�i�\���) = ?) and (month(��.�[�i�\���) = ?) ";
                    }

                    //sqlSTRING += "order by ��.�[�i�\��� DESC";

                    // �O���n���� 2015/11/19
                    sqlSTRING += "union ";
                    sqlSTRING += "select ��.ID,��.�`���V��,��.�[�i�\���,��.�z�z�J�n��,";
                    sqlSTRING += "��.�z�z�I����, ��.����, 0 as �z�z��,";
                    sqlSTRING += "��.���� - 0 AS �c��,";
                    sqlSTRING += "0 as �����z�z��,";
                    sqlSTRING += "��.���� - ��.�O���ϑ����� AS �����c��,";
                    sqlSTRING += "��.�O���˗����x��, ��.�O���ϑ����� ";
                    sqlSTRING += "from �� ";
                    sqlSTRING += "where (�󒍎��ID = 1) ";
                    sqlSTRING += "and (��.�O���˗����x�� > '2000/01/01') ";

                    if (temprb == 1)
                    {
                        sqlSTRING += "and (year(��.�[�i�\���) = ?) and (month(��.�[�i�\���) = ?) ";
                    }

                    sqlSTRING += "order by ��.�[�i�\��� DESC";
                    
                                        
                    //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    if (temprb == 1)
                    {
                        SCom.Parameters.AddWithValue("@year", tempYear);
                        SCom.Parameters.AddWithValue("@month", tempMonth);
                        SCom.Parameters.AddWithValue("@year", tempYear);
                        SCom.Parameters.AddWithValue("@month", tempMonth);
                    }

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();
                    
                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            // �`���V������ 2015/09/18
                            if (cName != string.Empty)
                            {
                                if (!dR["�`���V��"].ToString().Contains(cName))
                                {
                                    continue;
                                }
                            }

                            // �󒍔ԍ����� 2015/09/18
                            if (orderNum != 0)
                            {
                                if (!dR["ID"].ToString().Contains(orderNum.ToString()))
                                {
                                    continue;
                                }
                            }

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = dR["ID"].ToString();
                            tempDGV[1, iX].Value = dR["�`���V��"].ToString();

                            if (dR["�[�i�\���"] == DBNull.Value)
                            {
                                tempDGV[2, iX].Value = "";
                            }
                            else
                            {

                                tempDGV[2, iX].Value = DateTime.Parse(dR["�[�i�\���"].ToString()).ToShortDateString();
                            }

                            tempDGV[3, iX].Value = DateTime.Parse(dR["�z�z�J�n��"].ToString()).ToShortDateString() + "�`" + DateTime.Parse(dR["�z�z�I����"].ToString()).ToShortDateString();
                            tempDGV[4, iX].Value = int.Parse(dR["����"].ToString());
                            tempDGV[5, iX].Value = int.Parse(dR["�z�z��"].ToString());
                            tempDGV[6, iX].Value = int.Parse(dR["�c��"].ToString());

                            if (dR["�����z�z��"] == DBNull.Value)
                            {
                                tempDGV[7, iX].Value = 0;
                                tempDGV[8, iX].Value = int.Parse(dR["����"].ToString());

                            }
                            else
                            {
                                tempDGV[7, iX].Value = int.Parse(dR["�����z�z��"].ToString());
                                tempDGV[8, iX].Value = int.Parse(dR["�����c��"].ToString());
                            }

                            // �O����ɓn�����Ƃ� 2015/08/11
                            if (dR["�O���˗����x��"] != DBNull.Value)
                            {
                                // �ϑ������w�肪�Ȃ��Ƃ�
                                if (Utility.strToInt(dR["�O���ϑ�����"].ToString()) == global.FLGOFF)
                                {
                                    tempDGV[7, iX].Value = int.Parse(dR["����"].ToString()); // 2015/11/19
                                    tempDGV[8, iX].Value = "0";
                                }
                                else
                                { 
                                    // 2015/11/19
                                    tempDGV[7, iX].Value = dR["�O���ϑ�����"].ToString();

                                    // �ϑ������w��̂Ƃ��A��������ϑ��������������������������c��
                                    int m = Utility.strToInt(dR["����"].ToString()) - Utility.strToInt(dR["�O���ϑ�����"].ToString());
                                    tempDGV[8, iX].Value = m;
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

                    if (tempDGV.RowCount <= 27)
                    {
                        tempDGV.Columns[1].Width = 310;
                    }
                    else
                    {
                        tempDGV.Columns[1].Width = 293;
                    }

                    //�J�[�\���\����߂�
                    Cursor.Current = Cursors.Default;

                }
                catch (Exception e)
                {
                    //�J�[�\���\����߂�
                    Cursor.Current = Cursors.Default;

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
            int rb = 0;

            //if (MessageBox.Show("�`���V�݌ɊǗ��\��\�����܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
            //    return;

            if (radioButton1.Checked)
            {
                rb = 0;
            }
            else
            {
                rb = 1;
            }

            GridviewSet.ShowData(dataGridView2, int.Parse(txtYear.Text), int.Parse(txtMonth.Text), rb, txtChirashiName.Text, Utility.strToInt64(txtsOrderNum.Text));
            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("�Y���f�[�^������܂���ł���", MESSAGE_CAPTION);
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ShowPosting(dataGridView1, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�`���V�݌ɊǗ��\�𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            KanryoReport(dataGridView2);
        }

        private void KanryoReport(DataGridView tempDGV)
        {

            const int S_GYO = 4;        //�G�N�Z���t�@�C�����o���s�i���ׂ�4�s�ڂ���󎚁j
            const int S_ROWSMAX = 9;    //�G�N�Z���t�@�C����ő�l

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���݌ɊǗ��\, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        //oxlsSheet.Cells[S_GYO - 2, 1] = txtYear.Text + "�N" + txtMonth.Text + "�� �ғ����� " + "���ȏ�" ;
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[3, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[4, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 6] = tempDGV[5, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[6, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[7, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 9] = tempDGV[8, iX].Value.ToString();

                        //�Z�������֎������R�r��������
                        rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    }

                    //////�Z���㕔�֎������R�r��������
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

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = new RadioButton();

            if (sender == radioButton1)
            {
                txtYear.Enabled = false;
                txtMonth.Enabled = false;
                label1.Enabled = false;
                label2.Enabled = false;
            }
            else if (sender == radioButton2)
            {
                txtYear.Enabled = true;
                txtMonth.Enabled = true;
                label1.Enabled = true;
                label2.Enabled = true;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
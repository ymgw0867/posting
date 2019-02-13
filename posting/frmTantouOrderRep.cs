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
    public partial class frmTantouOrderRep : Form
    {
        const string MESSAGE_CAPTION = "�c�ƒS���ҕʎ󒍈ꗗ";
        const string COMBO_TOTAL = "���v";

        public frmTantouOrderRep()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.��' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            GridviewSet.Setting(dataGridView1);

            //�Ј��R���{
            Utility.ComboShain.load(comboBox1);

            //////comboBox1.Items.Add(COMBO_TOTAL);   //�Ј����v���ڂ�ǉ��@2010/1/27

            DispClear();
        }

        // �f�[�^�O���b�h�r���[�N���X
        private class GridviewSet
        {
            ///-----------------------------------------------------------------------------
            /// <summary>
            ///     �f�[�^�O���b�h�r���[�̒�`���s���܂� </summary>
            /// <param name="tempDGV">
            ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            ///-----------------------------------------------------------------------------
            public static void Setting(DataGridView tempDGV)
            {
                try
                {
                    //�t�H�[���T�C�Y��`

                    // ��X�^�C����ύX����
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    tempDGV.EnableHeadersVisualStyles = false;

                    // ��w�b�_�[�\���ʒu�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // ��w�b�_�[�t�H���g�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);
                    
                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 20;
                    tempDGV.RowTemplate.Height = 20;

                    // �S�̂̍���
                    tempDGV.Height = 460;

                    // ��s�̐F
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�󒍓�");
                    tempDGV.Columns.Add("col2", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col3", "�`���V��");
                    tempDGV.Columns.Add("col4", "�S��");
                    tempDGV.Columns.Add("col5", "�P��");
                    tempDGV.Columns.Add("col6", "����");
                    tempDGV.Columns.Add("col7", "�󒍍��v");
                    tempDGV.Columns.Add("col8", "�c�ƌ���");
                    tempDGV.Columns.Add("col9", "�e���P");
                    tempDGV.Columns.Add("col10", "�O����");
                    tempDGV.Columns.Add("col11", "�O����Q");
                    tempDGV.Columns.Add("col12", "�O����R");
                    tempDGV.Columns.Add("col13", "�e���Q");
                    tempDGV.Columns.Add("col14", "�e������");

                    tempDGV.Columns[0].Width = 100;
                    tempDGV.Columns[1].Width = 110;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 100;
                    tempDGV.Columns[11].Width = 100;
                    tempDGV.Columns[12].Width = 100;
                    tempDGV.Columns[13].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    tempDGV.Columns[0].DefaultCellStyle.Format = "yyyy/MM/dd";
                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0.00";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[11].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[12].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[13].DefaultCellStyle.Format = "#,##0";

                    //�`���V����̓I�[�g�T�C�Y���[�h
                    //tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static void ShowData(DataGridView tempDGV,DateTime tempDate1,DateTime tempDate2,int tempID)
            {
                string sqlSTRING = "";

                try
                {
                    if (tempID == 0)
                    {
                        tempDGV.Columns[3].Visible = true;
                    }
                    else
                    {
                        tempDGV.Columns[3].Visible = false;
                    }

                    Control.DataControl sdcon = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = sdcon.GetConnection();

                    //�f�[�^���[�_�[���擾����
                    OleDbDataReader dR;

                    sqlSTRING += "select ��.ID,��.�󒍓�,��.�`���V��,��.�P��,��.����,��.���z,��.�l���z,�Ј�.����, ";
                    sqlSTRING += "��.�O�������x��, ��.�O�������x��2, ��.�O�������x��3, ��.�O�������c�� ";
                    sqlSTRING += "from �� left join ���Ӑ� on ��.���Ӑ�ID = ���Ӑ�.ID ";
                    sqlSTRING += "left join �Ј� on ���Ӑ�.�S���Ј��R�[�h = �Ј�.ID ";
                    sqlSTRING += "where ";
                    sqlSTRING += "(��.�󒍓� >= ?) and ";
                    sqlSTRING += "(��.�󒍓� <= ?) ";

                    if (tempID != 0) sqlSTRING += "and (���Ӑ�.�S���Ј��R�[�h = ?)";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    SCom.Parameters.AddWithValue("@Date1", tempDate1);
                    SCom.Parameters.AddWithValue("@Date2", tempDate2);

                    if (tempID != 0) SCom.Parameters.AddWithValue("@SID", tempID);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //�O���b�h�r���[�ɕ\������
                    int iX = 0;
                    int Total = 0;
                    decimal TotalGai = 0;
                    decimal TotalArari = 0;
                    decimal TotalGai2 = 0;
                    decimal TotalArari2 = 0;
                    decimal TotalGai3 = 0;
                    decimal TotalArari3 = 0;
                    decimal TotalGai4 = 0;
                    decimal TotalArari4 = 0;
                    decimal TotalArariSai = 0;

                    DateTime jDate = DateTime.Today; ;

                    tempDGV.RowCount = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = Convert.ToDateTime(dR["�󒍓�"].ToString());
                        jDate = Convert.ToDateTime(dR["�󒍓�"].ToString());
                        tempDGV[1, iX].Value = dR["ID"].ToString();
                        tempDGV[2, iX].Value = dR["�`���V��"].ToString();
                        tempDGV[3, iX].Value = dR["����"].ToString();                        
                        tempDGV[4, iX].Value = double.Parse(dR["�P��"].ToString());
                        tempDGV[5, iX].Value = int.Parse(dR["����"].ToString(), System.Globalization.NumberStyles.Any);

                        // 2016/01/04 ���z�ɒl���𔽉f����
                        tempDGV[6, iX].Value = int.Parse(dR["���z"].ToString(), System.Globalization.NumberStyles.Any) - int.Parse(dR["�l���z"].ToString(), System.Globalization.NumberStyles.Any);

                        // �c�ƌ����@�O���� 2015/09/18
                        //decimal gaichuhi = ((decimal)Utility.strToInt(dR["����"].ToString())) * Utility.strToDecimal(dR["�O�������c��"].ToString());
                        decimal gaichuhi = Utility.strToDecimal(dR["�O�������c��"].ToString()); // 2015/12/06
                        tempDGV[7, iX].Value = gaichuhi;

                        // 2016/01/04 ���z�ɒl���𔽉f����
                        decimal aRari = ((decimal)Utility.strToLong(dR["���z"].ToString())) - ((decimal)Utility.strToLong(dR["�l���z"].ToString())) - gaichuhi;
                        tempDGV[8, iX].Value = aRari;

                        // �O���� 2015/09/18
                        decimal gaichuhi2 = Utility.strToDecimal(dR["�O�������x��"].ToString()); // 2015/12/06
                        tempDGV[9, iX].Value = gaichuhi2;

                        // �O����2 2016/10/24
                        decimal gaichuhi3 = Utility.strToDecimal(dR["�O�������x��2"].ToString());
                        tempDGV[10, iX].Value = gaichuhi3;

                        // �O����3 2016/10/24
                        decimal gaichuhi4 = Utility.strToDecimal(dR["�O�������x��3"].ToString());
                        tempDGV[11, iX].Value = gaichuhi4;

                        // 2016/01/04 ���z�ɒl���𔽉f����
                        decimal aRari2 = ((decimal)Utility.strToLong(dR["���z"].ToString())) - ((decimal)Utility.strToLong(dR["�l���z"].ToString())) - gaichuhi2 - gaichuhi3 - gaichuhi4;
                        tempDGV[12, iX].Value = aRari2;

                        // �e������ 2015/09/18
                        tempDGV[13, iX].Value = aRari - aRari2;

                        Total += (int.Parse(dR["���z"].ToString(), System.Globalization.NumberStyles.Any) - int.Parse(dR["�l���z"].ToString(), System.Globalization.NumberStyles.Any));
                        TotalGai += gaichuhi;   // �c�ƌ��� 2015/09/18
                        TotalArari += aRari;    // �e���P���v 2015/09/18
                        TotalGai2 += gaichuhi2; // �O����2���v 2015/09/18
                        TotalGai3 += gaichuhi3; // �O����3���v 2016/11/08
                        TotalGai4 += gaichuhi4; // �O����4���v 2016/11/08
                        TotalArari2 += aRari2;  // �e���Q���v 2015/09/18

                        TotalArariSai += (aRari -= aRari2);  // ���ٍ��v 2016/10/24
                        iX++;
                    }

                    //���v�s
                    if (tempDGV.RowCount == 0)
                    {
                        MessageBox.Show("�Y������f�[�^������܂���ł���", MESSAGE_CAPTION);
                    }
                    else
                    {
                        tempDGV.Rows.Add();
                        tempDGV[0, iX].Value = "";
                        tempDGV[1, iX].Value = "";
                        tempDGV[2, iX].Value = "���v : " + tempDGV.RowCount.ToString("#,##0") + " ��";
                        tempDGV[3, iX].Value = "";
                        tempDGV[4, iX].Value = "";
                        tempDGV[5, iX].Value = "";
                        tempDGV[6, iX].Value = Total;
                        tempDGV[7, iX].Value = TotalGai;        // 2015/09/18
                        tempDGV[8, iX].Value = TotalArari;      // 2015/09/18
                        tempDGV[9, iX].Value = TotalGai2;       // 2015/09/18
                        tempDGV[10, iX].Value = TotalGai3;    // 2015/09/18
                        tempDGV[11, iX].Value = TotalGai4;      // 2016/10/24
                        tempDGV[12, iX].Value = TotalArari2;      // 2016/10/24
                        tempDGV[13, iX].Value = TotalArariSai;  // 2015/09/18
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
                    if (comboBox1.Text != "")
                    {
                        MessageBox.Show("�S���҂̑I��������������܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
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

                if (comboBox1.Text == "")
                {
                    ShainID = 0;
                }
                else
                {

                    //if (comboBox1.Text == COMBO_TOTAL)
                    //{
                    //    ShainID = 0;
                    //}
                    //else
                    //{

                    Utility.ComboShain cmb1 = new Utility.ComboShain();
                    cmb1 = (Utility.ComboShain)comboBox1.SelectedItem;
                    ShainID = cmb1.ID;

                    //}
                }

                GridviewSet.ShowData(dataGridView1, StartDate, EndDate, ShainID);

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
            int TempID;

            if (comboBox1.Text == "")
            {
                TempID = 1;
            }
            else
            {
                TempID = 0;
            }

            EigyoReport(dataGridView1,TempID);
        }


        private void EigyoReport(DataGridView tempDGV,int tempID)
        {
            const int S_GYO = 4;        // �G�N�Z���t�@�C�����o���s�i���ׂ�3�s�ڂ���󎚁j
            const int S_ROWSMAX = 11;   // �G�N�Z���t�@�C����ő�l

            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���c�ƒS���ҕʎ󒍈ꗗ�V�[�g��, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        oxlsSheet.Cells[S_GYO - 2, 1] = this.comboBox1.Text;                // �S��
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();   // �󒍓�
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();   // �󒍔ԍ�

                        //���v�̂Ƃ�
                        if (tempID == 1)
                        {
                            oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString() + "�F" + tempDGV[3, iX].Value.ToString();
                        }
                        else
                        {
                            oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[2, iX].Value.ToString();
                        }

                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[4, iX].Value.ToString();   // �P��
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[5, iX].Value.ToString();   // ����
                        oxlsSheet.Cells[iX + S_GYO, 6] = tempDGV[6, iX].Value.ToString();   // ����
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[7, iX].Value.ToString();   // �c�ƌ��� 2015/09/18
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[8, iX].Value.ToString();   // �e���P   2015/09/18

                        // �O����P�C�Q�C�R���v 2016/10/24
                        double g = Utility.strToDouble(tempDGV[9, iX].Value.ToString()) +
                                Utility.strToDouble(tempDGV[10, iX].Value.ToString()) +
                                Utility.strToDouble(tempDGV[11, iX].Value.ToString());

                        oxlsSheet.Cells[iX + S_GYO, 9] = g.ToString();   // �O����  2016/10/24
                        oxlsSheet.Cells[iX + S_GYO, 10] = tempDGV[12, iX].Value.ToString();   // �e���Q  2015/09/18
                        oxlsSheet.Cells[iX + S_GYO, 11] = tempDGV[13, iX].Value.ToString();   // �e������  2015/09/18
                    }

                    //�Z���㕔�֎������R�r��������
                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�Z�������֎������R�r��������
                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

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
                MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        
        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
        }

    }
}
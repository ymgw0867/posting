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
using System.Linq;

namespace posting
{
    public partial class frmEiUriageRep : Form
    {
        const string MESSAGE_CAPTION = "�c�ƕʔ���\";

        public frmEiUriageRep()
        {
            InitializeComponent();

            jAdp.Fill(dts.��1);
            sAdp.Fill(dts.�V������);
            nAdp.Fill(dts.�V����);
            cAdp.Fill(dts.���Ӑ�);
            eAdp.Fill(dts.�Ј�);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.��1TableAdapter jAdp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.�V������TableAdapter sAdp = new darwinDataSetTableAdapters.�V������TableAdapter();
        darwinDataSetTableAdapters.�V����TableAdapter nAdp = new darwinDataSetTableAdapters.�V����TableAdapter();
        darwinDataSetTableAdapters.���Ӑ�TableAdapter cAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
        darwinDataSetTableAdapters.�Ј�TableAdapter eAdp = new darwinDataSetTableAdapters.�Ј�TableAdapter();

        private void form_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ő�T�C�Y
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            Setting(dataGridView1);

            //�Ј��R���{
            Utility.ComboShain.load(comboBox1);

            DispClear();
        }

        #region �O���b�h�r���[�J������`
        const string colNDt = "col0";           // ������
        const string colClient = "col1";        // ���Ӑ�
        const string colNyukin = "col2";        // ����
        const string colEganka = "col3";        // �c�ƌ���
        const string colGaichuhi = "col4";      // �O����
        const string colArari1 = "col5";        // �e��1
        const string colArari2 = "col6";        // �e��2
        const string colArariSai = "col7";      // �e������
        const string colGaichuhi2 = "col8";      // �O����2
        const string colGaichuhi3 = "col9";      // �O����3
        #endregion

        ///----------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̒�`���s���܂� </summary>
        /// <param name="tempDGV">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        ///----------------------------------------------------------------
        private static void Setting(DataGridView tempDGV)
        {
            try
            {
                //�t�H�[���T�C�Y��`

                // ��X�^�C����ύX����

                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // ��w�b�_�[�\���ʒu�w��
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // ��w�b�_�[�t�H���g�w��
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // �f�[�^�t�H���g�w��
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);
                    
                // �s�̍���
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // �S�̂̍���
                tempDGV.Height = 542;

                // ��s�̐F
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add(colNDt, "������");
                tempDGV.Columns.Add(colClient, "���Ӑ旪��");
                tempDGV.Columns.Add(colNyukin, "�����z");
                tempDGV.Columns.Add(colEganka, "�c�ƌ���");
                tempDGV.Columns.Add(colArari1, "�e���P");
                tempDGV.Columns.Add(colGaichuhi, "�O����");
                tempDGV.Columns.Add(colGaichuhi2, "�O����Q");
                tempDGV.Columns.Add(colGaichuhi3, "�O����R");
                tempDGV.Columns.Add(colArari2, "�e���Q");
                tempDGV.Columns.Add(colArariSai, "�e������");

                tempDGV.Columns[colNDt].Width = 110;
                tempDGV.Columns[colClient].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colNyukin].Width = 100;
                tempDGV.Columns[colEganka].Width = 100;
                tempDGV.Columns[colArari1].Width = 100;
                tempDGV.Columns[colGaichuhi].Width = 100;
                tempDGV.Columns[colGaichuhi2].Width = 100;
                tempDGV.Columns[colGaichuhi3].Width = 100;
                tempDGV.Columns[colArari2].Width = 100;
                tempDGV.Columns[colArariSai].Width = 100;

                tempDGV.Columns[colNDt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colNyukin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colEganka].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colArari1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colGaichuhi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colGaichuhi2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colGaichuhi3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colArari2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colArariSai].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                
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
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ShowData(DataGridView tempDGV,DateTime tempDate1,DateTime tempDate2,int tempID,Label t1,Label t2)
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

        ///---------------------------------------------------------------------------------------
        /// <summary>
        ///     �O���b�h�Ƀf�[�^��\�����܂� </summary>
        /// <param name="tempDGV">
        ///     dataGridView�I�u�W�F�N�g</param>
        /// <param name="tempDate1">
        ///     �J�n�N����</param>
        /// <param name="tempDate2">
        ///     �I���N����</param>
        /// <param name="tempID">
        ///     �Ј��ԍ�</param>
        ///---------------------------------------------------------------------------------------
        private void ShowDataLinq(DataGridView tempDGV, DateTime tempDate1, DateTime tempDate2, int tempID)
        {
            int nTotal = 0;
            int Total = 0;
            int iX = 0;
            int arari1Tl = 0;
            int arari2Tl = 0;
            int gaichuhiTl = 0;
            int gaichuhiTl2 = 0;
            int gaichuhiTl3 = 0;
            int arariSaiTl = 0;

            try
            {
                tempDGV.RowCount = 0;

                foreach (var t in dts.�V������.Where(a => a.�������� == global.FLGON && 
                                                          a.���� == 0 && a.Get��1Rows().Count() > 0 && 
                                                          a.Get�V����Rows().Count() > 0))
                {
                    bool hit = true;
                    int nyukin = 0;
                    DateTime nDate = DateTime.Today;

                    // �S���c�Ƃ�����
                    if (t.���Ӑ�Row == null || t.���Ӑ�Row.�Ј�Row == null || t.���Ӑ�Row.�Ј�Row.ID != tempID)
                    {
                        continue;
                    }

                    //if (t.���Ӑ�Row.�Ј�Row == null)
                    //{
                    //    continue;
                    //}

                    //if (t.���Ӑ�Row.�Ј�Row.ID != tempID)
                    //{
                    //    continue;
                    //}


                    // �S�Ă̓��������w��͈͓�������
                    foreach (var m in t.Get�V����Rows())
                    {
                        if (m.�����N���� < tempDate1 || m.�����N���� > tempDate2)
                        {
                            hit = false;
                            break;
                        }

                        nyukin += m.���z;
                        nDate = m.�����N����;
                    }

                    // �������͈͊O����F�l�O��
                    if (!hit)
                    {
                        continue;
                    }

                    tempDGV.Rows.Add();

                    tempDGV[colNDt, iX].Value = nDate.ToShortDateString();

                    if (t.���Ӑ�Row != null)
                    {
                        tempDGV[colClient, iX].Value = t.���Ӑ�Row.����;
                    }
                    else
                    {
                        tempDGV[colClient, iX].Value = string.Empty;
                    }

                    // ����ł��������� 2016/01/28
                    nyukin -= t.�����;
                    tempDGV[colNyukin, iX].Value = nyukin.ToString("#,##0");

                    decimal genka = 0;
                    decimal gaichuhi = 0;
                    decimal gaichuhi2 = 0;
                    decimal gaichuhi3 = 0;

                    // �e���Z��
                    foreach (var j in t.Get��1Rows())
                    {
                        //genka += ((decimal)j.���� * j.�O�������c��);
                        //gaichuhi += ((decimal)j.���� * j.�O�������x��);

                        // 2015/12/06 �O��������P�����猴�����z���͂֕ύX�ɔ���
                        genka += j.�O�������c��;
                        gaichuhi += j.�O�������x��; 
                        gaichuhi2 += j.�O�������x��2;     // 2016/10/24
                        gaichuhi3 += j.�O�������x��3;     // 2016/10/24
                    }

                    tempDGV[colEganka, iX].Value = genka.ToString("#,##0"); // �c�ƌ���
                    tempDGV[colArari1, iX].Value = ((decimal)nyukin - genka).ToString("#,##0"); // �e���P
                    tempDGV[colGaichuhi, iX].Value = gaichuhi.ToString("#,##0"); // �O����
                    tempDGV[colGaichuhi2, iX].Value = gaichuhi2.ToString("#,##0"); // �O����2
                    tempDGV[colGaichuhi3, iX].Value = gaichuhi3.ToString("#,##0"); // �O����3
                    tempDGV[colArari2, iX].Value = ((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3).ToString("#,##0"); // �e���Q
                    tempDGV[colArariSai, iX].Value = (((decimal)nyukin - genka) - ((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3)).ToString("#,##0"); // �e������

                    nTotal += nyukin;       // �����z���v
                    Total += (int)genka;    // �c�ƌ���
                    arari1Tl += (int)((decimal)nyukin - genka);     // �e���P
                    gaichuhiTl += (int)gaichuhi;
                    gaichuhiTl2 += (int)gaichuhi2;
                    gaichuhiTl3 += (int)gaichuhi3;
                    arari2Tl += (int)((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3);  // �e���Q
                    arariSaiTl += (int)(((decimal)nyukin - genka) - ((decimal)nyukin - gaichuhi - gaichuhi2 - gaichuhi3));  // �e������

                    iX++;
                }
                
                // �Y���f�[�^�Ȃ�
                if (tempDGV.RowCount == 0)
                {
                    MessageBox.Show("�Y����������ς݃f�[�^�A����і����m��f�[�^������܂���ł���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                
                // ���v�\��
                label5.Text = nTotal.ToString("#,##0");                 // ����
                lblEgenka.Text = Total.ToString("#,##0");               // �c�ƌ���
                lblArari1.Text = arari1Tl.ToString("#,##0");            // �e���P
                lblGaichuhi.Text = gaichuhiTl.ToString("#,##0");        // �O����
                lblGaichuhi2.Text = gaichuhiTl2.ToString("#,##0");      // �O����2
                lblGaichuhi3.Text = gaichuhiTl3.ToString("#,##0");      // �O����3
                lblArari2.Text = arari2Tl.ToString("#,##0");            // �e���Q
                lblArarisai.Text = arariSaiTl.ToString("#,##0");        // �e������

                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
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
                lblGaichuhi.Text = "";

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

                // �O���b�h�\��
                ShowDataLinq(dataGridView1, StartDate, EndDate, ShainID);

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
            const int S_GYO = 4;        //�G�N�Z���t�@�C�����o���s�i���ׂ�3�s�ڂ���󎚁j
            const int S_ROWSMAX = 8;    //�G�N�Z���t�@�C����ő�l
            string sMidashi;
            double gTL = 0;

            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�c�Ɣ���\, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
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
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[colNDt, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[colClient, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[colNyukin, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[colEganka, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[colArari1, iX].Value.ToString();

                        double g = Utility.strToDouble(tempDGV[colGaichuhi, iX].Value.ToString()) +
                                   Utility.strToDouble(tempDGV[colGaichuhi2, iX].Value.ToString()) +
                                   Utility.strToDouble(tempDGV[colGaichuhi3, iX].Value.ToString());

                        oxlsSheet.Cells[iX + S_GYO, 6] = g.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[colArari2, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[colArariSai, iX].Value.ToString();

                        //�Z�������֎������R�r��������
                        rng[0] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[iX + S_GYO, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        gTL += g;
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
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 4] = this.lblEgenka.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 5] = this.lblArari1.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 6] = gTL.ToString("#,##0");
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 7] = this.lblArari2.Text;
                    oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 8] = this.lblArarisai.Text;

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
                MessageBox.Show(e.Message, "�c�ƕʔ���\", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
            label5.Text = "";
            lblGaichuhi.Text = "";
        }

    }
}
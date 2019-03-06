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
    public partial class frmClientbyOrderRep : Form
    {
        const string MESSAGE_CAPTION = "�N���C�A���g�ʎ󒍓��e�ꗗ";

        public frmClientbyOrderRep()
        {
            InitializeComponent();

            // �f�[�^�ǂݍ���
            rAdp.Fill(dts.�V������);
            cAdp.Fill(dts.���Ӑ�);
            sAdp.Fill(dts.�Ј�);
            jAdp.Fill(dts.��1);
            kAdp.Fill(dts.�󒍎��);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.�V������TableAdapter rAdp = new darwinDataSetTableAdapters.�V������TableAdapter();
        darwinDataSetTableAdapters.���Ӑ�TableAdapter cAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
        darwinDataSetTableAdapters.�Ј�TableAdapter sAdp = new darwinDataSetTableAdapters.�Ј�TableAdapter();
        darwinDataSetTableAdapters.��1TableAdapter jAdp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.�󒍎��TableAdapter kAdp = new darwinDataSetTableAdapters.�󒍎��TableAdapter();

        bool mukouStatus = false;

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView1);

            //��ʃN���A
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
                    tempDGV.Height = 470;

                    // ��s�̐F
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�ԍ�");
                    tempDGV.Columns.Add("col2", "�N���C�A���g��");
                    tempDGV.Columns.Add("col3", "�󒍓��e");
                    tempDGV.Columns.Add("col4", "�S����");
                    tempDGV.Columns.Add("col5", "���s��");
                    tempDGV.Columns.Add("col6", "������z");
                    tempDGV.Columns.Add("col7", "�e��");
                    tempDGV.Columns.Add("col8", "�����\���");
                    tempDGV.Columns.Add("col9", "�z�z�J�n��");
                    tempDGV.Columns.Add("col10", "�z�z�I����");
                    tempDGV.Columns.Add("col11", "�敪");
                    tempDGV.Columns.Add("col12", "ID");


                    //tempDGV.Columns[1].Frozen = true;   // 2015/07/22

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 90;
                    tempDGV.Columns[4].Width = 90;
                    tempDGV.Columns[5].Width = 90;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 60;
                    tempDGV.Columns[11].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[3].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[5].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    //������ID���\���Ƃ���
                    //tempDGV.Columns[10].Visible = false;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // �ҏW�s�Ƃ���
                    tempDGV.ReadOnly = false;

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
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
        }
        
        private void ShowDataLinq(DataGridView g)
        {
            decimal sArari = 0;
            decimal sUriage = 0;

            try
            {
                // �O���b�h�r���[�ɕ\������
                int iX = 0;

                g.RowCount = 0;

                //var s = dts.�V������
                //        .Where(a => a.���� == global.FLGOFF && a.Get��1Rows().Count() > 0)
                //        .OrderBy(a => a.�������z).ThenBy(a => a.���������s��);


                var s = dts.��1.Where(a => a.�V������Row != null).OrderBy(a => a.���Ӑ�ID).ThenBy(a => a.�󒍎��ID);

                // ���������s��

                s = s.Where(a => a.���������s�� >= hDate.Value.Date).OrderBy(a => a.���Ӑ�ID).ThenBy(a => a.�󒍎��ID);

                if (hDate2.Checked)
                {
                    s = s.Where(a => a.���������s�� <= hDate2.Value.Date).OrderBy(a => a.���Ӑ�ID).ThenBy(a => a.�󒍎��ID);
                }

                // ���я�
                if (rBtnOrder.Checked == true)
                {
                    s = s.OrderBy(a => a.�󒍎��ID).ThenBy(a => a.���Ӑ�ID);
                }

                if (s.Count() == 0)
                {
                    MessageBox.Show("�Y������󒍃f�[�^�͂���܂���","��������",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                
                // �O���b�h�֕\��
                foreach (var t in s)
                {
                    // �N���C�A���g���w��
                    if (txtsClientName.Text.Trim() != string.Empty)
                    {
                        // ���Ӑ於
                        if (t.���Ӑ�Row == null)
                        {
                            continue;
                        }
                        else if (!t.���Ӑ�Row.����.Contains(txtsClientName.Text.Trim()))
                        {
                            continue;
                        }
                    }

                    g.Rows.Add();

                    g[0, iX].Value = t.���Ӑ�ID.ToString();    // ���Ӑ�ID

                    // ���Ӑ於�E�S����
                    if (t.���Ӑ�Row == null)
                    {
                        g[1, iX].Value = string.Empty;
                        g[3, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[1, iX].Value = t.���Ӑ�Row.����;

                        if (t.���Ӑ�Row.�Ј�Row == null)
                        {
                            g[3, iX].Value = string.Empty;
                        }
                        else
                        {
                            g[3, iX].Value = t.���Ӑ�Row.�Ј�Row.����;
                        }
                    }

                    // �󒍓��e
                    if (t.�󒍎��Row != null)
                    {
                        g[2, iX].Value = t.�󒍎��Row.����;
                    }
                    else
                    {
                        g[2, iX].Value = "";
                    }

                    g[4, iX].Value =  t.���������s��.ToShortDateString();     // ���������s��
                    g[5, iX].Value = t.������z.ToString("#,##0");            // ������z
                    //g[6, iX].Value = (t.������z - t.�O�������x�� - t.�O�������x��2 - t.�O�������x��3).ToString("#,##0"); // �e��
                    g[6, iX].Value = (t.���z - t.�l���z - t.�O�������x�� - t.�O�������x��2 - t.�O�������x��3).ToString("#,##0");   // �e�� 2017/04/17

                    // �����\���
                    if (t.�V������Row.Is�x������Null())
                    {
                        g[7, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[7, iX].Value = t.�V������Row.�x������.ToShortDateString();
                    }

                    // �z�z�J�n���F2017/05/24
                    if (t.Is�z�z�J�n��Null())
                    {
                        g[8, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[8, iX].Value = t.�z�z�J�n��.ToShortDateString();
                    }

                    // �z�z�I�����F2017/05/24
                    if (t.Is�z�z�I����Null())
                    {
                        g[9, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[9, iX].Value = t.�z�z�I����.ToShortDateString();
                    }

                    // �敪�F2017/05/24
                    g[10, iX].Value = "";
                    if (t.�V������Row != null)
                    {
                        if (t.�z�z�J�n�� > t.�����\���)
                        {
                            if (t.�V������Row.�������� == global.FLGON)
                            {
                                g[10, iX].Value = global.FLGON.ToString();
                            }
                            else
                            {
                                g[10, iX].Value = "2";
                            }
                        }
                    }

                    if (!t.Is�z�z�I����Null())
                    {
                        if (t.�z�z�I����.Year == DateTime.Today.Year && t.�z�z�I����.Month == DateTime.Today.Month)
                        {
                            int syymm = t.���������s��.Year * 100 + t.���������s��.Month;
                            int yymm = DateTime.Today.Year * 100 + DateTime.Today.Month;

                            if (syymm > yymm)
                            {
                                g[10, iX].Value = "3";
                            }
                        }
                    }

                    // �󒍊m�菑ID�F2017/05/24
                    g[11, iX].Value = t.ID.ToString();


                    iX++;

                    sUriage += t.������z;
                    //sArari += (t.������z - t.�O�������x�� - t.�O�������x��2 - t.�O�������x��3);
                    sArari += (t.���z - t.�l���z - t.�O�������x�� - t.�O�������x��2 - t.�O�������x��3);    // 2017/04/17
                }
                
                g.CurrentCell = null;

                txtUriTotal.Text = sUriage.ToString("#,##0");
                txtArariTotal.Text = sArari.ToString("#,##0");
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
                rBtnOrder.Checked = true;
                btnPrn.Enabled = false;
                hDate.Checked = false;
                hDate2.Checked = false;
                txtsClientName.Text = "";
                rBtnOrder.Checked = false;
                rBtnClient.Checked = true;
                linkLabel1.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                //��ʕ\��
                mukouStatus = false;
                ShowDataLinq(dataGridView1);
                mukouStatus = true;

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                    linkLabel1.Visible = true;
                }
                else
                {
                    btnPrn.Enabled = false;
                    linkLabel1.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "�I��", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            finally
            {
                Cursor = Cursors.Default;
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
            object [,] rtnArray = new object[dataGridView1.RowCount, dataGridView1.ColumnCount];
            getPrnArray(dataGridView1, ref rtnArray);
            EigyoReport(dataGridView1, rtnArray);
        }

        private void getPrnArray(DataGridView tempDGV, ref object [,] rtnArray)
        {
            // �O���b�h����z��Ɏ�荞��
            for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
            {
                rtnArray[iX, 0] = tempDGV[0, iX].Value.ToString();
                rtnArray[iX, 1] = tempDGV[1, iX].Value.ToString();
                rtnArray[iX, 2] = tempDGV[2, iX].Value.ToString();
                rtnArray[iX, 3] = tempDGV[3, iX].Value.ToString();
                rtnArray[iX, 4] = tempDGV[4, iX].Value.ToString();
                rtnArray[iX, 5] = tempDGV[5, iX].Value.ToString();
                rtnArray[iX, 6] = tempDGV[6, iX].Value.ToString();
                rtnArray[iX, 7] = tempDGV[7, iX].Value.ToString();
                rtnArray[iX, 8] = tempDGV[8, iX].Value.ToString();
                rtnArray[iX, 9] = tempDGV[9, iX].Value.ToString();
                rtnArray[iX, 10] = tempDGV[10, iX].Value.ToString();
                rtnArray[iX, 11] = tempDGV[11, iX].Value.ToString();
            }
        }

        private void EigyoReport(DataGridView tempDGV, object [,] rtnArray)
        {
            const int S_GYO = 6;        // �G�N�Z���t�@�C�����׈���J�n�s
            //const int S_ROWSMAX = 8;    // �G�N�Z���t�@�C����ő�l

            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;
                
                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���N���C�A���g�ʎ󒍈ꗗ, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] aRng = new Microsoft.Office.Interop.Excel.Range[2];
                Excel.Range rng;

                try
                {
                    // �z�񂩂�V�[�g�Z���Ɉꊇ���ăf�[�^���Z�b�g���܂�
                    rng = oxlsSheet.Range[oxlsSheet.Cells[6, 1], oxlsSheet.Cells[6 + rtnArray.GetLength(0) - 1, oxlsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;

                    // ���v������
                    oxlsSheet.Cells[S_GYO - 3, 10] = int.Parse(txtUriTotal.Text, System.Globalization.NumberStyles.Any);
                    oxlsSheet.Cells[S_GYO - 3, 12] = int.Parse(txtArariTotal.Text, System.Globalization.NumberStyles.Any);

                    // ���o��
                    oxlsSheet.Cells[3, 1] = getFromToDate();

                    ////////�Z���㕔�֎������R�r��������
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�Z�������֎������R�r��������
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂Ɏ����c�r��������
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̍��[�c�r��
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̉E�[�c�r��
                    aRng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, oxlsSheet.UsedRange.Columns.Count];
                    aRng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count];
                    oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    
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
                    saveFileDialog1.Title = "�N���C�A���g�ʎ󒍈ꗗ";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    DateTime dt = DateTime.Now;
                    saveFileDialog1.FileName = "�N���C�A���g�ʎ󒍈ꗗ " + getFromToDate();
                    saveFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xlsx)|*.xlsx|�S�Ẵt�@�C��(*.*)|*.*";

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
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {
                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();

                    // COM �I�u�W�F�N�g�̎Q�ƃJ�E���g��������� 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    oXls = null;
                    oXlsBook = null;
                    oxlsSheet = null;

                    GC.Collect();

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

        private string getFromToDate()
        {
            string msg = hDate.Value.Year.ToString() + "�N" + hDate.Value.Month.ToString() + "��" + hDate.Value.Day.ToString() + "���`";
            
            if (hDate2.Checked)
            {
                msg += hDate2.Value.Year.ToString() + "�N" + hDate2.Value.Month.ToString() + "��" + hDate2.Value.Day.ToString() + "��";
            }
            
            return msg;
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.BackColor = Color.White;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (mukouStatus)
            {
                // �����ς݁A���l
                if (e.ColumnIndex == 9 || e.ColumnIndex == 11)
                {
                    // ID�擾
                    int sID = Utility.strToInt(dataGridView1["col11", e.RowIndex].Value.ToString());

                    // �f�[�^�擾
                    var s = dts.�V������.Single(a => a.ID == sID);

                    // �����σ`�F�b�N
                    if (e.ColumnIndex == 9)
                    {
                        if (dataGridView1["col10", e.RowIndex].Value.ToString() == "True")
                        {
                            s.�������� = global.FLGON;
                        }
                        else
                        {
                            s.�������� = global.FLGOFF;
                        }
                    }
                    
                    // ���l
                    if (e.ColumnIndex == 11)
                    {
                        if (dataGridView1["col12", e.RowIndex].Value == null)
                        {
                            s.���l = string.Empty;
                        }
                        else
                        {
                            s.���l = dataGridView1["col12", e.RowIndex].Value.ToString();
                        }
                    }
                    
                    s.�ύX�N���� = DateTime.Now;
                    s.���[�U�[ID = global.loginUserID;

                    // �V�������f�[�^�X�V
                    rAdp.Update(dts.�V������);

                    // �f�[�^�ǂݍ���
                    rAdp.Fill(dts.�V������);
                }
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCellAddress.X == 9)
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DateTime sDt;
            DateTime eDt;

            // �����J�n�N����
            sDt = hDate.Value.Date;

            // �����I���N����
            if (hDate2.Checked)
            {
                eDt = hDate2.Value.Date;
            }
            else
            {
                eDt = new DateTime(2900, 12, 31, 0, 0, 0);
            }

            // �󒍓��e�ʏW�v��ʕ\��
            frmClientbyOrderGrpRep frm = new posting.frmClientbyOrderGrpRep(sDt, eDt, txtsClientName.Text, getFromToDate());
            frm.ShowDialog();
        }
    }
}
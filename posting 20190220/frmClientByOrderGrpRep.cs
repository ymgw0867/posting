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
    public partial class frmClientbyOrderGrpRep : Form
    {
        const string MESSAGE_CAPTION = "�N���C�A���g�ʎ󒍓��e�ʏW�v";

        public frmClientbyOrderGrpRep(DateTime sDt, DateTime eDt, string cName, string msg)
        {
            InitializeComponent();

            // �f�[�^�ǂݍ���
            rAdp.Fill(dts.�V������);
            cAdp.Fill(dts.���Ӑ�);
            sAdp.Fill(dts.�Ј�);
            jAdp.Fill(dts.��1);
            kAdp.Fill(dts.�󒍎��);

            _sDt = sDt;
            _eDt = eDt;
            _msg = msg;
            _cName = cName;

            this.Text = "�N���C�A���g�ʎ󒍓��e�ʏW�v  " + msg;
        }

        DateTime _sDt;
        DateTime _eDt;
        string _msg = string.Empty;
        string _cName = string.Empty;

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

            // �W�v���ʕ\��
            ShowDataLinq(dataGridView1);
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
                    tempDGV.Columns.Add("col6", "������z");
                    tempDGV.Columns.Add("col7", "�e��");
                    

                    //tempDGV.Columns[1].Frozen = true;   // 2015/07/22

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 100;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

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
            try
            {
                // �O���b�h�r���[�ɕ\������
                int iX = 0;

                g.RowCount = 0;

                var j = dts.��1
                    .Where(a => a.�V������Row != null &&
                    a.���������s�� >= DateTime.Parse(_sDt.ToShortDateString()) &&
                    a.���������s�� <= DateTime.Parse(_eDt.ToShortDateString()))
                    .GroupBy(a => a.���Ӑ�ID)
                    .Select(b => new
                    {
                        sID = b.Key,
                        ss = b.GroupBy(a => a.�󒍎��ID)
                        .Select(c => new
                        {
                            orderID = c.Key,
                            sumUri = c.Sum(a => a.���z),
                            sumNebiki = c.Sum(a => a.�l���z),
                            sumGenka1 = c.Sum(a => a.�O�������x��),
                            sumGenka2 = c.Sum(a => a.�O�������x��2),
                            sumGenka3 = c.Sum(a => a.�O�������x��3)
                        })
                        .OrderBy(a => a.orderID)
                    })
                    .OrderBy(a => a.sID);

                foreach (var t in j)
                {
                    string uID = t.sID.ToString();    // ���Ӑ�ID
                    string uName = "";

                    if (dts.���Ӑ�.Any(a => a.ID == t.sID))
                    {
                        foreach (var item in dts.���Ӑ�.Where(a => a.ID == t.sID))
                        {
                            uName = item.����;
                        }
                    }

                    // �N���C�A���g���w��
                    if (_cName != string.Empty)
                    {
                        // ���Ӑ於
                        if (!uName.Contains(_cName))
                        {
                            continue;
                        }
                    }

                    foreach (var d in t.ss)
                    {
                        g.Rows.Add();

                        g[0, iX].Value = uID;    // ���Ӑ�ID
                        g[1, iX].Value = uName;

                        if (dts.�󒍎��.Any(a => a.ID == d.orderID))
                        {
                            foreach (var item in dts.�󒍎��.Where(a => a.ID == d.orderID))
                            {
                                g[2, iX].Value = item.����;
                            }
                        }

                        g[3, iX].Value = d.sumUri.ToString("#,##0");
                        g[4, iX].Value = (d.sumUri - d.sumNebiki - d.sumGenka1 - d.sumGenka2 - d.sumGenka3).ToString("#,##0");      // 2017/04/17
                        //g[4, iX].Value = (d.sumUri - d.sumGenka1 - d.sumGenka2 - d.sumGenka3).ToString("#,##0");

                        iX++;
                    }          
                }
                
                g.CurrentCell = null;

                if (g.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                }
                else
                {
                    btnPrn.Enabled = false;
                }
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
                btnPrn.Enabled = false;
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

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���N���C�A���g�ʎ󒍕ʏW�v, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] aRng = new Microsoft.Office.Interop.Excel.Range[2];
                Excel.Range rng;

                try
                {
                    // ���o��
                    oxlsSheet.Cells[3, 1] = _msg;

                    // �z�񂩂�V�[�g�Z���Ɉꊇ���ăf�[�^���Z�b�g���܂�
                    rng = oxlsSheet.Range[oxlsSheet.Cells[6, 1], oxlsSheet.Cells[6 + rtnArray.GetLength(0) - 1, oxlsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;
                    
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
                    saveFileDialog1.Title = "�N���C�A���g�ʎ󒍕ʏW�v";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    DateTime dt = DateTime.Now;
                    saveFileDialog1.FileName = this.Text;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void frmClientbyOrderGrpRep_Shown(object sender, EventArgs e)
        {
        }
    }
}
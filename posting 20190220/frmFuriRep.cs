using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmFuriRep : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "�U��\";

        public frmFuriRep()
        {
            InitializeComponent();

            // �f�[�^�Ǎ�
            adp.Fill(dts.��1);
            cAdp.Fill(dts.���Ӑ�);
            gAdp.Fill(dts.�O����);
            sAdp.Fill(dts.�Ј�);
            hAdp.Fill(dts.���^);
            fAdp.Fill(dts.�U��\);
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            gridSetting(dataGridView2);

            button1.Enabled = false;
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.��1TableAdapter adp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.���Ӑ�TableAdapter cAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
        darwinDataSetTableAdapters.�O����TableAdapter gAdp = new darwinDataSetTableAdapters.�O����TableAdapter();
        darwinDataSetTableAdapters.�Ј�TableAdapter sAdp = new darwinDataSetTableAdapters.�Ј�TableAdapter();
        darwinDataSetTableAdapters.���^TableAdapter hAdp = new darwinDataSetTableAdapters.���^TableAdapter();
        darwinDataSetTableAdapters.�U��\TableAdapter fAdp = new darwinDataSetTableAdapters.�U��\TableAdapter();
        
        /// <summary>
        /// �f�[�^�O���b�h�r���[�̒�`���s���܂�
        /// </summary>
        /// <param name="tempDGV">�f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        private void gridSetting(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // �f�[�^�t�H���g�w��
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // �s�̍���
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // �S�̂̍���
                tempDGV.Height = 500;

                // ��s�̐F
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add("col1", "�˗���");
                tempDGV.Columns.Add("col12", "�󒍔ԍ�");
                tempDGV.Columns.Add("col2", "�N���C�A���g��");
                tempDGV.Columns.Add("col3", "�`���V��");
                tempDGV.Columns.Add("col4", "�T�C�Y");
                tempDGV.Columns.Add("col5", "�󒍒P��");
                tempDGV.Columns.Add("col6", "����");
                tempDGV.Columns.Add("col7", "�\��\�̔z�z����");
                tempDGV.Columns.Add("col12", "�n������");
                tempDGV.Columns.Add("col13", "�n�����S����");
                tempDGV.Columns.Add("col8", "");
                tempDGV.Columns.Add("col9", "�U���");
                tempDGV.Columns.Add("col10", "�t������");
                tempDGV.Columns.Add("col11", "�c�ƒS��");

                tempDGV.Columns[0].Width = 90;
                tempDGV.Columns[1].Width = 130;
                tempDGV.Columns[2].Width = 200;
                tempDGV.Columns[3].Width = 300;
                tempDGV.Columns[4].Width = 100;
                tempDGV.Columns[5].Width = 100;
                tempDGV.Columns[6].Width = 100;
                tempDGV.Columns[7].Width = 160;
                tempDGV.Columns[8].Width = 90;
                tempDGV.Columns[9].Width = 110;
                tempDGV.Columns[10].Width = 60;
                tempDGV.Columns[11].Width = 200;
                tempDGV.Columns[12].Width = 100;
                tempDGV.Columns[13].Width = 100;

                //tempDGV.Columns[0].Visible = false;
                //tempDGV.Columns[9].Visible = false;
                //tempDGV.Columns[10].Visible = false;
                //tempDGV.Columns[11].Visible = false;
                //tempDGV.Columns[12].Visible = false;

                tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0.00";
                tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                tempDGV.Columns[12].DefaultCellStyle.Format = "#,##0";

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                tempDGV.Columns[2].Frozen = true;

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

        /// -----------------------------------------------------------------
        /// <summary>
        ///     �O���b�h�Ƀf�[�^��\�� </summary>
        /// <param name="g">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g </param>
        /// -----------------------------------------------------------------
        private void gridShowData_org(DataGridView g)
        {
            int iX;

            try
            {
                // �J�[�\���\����ҋ@���
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // �U��\�f�[�^�i�荞�݌���
                var s = dts.��1
                    .Where(a => !a.Is�O���˗����x��Null() && a.�󒍎��ID == 1 &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�J�n�� && a.�z�z�J�n�� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�I���� && a.�z�z�I���� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.�z�z�J�n�� <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.�z�z�I����))
                    .OrderBy(a => a.�O���˗����x��);

                // �O���b�h�r���[�ɕ\������
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.�O���˗����x��.ToShortDateString();
                    g[1, iX].Value = t.ID;

                    if (t.���Ӑ�Row == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.���Ӑ�Row.����;
                    }

                    g[3, iX].Value = t.�`���V��;

                    if (t.���^Row == null)
                    {
                        g[4, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[4, iX].Value = t.���^Row.����;
                    }

                    g[5, iX].Value = t.�P��;
                    g[6, iX].Value = t.����;

                    string sDay = string.Empty;
                    string eDay = string.Empty;

                    if (!t.Is�z�z�J�n��Null())
                    {
                        sDay = t.�z�z�J�n��.Month + "/" + t.�z�z�J�n��.Day;
                    }

                    if (!t.Is�z�z�I����Null())
                    {
                        eDay = t.�z�z�I����.Month + "/" + t.�z�z�I����.Day;
                    }

                    g[7, iX].Value = sDay + " �` " + eDay;

                    if (t.Is�O���n����Null())
                    {
                        g[8, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[8, iX].Value = t.�O���n����.ToShortDateString();
                    }

                    if (t.�O���󂯓n���S���� == null)
                    {
                        g[9, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[9, iX].Value = t.�O���󂯓n���S����;
                    }

                    g[10, iX].Value = "��";

                    if (t.�O����RowBy�O����_��1 == null)
                    {
                        g[11, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[11, iX].Value = t.�O����RowBy�O����_��1.���� + " �l";
                    }

                    g[12, iX].Value = t.�O�������x��;

                    if (t.���Ӑ�Row != null && t.���Ӑ�Row.�Ј�Row != null)
                    {
                        g[13, iX].Value = t.���Ӑ�Row.�Ј�Row.����;
                    }
                    else
                    {
                        g[13, iX].Value = string.Empty;
                    }

                    iX++;
                }

                // �O����ʖ����v
                var ss = dts.��1
                    .Where(a => !a.Is�O���˗����x��Null() && a.�󒍎��ID == 1 &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�J�n�� && a.�z�z�J�n�� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�I���� && a.�z�z�I���� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.�z�z�J�n�� <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.�z�z�I����))
                    .GroupBy(a => new { a.�O����ID�x��, a.�O����RowBy�O����_��1.���� })
                    .Select(gg => new
                    {
                        gg.Key.�O����ID�x��,
                        gg.Key.����,
                        ���v = gg.Sum(a => a.����)
                    });

                decimal ttl = 0;
                foreach (var tt in ss)
                {
                    g.Rows.Add();
                    g[11, iX].Value = tt.���� + " �l";
                    g[12, iX].Value = tt.���v.ToString("#,##0");
                    ttl += tt.���v;
                    iX++;
                }

                // �����v
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[11, iX].Value = "���v";
                    g[12, iX].Value = ttl.ToString("#,##0");

                    // �J�����g�J�[�\��
                    dataGridView2.CurrentCell = null;
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


        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     �O���b�h�Ƀf�[�^��\�� 2017/01/30 �u�U��\�v�e�[�u���Z�b�g </summary>
        /// <param name="g">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g </param>
        /// ---------------------------------------------------------------------------
        private void gridShowData(DataGridView g)
        {
            int iX;

            try
            {
                // �J�[�\���\����ҋ@���
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // �U��\�f�[�^�i�荞�݌���
                var s = dts.�U��\
                    .Where(a => !a.Is�O���˗����x��Null() &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�J�n�� && a.�z�z�J�n�� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�I���� && a.�z�z�I���� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.�z�z�J�n�� <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.�z�z�I����))
                    .OrderBy(a => a.�O���˗����x��);

                // �O���b�h�r���[�ɕ\������
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.�O���˗����x��.ToShortDateString();
                    g[1, iX].Value = t.ID;

                    if (t.���Ӑ�Row == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.���Ӑ�Row.����;
                    }

                    g[3, iX].Value = t.�`���V��;

                    if (t.���^Row == null)
                    {
                        g[4, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[4, iX].Value = t.���^Row.����;
                    }

                    g[5, iX].Value = t.�P��;
                    g[6, iX].Value = t.����;

                    string sDay = string.Empty;
                    string eDay = string.Empty;

                    if (!t.Is�z�z�J�n��Null())
                    {
                        sDay = t.�z�z�J�n��.Month + "/" + t.�z�z�J�n��.Day;
                    }

                    if (!t.Is�z�z�I����Null())
                    {
                        eDay = t.�z�z�I����.Month + "/" + t.�z�z�I����.Day;
                    }

                    g[7, iX].Value = sDay + " �` " + eDay;

                    if (t.Is�O���n����Null())
                    {
                        g[8, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[8, iX].Value = t.�O���n����.ToShortDateString();
                    }

                    if (t.�O���󂯓n���S���� == null)
                    {
                        g[9, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[9, iX].Value = t.�O���󂯓n���S����;
                    }

                    g[10, iX].Value = "��";

                    if (t.�O����Row == null)
                    {
                        g[11, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[11, iX].Value = t.�O����Row.���� + " �l";
                    }

                    g[12, iX].Value = t.�O�������x��;

                    if (t.���Ӑ�Row != null && t.���Ӑ�Row.�Ј�Row != null)
                    {
                        g[13, iX].Value = t.���Ӑ�Row.�Ј�Row.����;
                    }
                    else
                    {
                        g[13, iX].Value = string.Empty;
                    }

                    iX++;
                }

                // �O����ʖ����v
                var ss = dts.�U��\
                    .Where(a => !a.Is�O���˗����x��Null() &&
                    (DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�J�n�� && a.�z�z�J�n�� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     DateTime.Parse(iraiDtS.Value.ToShortDateString()) <= a.�z�z�I���� && a.�z�z�I���� <= DateTime.Parse(iraiDtE.Value.ToShortDateString()) ||
                     a.�z�z�J�n�� <= DateTime.Parse(iraiDtS.Value.ToShortDateString()) && DateTime.Parse(iraiDtE.Value.ToShortDateString()) <= a.�z�z�I����))
                    .GroupBy(a => new { a.�O����ID�x��, a.�O����Row.���� })
                    .Select(gg => new
                    {
                        gg.Key.�O����ID�x��,
                        gg.Key.����,
                        ���v = gg.Sum(a => a.����)
                    });

                decimal ttl = 0;
                foreach (var tt in ss)
                {
                    g.Rows.Add();
                    g[11, iX].Value = tt.���� + " �l";
                    g[12, iX].Value = tt.���v.ToString("#,##0");
                    ttl += tt.���v;
                    iX++;
                }

                // �����v
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[11, iX].Value = "���v";
                    g[12, iX].Value = ttl.ToString("#,##0");

                    // �J�����g�J�[�\��
                    dataGridView2.CurrentCell = null;
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

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�I�����܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // �z�z���͈̓`�F�b�N
            if (DateTime.Parse(iraiDtS.Value.ToShortDateString()) > DateTime.Parse(iraiDtE.Value.ToShortDateString()))
            {
                MessageBox.Show("�z�z���͈͂�����������܂���", "�z�z���͈̓G���[", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // �f�[�^�\��
            gridShowData(dataGridView2);
            
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
            if (MessageBox.Show("�U��\�𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            KanryoReport(dataGridView2);
        }

        private void KanryoReport(DataGridView g)
        {
            const int S_GYO = 4;        // �G�N�Z���t�@�C�����o���s�i���ׂ�4�s�ڂ���󎚁j
            const int S_ROWSMAX = 14;   // �G�N�Z���t�@�C����ő�l
            int r = S_GYO;              // �������s
            bool tl = true;

            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���U��\, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    // �Ώ۔z�z���͈�
                    string hs = string.Empty;
                    string he = string.Empty;

                    if (iraiDtS.Checked)
                    {
                        hs = iraiDtS.Value.ToShortDateString();
                    }

                    if (iraiDtE.Checked)
                    {
                        he = iraiDtE.Value.ToShortDateString();
                    }

                    oxlsSheet.Cells[1, 3] = "�z�z���F " + hs + " �` " + he;

                    // ����
                    //for (int i = 0; i < 10; i++)
                    //{
                        for (int iX = 0; iX <= g.RowCount - 1; iX++)
                        {
                            if (g[0, iX].Value != null)
                            {
                                // 1�s��
                                oxlsSheet.Cells[r, 1] = " ";

                                //�Z�������֎������R�r��������
                                rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                                rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                r++;

                                // ����
                                oxlsSheet.Cells[r, 1] = g[0, iX].Value.ToString();     // �˗���
                                oxlsSheet.Cells[r, 2] = g[1, iX].Value.ToString();     // �󒍔ԍ�
                                oxlsSheet.Cells[r, 3] = g[2, iX].Value.ToString();     // �N���C�A���g��
                                oxlsSheet.Cells[r, 4] = g[3, iX].Value.ToString();     // �`���V��
                                oxlsSheet.Cells[r, 5] = g[4, iX].Value.ToString();     // �T�C�Y
                                oxlsSheet.Cells[r, 6] = g[5, iX].Value.ToString();     // �󒍒P��
                                oxlsSheet.Cells[r, 7] = g[6, iX].Value.ToString();     // ����
                                oxlsSheet.Cells[r, 8] = g[7, iX].Value.ToString();     // �\��\�̔z�z����
                                oxlsSheet.Cells[r, 9] = g[8, iX].Value.ToString();     // �n������
                                oxlsSheet.Cells[r, 10] = g[9, iX].Value.ToString();     // �n�����S����
                                oxlsSheet.Cells[r, 11] = g[10, iX].Value.ToString();     // 
                                oxlsSheet.Cells[r, 12] = g[11, iX].Value.ToString();    // �U���
                                oxlsSheet.Cells[r, 13] = g[12, iX].Value.ToString();   // �U��P��
                                oxlsSheet.Cells[r, 14] = g[13, iX].Value.ToString();   // �c�ƒS��
                            }
                            else
                            {
                                if (tl)
                                {
                                    // 1�s��
                                    oxlsSheet.Cells[r, 1] = " ";

                                    //�Z�������֎������R�r��������
                                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    r++;
                                }

                                // �U���ʍ��v
                                oxlsSheet.Cells[r, 12] = g[11, iX].Value.ToString();    // �U���
                                oxlsSheet.Cells[r, 13] = g[12, iX].Value.ToString();   // �U�荇�v
                                tl = false;
                            }

                            //�Z�������֎������R�r��������
                            rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                            r++;
                        }
                    //}
                    
                    //////�Z���㕔�֎������R�r��������
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    ////�Z�������֎������R�r��������
                    //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    // �\�S�̂Ɏ����c�r��������
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    // �\�S�̂̍��[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    // �\�S�̂̉E�[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //// �O����ʁF�x���z
                    //var s = dts.��1
                    //    .Where(a => !a.Is�O���˗����x��Null())
                    //    .GroupBy(a => new { a.�O����ID�x��, a.�O����Row.���� })
                    //    .Select(gg => new
                    //    {
                    //        gg.Key.�O����ID�x��,
                    //        gg.Key.����,
                    //        ���v = gg.Sum(a => (decimal)a.���� * a.�O�������x��)
                    //    });

                    //r++;
                    //foreach (var t in s)
                    //{
                    //    oxlsSheet.Cells[r, 10] = t.����;    // �U���
                    //    oxlsSheet.Cells[r, 11] = t.���v;   // �U�荇�v
                    //    r++;
                    //}

                    // �}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;

                    // �m�F�̂���Excel�̃E�B���h�E��\������
                    oXls.Visible = true;

                    // ���
                    oxlsSheet.PrintPreview(true);

                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    // �ۑ�����
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    // �_�C�A���O�{�b�N�X�̏����ݒ�
                    saveFileDialog1.Title = MESSAGE_CAPTION;
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = MESSAGE_CAPTION;
                    saveFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xls)|*.xls|�S�Ẵt�@�C��(*.*)|*.*";

                    // �_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
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

                    // Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excel���I��
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    // Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excel���I��
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
    }
}
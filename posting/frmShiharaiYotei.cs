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
using MyLibrary;

namespace posting
{
    public partial class frmShiharaiYotei : Form
    {
        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "�O���x���\��\";

        public frmShiharaiYotei()
        {
            InitializeComponent();

            // �f�[�^�Ǎ�
            adp.Fill(dts.��1);
            sAdp.Fill(dts.�O���x��);
            gAdp.Fill(dts.�O����);
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            gridSettingShiharai(dataGridView2);
            gridSetting(dataGridView1);

            // �N��
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            //button1.Enabled = false;
            linkLabel1.Visible = false;
            linkLabel2.Visible = false;

            // �x���\��\��
            showShiharaiYotei();
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.��1TableAdapter adp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.�O���x��TableAdapter sAdp = new darwinDataSetTableAdapters.�O���x��TableAdapter();
        darwinDataSetTableAdapters.�O����TableAdapter gAdp = new darwinDataSetTableAdapters.�O����TableAdapter();

        Utility.����ŗ� cTax = new Utility.����ŗ�();

        ///--------------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̒�`���s���܂� </summary>
        /// <param name="tempDGV">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        ///--------------------------------------------------------------------
        private void gridSettingShiharai(DataGridView tempDGV)
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
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // �s�̍���
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // �S�̂̍���
                tempDGV.Height = 222;

                // ��s�̐F
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add("col1", "�\���");
                tempDGV.Columns.Add("col2", "�R�[�h");
                tempDGV.Columns.Add("col3", "�x���於");
                tempDGV.Columns.Add("col5", "�x���z");
                //tempDGV.Columns.Add("col6", "����Ŋz");
                //tempDGV.Columns.Add("col7", "�ō����z");

                tempDGV.Columns[0].Width = 110;
                tempDGV.Columns[1].Width = 60;
                tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[3].Width = 120;
                //tempDGV.Columns[4].Width = 120;
                //tempDGV.Columns[5].Width = 120;

                tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                //tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                //tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                //tempDGV.Columns[2].Frozen = true;

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

        ///--------------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̒�`���s���܂� </summary>
        /// <param name="tempDGV">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        ///--------------------------------------------------------------------
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
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // �f�[�^�t�H���g�w��
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // �s�̍���
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // �S�̂̍���
                tempDGV.Height = 222;

                // ��s�̐F
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add("col1", "�\���");
                tempDGV.Columns.Add("col2", "�R�[�h");
                tempDGV.Columns.Add("col3", "�x���於");
                tempDGV.Columns.Add("col4", "���e");
                tempDGV.Columns.Add("col5", "�x���z");
                //tempDGV.Columns.Add("col7", "����Ŋz");  // 2016/02/01
                //tempDGV.Columns.Add("col8", "�ō����z");  // 2016/02/01
                tempDGV.Columns.Add("col6", "�x�����{��");
                
                tempDGV.Columns[0].Width = 110;
                tempDGV.Columns[1].Width = 60;
                tempDGV.Columns[2].Width = 180;
                tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[4].Width = 120;

                //tempDGV.Columns[5].Width = 90;    // 2016/02/01
                //tempDGV.Columns[6].Width = 120;   // 2016/02/01
                
                tempDGV.Columns[5].Width = 110;
                
                tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                //tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";  // 2016/02/01
                //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";  // 2016/02/01

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                //tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;     // 2016/02/01
                //tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;     // 2016/02/01
                
                tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
        private void gridShowShiharai(DataGridView g)
        {
            int iX;
            decimal tl = 0;
            decimal tlTax = 0;
            decimal tlZeikomi = 0;

            try
            {
                //�J�[�\���\����ҋ@���
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // �x���\��\�f�[�^�����F2015/11/19 a.�O����Row != null �������ɒǉ�
                var s = dts.��1
                    .Where(a => a.�O����RowBy�O����_��1 != null &&
                           !a.Is�O���x�����x��Null() &&
                                 a.�O���x�����x��.Year == int.Parse(txtYear.Text) && a.�O���x�����x��.Month == int.Parse(txtMonth.Text))
                    .GroupBy(a => new { a.�O���x�����x��, a.�O����ID�x��, a.�O����RowBy�O����_��1.���� })
                    .Select(gg => new
                    {
                        gg.Key.�O���x�����x��,
                        gg.Key.�O����ID�x��,
                        gg.Key.����,
                        //�x���z = gg.Sum(a => a.�O�������x�� * a.����)

                        // 2015/12/06 �O���������u�P�����猴�����z���͂֕ύX�v�ɔ���
                        �x���z = gg.Sum(a => a.�O�������x��)
                    })
                    .OrderBy(a => a.�O���x�����x��).ThenBy(a => a.�O����ID�x��);

                // �O���b�h�r���[�ɕ\������
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.�O���x�����x��.ToShortDateString();
                    g[1, iX].Value = t.�O����ID�x��;

                    if (t.���� == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.����;
                    }

                    g[3, iX].Value = t.�x���z.ToString("#,##0");

                    // ����Ŋz�\���p�~�@2016/02/01
                    //// ����Ŋz�擾
                    //decimal KingakuTax = 0;
                    //foreach (var it in dts.��1.Where(a => !a.Is�O���x�����x��Null() && a.�O���x�����x�� == t.�O���x�����x�� && a.�O����ID�x�� == t.�O����ID�x��))
                    //{
                    //    //�ŗ��Ď擾
                    //    cTax.Ritsu = Utility.GetTaxRT(it.�󒍓�);

                    //    //����Ŋz�v�Z 
                    //    //KingakuTax += Utility.GetTax(((decimal)it.���� * it.�O�������x��), cTax.Ritsu);
                    //    // 2015/12/06 �O���������u�P�����猴�����z���͂֕ύX�v�ɔ���
                    //    KingakuTax += Utility.GetTax(it.�O�������x��, cTax.Ritsu);
                    //}

                    // ����Ŋz�\���p�~�@2016/02/01
                    //// �����
                    //g[4, iX].Value = KingakuTax.ToString("#,##0");

                    // ����Ŋz�\���p�~�@2016/02/01
                    //// �ō����z
                    //g[5, iX].Value = (t.�x���z + KingakuTax).ToString("#,##0");

                    iX++;

                    tl += t.�x���z;

                    // ����Ŋz�\���p�~�@2016/02/01
                    //tlTax += KingakuTax;
                    //tlZeikomi += (t.�x���z + KingakuTax);
                }

                // �����v
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[2, iX].Value = "���v";
                    g[3, iX].Value = tl.ToString("#,##0");

                    //g[4, iX].Value = tlTax.ToString("#,##0");     // 2016/02/01
                    //g[5, iX].Value = tlZeikomi.ToString("#,##0"); // 2016/02/01

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

        /// -----------------------------------------------------------------
        /// <summary>
        ///     �O���b�h�Ƀf�[�^��\�� </summary>
        /// <param name="g">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g </param>
        /// -----------------------------------------------------------------
        private void gridShowData(DataGridView g, DateTime sDt, int sID)
        {
            int iX;
            decimal tl = 0;
            decimal tlTax = 0;
            decimal tlZeikomi = 0;

            try
            {
                //�J�[�\���\����ҋ@���
                Cursor.Current = Cursors.WaitCursor;
                g.RowCount = 0;

                // �w��̊O����̎x���\��f�[�^����
                var s = dts.��1
                    .Where(a => !a.Is�O���x�����x��Null() && a.�O���x�����x�� == sDt && a.�O����ID�x�� == sID)
                    .OrderBy(a => a.�O���x�����x��);

                // �O���b�h�r���[�ɕ\������
                iX = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();

                    g[0, iX].Value = t.�O���x�����x��.ToShortDateString();
                    g[1, iX].Value = t.�O����ID�x��;

                    if (t.�O����RowBy�O����_��1 == null)
                    {
                        g[2, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[2, iX].Value = t.�O����RowBy�O����_��1.����;
                    }

                    g[3, iX].Value = t.�`���V��;

                    //decimal kingaku = (decimal)t.���� * t.�O�������x��;

                    // 2015/12/06 �O���������u�P�����猴�����z���͂֕ύX�v�ɔ���
                    decimal kingaku = t.�O�������x��;
                    g[4, iX].Value = kingaku.ToString("#,##0");
                    
                    //�ŗ��Ď擾
                    cTax.Ritsu = Utility.GetTaxRT(t.�󒍓�);

                    // ����Ŋz�\���p�~�@2016/02/01
                    ////����Ŋz�v�Z 
                    //decimal KingakuTax = Utility.GetTax(kingaku, cTax.Ritsu);

                    // ����Ŋz�\���p�~�@2016/02/01
                    //// �����
                    //g[5, iX].Value = KingakuTax.ToString("#,##0");

                    // ����Ŋz�\���p�~�@2016/02/01
                    //// �ō����z
                    //g[6, iX].Value = (kingaku + KingakuTax).ToString("#,##0");

                    // �O���x����
                    if (t.�O���x��Row == null)
                    {
                        g[5, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[5, iX].Value = t.�O���x��Row.���t.ToShortDateString();
                    }

                    iX++;

                    tl += kingaku;

                    // ����Ŋz�\���p�~�@2016/02/01
                    //tlTax += KingakuTax;
                    //tlZeikomi += (kingaku + KingakuTax);
                }

                // �����v
                if (g.Rows.Count > 0)
                {
                    g.Rows.Add();
                    g[2, iX].Value = "���v";
                    g[4, iX].Value = tl.ToString("#,##0");

                    // 2016/02/01
                    //g[5, iX].Value = tlTax.ToString("#,##0");
                    //g[6, iX].Value = tlZeikomi.ToString("#,##0");

                    // �J�����g�J�[�\��
                    dataGridView1.CurrentCell = null;

                    // CSV
                    linkLabel2.Visible = true;
                }
                else
                {
                    linkLabel2.Visible = false;
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
            // �N���`�F�b�N
            if (Utility.strToInt(txtMonth.Text) < 1 || Utility.strToInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("�w�茎������������܂���", "�w�茎�G���[", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            // �f�[�^�\��
            gridShowShiharai(dataGridView2);

            // ���׃O���b�h�N���A
            dataGridView1.Rows.Clear();
            
            // �Y���f�[�^�Ȃ�
            if (dataGridView2.RowCount == 0)
            {
                label5.Text = "�Y���f�[�^������܂���ł���";
                label5.Visible = true;
                //button1.Enabled = false;
            }
            else
            {
                label5.Text = "";
                label5.Visible = false;
                //button1.Enabled = true;
            }
        }

        private void showShiharaiYotei()
        {
            // �f�[�^�\��
            gridShowShiharai(dataGridView2);

            // ���׃O���b�h�N���A
            dataGridView1.Rows.Clear();
            linkLabel2.Visible = false;

            // �Y���f�[�^�Ȃ�
            if (dataGridView2.RowCount == 0)
            {
                label5.Text = "�Y���f�[�^������܂���ł���";
                label5.Visible = true;
                //button1.Enabled = false;
                linkLabel1.Visible = false;
            }
            else
            {
                label5.Text = "";
                label5.Visible = false;
                //button1.Enabled = true;
                linkLabel1.Visible = true;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int sID = 0;
            DateTime sDt = DateTime.Parse("1900/01/01");

            if (dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value == null)
            {
                return;
            }
            else
            {
                // �O����R�[�h�擾
                sID = Utility.strToInt(dataGridView2[1, dataGridView2.SelectedRows[0].Index].Value.ToString());
            }

            if (dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value == null)
            {
                return;
            }
            else
            {
                // �x�����擾
                if (!DateTime.TryParse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString(), out sDt))
                {
                    return;
                }
            }

            // �w��O���斾�׃f�[�^�\��
            gridShowData(dataGridView1, sDt, sID);
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void txtMonth_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // �O��
            yearMonthBefore();
            showShiharaiYotei();
        }

        private void yearMonthBefore()
        {
            int m = Utility.strToInt(txtMonth.Text);
            m--;

            if (m == 0)
            {
                txtYear.Text = (Utility.strToInt(txtYear.Text) - 1).ToString();
                txtMonth.Text = "12";
            }
            else
            {
                txtMonth.Text = m.ToString();
            } 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // ����
            yearMonthAfter();
            showShiharaiYotei();
        }

        private void yearMonthAfter()
        {
            int m = Utility.strToInt(txtMonth.Text);
            m++;

            if (m > 12)
            {
                txtYear.Text = (Utility.strToInt(txtYear.Text) + 1).ToString();
                txtMonth.Text = (m - 12).ToString();
            }
            else
            {
                txtMonth.Text = m.ToString();
            }
        }

        private void txtMonth_TextChanged(object sender, EventArgs e)
        {
            yearMonthChange();
        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            yearMonthChange();
        }

        private void yearMonthChange()
        {
            // �N���`�F�b�N
            if (Utility.strToInt(txtYear.Text) < 2014)
            {
                txtYear.Focus();
                return;
            }

            if (Utility.strToInt(txtMonth.Text) < 1 || Utility.strToInt(txtMonth.Text) > 12)
            {
                txtMonth.Focus();
                return;
            }

            // �x���\��\��
            showShiharaiYotei();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, MESSAGE_CAPTION);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION + "����");
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmShowTable : Form
    {
        const string MESSAGE_CAPTION = "�e�[�u���f�[�^�\��";
        const string TAB_SEIKYUSHO = "������";
        const string TAB_NYUKIN = "����";
        const string TAB_JYUCHU = "��";
        const string TAB_HAIFUAREA = "�z�z�G���A";
        const string TAB_HAIFUSHIJI = "�z�z�w��";

        public frmShowTable()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            this.tabControl1.TabPages[0].Text = TAB_SEIKYUSHO;
            this.tabControl1.TabPages[1].Text = TAB_NYUKIN;
            this.tabControl1.TabPages[2].Text = TAB_JYUCHU;
            this.tabControl1.TabPages[3].Text = TAB_HAIFUAREA;
            this.tabControl1.TabPages[4].Text = TAB_HAIFUSHIJI;

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.�z�z�w��' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            this.�z�z�w��TableAdapter.Fill(this.darwinDataSet.�z�z�w��);
            
            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.�z�z�G���A' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            this.�z�z�G���ATableAdapter.Fill(this.darwinDataSet.�z�z�G���A);
            
            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.����' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            this.����TableAdapter.Fill(this.darwinDataSet.����);

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.������' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            this.������TableAdapter.Fill(this.darwinDataSet.������);

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView1);
            GridviewSet.Setting2(dataGridView2);
            GridviewSet.Setting3(dataGridView3);
            GridviewSet.Setting4(dataGridView4);
            GridviewSet.Setting5(dataGridView5);

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
                    tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", 9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 613;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    //tempDGV.Columns.Add("col1", "����");
                    //tempDGV.Columns.Add("col2", "����");
                    //tempDGV.Columns.Add("col3", "���l");

                    //tempDGV.Columns[0].Width = 60;
                    //tempDGV.Columns[1].Width = 200;
                    //tempDGV.Columns[2].Width = 100;
                    ////tempDGV.Columns[3].Width = 160;
                    //tempDGV.Columns[4].Width = 100;
                    //tempDGV.Columns[5].Width = 100;

                    //tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
                    tempDGV.Height = 613;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

            public static void Setting3(DataGridView tempDGV)
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
                    tempDGV.Height = 613;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns[6].Width = 200;
                    tempDGV.Columns[29].Width = 110;
                    tempDGV.Columns[35].Width = 110;
                    tempDGV.Columns[37].Width = 120;

                    // �e��z�u
                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[18].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[20].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[21].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[22].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[23].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[24].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[25].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[26].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[27].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[28].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[29].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[30].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[31].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[32].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[33].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[34].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[35].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    tempDGV.Columns[36].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[37].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[38].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[39].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[40].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[41].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[42].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[43].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

            public static void Setting4(DataGridView tempDGV)
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
                    tempDGV.Height = 613;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

            public static void Setting5(DataGridView tempDGV)
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
                    tempDGV.Height = 613;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns[5].Width = 110;
                    tempDGV.Columns[6].Width = 110;
                    tempDGV.Columns[7].Width = 110;
                    tempDGV.Columns[8].Width = 110;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
         }

        /// <summary>
        /// ��ʂ��N���A����
        /// </summary>
        private void DispClear()
        {

            try
            {
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }

        private void btnCsv_Click(object sender, EventArgs e)
        {
            string iX;

            iX = tabControl1.SelectedIndex.ToString();

            switch (iX)
            {
                case "0":
                    dataGridViewCSV = dataGridView1;
                    break;

                case "1":
                    dataGridViewCSV = dataGridView2;
                    break;

                case "2":
                    dataGridViewCSV = dataGridView3;
                    break;

                case "3":
                    dataGridViewCSV = dataGridView4;
                    break;

                case "4":
                    dataGridViewCSV = dataGridView5;
                    break;

                default:
                    break;
            }

            MyLibrary.CsvOut.GridView(dataGridViewCSV, tabControl1.SelectedTab.Text);

        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0) dataGridView1.CurrentCell = null;
            if (tabControl1.SelectedIndex == 1) dataGridView2.CurrentCell = null;
            if (tabControl1.SelectedIndex == 2) dataGridView3.CurrentCell = null;
            if (tabControl1.SelectedIndex == 3) dataGridView4.CurrentCell = null;
            if (tabControl1.SelectedIndex == 4) dataGridView5.CurrentCell = null;
        }

    }
}
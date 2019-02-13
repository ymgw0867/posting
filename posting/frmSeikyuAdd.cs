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
    public partial class frmSeikyuAdd : Form
    {
        Entity.������ cMaster = new Entity.������();
        Entity.���� cNyukin = new Entity.����();

        Utility.formMode fMode = new Utility.formMode();

        const string MESSAGE_CAPTION = "�����f�[�^�쐬";

        private DateTime sD;
        private DateTime eD;

        public frmSeikyuAdd(int pMode)
        {
            
            InitializeComponent();

            //�������[�h�@0:�o�^�A1:�X�V
            this.fMode.Mode  = pMode;
        }

        private void frmSeikyuAdd_Load(object sender, EventArgs e)
        {

            sD = Convert.ToDateTime("1900/01/01");
            eD = Convert.ToDateTime("9999/12/31");

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            //GridviewSet.Setting(dataGridView2);

            GridviewSet.Setting2(dataGridView1);
            GridviewSet.Setting3(dataGridView3,fMode.Mode);

            switch (fMode.Mode)
            {
                case 0: //�V�K�o�^
                    this.Text = "�������f�[�^�o�^";
                    label15.Text = "�y�������N���C�A���g�z" + Environment.NewLine + "�z�z�I�����F�N���C�A���g��";

                    ListClient.load(listBox1,"",sD,eD);

                    //�����I�u�W�F�N�g��\��
                    panel1.Hide();
                    label20.Hide();
                    label23.Hide();
                    label25.Hide();
                    txtZan.Hide();
                    checkBox1.Hide();
                    dataGridView4.Hide();

                    break;

                case 1: //�ҏW
                    this.Text = "�������f�[�^�ҏW";
                    label15.Text = "�y�������f�[�^�z" + Environment.NewLine + "�z�z�I�����F�N���C�A���g��";
                    radioButton1.Checked = true;
                    radioButton2.Checked = false;

                    ListClient.loadData(listBox1, getRbtn(), "",sD,eD);

                    //�����I�u�W�F�N�g�\��
                    panel1.Show();
                    label20.Show();
                    label23.Show();
                    label25.Show();
                    txtZan.Show();
                    checkBox1.Show();
                    dataGridView4.Show();

                    GridviewSet.Setting4(dataGridView4);
                    
                    break;

            }

            DispClear();

            label1.Text = "�y�󒍃f�[�^�z";
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
                    tempDGV.Height = 145;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "����");
                    tempDGV.Columns.Add("col2", "����");
                    tempDGV.Columns.Add("col3", "�S����");
                    tempDGV.Columns.Add("col4", "��");
                    tempDGV.Columns.Add("col5", "�Z��");
                    tempDGV.Columns.Add("col6", "���S����");
                    tempDGV.Columns.Add("col7", "TEL");
                    tempDGV.Columns.Add("col8", "FAX");

                    tempDGV.Columns[0].Width = 40;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 300;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 80;
                    //tempDGV.Columns[8].Width = 80;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    //tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";

                    //tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    //tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

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
                    tempDGV.Height = 145;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col2", "�󒍓�");
                    tempDGV.Columns.Add("col3", "�`���V��");
                    tempDGV.Columns.Add("col4", "�󒍎��");
                    tempDGV.Columns.Add("col5", "�P��");
                    tempDGV.Columns.Add("col6", "����");
                    tempDGV.Columns.Add("col7", "������z");
                    tempDGV.Columns.Add("col8", "�����");
                    tempDGV.Columns.Add("col9", "�ō����z");
                    tempDGV.Columns.Add("col10", "�l���z");
                    tempDGV.Columns.Add("col11", "�������z");
                    tempDGV.Columns.Add("col12", "�T�C�Y");
                    tempDGV.Columns.Add("col13", "�z�z�J�n��");
                    tempDGV.Columns.Add("col14", "�z�z�I����");
                    tempDGV.Columns.Add("col15", "�����\���");
                    tempDGV.Columns.Add("col16", "�z�z�`��");
                    tempDGV.Columns.Add("col17", "�ŗ�");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;
                    tempDGV.Columns[11].Width = 90;
                    tempDGV.Columns[12].Width = 110;
                    tempDGV.Columns[13].Width = 110;
                    tempDGV.Columns[14].Width = 110;
                    tempDGV.Columns[15].Width = 120;
                    tempDGV.Columns[16].Width = 60;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void Setting3(DataGridView tempDGV,int tempMode)
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
                    tempDGV.Height = 253;

                    switch (tempMode)
                    {
                        case 0:
                            tempDGV.Width = 970;
                            break;

                        case 1:
                            tempDGV.Width = 740;
                            break;
                    }

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�z�z��");
                    tempDGV.Columns.Add("col2", "�z�z��");
                    tempDGV.Columns.Add("col3", "�`��");
                    tempDGV.Columns.Add("col4", "�`���V��");
                    tempDGV.Columns.Add("col5", "�P��");
                    tempDGV.Columns.Add("col6", "����");
                    tempDGV.Columns.Add("col7", "������z");
                    tempDGV.Columns.Add("col8", "�����");
                    tempDGV.Columns.Add("col9", "�ō����z");
                    tempDGV.Columns.Add("col10", "�l���z");
                    tempDGV.Columns.Add("col11", "�������z");
                    tempDGV.Columns.Add("col12", "��ID");
                    tempDGV.Columns.Add("col13", "�ŗ�");

                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 100;
                    tempDGV.Columns[2].Width = 120;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 80;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 90;
                    tempDGV.Columns[8].Width = 90;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;

                    tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

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
                    tempDGV.Height = 73;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "ID");
                    tempDGV.Columns.Add("col2", "������");
                    tempDGV.Columns.Add("col3", "���z");

                    tempDGV.Columns[0].Visible = false;

                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 137;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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
                    tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// �����Ώۓ��Ӑ�\��
            /// </summary>
            /// <param name="tempDGV"></param>
            public static void ShowData(DataGridView tempDGV)
            {

                string sqlSTRING = "";
                int iX;

                try
                {
                    tempDGV.RowCount = 0;
                    
                    //�f�[�^���[�_�[���擾����
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select DISTINCT ��.���Ӑ�ID, ���Ӑ�.*,�Ј�.���� ";
                    sqlSTRING += "from (�� left join ���Ӑ� ";
                    sqlSTRING += "on ��.���Ӑ�ID = ���Ӑ�.ID) left join �Ј� ";
                    sqlSTRING += "on ���Ӑ�.�S���Ј��R�[�h = �Ј�.ID ";
                    sqlSTRING += "where  (��.������ = 1) AND (��.������ID = 0) ";
                    sqlSTRING += "order by ��.���Ӑ�ID";
                
                    //���������Ӑ�̃f�[�^���[�_�[���擾����
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);
                    
                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = long.Parse(dR["���Ӑ�ID"].ToString());
                            tempDGV[1, iX].Value = dR["����"].ToString() + "";
                            tempDGV[2, iX].Value = dR["����"].ToString();
                            tempDGV[3, iX].Value = dR["������X�֔ԍ�"].ToString() + "";
                            tempDGV[4, iX].Value = dR["������Z��1"].ToString() + dR["������Z��2"].ToString() + "";
                            tempDGV[5, iX].Value = dR["�S���Җ�"].ToString();
                            tempDGV[6, iX].Value = dR["�d�b�ԍ�"].ToString();
                            tempDGV[7, iX].Value = dR["FAX�ԍ�"].ToString();

                            iX++;

                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();

                    fCon.Close();

                    //if (tempDGV.RowCount <= 12)
                    //{
                    //    tempDGV.Columns[3].Width = 97;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[3].Width = 80;
                    //}

                }
                catch (Exception e)
                {
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

        private void ShowPosting(DataGridView tempDGV,long tempID)
        {

            string mySql = "";
            OleDbDataReader dR;
            int iX = 0;

            mySql += "select ��.*, �󒍎��.���� as �󒍎�ʖ���,���^.���� as ���^����,�z�z�`��.���� as �z�z�`�Ԗ��� ";
            mySql += "from ((�� left join �󒍎�� ";
            mySql += "on ��.�󒍎��ID = �󒍎��.ID) left join ���^ ";
            mySql += "on ��.���^ = ���^.ID) left join �z�z�`�� ";
            mySql += "on ��.�z�z�`�� = �z�z�`��.ID ";
            mySql += "where ";
            mySql += "(��.���Ӑ�ID = " + tempID.ToString() + ") and ";
            mySql += "(��.������ = 1) and ";
            mySql += "(��.������ID = 0) ";
            mySql += "order by ��.ID desc";

            Control.FreeSql fCon = new Control.FreeSql();
            dR = fCon.free_dsReader(mySql);

            tempDGV.RowCount = 0;
            button1.Enabled = false;

            while (dR.Read())
            {
                tempDGV.Rows.Add();
                tempDGV[0, iX].Value = dR["ID"].ToString();
                tempDGV[1, iX].Value = DateTime.Parse(dR["�󒍓�"].ToString()).ToShortDateString();
                tempDGV[2, iX].Value = dR["�`���V��"].ToString();
                
                if (dR["�󒍎�ʖ���"] == DBNull.Value)
                {
                    tempDGV[3, iX].Value = "";
                }
                else
                {
                    tempDGV[3, iX].Value = dR["�󒍎�ʖ���"].ToString();
                }

                tempDGV[4, iX].Value = double.Parse(dR["�P��"].ToString());
                tempDGV[5, iX].Value = int.Parse(dR["����"].ToString());
                tempDGV[6, iX].Value = double.Parse(dR["���z"].ToString());
                tempDGV[7, iX].Value = double.Parse(dR["�����"].ToString());
                tempDGV[8, iX].Value = double.Parse(dR["�ō����z"].ToString());
                tempDGV[9, iX].Value = double.Parse(dR["�l���z"].ToString());
                tempDGV[10, iX].Value = double.Parse(dR["������z"].ToString());

                if (dR["���^����"] == DBNull.Value)
                {
                    tempDGV[11, iX].Value = "";
                }
                else
                {
                    tempDGV[11, iX].Value = dR["���^����"].ToString();
                }

                if (dR["�z�z�J�n��"] == DBNull.Value)
                {
                    tempDGV[12, iX].Value = "";
                }
                else
                {
                    tempDGV[12, iX].Value = DateTime.Parse(dR["�z�z�J�n��"].ToString()).ToShortDateString();
                }

                if (dR["�z�z�I����"] == DBNull.Value)
                {
                    tempDGV[13, iX].Value = "";
                }
                else
                {
                    tempDGV[13, iX].Value = DateTime.Parse(dR["�z�z�I����"].ToString()).ToShortDateString();
                }

                if (dR["�����\���"] == DBNull.Value)
                {
                    tempDGV[14, iX].Value = "";
                }
                else
                {
                    tempDGV[14, iX].Value = DateTime.Parse(dR["�����\���"].ToString()).ToShortDateString();
                }

                if (dR["�z�z�`�Ԗ���"] == DBNull.Value)
                {
                    tempDGV[15, iX].Value = "";
                }
                else
                {
                    tempDGV[15, iX].Value = dR["�z�z�`�Ԗ���"].ToString();
                }

                tempDGV[16, iX].Value = dR["�ŗ�"].ToString();

                iX++;
            }

            //if (tempDGV.RowCount <= 8)
            //{
            //    tempDGV.Columns[6].Width = 252;
            //}
            //else
            //{
            //    tempDGV.Columns[6].Width = 235;
            //}

            tempDGV.CurrentCell = null;

            dR.Close();
            fCon.Close();

            switch (fMode.Mode)
            {
                case 0:
                    label1.Text = "�y" + listBox1.Text + " �󒍃f�[�^ �z";
                    break;

                case 1:
                    label1.Text = "�y" + listBox1.Text.Substring(11,listBox1.Text.Length -11) + " �󒍃f�[�^ �z";
                    break;
            }

            dataGridView3.RowCount = 0;

            if (tempDGV.RowCount > 0)
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }

            button7.Enabled = true;
        }

        private void ShowClient(int tempID)
        {
            OleDbDataReader dR;
            Control.���Ӑ� cCli = new Control.���Ӑ�();
            dR = cCli.FillBy("where ID = " + tempID.ToString());

            while(dR.Read())
            {
                cMaster.���Ӑ�ID = int.Parse(dR["ID"].ToString());

                txtClient.Text = dR["����"].ToString() + "";                                            //���Ӑ於
                txtKeisho.Text = dR["�h��"].ToString() + "";                                            //�h��
                txtZipCode.Text = dR["�X�֔ԍ�"].ToString() + "";                                       //�X�֔ԍ�
                txtAddress.Text = dR["������Z��1"].ToString() + "�@" + dR["������Z��2"].ToString();   //�Z��
                txtTantou.Text = dR["�S���Җ�"].ToString() + "";                                        //���S����
                txtTel.Text = dR["�d�b�ԍ�"].ToString() + "";                                           //TEL
                txtFax.Text = dR["FAX�ԍ�"].ToString() + "";                                            //FAX
            }

            dR.Close();
            cCli.Close();

        }

        /// <summary>
        /// �������f�[�^�ǂݍ���
        /// </summary>
        /// <param name="tempID">������ID</param>
        private void ShowSeikyuData(int tempID)
        {
            int iX;
            OleDbDataReader dR,dRToku;
            Control.������ cData = new Control.������();
            dR = cData.FillBy("where ID = " + tempID.ToString());

            while (dR.Read())
            {
                GetData(dR);

                //���Ӑ�}�X�^�擾
                Control.���Ӑ� cCli = new Control.���Ӑ�();
                dRToku = cCli.FillBy("where ID = " + cMaster.���Ӑ�ID.ToString());

                //���Ӑ���\��
                while (dRToku.Read())
	            {
                    txtClient.Text = dRToku["����"].ToString() + "";    //���Ӑ於
                    txtKeisho.Text = dRToku["�h��"].ToString() + "";    //�h��
                    txtZipCode.Text = dRToku["�X�֔ԍ�"].ToString() + "";   //�X�֔ԍ�
                    txtAddress.Text = dRToku["������Z��1"].ToString() + "�@" + dRToku["������Z��2"].ToString();   //�Z��
                    txtTantou.Text = dRToku["�S���Җ�"].ToString() + "";    //���S����
                    txtTel.Text = dRToku["�d�b�ԍ�"].ToString() + "";   //TEL
                    txtFax.Text = dRToku["FAX�ԍ�"].ToString() + "";    //FAX
	            }

                dRToku.Close();
                cCli.Close();
            }

            dR.Close();
            cData.Close();

            //�������f�[�^�\��
            nDate.Checked = true;

            tabControl1.TabPages[0].Text = "��������" + cMaster.ID.ToString();

            nDate.Value = cMaster.�����\���;
            label18.Text = cMaster.�ŗ�.ToString();
            label8.Text = cMaster.������z.ToString("#,##0");
            label9.Text = cMaster.�����.ToString("#,##0");
            label10.Text = cMaster.�l���z.ToString("#,##0");
            label11.Text = cMaster.�������z.ToString("#,##0");
            txtMemo.Text = cMaster.���l;

            switch (cMaster.�����敪)
	        {
		        case 0:
                    checkBox1.Checked = false;
                    label24.Hide();
                    break;

                case 1:
                    checkBox1.Checked = true;
                    label24.Show();
                    break;
	        }

            txtZan.Text = cMaster.�����c.ToString("#,##0");

            //�������ו\��
            dataGridView3.RowCount = 0;

            Control.FreeSql fCon = new Control.FreeSql();

            string mySql = "";
            mySql += "select ��.*, �󒍎��.���� as �󒍎�ʖ���,���^.���� as ���^����,�z�z�`��.���� as �z�z�`�Ԗ��� ";
            mySql += "from ((�� left join �󒍎�� ";
            mySql += "on ��.�󒍎��ID = �󒍎��.ID) left join ���^ ";
            mySql += "on ��.���^ = ���^.ID) left join �z�z�`�� ";
            mySql += "on ��.�z�z�`�� = �z�z�`��.ID ";
            mySql += "where ";
            mySql += "(��.������ID = " + cMaster.ID.ToString() + ")";
            mySql += "order by ��.ID desc";

            dR = fCon.free_dsReader(mySql);

            while (dR.Read())
            {
                dataGridView3.Rows.Add();

                iX = dataGridView3.Rows.Count;

                if (dR["�z�z�J�n��"] == DBNull.Value)
                {
                    dataGridView3[0, iX - 1].Value = "";
                }
                else
                {
                    dataGridView3[0, iX - 1].Value = DateTime.Parse(dR["�z�z�J�n��"].ToString()).ToShortDateString();
                }

                if (dR["���^����"] == DBNull.Value)
                {
                    dataGridView3[1, iX - 1].Value = "";
                }
                else
                {
                    dataGridView3[1, iX - 1].Value = dR["���^����"].ToString();
                }

                if (dR["�z�z�`�Ԗ���"] == DBNull.Value)
                {
                    dataGridView3[2, iX - 1].Value = "";
                }
                else
                {
                    dataGridView3[2, iX - 1].Value = dR["�z�z�`�Ԗ���"].ToString();
                }

                dataGridView3[3, iX - 1].Value = dR["�`���V��"].ToString();
                dataGridView3[4, iX - 1].Value = double.Parse(dR["�P��"].ToString());
                dataGridView3[5, iX - 1].Value = int.Parse(dR["����"].ToString());
                dataGridView3[6, iX - 1].Value = double.Parse(dR["���z"].ToString());
                dataGridView3[7, iX - 1].Value = double.Parse(dR["�����"].ToString());
                dataGridView3[8, iX - 1].Value = double.Parse(dR["�ō����z"].ToString());
                dataGridView3[9, iX - 1].Value = double.Parse(dR["�l���z"].ToString());
                dataGridView3[10, iX - 1].Value = double.Parse(dR["������z"].ToString());
                dataGridView3[11, iX - 1].Value = dR["ID"].ToString();
                dataGridView3[12, iX - 1].Value = dR["�ŗ�"].ToString();

            }

            dR.Close();
            fCon.Close();

            dataGridView3.CurrentCell = null;

            //��������\��
            ShowNyukin();

            //�{�^���\��
            button1.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;

        }

        private void ShowNyukin()
        {
            //��������\��
            int iX;
            OleDbDataReader dR;
            dataGridView4.RowCount = 0;
            Control.���� cNyukin = new Control.����();
            dR = cNyukin.FillBy("where ������ID = " + cMaster.ID.ToString());

            while (dR.Read())
            {
                dataGridView4.Rows.Add();

                iX = dataGridView4.Rows.Count;
                dataGridView4[0, iX - 1].Value = int.Parse(dR["ID"].ToString());
                dataGridView4[1, iX - 1].Value = DateTime.Parse(dR["�����N����"].ToString()).ToShortDateString();
                dataGridView4[2, iX - 1].Value = int.Parse(dR["���z"].ToString(), System.Globalization.NumberStyles.Any);
            }

            if (dataGridView4.RowCount > 3)
            {
                dataGridView4.Columns[2].Width = 120;
            }
            else
            {
                dataGridView4.Columns[2].Width = 137;
            }

            dR.Close();
            cNyukin.Close();

            dataGridView4.CurrentCell = null;
        }

        private void GetData(OleDbDataReader dR)
        {
            cMaster.ID = int.Parse(dR["ID"].ToString());
            cMaster.���Ӑ�ID = int.Parse(dR["���Ӑ�ID"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.�������z = int.Parse(dR["�������z"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.����� = int.Parse(dR["�����"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.�l���z = int.Parse(dR["�l���z"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.������z = int.Parse(dR["������z"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.�ŗ� = int.Parse(dR["�ŗ�"].ToString());
            cMaster.�����\��� = DateTime.Parse(dR["�����\���"].ToString());
            cMaster.���s�� = DateTime.Parse(dR["���s��"].ToString());
            cMaster.�����c = int.Parse(dR["�����c"].ToString(),System.Globalization.NumberStyles.Any);
            cMaster.�����敪 = int.Parse(dR["�����敪"].ToString());
            cMaster.�U������ID1 = int.Parse(dR["�U������ID1"].ToString());
            cMaster.�U������ID2 = int.Parse(dR["�U������ID2"].ToString());
            cMaster.���l = dR["���l"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("�������𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int G_COUNT = 12; //�������̖��׍s��

            int pCnt;

            //�y�[�W�J�E���g
            pCnt = dataGridView3.Rows.Count / G_COUNT + 1;

            for (int i = 1; i <= pCnt; i++)
            {
                SeikyuReport(pCnt, i, G_COUNT);
            }

        }

        private void SeikyuReport(int tempPage, int tempCurrentPage, int tempFixRows)
        {

            const int S_GYO = 17;    //�G�N�Z���t�@�C�����ׂ�8�s�ڂ����
            int dgvIndex;
            int i;
            DateTime sDate;

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z��������, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    //�y�[�W��
                    oxlsSheet.Cells[2, 11] = tempCurrentPage.ToString() + "/" + tempPage.ToString();

                    //���Ӑ於
                    oxlsSheet.Cells[1, 1] = txtClient.Text;

                    //�h��
                    oxlsSheet.Cells[1, 5] = txtKeisho.Text;

                    //�X�֔ԍ�
                    oxlsSheet.Cells[3, 1] = "�� " + txtZipCode.Text;

                    //�Z��
                    oxlsSheet.Cells[3, 2] = txtAddress.Text;

                    //�S���Җ�
                    oxlsSheet.Cells[4, 1] = txtTantou.Text;

                    //TEL
                    //oxlsSheet.Cells[5, 1] = "TEL�F " + txtTel.Text;

                    //FAX
                    //oxlsSheet.Cells[5, 3] = "FAX�F " + txtFax.Text;

                    //�������z
                    if (tempCurrentPage == 1)
                    {
                        oxlsSheet.Cells[13, 3] = int.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                    }
                    else
                    {
                        oxlsSheet.Cells[13, 3] = "�\";
                    }

                    //���l
                    oxlsSheet.Cells[30, 2] = txtMemo.Text;

                    //�x������
                    oxlsSheet.Cells[35, 2] = nDate.Value.ToLongDateString();

                    //�l�����o��
                    if (label10.Text == "0")
                    {
                        oxlsSheet.Cells[S_GYO - 1, 10] = ""; 
                    }
                    else
                    {
                        oxlsSheet.Cells[S_GYO - 1, 10] = "�l���z"; 
                    }

                    //��������
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //�f�[�^�O���b�h�r���[�̍s�C���f�b�N�X�����߂�

                        //�z�z��
                        sDate = DateTime.Parse(dataGridView3[0, dgvIndex].Value.ToString());
                        oxlsSheet.Cells[i + S_GYO, 1] = sDate.Month.ToString() + "/" + sDate.Day.ToString() + "�`";

                        //�z�z��(�T�C�Y)
                        oxlsSheet.Cells[i + S_GYO, 2] = dataGridView3[1, dgvIndex].Value.ToString();

                        //�`��
                        oxlsSheet.Cells[i + S_GYO, 3] = dataGridView3[2, dgvIndex].Value.ToString();

                        //�`���V��
                        oxlsSheet.Cells[i + S_GYO, 4] = dataGridView3[3, dgvIndex].Value.ToString();

                        //����
                        oxlsSheet.Cells[i + S_GYO, 6] = int.Parse(dataGridView3[5, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //�P��
                        oxlsSheet.Cells[i + S_GYO, 7] = double.Parse(dataGridView3[4, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //���z
                        oxlsSheet.Cells[i + S_GYO, 8] = int.Parse(dataGridView3[6, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //�����
                        oxlsSheet.Cells[i + S_GYO, 9] = int.Parse(dataGridView3[7, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //�l���z
                        oxlsSheet.Cells[i + S_GYO, 10] = int.Parse(dataGridView3[9, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //�������z
                        oxlsSheet.Cells[i + S_GYO, 11] = int.Parse(dataGridView3[10, dgvIndex].Value.ToString(),System.Globalization.NumberStyles.Any);

                        //�O���b�h�ŏI�s�̂Ƃ��I��
                        if (dgvIndex == (dataGridView3.Rows.Count - 1)) break;

                        //������׍ő�s�̂Ƃ��I��
                        if (i == (tempFixRows - 1)) break;

                        i++;
                    }

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
                    saveFileDialog1.FileName = "������" + "_" + txtClient.Text + DateTime.Today.Year.ToString() + "�N" + DateTime.Today.Month.ToString() + "��";

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

        //���Ӑ�R���{�{�b�N�X�N���X
        private class ListClient
        {
            private int F_ID;
            private string F_Name;
            private int F_tID;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            public int tID
            {
                get
                {
                    return F_tID;
                }
                set
                {
                    F_tID = value;
                }
            }

            //���������Ӑ�}�X�^�[���[�h
            public static void load(ListBox tempObj, string tempClient, DateTime sDate, DateTime eDate)
            {
                try
                {
                    OleDbDataReader dR;
                    ListClient lst1;
                    string sqlSTRING;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    //�f�[�^���[�_�[���擾����
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    sqlSTRING = "";
                    sqlSTRING += "select ��.���Ӑ�ID,���Ӑ�.���� as ���Ӑ於��,��.�z�z�I���� ";
                    sqlSTRING += "from �� inner join ���Ӑ� ";
                    sqlSTRING += "on ��.���Ӑ�ID = ���Ӑ�.ID ";
                    sqlSTRING += "where  (��.������ = 1) AND (��.������ID = 0) ";

                    if (tempClient != "")    //2009.09.09 �N���C�A���g����
                    {
                        sqlSTRING += "and (���Ӑ�.���� like '%" + tempClient + "%')";
                    }

                    sqlSTRING += " and (��.�z�z�I���� >= '" + sDate + "') and (��.�z�z�I���� <= '" + eDate + "') ";
                    sqlSTRING += "order by ��.�z�z�I���� desc,��.���Ӑ�ID";

                    //���������Ӑ�̃f�[�^���[�_�[���擾����
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);

                    while (dR.Read())
                    {
                        lst1 = new ListClient();
                        lst1.ID = Int32.Parse(dR["���Ӑ�ID"].ToString());
                        lst1.Name = DateTime.Parse(dR["�z�z�I����"].ToString()).ToShortDateString() + " " + dR["���Ӑ於��"].ToString() + "";
                        tempObj.Items.Add(lst1);
                    }

                    dR.Close();
                    Con.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���������Ӑ惊�X�g�{�b�N�X���[�h");
                }
            }

            //�������f�[�^���[�h
            public static void loadData(ListBox tempObj, int tempS, string tempClient, DateTime sDate, DateTime eDate)
            {
                try
                {
                    OleDbDataReader dR;
                    ListClient lst1;
                    string sqlSTRING;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    //�f�[�^���[�_�[���擾����
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    sqlSTRING = "";
                    sqlSTRING += "select ������.*,���Ӑ�.����,��.�z�z�I���� ";
                    sqlSTRING += "from ������ inner join ���Ӑ� ";
                    sqlSTRING += "on ������.���Ӑ�ID = ���Ӑ�.ID ";
                    sqlSTRING += "inner join �� ";
                    sqlSTRING += "on ������.ID = ��.������ID ";
                    sqlSTRING += "where (1 = 1) ";

                    if (tempS == 1)
                    {
                        sqlSTRING += "and (������.�����敪 = 0) ";
                    }

                    if (tempClient != "")    //2009.09.09 �N���C�A���g����
                    {
                        sqlSTRING += "and (���Ӑ�.���� like '%" + tempClient + "%')";
                    }

                    sqlSTRING += " and (��.�z�z�I���� >= '" + sDate + "') and (��.�z�z�I���� <= '" + eDate + "') ";

                    sqlSTRING += "order by ��.�z�z�I���� desc";

                    //�������̃f�[�^���[�_�[���擾����
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);

                    while (dR.Read())
                    {
                        lst1 = new ListClient();
                        lst1.ID = Int32.Parse(dR["ID"].ToString());
                        lst1.Name = DateTime.Parse(dR["�z�z�I����"].ToString()).ToShortDateString() + " " + dR["����"].ToString() + "";

                        if (dR["�����敪"].ToString() == "1") lst1.Name += " �y�����z";

                        lst1.tID = Int32.Parse(dR["���Ӑ�ID"].ToString());

                        tempObj.Items.Add(lst1);
                    }

                    dR.Close();
                    Con.Close();
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���������X�g�{�b�N�X���[�h");
                }
            }


        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBox1.SelectedIndex == -1) return;

            if (MessageBox.Show(listBox1.Text  + "�@���I������܂����B��낵���ł���", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            ListClient lst1 = new ListClient();
            lst1 = (ListClient)listBox1.SelectedItem;

            switch (fMode.Mode)
            {
                case 0: //�V�K�o�^
                    ShowPosting(dataGridView1, lst1.ID);    //�������󒍃f�[�^�\��
                    ShowClient(lst1.ID);                    //�N���C�A���g���\��
                    break;

                case 1: //�ҏW
                    ShowPosting(dataGridView1, lst1.tID);   //�������󒍃f�[�^�\��
                    ShowSeikyuData(lst1.ID);                //�����f�[�^�\��
                    break;
            }

            panel2.Hide();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("�󒍃f�[�^��I�����Ă�������", "�������דo�^", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�I������Ă���󒍃f�[�^�𐿋����ׂɒǉ����܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {

                int iX;

                dataGridView3.Rows.Add();
                iX = dataGridView3.Rows.Count;

                dataGridView3[0, iX-1].Value = dataGridView1[12, r.Index].Value.ToString();
                dataGridView3[1, iX-1].Value = dataGridView1[11, r.Index].Value.ToString();
                dataGridView3[2, iX-1].Value = dataGridView1[15, r.Index].Value.ToString();
                dataGridView3[3, iX-1].Value = dataGridView1[2, r.Index].Value.ToString();
                dataGridView3[4, iX-1].Value = double.Parse(dataGridView1[4, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[5, iX-1].Value = int.Parse(dataGridView1[5, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[6, iX-1].Value = double.Parse(dataGridView1[6, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[7, iX-1].Value = double.Parse(dataGridView1[7, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[8, iX-1].Value = double.Parse(dataGridView1[8, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[9, iX-1].Value = double.Parse(dataGridView1[9, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[10, iX-1].Value = double.Parse(dataGridView1[10, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);
                dataGridView3[11, iX - 1].Value = dataGridView1[0, r.Index].Value.ToString();
                dataGridView3[12, iX - 1].Value = int.Parse(dataGridView1[16, r.Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                //�����\���
                if (dataGridView1[14, r.Index].Value.ToString() == "")
                {
                    nDate.Checked = false;
                }
                else
                {
                    nDate.Checked = true;
                    nDate.Value = DateTime.Parse(dataGridView1[14, r.Index].Value.ToString());
                }

                //�ŗ�
                if (label18.Text == "") label18.Text = dataGridView3[12, iX - 1].Value.ToString();

            }

            //�������v���z�v�Z
            SumKingaku();
            
            button1.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;

            //�󒍃f�[�^����s�폜����
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(r);
            }

        }

        private void SumKingaku()
        {

            double sTotal = 0;
            double sTax = 0;
            double sNebiki = 0;
            double sSeikyu = 0;

            //���z�v�Z
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                sTotal += double.Parse(dataGridView3[6, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                sTax += double.Parse(dataGridView3[7, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                sNebiki += double.Parse(dataGridView3[9, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                sSeikyu += double.Parse(dataGridView3[10, i].Value.ToString(), System.Globalization.NumberStyles.Any);
            }

            label8.Text = sTotal.ToString("#,##0");
            label9.Text = sTax.ToString("#,##0");
            label10.Text = sNebiki.ToString("#,##0");
            label11.Text = sSeikyu.ToString("#,##0");

            dataGridView3.CurrentCell = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�󒍃f�[�^�S�Ă�I����Ԃɂ��܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Selected = true;
            } 

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try 
	        {	        
                if (dataGridView3.SelectedRows.Count == 0)
                {
                    MessageBox.Show("�������ׂ�I�����Ă�������", "�����f�[�^�폜", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (MessageBox.Show("�I������Ă��鐿�����ׂ��폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                foreach (DataGridViewRow r in dataGridView3.SelectedRows)
                {
                    string mySql = "";
                    Control.FreeSql fCon = new Control.FreeSql();

                    //�ҏW�����̂Ƃ��󒍃f�[�^�̐������敪��������
                    if (fMode.Mode == 1)
                    {

                        mySql += "update �� set ";
                        mySql += "������ID = 0,";
                        mySql += "���������s�� = null,";
                        mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                        mySql += "where ID = " + dataGridView3[11, dataGridView3.SelectedRows[0].Index].Value.ToString();

                        if (fCon.Execute(mySql) == false)
                        {
                            fCon.Close();
                            throw new Exception("�󒍃f�[�^�̍X�V�Ɏ��s���܂���");
                        }
                    }

                    fCon.Close();

                    //�f�[�^�O���b�h�s�폜
                    dataGridView3.Rows.Remove(r);
                }

                //�������z�Čv�Z
                SumKingaku();

                //�������בS�Ă��폜�����ꍇ
                DateTime StartDate;
                DateTime EndDate;

                if (dataGridView3.RowCount == 0)
                {
                    switch (fMode.Mode)
                    {
                        case 0: //�o�^

                            button1.Enabled = false;
                            button5.Enabled = false;
                            button6.Enabled = false;

                            break;

                        case 1: //�ҏW
                            MessageBox.Show("�������ׂ��S�č폜����܂����̂Ő������f�[�^�͍폜����܂��B", "�������폜", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            Control.������ cSe = new Control.������();
                            cSe.DataDelete(cMaster.ID);
                            cSe.Close();

                            //���������X�g�{�b�N�X�����[�h
                            if (tDate.Checked == false)
                            {
                                StartDate = sD;
                            }
                            else
                            {
                                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
                            }

                            if (tDate2.Checked == false)
                            {
                                EndDate = eD;
                            }
                            else
                            {
                                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
                            }
    
                            ListClient.loadData(listBox1, getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
                            DispClear();

                            break;
                    }

                }	
            }

	        catch (Exception ex)
	        {
                MessageBox.Show(ex.Message, "�������׍폜", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void DispClear()
        {
            switch (fMode.Mode)
            {
                case 0:
                    panel2.Hide();
                    break;

                case 1:
                    panel2.Show();
                    break; 
            }

            label24.Hide();

            textBox1.Text = "";
            tDate.Checked = false;
            tDate2.Checked = false;

            tabControl1.TabPages[0].Text = "�������f�[�^";

            txtClient.Text = "";
            txtKeisho.Text = "";
            txtZipCode.Text = "";
            txtAddress.Text = "";
            txtTantou.Text = "";
            txtTel.Text = "";
            txtFax.Text = "";
            nDate.Checked = false;

            label8.Text = "";
            label9.Text = "";
            label10.Text = "";
            label11.Text = "";
            label18.Text = "";

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;

            dataGridView1.RowCount = 0;
            dataGridView3.RowCount = 0;

            txtMemo.Text = "";

            label1.Text = "�y�󒍃f�[�^�z";

            //�������
            if (fMode.Mode == 1)
            {
                txtKingaku.Text = "";
                txtZan.Text = "";
                dataGridView4.RowCount = 0;
                button9.Enabled = false;
                button10.Enabled = false;
                nyukinMode = 0;
                checkBox1.Checked = false;
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("��ʂ��������܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            
            DispClear();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {

                if (fDataCheck() == true)
                {
                    Control.DataControl Con;
                    OleDbConnection cn;
                    OleDbTransaction tran;
                    OleDbCommand SCom;

                    switch (fMode.Mode)
                    {
                        case 0:

                            //�V�K�o�^
                            if (MessageBox.Show("�����f�[�^��o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                return;

                            //ID���̔�
                            string sqlStr = "";
                            int gID = (int)(1);

                            sqlStr = "select max(ID) as ID from ������ ";
                            OleDbDataReader dR;
                            Control.FreeSql fCon = new Control.FreeSql();
                            dR = fCon.free_dsReader(sqlStr);

                            while (dR.Read())
                            {
                                if (dR["ID"] == DBNull.Value)
                                {
                                    gID = (int)(1);
                                }
                                else
                                {
                                    gID = Int32.Parse(dR["ID"].ToString()) + 1;
                                }
                            }

                            dR.Close();
                            fCon.Close();

                            //ID��ݒ�
                            cMaster.ID = gID;

                            //�o�^����
                            Con = new Control.DataControl();
                            cn = new OleDbConnection();

                            cn = Con.GetConnection();

                            //�g�����U�N�V�����J�n
                            tran = cn.BeginTransaction();

                            SCom = new OleDbCommand();
                            SCom.Connection = cn;
                            SCom.Transaction = tran;

                            try
                            {
                                //�������f�[�^�o�^����
                                sqlStr = "";
                                sqlStr += "insert into ������ ";
                                sqlStr += "(ID,���Ӑ�ID,�������z,�����,�l���z,������z,�ŗ�,�����\���,���s��,";
                                sqlStr += "�����c,�����敪,�U������ID1,�U������ID2,���l,�o�^�N����,�ύX�N����) ";
                                sqlStr += "values (";
                                sqlStr += cMaster.ID + ",";
                                sqlStr += cMaster.���Ӑ�ID + ",";
                                sqlStr += cMaster.�������z + ",";
                                sqlStr += cMaster.����� + ",";
                                sqlStr += cMaster.�l���z + ",";
                                sqlStr += cMaster.������z + ",";
                                sqlStr += cMaster.�ŗ� + ",";
                                sqlStr += "'" + cMaster.�����\��� + "',";
                                sqlStr += "'" + cMaster.���s�� + "',";
                                sqlStr += cMaster.�����c + ",";
                                sqlStr += cMaster.�����敪 + ",";
                                sqlStr += cMaster.�U������ID1 + ",";
                                sqlStr += cMaster.�U������ID2 + ",";
                                sqlStr += "'" + cMaster.���l + "',";
                                sqlStr += "'" + cMaster.�o�^�N���� + "',";
                                sqlStr += "'" + cMaster.�ύX�N���� + "')";

                                SCom.CommandText = sqlStr;

                                SCom.ExecuteNonQuery();

                                //�󒍃f�[�^�X�V
                                string sID;

                                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                                {
                                    sID = dataGridView3[11, i].Value.ToString();    //��ID

                                    sqlStr = "";
                                    sqlStr += "update �� ";
                                    sqlStr += "set ";
                                    sqlStr += "������ID = " + gID.ToString() + ",";
                                    sqlStr += "���������s�� = '" + DateTime.Today + "',";
                                    sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
                                    sqlStr += "where (��.ID = " + sID + ") ";

                                    SCom.CommandText = sqlStr;

                                    SCom.ExecuteNonQuery();
                                }

                                tran.Commit();

                                MessageBox.Show("�V�K�o�^����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();

                                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                MessageBox.Show("�V�K�o�^�Ɏ��s���܂����B���[���o�b�N���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            }

                            cn.Close();

                            Con.Close();

                            break;

                        case 1: //�X�V
                            if (MessageBox.Show("�X�V���܂��B��낵���ł����H", "�X�V�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                return;

                            //�X�V��������
                            Con = new Control.DataControl();
                            cn = new OleDbConnection();
                            cn = Con.GetConnection();

                            //�g�����U�N�V�����J�n
                            tran = cn.BeginTransaction();

                            SCom = new OleDbCommand();
                            SCom.Connection = cn;
                            SCom.Transaction = tran;

                            try
                            {
                                //�������f�[�^�X�V
                                sqlStr = "";
                                sqlStr += "update ������ set ";
                                sqlStr += "���Ӑ�ID = " + cMaster.���Ӑ�ID + ",";
                                sqlStr += "�������z = " + cMaster.�������z + ",";
                                sqlStr += "����� = " + cMaster.����� + ",";
                                sqlStr += "�l���z = " + cMaster.�l���z + ",";
                                sqlStr += "������z = " + cMaster.������z + ",";
                                sqlStr += "�ŗ� = " + cMaster.�ŗ� + ",";
                                sqlStr += "�����\��� = '" + cMaster.�����\��� + "',";
                                sqlStr += "���s�� = '" + cMaster.���s�� + "',";
                                sqlStr += "�����c = " + cMaster.�����c + ",";
                                sqlStr += "�����敪 = " + cMaster.�����敪 + ",";
                                sqlStr += "�U������ID1 = " + cMaster.�U������ID1 + ",";
                                sqlStr += "�U������ID2 = " + cMaster.�U������ID2 + ",";
                                sqlStr += "���l = '" + cMaster.���l + "',";
                                sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
                                sqlStr += "where ID = " + cMaster.ID;

                                SCom.CommandText = sqlStr;

                                // SQL�̎��s
                                SCom.ExecuteNonQuery();

                                //�󒍃f�[�^�X�V
                                string sID;

                                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                                {
                                    sID = dataGridView3[11, i].Value.ToString();    //��ID

                                    sqlStr = "";
                                    sqlStr += "update �� ";
                                    sqlStr += "set ";
                                    sqlStr += "������ID = " + cMaster.ID.ToString() + ",";
                                    sqlStr += "���������s�� = '" + DateTime.Today + "',";
                                    sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
                                    sqlStr += "where (��.ID = " + sID + ") ";

                                    SCom.CommandText = sqlStr;

                                    SCom.ExecuteNonQuery();
                                }

                                tran.Commit();
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();

                                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                MessageBox.Show("�X�V�Ɏ��s���܂����B���[���o�b�N���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }

                            cn.Close();
                            Con.Close();
                            break;
                    }

                    //���Ӑ惊�X�g�������[�h
                    DateTime StartDate;
                    DateTime EndDate;

                    if (tDate.Checked == false)
                    {
                        StartDate = sD;
                    }
                    else
                    {
                        StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
                    }

                    if (tDate2.Checked == false)
                    {
                        EndDate = eD;
                    }
                    else
                    {
                        EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
                    }
    
                    switch (fMode.Mode)
                    {

                        case 0:
                            ListClient.load(listBox1, textBox1.Text.Trim(), StartDate, EndDate);
                            break;

                        case 1:
                            ListClient.loadData(listBox1,getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
                            break;
                    }

                    //��ʏ�����
                    DispClear();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "�o�^����", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }    
        }

        //�o�^�f�[�^�`�F�b�N
        private Boolean fDataCheck()
        {

            try
            {

                //�ҏW���[�h�̂Ƃ�
                if (fMode.Mode == 1)
                {
                    if (int.Parse(txtZan.Text, System.Globalization.NumberStyles.Any) > 0 && checkBox1.Checked == true)
                    {
                        if (MessageBox.Show("�����c��������܂������������Ƀ`�F�b�N�������Ă��܂��B��낵���ł���", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            return false;
                    }
                }

                if (nDate.Checked == false)
                {
                    throw new Exception("�����\�����o�^���Ă�������");
                }

                //�N���X�Ƀf�[�^�Z�b�g
                cMaster.�������z = int.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                cMaster.����� = int.Parse(label9.Text, System.Globalization.NumberStyles.Any);
                cMaster.������z  = int.Parse(label8.Text, System.Globalization.NumberStyles.Any);
                cMaster.�l���z = int.Parse(label10.Text, System.Globalization.NumberStyles.Any);
                cMaster.�ŗ� = int.Parse(label18.Text, System.Globalization.NumberStyles.Any);
                cMaster.�����\��� = DateTime.Parse(nDate.Value.ToShortDateString());

                switch (fMode.Mode)
	            {
                    case 0:
                        cMaster.���s�� = DateTime.Today;
                        cMaster.�����c = int.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                        cMaster.�����敪 = 0;
                        break;

                    case 1:
                        cMaster.�����c = int.Parse(txtZan.Text, System.Globalization.NumberStyles.Any);

                        if (checkBox1.Checked == true)
                        {
                            cMaster.�����敪 = 1;
                        }
                        else
                        {
                            cMaster.�����敪 = 0;
                        }
                        break;
	            } 

                cMaster.�U������ID1 = 0;
                cMaster.�U������ID2 = 0;
                cMaster.���l = txtMemo.Text;
                
                if (fMode.Mode == 0) cMaster.�o�^�N���� = DateTime.Today;
                
                cMaster.�ύX�N���� = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "�o�^", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�\�����̐����f�[�^���폜���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�폜
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();
            cn = Con.GetConnection();

            OleDbTransaction tran;

            //�g�����U�N�V�����J�n
            tran = cn.BeginTransaction();

            OleDbCommand SCom = new OleDbCommand();

            SCom.Connection = cn;
            SCom.Transaction = tran;

            string sqlSTR;

            try
            {
                //�������f�[�^�폜
                sqlSTR = "";
                sqlSTR += "delete from ������ ";
                sqlSTR += "where ID = " + cMaster.ID.ToString();

                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                //�󒍃f�[�^�X�V�E������ID������
                string sID;

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    sID = dataGridView3[11, i].Value.ToString();    //��ID

                    sqlSTR = "";
                    sqlSTR += "update �� set ";
                    sqlSTR += "������ID = 0,";
                    sqlSTR += "���������s�� = null,";
                    sqlSTR += "�ύX�N���� = '" + DateTime.Today + "' ";
                    sqlSTR += "where ID = " + sID.ToString();

                    SCom.CommandText = sqlSTR;

                    SCom.ExecuteNonQuery();
                }

                //�����f�[�^�폜
                sqlSTR = "";
                sqlSTR += "delete from ���� ";
                sqlSTR += "where ������ID = " + cMaster.ID.ToString();

                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                tran.Commit();

                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                tran.Rollback();

                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MessageBox.Show("�폜�Ɏ��s���܂����B���[���o�b�N���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            cn.Close();
            Con.Close();

            //���Ӑ惊�X�g�������[�h
            DateTime StartDate;
            DateTime EndDate;

            if (tDate.Checked == false)
            {
                StartDate = sD;
            }
            else
            {
                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
            }

            if (tDate2.Checked == false)
            {
                EndDate = eD;
            }
            else
            {
                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
            }
            
            switch (fMode.Mode)
            {
                case 0:
                    ListClient.load(listBox1, textBox1.Text.Trim(), StartDate, EndDate);
                    break;

                case 1: 
                    ListClient.loadData(listBox1,getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
                    break;
            } 
            
            DispClear();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (Utility.NumericCheck(txtKingaku.Text) == false)
            {
                MessageBox.Show("�����z�͐����œ��͂��Ă�������","�����z�G���[",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                txtKingaku.Focus();
                return;
            }

            if (int.Parse(txtKingaku.Text,System.Globalization.NumberStyles.Any) > int.Parse(txtZan.Text,System.Globalization.NumberStyles.Any))
            {
                if (MessageBox.Show("�����c���𒴂��Ă��܂��B��낵���ł���", "�ߓ���", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    txtKingaku.Focus();
                    return;
                }
            }

            if (MessageBox.Show(dateTimePicker1.Value.ToShortDateString() + "�̓�������o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            
            Control.���� nCon = new Control.����();

            switch (nyukinMode)
            {
                case 0: //�V�K�o�^

                    //�N���X�Ƀf�[�^���Z�b�g����
                    cNyukin.������ID = cMaster.ID;
                    cNyukin.�����N���� = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                    cNyukin.���z = int.Parse(txtKingaku.Text, System.Globalization.NumberStyles.Any);
                    cNyukin.���l = "";
                    cNyukin.�o�^�N���� = DateTime.Today;
                    cNyukin.�ύX�N���� = DateTime.Today;

                    //�����f�[�^�o�^
                    nCon.DataInsert(cNyukin);
                    break;

                case 1:

                    //�N���X�Ƀf�[�^���Z�b�g����
                    cNyukin.�����N���� = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                    cNyukin.���z = int.Parse(txtKingaku.Text, System.Globalization.NumberStyles.Any);
                    cNyukin.�ύX�N���� = DateTime.Today;

                    //�����f�[�^�o�^
                    nCon.DataUpdate(cNyukin);
                    break;
            }

            nCon.Close();

            //��������\��
            ShowNyukin();

            //�c���\��
            ShowSeikyuZan();

            //�����������[�h������
            nyukinMode = 0;

            button10.Enabled = false;
            txtKingaku.Text = "";

        }

        //�����c���\��
        private void ShowSeikyuZan()
        {
            int sKin = 0;
            int sZan;

            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                sKin += int.Parse(dataGridView4[2, i].Value.ToString(), System.Globalization.NumberStyles.Any);
            }

            sZan = int.Parse(label11.Text, System.Globalization.NumberStyles.Any) - sKin;
            txtZan.Text = sZan.ToString("#,##0");

        }

        private void txtKingaku_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == textBox1) txtObj = textBox1;
            if (sender == txtClient) txtObj = txtClient;
            if (sender == txtKeisho) txtObj = txtKeisho;
            if (sender == txtZipCode) txtObj = txtZipCode;
            if (sender == txtAddress) txtObj = txtAddress;
            if (sender == txtTantou) txtObj = txtTantou;
            if (sender == txtTel) txtObj = txtTel;
            if (sender == txtFax) txtObj = txtFax;
            if (sender == txtKingaku) txtObj = txtKingaku;
            if (sender == txtMemo) txtObj = txtMemo;

            txtObj.BackColor = Color.LightGray;
            txtObj.SelectAll();
        }

        private void txtKingaku_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == textBox1) txtObj = textBox1;
            if (sender == txtClient) txtObj = txtClient;
            if (sender == txtKeisho) txtObj = txtKeisho;
            if (sender == txtZipCode) txtObj = txtZipCode;
            if (sender == txtAddress) txtObj = txtAddress;
            if (sender == txtTantou) txtObj = txtTantou;
            if (sender == txtTel) txtObj = txtTel;
            if (sender == txtFax) txtObj = txtFax;
            if (sender == txtKingaku) txtObj = txtKingaku;
            if (sender == txtMemo) txtObj = txtMemo;

            txtObj.BackColor = Color.White;
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView4.SelectedRows.Count == 0) return;

            //�������I��
            if (MessageBox.Show(dataGridView4[1, dataGridView4.SelectedRows[0].Index].Value.ToString() + "�̓����f�[�^���I������܂����B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�擾
            GetDataNyukin(int.Parse(dataGridView4[0, dataGridView4.SelectedRows[0].Index].Value.ToString()));

            //�Ώۃf�[�^�̕\��

            dateTimePicker1.Value = cNyukin.�����N����;
            txtKingaku.Text = cNyukin.���z.ToString("#,##0");

            //�폜�{�^���\��
            button10.Enabled = true;

        }

        /// <summary>
        /// �����f�[�^�擾
        /// </summary>
        /// <param name="tempID">ID</param>
        private void GetDataNyukin(int tempID)
        {
            OleDbDataReader dR;
            Control.���� sNyukin = new Control.����();
            dR = sNyukin.FillBy("where ID = " + tempID.ToString());

            while (dR.Read())
            {
                cNyukin.ID = int.Parse(dR["ID"].ToString());
                cNyukin.������ID = int.Parse(dR["������ID"].ToString());
                cNyukin.�����N���� = DateTime.Parse(dR["�����N����"].ToString());
                cNyukin.���z = int.Parse(dR["���z"].ToString(),System.Globalization.NumberStyles.Any);
                cNyukin.���l = dR["���l"].ToString();
            }

            dR.Close();
            sNyukin.Close();

            nyukinMode = 1;
        }

        //�����������[�h(0:�o�^,1:�ҏW)
        private int nMode;

        private int nyukinMode
        {
            get { return nMode; }
            set { nMode = value; }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(dateTimePicker1.Value.ToShortDateString() + "�̓��������폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            Control.���� nCon = new Control.����();
            nCon.DataDelete(cNyukin.ID);
            nCon.Close();

            //��������\��
            ShowNyukin();

            //�c���\��
            ShowSeikyuZan();

            //�����������[�h������
            nyukinMode = 0;

            button10.Enabled = false;
            txtKingaku.Text = "";
        }

        private int getRbtn()
        {
            if (radioButton1.Checked == true)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            DateTime StartDate;
            DateTime EndDate;

            if (tDate.Checked == false)
            {
                StartDate = sD;
            }
            else
            {
                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());
            }

            if (tDate2.Checked == false)
            {
                EndDate = eD;
            }
            else
            {
                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
            }
            
            ListClient.loadData(listBox1, getRbtn(),textBox1.Text.Trim(),StartDate,EndDate);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DateTime StartDate;
            DateTime EndDate;

            if (tDate.Checked == false)
                StartDate = sD;
            else
                StartDate = Convert.ToDateTime(tDate.Value.ToShortDateString());

            if (tDate2.Checked == false)
                EndDate = eD;
            else
                EndDate = Convert.ToDateTime(tDate2.Value.ToShortDateString());
            
            switch (fMode.Mode)
            {
                case 0:
                    ListClient.load(listBox1, textBox1.Text.Trim(), StartDate, EndDate);
                    break;

                case 1:
                    ListClient.loadData(listBox1, getRbtn(), textBox1.Text.Trim(),StartDate,EndDate);
                    break;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
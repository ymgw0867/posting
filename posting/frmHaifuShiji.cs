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
    public partial class frmHaifuShiji: Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�z�z�w�� cMaster = new Entity.�z�z�w��();
        Entity.�z�z�G���A cArea = new Entity.�z�z�G���A();

        const string MESSAGE_CAPTION = "�z�z�w���E�񍐏�";
        const int COLUMN_KANRYO = 12;   //�����敪��  2015/07/09(11 �� 12)
        const string KANRYO_STATUS = "True";    // �����敪
        const string STATUS_KANRYO = "1";       // ����
        const string STATUS_MIKANRYO = "0";     // ������

        bool STATUS_MAISU = false;      // �z�z�����v�Z�X�e�[�^�X

        public frmHaifuShiji()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            try
            {
                //�O���b�h�ݒ�
                GridviewSet.Setting(dataGridView1);
                GridviewSet.Setting2(dataGridView2);

                //�z�z���R���{
                Utility.ComboStaff.load(cmbsStaff);

                DispClear();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }
            
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
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 166;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�w���ԍ�");
                    tempDGV.Columns.Add("col2", "�`���V��");
                    tempDGV.Columns.Add("col3", "�z�z��");
                    tempDGV.Columns.Add("col4", "���͓�");
                    tempDGV.Columns.Add("col5", "�z�z��");
                    tempDGV.Columns.Add("col6", "��ʔ�");
                    tempDGV.Columns.Add("col7", "�o�^�N����");
                    tempDGV.Columns.Add("col8", "�ύX�N����");
                    tempDGV.Columns.Add("col9", "�{���̒��ӎ���");
                    tempDGV.Columns.Add("col10", "���[�U�[");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 300;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 120;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 110;
                    tempDGV.Columns[7].Width = 110;
                    tempDGV.Columns[8].Width = 614;
                    tempDGV.Columns[9].Width = 100;

                    tempDGV.Columns[2].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[3].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

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
                    //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //�z�z�w�����׃f�[�^�O���b�h
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
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 252;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "PID");
                    tempDGV.Columns.Add("col2", "�`���V��");
                    tempDGV.Columns.Add("col3", "�z�z�敪");
                    tempDGV.Columns.Add("col4", "�z�z�`��");
                    tempDGV.Columns.Add("col5", "�z�z��Z��");
                    tempDGV.Columns.Add("col6", "�}�ԋL��");
                    tempDGV.Columns.Add("col7", "ID");
                    tempDGV.Columns.Add("col8", "�P��");
                    tempDGV.Columns.Add("col9", "�\�薇��");
                    tempDGV.Columns.Add("col10", "�\�薇���v");
                    tempDGV.Columns.Add("col11", "�z�z����");
                    tempDGV.Columns.Add("col18", "�z�z�����v");     // 2015/07/15

                    // 2016/10/31
                    DataGridViewCheckBoxColumn cbc2 = new DataGridViewCheckBoxColumn();
                    cbc2.Name = "col12";
                    cbc2.HeaderText = "�z�z����";
                    cbc2.TrueValue = "1";
                    cbc2.FalseValue = "0";
                    tempDGV.Columns.Add(cbc2);

                    tempDGV.Columns.Add("col13", "�񍐖���");
                    tempDGV.Columns.Add("col14", "delete");
                    tempDGV.Columns.Add("col15", "Add");
                    tempDGV.Columns.Add("col16", "���z�z���L��");
                    tempDGV.Columns.Add("col17", "�}�ԗL��");

                    // 2016/11/14 ���z�z�}���V�������͉�ʕ\���{�^��
                    DataGridViewButtonColumn dbt = new DataGridViewButtonColumn();
                    dbt.UseColumnTextForButtonValue = true;
                    dbt.Text = "��";
                    dbt.Name = "col19";
                    dbt.HeaderText = "���z�z";
                    tempDGV.Columns.Add(dbt);

                    // 2016/12/26 ���׍폜�{�^��
                    DataGridViewButtonColumn delBtn = new DataGridViewButtonColumn();
                    delBtn.UseColumnTextForButtonValue = true;
                    delBtn.Text = "Del";
                    delBtn.Name = "col20";
                    delBtn.HeaderText = "�폜";
                    tempDGV.Columns.Add(delBtn);
                    
                    //tempDGV.Columns.Add("col18", "�Ԓn��");
                    //tempDGV.Columns.Add("col19", "�}���V������");
                    //tempDGV.Columns.Add("col20", "���R");
                    //tempDGV.Columns.Add("col21", "���̑����e");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 120;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 200;
                    tempDGV.Columns[5].Width = 200;
                    tempDGV.Columns[6].Width = 60;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 100;
                    tempDGV.Columns[11].Width = 100;    // 2015/07/15
                    tempDGV.Columns[13].Width = 100;
                    tempDGV.Columns[16].Width = 200;
                    tempDGV.Columns[18].Width = 60;     // 2016/11/14
                    tempDGV.Columns[19].Width = 60;     // 2016/12/26

                    //tempDGV.Columns[17].Width = 120;
                    //tempDGV.Columns[18].Width = 160;
                    //tempDGV.Columns[19].Width = 70;
                    //tempDGV.Columns[20].Width = 200;

                    tempDGV.Columns[1].Frozen = true;       // 2016/10/27

                    tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[12].Visible = false;    // 2015/07/15
                    tempDGV.Columns[13].Visible = false;    // 2015/07/15
                    tempDGV.Columns[14].Visible = false;    // 2015/07/15
                    tempDGV.Columns[15].Visible = false;    // 2015/07/15
                    tempDGV.Columns[16].Visible = false;    // 2015/07/15
                    tempDGV.Columns[17].Visible = false;    // 2015/07/15

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[11].DefaultCellStyle.Format = "#,##0";    // 2015/07/15
                    tempDGV.Columns[13].DefaultCellStyle.Format = "#,##0";    // 2015/07/15

                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;    // 2015/07/15
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;    // 2015/07/15

                    //tempDGV.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    //tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //tempDGV.MultiSelect = true;

                    tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // �P���Ɣz�z�����̕ҏW�Ƃ��� 2016/10/27
                    //tempDGV.ReadOnly = true;

                    foreach (DataGridViewColumn d in tempDGV.Columns)
                    {
                        if (d.Name == "col8" || d.Name == "col11" || d.Name == "col12" || d.Name == "col19" || d.Name == "col20")
                        {
                            d.ReadOnly = false;
                        }
                        else
                        {
                            d.ReadOnly = true;
                        }
                    }
                    
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

            ///------------------------------------------------------------------------------
            /// <summary>
            ///     �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾���� </summary>
            /// <param name="dgv">
            ///     �ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            ///------------------------------------------------------------------------------
            public static Boolean GetData(DataGridView dgv, ref Entity.�z�z�w�� tempC, int iX)
            {
                string sqlStr;
                Control.�z�z�w�� cShiji = new Control.�z�z�w��();
                OleDbDataReader dr;

                sqlStr = " where �z�z�w��.ID = " + (int)dgv[0, iX].Value;
                dr = cShiji.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.�z�z�� = Convert.ToDateTime(dr["�z�z��"].ToString());
                        tempC.���͓� = Convert.ToDateTime(dr["���͓�"].ToString());
                        tempC.�z�z��ID = Int32.Parse(dr["�z�z��ID"].ToString());
                        tempC.��ʔ� = Int32.Parse(dr["��ʔ�"].ToString());
                        tempC.��ʋ�ԊJ�n = dr["��ʋ�ԊJ�n"].ToString();
                        tempC.��ʋ�ԏI�� = dr["��ʋ�ԏI��"].ToString();
                        tempC.�z�z�J�n���� = dr["�z�z�J�n����"].ToString();
                        tempC.�z�z�I������ = dr["�z�z�I������"].ToString();
                        tempC.�I�����|�[�g = dr["�I�����|�[�g"].ToString();
                        tempC.���z�z�敪 = dr["���z�z�敪"].ToString();
                        tempC.���z�z���R = dr["���z�z���R"].ToString();
                        tempC.���ӎ��� = dr["���ӎ���"].ToString();
                    }
                }
                else
                {
                    dr.Close();
                    cShiji.Close();
                    return false;
                }

                dr.Close();
                cShiji.Close();
                return true;
            }

            public static Boolean GetDataItem(int tempID, ref Entity.�z�z�G���A tempC)
            {
                string sqlStr;

                Control.�z�z�G���A cArea = new Control.�z�z�G���A();
                OleDbDataReader dr;

                sqlStr = " where �z�z�G���A.ID = " + tempID;
                dr = cArea.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Int32.Parse(dr["ID"].ToString());
                        tempC.����ID = Int32.Parse(dr["����ID"].ToString());
                        tempC.�\�薇�� = Int32.Parse(dr["�\�薇��"].ToString());
                        tempC.��ID = Int32.Parse(dr["��ID"].ToString());
                        tempC.�z�z�w��ID = Int32.Parse(dr["�z�z�w��ID"].ToString());
                        tempC.�z�z�P�� = double.Parse(dr["�z�z�P��"].ToString());
                        tempC.�z�z�� = dr["�z�z��"].ToString();
                        tempC.���z�z���� = Int32.Parse(dr["���z�z����"].ToString());
                        tempC.���c�� = Int32.Parse(dr["���c��"].ToString());
                        tempC.�񍐖��� = Int32.Parse(dr["�񍐖���"].ToString());
                        tempC.�񍐎c�� = Int32.Parse(dr["���z�z�敪"].ToString());
                        tempC.���z�敪 = Int32.Parse(dr["���z�敪"].ToString());
                        tempC.�����敪 = Int32.Parse(dr["�����敪"].ToString());
                        tempC.�X�e�[�^�X = Int32.Parse(dr["�X�e�[�^�X"].ToString());
                    }
                }
                else
                {
                    dr.Close();
                    cArea.Close();
                    return false;
                }

                dr.Close();
                cArea.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV,int tempID,string tempsID,string tempCName)
            {
                string sqlSTRING = "";
                int iX;
                string wID = "0";
                
                try
                {
                    tempDGV.RowCount = 0;

                    //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����

                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    //�f�[�^���[�_�[���擾����
                    OleDbDataReader dR;

                    sqlSTRING += "select �z�z�w��.*,�z�z��.����,x.�`���V��, ���O�C�����[�U�[.���O�C�����[�U�[ ";
                    sqlSTRING += "from �z�z�w�� left join �z�z�� on �z�z�w��.�z�z��ID = �z�z��.ID ";

                    //�`���V���w��̂Ƃ�
                    if (tempCName.Length == 0)
                    {
                        sqlSTRING += "left join ";
                    }
                    else
                    {
                        sqlSTRING += "inner join ";
                    }

                    sqlSTRING += "(SELECT DISTINCT �z�z�w��.ID, ��.�`���V�� ";
                    sqlSTRING += "from �z�z�w�� inner join �z�z�G���A ";
                    sqlSTRING += "ON �z�z�w��.ID = �z�z�G���A.�z�z�w��ID inner join ";
                    sqlSTRING += "�� ON �z�z�G���A.��ID = ��.ID ";

                    //�`���V���w��̂Ƃ�
                    if (tempCName.Length > 0)
                    {
                        sqlSTRING += "where (��.�`���V�� like ?)";
                    }
                    
                    sqlSTRING += ") AS x ";
                    sqlSTRING += "ON �z�z�w��.ID = x.ID ";

                    // 2016/09/26
                    sqlSTRING += "left join ���O�C�����[�U�[ on ";
                    sqlSTRING += "�z�z�w��.���[�U�[ID = ���O�C�����[�U�[.ID ";

                    sqlSTRING += "where (1 = 1) ";

                    //////sqlSTRING += "(�z�z�w��.�z�z�� >= ?) and ";
                    //////sqlSTRING += "(�z�z�w��.�z�z�� <= ?) and ";
                    //////sqlSTRING += "(�z�z�w��.���͓� >= ?) and ";
                    //////sqlSTRING += "(�z�z�w��.���͓� <= ?) ";

                    //�z�z���w��̂Ƃ�
                    if (tempID != -1)
                    {
                        sqlSTRING += "and (�z�z�w��.�z�z��ID = ?) ";
                    }

                    //�z�z�w��ID�w��̂Ƃ�
                    if (tempsID.Length > 0)
                    {
                        sqlSTRING += "and (�z�z�w��.ID = ?) ";
                    }

                    // �z�z����1�N�ȓ��̈Č��Ƃ��� 2018/02/20
                    sqlSTRING += "and (�z�z�w��.�z�z�� >= ?) ";

                    sqlSTRING += "order by �z�z�w��.ID desc ";

                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    // �`���V���w��p�����[�^
                    if (tempCName.Length > 0)
                    {
                        SCom.Parameters.AddWithValue("@temCName","%" + tempCName + "%");
                    }

                    ////////�z�z���p�����[�^
                    //////SCom.Parameters.AddWithValue("@sDate", tempsDate);
                    //////SCom.Parameters.AddWithValue("@eDate", tempeDate);

                    ////////���͓��p�����[�^
                    //////SCom.Parameters.AddWithValue("@nsDate", tempnsDate);
                    //////SCom.Parameters.AddWithValue("@neDate", tempneDate);

                    // �z�z��ID�p�����[�^
                    if (tempID != -1)
                    {
                        SCom.Parameters.AddWithValue("@SID", tempID);
                    }

                    // �z�z�w��ID�p�����[�^
                    if (tempsID.Length > 0)
                    {
                        SCom.Parameters.AddWithValue("@tempsID",tempsID);
                    }

                    // �z�z���p�����[�^ : 2018/02/20
                    DateTime hDt = DateTime.Today.AddYears(-1);
                    SCom.Parameters.AddWithValue("@hdt", hDt);

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();

                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    //�\���p�f�[�^�̈�̃C���X�^���X�쐬
                    DataTemp d = new DataTemp();
                    dClear(d);

                    //�f�[�^��ǂݍ���
                    if (dR.HasRows == true)
                    {
                        while (dR.Read())
                        {
                            //�ŏ��̃f�[�^�ł͂Ȃ�ID�Ńu���[�N�����������Ƃ�
                            if ((wID != "0") && (wID != dR["ID"].ToString()))
                            {
                                //�O���b�h�֕\��
                                AddDataGrid(tempDGV, d, iX);

                                //�f�[�^�̈揉����
                                dClear(d);

                                iX++;
                            }

                            //�\�����ڂ̎擾
                            d.ID = dR["ID"].ToString();

                            if (d.CName.Length > 0)
                            {
                                d.CName += ", " + dR["�`���V��"].ToString();
                            }
                            else
                            {
                                d.CName = dR["�`���V��"].ToString();
                            }

                            d.HDate = dR["�z�z��"].ToString();
                            d.IDate = dR["���͓�"].ToString();
                            d.Name = dR["����"].ToString() + "";
                            d.Kotsuhi = dR["��ʔ�"].ToString();
                            d.AddDate = dR["�o�^�N����"].ToString();
                            d.UppDate = dR["�ύX�N����"].ToString();
                            d.Memo = dR["���ӎ���"].ToString();

                            // 2016/09/26
                            if (dR["���O�C�����[�U�["] == DBNull.Value)
                            {
                                d.loginUser = string.Empty;
                            }
                            else
                            {
                                d.loginUser = dR["���O�C�����[�U�["].ToString();
                            }

                            wID = dR["ID"].ToString();
                        }

                        //�ŏI�f�[�^�̃O���b�h�\��
                        AddDataGrid(tempDGV, d, iX);
                    }
                    else
                    {
                        MessageBox.Show("�Y������f�[�^������܂���", "�z�z�w���f�[�^����",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    }

                    dR.Close();

                    Con.Close();

                    cn.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }

                tempDGV.CurrentCell = null;

            }

            private static void dClear(DataTemp d)
            {
                d.ID = "";
                d.CName = "";
                d.HDate = "";
                d.IDate = "";
                d.Kotsuhi = "";
                d.Name = "";
                d.AddDate = "";
                d.UppDate = "";
            }

            //�O���b�h�r���[�\������
            private static void AddDataGrid(DataGridView tempDGV, DataTemp d,int iX)
            {

                tempDGV.Rows.Add();

                tempDGV[0, iX].Value = int.Parse(d.ID);
                tempDGV[1, iX].Value = d.CName;
                tempDGV[2, iX].Value = Convert.ToDateTime(d.HDate);
                tempDGV[3, iX].Value = Convert.ToDateTime(d.IDate);
                tempDGV[4, iX].Value = d.Name + "";
                tempDGV[5, iX].Value = Int32.Parse(d.Kotsuhi);
                tempDGV[6, iX].Value = Convert.ToDateTime(d.AddDate);
                tempDGV[7, iX].Value = Convert.ToDateTime(d.UppDate);
                tempDGV[8, iX].Value = d.Memo;
                tempDGV[9, iX].Value = d.loginUser;     // 2016/09/26
            }

            ///--------------------------------------------------------------------
            /// <summary>
            ///     �z�z�G���A�f�[�^�\�� </summary>
            /// <param name="tempDGV">
            ///     �z�z�w���f�[�^�O���b�h�r���[</param>
            /// <param name="tempDGV2">
            ///     �z�z�G���A�O���b�h�r���[</param>
            ///--------------------------------------------------------------------
            public static void ShowDataItem(DataGridView tempDGV, DataGridView tempDGV2,int iX)
            {
                //int iX = 0;

                try
                {
                    tempDGV2.RowCount = 0;

                    //�z�z�G���A�f�[�^�̃f�[�^���[�_�[���擾����
                    string sqlStr;

                    Control.�z�z�G���A cArea = new Control.�z�z�G���A();
                    OleDbDataReader dr;

                    sqlStr = " where �z�z�G���A.�z�z�w��ID = " + (int)tempDGV[0, iX].Value + " ";
                    sqlStr += "order by �z�z�G���A.��ID,�z�z�G���A.����ID";

                    dr = cArea.FillBy(sqlStr);

                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dr.Read())
                    {
                        tempDGV2.Rows.Add();

                        tempDGV2[0, iX].Value = dr["ID"];

                        iX++;
                    }

                    if (iX > 0)
                    {
                        tempDGV2.CurrentCell = null;    // 2017/10/03
                    }

                    dr.Close();
                    cArea.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }

                tempDGV2.CurrentCell = null;

            }

            /// ----------------------------------------------------------------------
            /// <summary>
            ///     �`���V�ʂ̗\�薇���Ɣz�z�����\�� </summary>
            /// <param name="tempDGV">
            ///     �f�[�^�O���b�h��</param>
            /// ----------------------------------------------------------------------
            public static void MaisuSubTotal(DataGridView tempDGV)
            {
                if (tempDGV.Rows.Count == 0) return; 
                
                string cName = "";
                int cSum = 0;
                int hSum = 0;
                int i;

                // �\�薇���Ɣz�z�����̍��v�\���@2015/07/15
                for (i = 0; i < tempDGV.Rows.Count ; i++)
                {
                    if ((cName != "") && (cName != tempDGV[1, i].Value.ToString()))
                    {
                        tempDGV[9, i - 1].Value = cSum;     // �\�薇��
                        tempDGV[11, i - 1].Value = hSum;    // �z�z����
                        cSum = 0;
                        hSum = 0;
                    }

                    cSum += Utility.nullToInt(tempDGV[8, i].Value);       // �\�薇��
                    hSum += Utility.nullToInt(tempDGV[10, i].Value);      // �z�z����
                    cName = tempDGV[1, i].Value.ToString();
                }

                tempDGV[9, i - 1].Value = cSum;
                tempDGV[11, i - 1].Value = hSum;
            }

            ///--------------------------------------------------------
            /// <summary>
            ///     �����敪��[1]�̃f�[�^�͐Ԗ��� </summary>
            /// <param name="tempDGV">
            ///     �f�[�^�O���b�h�r���[��</param>
            /// <param name="ex">
            ///     �s�ԍ�</param>
            ///--------------------------------------------------------
            public static void KanryoColorShow(DataGridView tempDGV,int ex)
            {
                //�����敪
                if (tempDGV[COLUMN_KANRYO, ex].Value.ToString() == STATUS_KANRYO)
                {
                    tempDGV.Rows[ex].DefaultCellStyle.ForeColor = Color.Red;
                }
                else
                {
                    tempDGV.Rows[ex].DefaultCellStyle.ForeColor = Color.Black;
                }

                tempDGV.CurrentCell = null;
            }
        }

        private class DataTemp
        {
            //�\���f�[�^
            private string F_ID;        //ID

            public string ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }

            private string F_CName;     //�`���V��

            public string CName
            {
                get { return F_CName; }
                set { F_CName = value; }
            }

            private string F_HDate;     //�z�z��

            public string HDate
            {
                get { return F_HDate; }
                set { F_HDate = value; }
            }

            private string F_IDate;     //���͓�

            public string IDate
            {
                get { return F_IDate; }
                set { F_IDate = value; }
            }

            private string F_Name;         //�z�z������

            public string Name
            {
                get { return F_Name; }
                set { F_Name = value; }
            }

            private string F_Kotsuhi;   //��ʔ�

            public string Kotsuhi
            {
                get { return F_Kotsuhi; }
                set { F_Kotsuhi = value; }
            }

            private string F_AddDate;   //�o�^�N����

            public string AddDate
            {
                get { return F_AddDate; }
                set { F_AddDate = value; }
            }
            private string F_UppDate;

            public string UppDate       //�ύX�N����
            {
                get { return F_UppDate; }
                set { F_UppDate = value; }
            }

            private string F_Memo;      //�{���̒��ӎ���

            public string Memo
            {
                get { return F_Memo; }
                set { F_Memo = value; }
            }

            // ���O�C�����[�U�[
            public string loginUser { get; set; }
        }

        // �O���b�h����f�[�^��I��
        private void GridEnter(int tempiX)
        {
            try
            {
                // �z�z�w���f�[�^���擾����
                if (!GridviewSet.GetData(dataGridView1, ref cMaster, tempiX))
                {
                    MessageBox.Show("�Y������f�[�^���o�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // �z�z�w���f�[�^�l��\��
                jDate.Value = cMaster.�z�z��;

                // �z�z�����\��
                txtStaffID.Text = cMaster.�z�z��ID.ToString();

                OleDbDataReader dR;
                Control.�z�z�� cStaff = new Control.�z�z��();
                dR = cStaff.FillBy("where ID = " + cMaster.�z�z��ID.ToString());
                while (dR.Read())
                {
                    label13.Text = dR["����"].ToString();                        
                }

                dR.Close();
                cStaff.Close();

                // ���͓�
                nDate.Value = cMaster.���͓�;
                txthID.Text = cMaster.ID.ToString();
                txtKoutsu.Text = cMaster.��ʔ�.ToString("#,##0");
                txtChuui.Text = cMaster.���ӎ���;

                Control.�� cOrder = new Control.��();

                // �{�^�����
                btnDel.Enabled = true;
                btnClr.Enabled = true;

                fMode.Mode = 1;     // �t�H�[�����[�h�X�e�[�^�X:�ύX�폜

                jDate.Focus();

                GridViewEnable(dataGridView1, false);

                button1.Enabled = false;

                // ���o�^�`���V�f�[�^�̃X�e�[�^�X��߂�
                StatusBack();

                // �z�z�G���A�f�[�^�\��
                GridviewSet.ShowDataItem(dataGridView1, dataGridView2,tempiX);

                // �`���V�ʖ����\��
                GridviewSet.MaisuSubTotal(dataGridView2);

                // �z�z�w��������{�^��
                if (dataGridView2.Rows.Count > 0)
                {
                    button4.Enabled = true;
                }

                // �x���T�����דo�^�{�^��
                button6.Enabled = true;

                // �V��\��
                tenkouUpdate();

                // ���������ׂ����鎞
                if (GetHaifuKanryo(dataGridView2) == false)
                {
                    // ���͓��������̓��t�Ƃ���
                    nDate.Value = DateTime.Parse(DateTime.Today.ToShortDateString());
                     
                    // �z�z����O���̓��t�Ƃ��� 2010/1/18
                    jDate.Value = DateTime.Parse(DateTime.Today.AddDays(-1).ToShortDateString());
                }

                // ��@���L�����𒍈ӎ����ɕ\�� 2015/11/27
                txtChuui.Text = setTokkijikou(cMaster.ID);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "�f�[�^�\��", MessageBoxButtons.OK);
            }
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     �z�z�w�����ɑΉ�����󒍃f�[�^�̓��L�������擾���� </summary>
        /// <param name="sID">
        ///     �z�z�w��ID</param>
        /// <returns>
        ///     ���L����������</returns>
        ///--------------------------------------------------------------------------------------
        private string setTokkijikou(int sID)
        {
            darwinDataSetTableAdapters.��1TableAdapter dAdp = new darwinDataSetTableAdapters.��1TableAdapter();
            darwinDataSetTableAdapters.�z�z�G���ATableAdapter aAdp = new darwinDataSetTableAdapters.�z�z�G���ATableAdapter();
 
            darwinDataSet dts = new darwinDataSet();

            dAdp.Fill(dts.��1);
            aAdp.Fill(dts.�z�z�G���A);

            var s = dts.�z�z�G���A.Where(a => a.�z�z�w��ID == sID && a.��1Row != null)
                                .Select(b => new
                                {
                                    sjID = b.�z�z�w��ID,
                                    sjMemo = b.��1Row.���L����
                                });

            string wMemo = string.Empty;
            string w = string.Empty;

            foreach (var t in s.OrderBy(a => a.sjMemo))
            {
                if (w != t.sjMemo)
                {
                    if (wMemo == string.Empty)
                    {
                        wMemo = t.sjMemo;
                    }
                    else
                    {
                        ////�����̓��L�����̏ꍇ�͘A�����Ēǉ����܂�
                        //wMemo += ("�@�@" + t.sjMemo);

                        //�����̓��L�����̏ꍇ�͉��s���Ēǉ����܂� 2016/04/04
                        wMemo += Environment.NewLine + t.sjMemo;
                    }
                }

                w = t.sjMemo; 
            }

            // ���L������Ԃ�
            return wMemo;
        }
        
        // �z�z�G���A�O���b�h����f�[�^��I��
        private void GridEnter_Haifu(int tempRow)
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView2[1, tempRow].Value + " " + dataGridView2[4, tempRow].Value + " ���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�`���V�f�[�^�I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {

                    //�`���V�f�[�^���擾����
                    //if (GridviewSet.GetDataItem(Int32.Parse(dataGridView2[1, dataGridView2.SelectedRows[iX].Index].Value.ToString()),ref cArea) == false)
                    //{
                    //    MessageBox.Show("�Y������f�[�^���o�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    //    return;
                    //}

                    //�`���V�f�[�^�ҏW��ʂ�\��
                    frmHaifuShijiSUb2 frm = new frmHaifuShijiSUb2();
                    iX = tempRow;

                    if (txthID.Text.ToString() == "")
                    {
                        frm.SID = 0;
                    }
                    else
                    {
                        frm.SID = Int32.Parse(txthID.Text.ToString(), System.Globalization.NumberStyles.Any);
                    }

                    frm.hDate = jDate.Value.ToShortDateString();
                    frm.staffName = label13.Text;
                    frm.ID = Int32.Parse(dataGridView2[0, iX].Value.ToString());
                    frm.cName = dataGridView2[1, iX].Value.ToString();
                    frm.fJyouken = dataGridView2[2, iX].Value.ToString();
                    frm.fKeitai = dataGridView2[3, iX].Value.ToString();
                    frm.Add = dataGridView2[4, iX].Value.ToString();
                    frm.Tanka = double.Parse(dataGridView2[7, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                    frm.yMaisu = Int32.Parse(dataGridView2[8, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                    frm.hMaisu = Int32.Parse(dataGridView2[10, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                    frm.kanryo = Int32.Parse(dataGridView2[12, iX].Value.ToString(), System.Globalization.NumberStyles.Any);    // 2015/07/15
                    frm.Edaban = dataGridView2[5, iX].Value.ToString();
                    frm.Edaban_Status = int.Parse(dataGridView2[17, iX].Value.ToString());    // 2015/07/15

                    //���z�z���֘A
                    frm.Mihaifu_Status = int.Parse(dataGridView2[16, iX].Value.ToString());    // 2015/07/15
                    //frm.Banchi = dataGridView2[17, iX].Value.ToString();
                    //frm.Manshon = dataGridView2[18, iX].Value.ToString();
                    //frm.Riyu =  int.Parse(dataGridView2[19, iX].Value.ToString());
                    //frm.Sonota = dataGridView2[20, iX].Value.ToString();

                    //�ҏW���
                    frm.ShowDialog();

                    //�l��߂�
                    //dataGridView2[7, iX].Value = frm.Tanka;     //�P��
                    //dataGridView2[10, iX].Value = frm.hMaisu;   //�z�z����
                    //dataGridView2[12, iX].Value = frm.kanryo;   //�����敪    // 2015/07/15
                    dataGridView2[5, iX].Value = frm.Edaban;    //�}�ԋL��

                    //dataGridView2[17, iX].Value = frm.Banchi;   //�Ԓn�E��
                    //dataGridView2[18, iX].Value = frm.Manshon;  //�}���V������
                    //dataGridView2[19, iX].Value = frm.Riyu;     //���R
                    //dataGridView2[20, iX].Value = frm.Sonota;   //���̑�

                    //�����f�[�^�͐ԕ\��
                    GridviewSet.KanryoColorShow(dataGridView2, iX);

                    frm.Dispose();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�f�[�^�\��", MessageBoxButtons.OK);
                }
            }

            //�o�^��ʏI����A���̍s�փJ�[�\�����ړ�����@2009/10/29
            if (tempRow == dataGridView2.RowCount - 1)
            {
                dataGridView2.CurrentCell = dataGridView2[1, 0];�@//�ŉ��s�̂Ƃ��͍ŏ�s�ֈړ�����
            }
            else
            {
                dataGridView2.CurrentCell = dataGridView2[1, tempRow + 1];
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            // Enterkey�ȊO�͑ΏۊO
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            //�m�F
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += "�w���ԍ� " + dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value + " ���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�z�z�w���E�񍐏��I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }

            GridEnter(dataGridView1.SelectedRows[iX].Index);
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //GridEnter();
        }

        /// <summary>
        /// ��ʂ��N���A����
        /// </summary>
        private void DispClear()
        {

            try
            {
                //StartDate.Value = DateTime.Today;
                //StartDate.Checked = false;
                //EndDate.Value = DateTime.Today;
                //EndDate.Checked = false;

                //nStartDate.Value = DateTime.Today;
                //nStartDate.Checked = false;
                //nEndDate.Value = DateTime.Today;
                //nEndDate.Checked = false;

                cmbsStaff.SelectedIndex = -1;

                button1.Enabled = true;

                txthID.Text = "";

                jDate.Value = DateTime.Today.AddDays(-1);   //�f�t�H���g�͑O�� 2010/1/18
                nDate.Value = DateTime.Today;

                //mStartTime.Text = "";
                //mEndTime.Text = "";

                //cmbStaff.SelectedIndex = -1;

                txtStaffID.Text = "";
                label13.Text = "";

                txtKoutsu.Text = "0";

                //cmbRiyu.SelectedIndex = -1;
                //txtMemo2.Text = "";

                txtChuui.Text = "";
                txtTenkou.Text = "";

                btnDel.Enabled = false;
                //btnClr.Enabled = false;

                dataGridView1.Enabled = true;

                //txtCode.Focus();
                jDate.Focus();

                dataGridView2.RowCount = 0;

                button4.Enabled = false;
                button6.Enabled = false;

                fMode.Mode = 0;

                GridViewEnable(dataGridView1, true);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }

        /// <summary>
        /// ���o�^�`���V�f�[�^�̃X�e�[�^�X��0�ɖ߂�
        /// </summary>
        private void StatusBack()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2[14, i].Value.ToString() == "1")
                {
                    //�f�[�^�x�[�X�X�V
                    HaihuAreaUpdate(Int32.Parse(dataGridView2[0, i].Value.ToString()),
                     (int)(0), (int)(0), (int)(0), (int)(0), (int)(0), (int)(0),"");

                    //���z�z���폜
                    string sqlStr;
                    Control.FreeSql fCon = new Control.FreeSql();
                    sqlStr = "";
                    sqlStr += "delete from ���z�z��� ";
                    sqlStr += "where �z�z�G���AID = " + Int32.Parse(dataGridView2[0, i].Value.ToString());

                    if (fCon.Execute(sqlStr) == false)
                    {
                        MessageBox.Show("���z�z���̍폜�Ɏ��s���܂���", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                    fCon.Close();
                }
            }

        }

        /// <summary>
        /// �f�[�^�O���b�h�r���[�X�e�[�^�X�\���ؑ�
        /// </summary>
        /// <param name="tempDGV">�f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        /// <param name="tempBool">��ԁitrue:false�j</param>
        private void GridViewEnable(DataGridView tempDGV,bool tempBool)
        {
            if (tempBool == true)
            {
                for (int i = 0; i < tempDGV.Rows.Count; i++)
                {
                    tempDGV.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
                }

                tempDGV.Enabled = true;
            }
            else
            {
                for (int i = 0; i < tempDGV.Rows.Count; i++)
                {
                    tempDGV.Rows[i].DefaultCellStyle.ForeColor = Color.LightGray;
                }

                tempDGV.Enabled = false;
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�I������Ă���f�[�^��ύX���Ȃ��ŏI�����܂��B��낵���ł����H","�m�F",MessageBoxButtons.YesNo,MessageBoxIcon.Question)== DialogResult.No )
                return;

            //���o�^�`���V�f�[�^�𖢑I����Ԃɖ߂�
            StatusBack();

            //��ʕ\����������
            DispClear();
        }

        //�z�z�������G���A�f�[�^�����邩���f
        private bool GetHaifuKanryo(DataGridView dGv)
        {
            //int iX = 0;

            bool rtn = true;

            foreach (DataGridViewRow  r in dGv.Rows)
            {
                // �����E�������̔��f���C���F2017/05/18
                if (dGv["col12", dGv.Rows[r.Index].Index].Value.ToString() == STATUS_MIKANRYO)
                {
                    rtn = false;
                    break;
                }
            }

            return rtn;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                //// �z�z���͕K�{���́F2018/01/04
                //if (Utility.strToInt(txtStaffID.Text) == 0)
                //{
                //    MessageBox.Show("�z�z���������͂ł�", "�m�F", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    return;
                //}

                // �S�Ă̖��ׂ��z�z���������ׂ� 2018/03/02
                bool kanryo = true;
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    if (dataGridView2[12, i].Value.ToString() != STATUS_KANRYO)
                    {
                        kanryo = false;
                        break;
                    }
                }

                // �S�Ă̖��ׂ��z�z�����Ŕz�z���������͂̂Ƃ��A���[�g�\�� 2018/03/02
                if (kanryo && (Utility.strToInt(txtStaffID.Text) == 0))
                {
                    MessageBox.Show("�z�z���������͂ł�", "�m�F", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                
                // �}�Ԗ����̓`�F�b�N�A���[�g�F2018/01/11
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    if (dataGridView2[17, i].Value.ToString() == global.FLGON.ToString())
                    {
                        if (dataGridView2[5, i].Value.ToString() == string.Empty)
                        {
                            string cName = dataGridView2[1, i].Value.ToString();

                            if (MessageBox.Show("�}�Ԗ����͂̔z�z���ׂ�����܂��B���s���Ă�낵���ł����H" + Environment.NewLine + Environment.NewLine + cName, "�}�Ԗ����̓f�[�^����", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                            {
                                dataGridView2.Focus();
                                dataGridView2.CurrentCell = dataGridView2[5, i];
                                return;
                            }
                        }
                    }
                }

                // �z�z�����O�Ŕz�z�����`�F�b�N�F2018/02/20
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    if (Utility.strToInt(dataGridView2[10, i].Value.ToString()) == global.FLGOFF)
                    {
                        if (dataGridView2[12, i].Value.ToString() == STATUS_KANRYO)
                        {
                            string cName = dataGridView2[1, i].Value.ToString();

                            if (MessageBox.Show("�z�z�����O�̒��ڂ�����܂��B��낵���ł����H" + Environment.NewLine + Environment.NewLine + cName, "�z�z�����O���ڂ���", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                            {
                                dataGridView2.Focus();
                                dataGridView2.CurrentCell = dataGridView2[10, i];
                                return;
                            }
                        }
                    }
                }

                if (!GetHaifuKanryo(dataGridView2))
                {
                    if (MessageBox.Show("�������̔z�z���ׂ�����܂��B���s���Ă�낵���ł����H", "�������f�[�^����", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) return;
                }

                if (MessageBox.Show("�o�^���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                            
                if (fDataCheck())
                {
                    Control.DataControl Con;
                    OleDbConnection cn;
                    OleDbTransaction tran;
                    OleDbCommand SCom;

                    switch (fMode.Mode)
                    {
                        case 0: 
                        
                            //ID���̔�
                            string sqlStr = "";
                            int gID = (int)(1);

                            sqlStr = "select max(ID) as ID from �z�z�w�� ";
                            OleDbDataReader dR;
                            Control.FreeSql fCon = new Control.FreeSql();
                            dR = fCon.free_dsReader(sqlStr);

                            while (dR.Read())
                            {
                                if (dR["ID"] == DBNull.Value )
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
                                // �z�z�w���f�[�^�o�^����                                
                                sqlStr = "";
                                sqlStr += "insert into �z�z�w�� ";
                                sqlStr += "(ID,�z�z��,���͓�,�z�z��ID,��ʔ�,��ʋ�ԊJ�n,��ʋ�ԏI��,";
                                sqlStr += "�z�z�J�n����,�z�z�I������,�I�����|�[�g,���z�z�敪,���z�z���R,";
                                sqlStr += "�o�^�N����,�ύX�N����,���[�U�[ID) ";
                                sqlStr += "values (";
                                sqlStr += cMaster.ID + ",";
                                sqlStr += "'" + cMaster.�z�z�� + "',";
                                sqlStr += "'" + cMaster.���͓� + "',";
                                sqlStr += cMaster.�z�z��ID + ",";
                                sqlStr += cMaster.��ʔ� + ",";
                                sqlStr += "'" + cMaster.��ʋ�ԊJ�n + "',";
                                sqlStr += "'" + cMaster.��ʋ�ԏI�� + "',";
                                sqlStr += "'" + cMaster.�z�z�J�n���� + "',";
                                sqlStr += "'" + cMaster.�z�z�I������ + "',";
                                sqlStr += "'" + cMaster.�I�����|�[�g + "',";
                                sqlStr += "'" + cMaster.���z�z�敪 + "',";
                                sqlStr += "'" + cMaster.���z�z���R + "',";
                                sqlStr += "'" + cMaster.�o�^�N���� + "',";
                                sqlStr += "'" + cMaster.�ύX�N���� + "',";
                                sqlStr += global.loginUserID.ToString() + ")";  // 2016/09/26 ���[�U�[ID�ǉ�

                                SCom.CommandText = sqlStr;

                                SCom.ExecuteNonQuery();

                                //�z�z�G���A�f�[�^�X�V
                                string sID, sTanka, sMaisu, sMaisu2, sKanryo, sStatus, sEdaban;

                                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                                {
                                    if (dataGridView2[14, i].Value.ToString() != "1") //�폜�t���O���f
                                    {
                                        sID = dataGridView2[0, i].Value.ToString();
                                        sTanka = Utility.nullToDouble(dataGridView2[7, i].Value).ToString();
                                        sMaisu = Utility.nullToInt(dataGridView2[10, i].Value).ToString();
                                        //sMaisu2 = dataGridView2[12, i].Value.ToString();
                                        //sKanryo = dataGridView2[11, i].Value.ToString();
                                        sMaisu2 = dataGridView2[13, i].Value.ToString();

                                        // 2016/10/31
                                        if (dataGridView2[12, i].Value.ToString() == STATUS_KANRYO)
                                        {
                                            sKanryo = STATUS_KANRYO;
                                        }
                                        else
                                        {
                                            sKanryo = STATUS_MIKANRYO;
                                        }

                                        sStatus = "2";
                                        sEdaban = dataGridView2[5, i].Value.ToString() + "";

                                        //mBanchi = dataGridView2[17, i].Value.ToString() + "";
                                        //mManshon = dataGridView2[18, i].Value.ToString() + "";
                                        //mRiyu = int.Parse(dataGridView2[19, i].Value.ToString(),System.Globalization.NumberStyles.Any);
                                        //mSonota = dataGridView2[20, i].Value.ToString() + ""; 

                                        sqlStr = "";
                                        sqlStr += "update �z�z�G���A ";
                                        sqlStr += "set ";
                                        sqlStr += "�z�z�w��ID = " + gID.ToString() + ",";
                                        sqlStr += "�z�z�P�� = " + sTanka + ",";
                                        sqlStr += "���z�z���� = " + sMaisu + ",";
                                        sqlStr += "���c�� = �\�薇�� - " + sMaisu + ",";
                                        sqlStr += "�񍐖��� = " + sMaisu2 + ",";
                                        sqlStr += "�񍐎c�� = �\�薇�� - " + sMaisu2 + ",";
                                        sqlStr += "�����敪 = " + sKanryo + ",";
                                        sqlStr += "�X�e�[�^�X = " + sStatus + ",";
                                        sqlStr += "�}�ԋL�� = '" + sEdaban + "',";

                                        //sqlStr += "�Ԓn�� = '" + mBanchi + "',";
                                        //sqlStr += "�}���V������ = '" + mManshon + "',";
                                        //sqlStr += "���R = " + mRiyu.ToString() + ",";
                                        //sqlStr += "���̑����e = '" + mSonota + "',";

                                        sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
                                        sqlStr += "where (�z�z�G���A.ID = " + sID + ") and ";
                                        sqlStr += "(�X�e�[�^�X <> 0)";

                                        SCom.CommandText = sqlStr;

                                        SCom.ExecuteNonQuery();
                                    }
                                }

                                tran.Commit();
                                MessageBox.Show("�V�K�o�^����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            catch (Exception ex)
                            {
                                tran.Rollback();

                                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                MessageBox.Show("�V�K�o�^�Ɏ��s���܂����B���[���o�b�N���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                //���o�^�z�z�f�[�^�̃X�e�[�^�X��߂�
                                StatusBack();
                            }

                            cn.Close();

                            Con.Close();

                            break;

                        case 1: //�X�V

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
                                //�z�z�w���f�[�^�X�V
                                sqlStr = "";
                                sqlStr += "update �z�z�w�� set ";
                                sqlStr += "�z�z�� = '" + cMaster.�z�z�� + "',";
                                sqlStr += "���͓� = '" + cMaster.���͓� + "',";
                                sqlStr += "�z�z��ID = " + cMaster.�z�z��ID + ",";
                                sqlStr += "��ʔ� = " + cMaster.��ʔ� + ",";
                                sqlStr += "��ʋ�ԊJ�n = '" + cMaster.��ʋ�ԊJ�n + "',";
                                sqlStr += "��ʋ�ԏI�� = '" + cMaster.��ʋ�ԏI�� + "',";
                                sqlStr += "�z�z�J�n���� = '" + cMaster.�z�z�J�n���� + "',";
                                sqlStr += "�z�z�I������ = '" + cMaster.�z�z�I������ + "',";
                                sqlStr += "�I�����|�[�g = '" + cMaster.�I�����|�[�g + "',";
                                sqlStr += "���z�z�敪 = '" + cMaster.���z�z�敪 + "',";
                                sqlStr += "���z�z���R = '" + cMaster.���z�z���R + "',";
                                sqlStr += "���ӎ��� = '" + cMaster.���ӎ��� + "',";
                                sqlStr += "�ύX�N���� = '" + DateTime.Today + "',";
                                sqlStr += "���[�U�[ID = " + global.loginUserID.ToString() + " "; // 2016/09/26
                                sqlStr += "where ID = " + cMaster.ID;

                                SCom.CommandText = sqlStr;

                                // SQL�̎��s
                                SCom.ExecuteNonQuery();

                                //�z�z�G���A�f�[�^�X�V
                                string sID, sTanka, sMaisu, sMaisu2, sKanryo, sStatus, sEdaban;

                                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                                {
                                    if (dataGridView2[14, i].Value.ToString() != "1") //�폜�t���O���f
                                    {
                                        sID = dataGridView2[0, i].Value.ToString();
                                        sTanka = Utility.nullToDouble(dataGridView2[7, i].Value).ToString();
                                        sMaisu = Utility.nullToInt(dataGridView2[10, i].Value).ToString();
                                        //sMaisu2 = dataGridView2[12, i].Value.ToString();
                                        //sKanryo = dataGridView2[11, i].Value.ToString();
                                        sMaisu2 = dataGridView2[13, i].Value.ToString();

                                        // 2016/10/31
                                        if (dataGridView2[12, i].Value.ToString() == STATUS_KANRYO)
                                        {
                                            sKanryo = STATUS_KANRYO;
                                        }
                                        else
                                        {
                                            sKanryo = STATUS_MIKANRYO;
                                        }

                                        sStatus = "2";
                                        sEdaban = dataGridView2[5, i].Value.ToString() + "";

                                        //mBanchi = dataGridView2[17, i].Value.ToString() + "";
                                        //mManshon = dataGridView2[18, i].Value.ToString() + "";
                                        //mRiyu = int.Parse(dataGridView2[19, i].Value.ToString(), System.Globalization.NumberStyles.Any);
                                        //mSonota = dataGridView2[20, i].Value.ToString() + ""; 

                                        sqlStr = "";
                                        sqlStr += "update �z�z�G���A ";
                                        sqlStr += "set ";
                                        sqlStr += "�z�z�w��ID = " + txthID.Text + ",";
                                        sqlStr += "�z�z�P�� = " + sTanka + ",";
                                        sqlStr += "���z�z���� = " + sMaisu + ",";
                                        sqlStr += "���c�� = �\�薇�� - " + sMaisu + ",";
                                        sqlStr += "�񍐖��� = " + sMaisu2 + ",";
                                        sqlStr += "�񍐎c�� = �\�薇�� - " + sMaisu2 + ",";
                                        sqlStr += "�����敪 = " + sKanryo + ",";
                                        sqlStr += "�X�e�[�^�X = " + sStatus + ",";
                                        sqlStr += "�}�ԋL�� = '" + sEdaban + "',";

                                        //sqlStr += "�Ԓn�� = '" + mBanchi + "',";
                                        //sqlStr += "�}���V������ = '" + mManshon + "',";
                                        //sqlStr += "���R = " + mRiyu.ToString() + ",";
                                        //sqlStr += "���̑����e = '" + mSonota + "',";

                                        sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
                                        sqlStr += "where (�z�z�G���A.ID = " + sID + ") and ";
                                        sqlStr += "(�X�e�[�^�X <> 0)";

                                        SCom.CommandText = sqlStr;

                                        // SQL�̎��s
                                        SCom.ExecuteNonQuery();
                                    }
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

                    DispClear();

                    //�O���b�h�ĕ\��
                    //GridDataShow(StartDate, EndDate,nStartDate,nEndDate, cmbsStaff);
                    dataGridView1.RowCount = 0;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"�X�V����",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        //�o�^�f�[�^�`�F�b�N
        private Boolean fDataCheck()
        {

            try
            {

                //�o�^���[�h�̂Ƃ��A�R�[�h���`�F�b�N
                if (fMode.Mode == 0)
                {

                    //// �������H
                    //if (txtCode.Text == null)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("�R�[�h�͐����œ��͂��Ă�������");
                    //}

                    //str = this.txtCode.Text;

                    //if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    //{
                    //}
                    //else
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("�R�[�h�͐����œ��͂��Ă�������");
                    //}

                    //// �����͂܂��̓X�y�[�X�݂͕̂s��
                    //if ((this.txtCode.Text).Trim().Length < 1)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("�R�[�h����͂��Ă�������");
                    //}

                    ////�[���͕s��
                    //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                    //{
                    //    this.txtCode.Focus();
                    //    throw new Exception("�[���͓o�^�ł��܂���");
                    //}

                    ////�o�^�ς݃R�[�h�����ׂ�
                    //string sqlStr;
                    //Control.�� Order = new Control.��();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Order.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Order.Close();
                    //    throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    //}

                    //dr.Close();
                    //Order.Close();

                }

                //�N���X�Ƀf�[�^�Z�b�g
                cMaster.�z�z�� = jDate.Value;
                cMaster.���͓� = nDate.Value;

                if (txtStaffID.Text == "")
                {
                    cMaster.�z�z��ID = 0;
                }
                else
                {
                    cMaster.�z�z��ID = int.Parse(txtStaffID.Text);
                }

                cMaster.��ʔ� = Int32.Parse(txtKoutsu.Text, System.Globalization.NumberStyles.Any);
                cMaster.��ʋ�ԊJ�n = "";
                cMaster.��ʋ�ԏI�� = "";

                cMaster.�z�z�J�n���� = "";
                cMaster.�z�z�I������ = "";

                cMaster.�I�����|�[�g = "";

                cMaster.���z�z�敪 = "";
                cMaster.���z�z���R = "";

                cMaster.���ӎ��� = txtChuui.Text;

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

        private void txtEnter(object sender, EventArgs e)
        {
            TextBox objtxt = new TextBox();
            MaskedTextBox objMtxt = new MaskedTextBox();

            if (sender == txtStaffID)
            {
                objtxt = txtStaffID;
                if (txtStaffID.Text == "0")
                {
                    txtStaffID.Text = "";
                }
            }

            if (sender == txtKoutsu)
            {
                objtxt = txtKoutsu;

                dataGridView2.CurrentCell = null;   // 2017/10/03
            }

            if (sender == txtsID) objtxt = txtsID;
            if (sender == txtsCName) objtxt = txtsCName;

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

            objMtxt.SelectAll();
            objMtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {
            
            TextBox objtxt = new TextBox();
            MaskedTextBox objMtxt = new MaskedTextBox();

            if (sender == txtStaffID) objtxt = txtStaffID;            
            if (sender == txtKoutsu) objtxt = txtKoutsu;
            if (sender == txtsID) objtxt = txtsID;
            if (sender == txtsCName) objtxt = txtsCName;

            objtxt.BackColor = Color.White;
            objMtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //���o�^�`���V�f�[�^�̃X�e�[�^�X��߂�
            StatusBack();

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
                //�z�z�w���f�[�^�폜
                sqlSTR = "";
                sqlSTR += "delete from �z�z�w�� ";
                sqlSTR += "where ID = " + cMaster.ID.ToString();

                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                //���z�z���̍폜
                sqlSTR = "";
                sqlSTR += "delete a FROM ���z�z��� as a inner join �z�z�G���A as b ";
                sqlSTR += "on a.�z�z�G���AID = b.ID ";
                sqlSTR += "where b.�z�z�w��ID = " + cMaster.ID.ToString();
                SCom.CommandText = sqlSTR;

                SCom.ExecuteNonQuery();

                //�z�z�G���A�f�[�^�̏�����
                sqlSTR = "";
                sqlSTR += "update �z�z�G���A ";
                sqlSTR += "set ";
                sqlSTR += "�z�z�w��ID = 0,";
                //sqlSTR += "�z�z�P�� = 0,";
                sqlSTR += "���z�z���� = 0,";
                sqlSTR += "���c�� = �\�薇��,";
                sqlSTR += "�񍐖��� = 0,";
                sqlSTR += "�񍐎c�� = �\�薇��,";
                sqlSTR += "�����敪 = 0,";
                sqlSTR += "�X�e�[�^�X = 0,";
                sqlSTR += "�}�ԋL�� = '',";

                //sqlSTR += "�Ԓn�� = '',";
                //sqlSTR += "�}���V������ = '',";
                //sqlSTR += "���R = 0,";
                //sqlSTR += "���̑����e = '',";

                sqlSTR += "�ύX�N���� = '" + DateTime.Today + "' ";
                sqlSTR += "where �z�z�G���A.�z�z�w��ID = " + cMaster.ID.ToString();

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

            DispClear();

            //�O���b�h�ĕ\��
            //GridDataShow(StartDate, EndDate, nStartDate, nEndDate, cmbsStaff);
            dataGridView1.RowCount = 0;
        }

        private void btnCsv_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView1, MESSAGE_CAPTION);
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            GridEnter(0);
        }

        //private void cmbRiyuSet()
        //{
        //    cmbRiyu.Items.Clear();
        //    cmbRiyu.Items.Add("�Ǘ��l");
        //    cmbRiyu.Items.Add("�Z��");
        //    cmbRiyu.Items.Add("���d�\��");
        //    cmbRiyu.Items.Add("���̑�");
        //}


        private void button1_Click(object sender, EventArgs e)
        {
            if ((txtsID.Text.Trim().Length > 0) && (Utility.NumericCheck(txtsID.Text) == false))
            {
                MessageBox.Show("�����p�z�z�w��ID�͐����œ��͂��Ă�������", "�����z�z�w��ID", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            GridDataShow(cmbsStaff,txtsID,txtsCName);
        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            // �m�F
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += "�w���ԍ� " + dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value + " ���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�z�z�w���E�񍐏��I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }

            Cursor = Cursors.WaitCursor;
            GridEnter(dataGridView1.SelectedRows[iX].Index);
            Cursor = Cursors.Default;
        }

        //////private void GridDataShow(DateTimePicker d1, DateTimePicker d2, DateTimePicker d3, DateTimePicker d4, ComboBox tempCmb, TextBox tempID,TextBox tempCName)

        private void GridDataShow(ComboBox tempCmb, TextBox tempID,TextBox tempCName)
        {
            //DateTime Date1;
            //DateTime Date2;
            //DateTime Date3;
            //DateTime Date4;

            int SID;

            ////////�z�z�J�n��
            //////if (d1.Checked == true)
            //////{
            //////    Date1 = d1.Value;
            //////}
            //////else
            //////{
            //////    Date1 = Convert.ToDateTime("1900/01/01");
            //////}

            ////////�z�z�I����
            //////if (d2.Checked == true)
            //////{
            //////    Date2 = d2.Value;
            //////}
            //////else
            //////{
            //////    Date2 = Convert.ToDateTime("2999/12/31");
            //////}

            ////////���͊J�n��
            //////if (d3.Checked == true)
            //////{
            //////    Date3 = d3.Value;
            //////}
            //////else
            //////{
            //////    Date3 = Convert.ToDateTime("1900/01/01");
            //////}

            ////////���͏I����
            //////if (d4.Checked == true)
            //////{
            //////    Date4 = d4.Value;
            //////}
            //////else
            //////{
            //////    Date4 = Convert.ToDateTime("2999/12/31");
            //////}

            //�z�z��ID
            if (tempCmb.SelectedIndex != -1)
            {
                Utility.ComboStaff cmb1 = new Utility.ComboStaff();
                cmb1 = (Utility.ComboStaff)tempCmb.SelectedItem;
                SID = cmb1.ID;
            }
            else
            {
                SID = (int)(-1);
            }

            this.Cursor = Cursors.WaitCursor;

            //�O���b�h�\��
            GridviewSet.ShowData(dataGridView1,SID,tempID.Text,tempCName.Text);
            this.Cursor = Cursors.Default;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //���̃}�V���Ŕz�z�w�����s�������؂���
            int cFlg = 0;

            OleDbDataReader dRflg;
            Control.��Џ�� cSystem = new Control.��Џ��();
            dRflg = cSystem.Fill();

            while (dRflg.Read())
            {
                cFlg = int.Parse(dRflg["�z�z�t���O"].ToString());
            }

            dRflg.Close();
            cSystem.Close();

            if (cFlg == 1)
            {
                MessageBox.Show("���݁A���̃}�V���Ŕz�z�w���o�^���ł��B", "�N���`�F�b�N", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //���w���z�z�f�[�^�̓o�^
            frmHaifuShijiSUb fSub = new frmHaifuShijiSUb();

            long jID;
            int cID;
            string msgHD = "";
            string msgSTR = "";
            string d1, d2;

            if (fSub.ShowDialog(this) == DialogResult.OK)
            {

                for (int i = 0; i < fSub.Count; i++)
                {
                    //�����A������ID�A��������ID�̓o�^�̂Ƃ����b�Z�[�W�\��
                    //�@��ID�ƒ���ID���擾
                    OleDbDataReader dR,dR2;
                    Control.�z�z�G���A cArea = new Control.�z�z�G���A();
                    dR = cArea.FillBy("where ID = " + fSub[i].ToString());

                    while (dR.Read())
                    {
                        jID = long.Parse(dR["��ID"].ToString());
                        cID = int.Parse(dR["����ID"].ToString());

                        string sqlStr;
                        Control.FreeSql fCon = new Control.FreeSql();

                        sqlStr = "";
                        sqlStr += "select ��.�`���V��,����.����,�z�z�G���A.�z�z�w��ID,�z�z�G���A.�\�薇��,";
                        sqlStr += "�z�z�w��.�z�z��,�z�z�G���A.����ID ";
                        sqlStr += "from �z�z�G���A ";
                        sqlStr += "inner join �z�z�w�� on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                        sqlStr += "left join �� on �z�z�G���A.��ID = ��.ID ";
                        sqlStr += "left join ���� on �z�z�G���A.����ID = ����.ID ";
                        sqlStr += "where ";
                        sqlStr += "(�z�z�G���A.��ID = " + jID.ToString() + ") and ";
                        sqlStr += "(�z�z�G���A.����ID = " + cID.ToString() + ") and ";
                        sqlStr += "(�z�z�G���A.�z�z�w��ID != 0) ";

                        dR2 = fCon.free_dsReader(sqlStr);

                        msgHD = "";
                        msgSTR = "";

                        while(dR2.Read())
                        {
                            if (dR2["�`���V��"] == DBNull.Value )
                            {
                                d1 = "�`���V���F�@";
                            }
                            else
                            {
                                d1 = "�`���V���F" + dR2["�`���V��"].ToString() + "�@";
                            }

                            if (dR2["����"] == DBNull.Value)
                            {
                                d2 = dR2["����ID"].ToString() + "�@";
                            }
                            else
                            {
                                d2 = dR2["����ID"].ToString() + "�F" + dR2["����"].ToString();
                            }

                            msgHD = d1 + d2 + Environment.NewLine + Environment.NewLine;
                            
                            msgSTR += "�w���ԍ��F" + dR2["�z�z�w��ID"].ToString() + "�@";
                            msgSTR += "�z�z���F" + DateTime.Parse(dR2["�z�z��"].ToString()).ToShortDateString() + "�@";
                            msgSTR += "�z�z�����F" + dR2["�\�薇��"].ToString() + Environment.NewLine;
                        }

                        dR2.Close();
                        fCon.Close();

                        //�A���b�Z�[�W�\��
                        if (msgSTR != "")
                        {
                            MessageBox.Show(msgHD + msgSTR, "�ߋ��̓��l�̔z�z�w��",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        }

                    }

                    dR.Close();
                    cArea.Close();



                    //�O���b�h���׍s�ǉ�
                    HaifuItemAdd(int.Parse(fSub[i].ToString()));

                    //�z�z�G���A�f�[�^�X�e�[�^�X��[1]�ɏ�������
                    HaihuStatusUpdate(int.Parse(fSub[i].ToString()));   
                }

                button4.Enabled = false;
                
            }
            else
            {
            }

            fSub.Dispose();

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            string sqlSTRING;

            if (e.ColumnIndex == 0)
            {

                //dataGridView2[2, e.RowIndex].Value = GetTownName(dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString());

                OleDbDataReader dR;

                sqlSTRING = "";
                sqlSTRING += "select �z�z�G���A.ID,�z�z�G���A.��ID,��.�`���V��,";
                sqlSTRING += "�z�z�G���A.����ID,����.���� as ����,��.�z�z�J�n��,��.�z�z�I����,";
                sqlSTRING += "�z�z�G���A.���z�敪,�z�z�G���A.�z�z�P��,�z�z�G���A.�����敪,";
                sqlSTRING += "��.�z�z����,�z�z�`��.���� as �z�z�`��,�z�z�G���A.�\�薇��,";
                sqlSTRING += "�z�z�G���A.���z�z����,�z�z�G���A.�񍐖���,�z�z�G���A.�}�ԋL��,";
                sqlSTRING += "��.���z�z���L��,��.�}�ԗL�� ";

                //sqlSTRING += "�z�z�G���A.�Ԓn��,�z�z�G���A.�}���V������,�z�z�G���A.���R,�z�z�G���A.���̑����e ";
            
                sqlSTRING += "from ((�z�z�G���A left join �� on �z�z�G���A.��ID = ��.ID) ";
                sqlSTRING += "left join ���� on �z�z�G���A.����ID = ����.ID) ";
                sqlSTRING += "left join �z�z�`�� on ��.�z�z�`�� = �z�z�`��.ID ";
                sqlSTRING += "where �z�z�G���A.ID = " + dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString();

                //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����
                Control.FreeSql cArea = new Control.FreeSql();
                dR = cArea.free_dsReader(sqlSTRING);

                //�O���b�h�r���[�ɕ\������
                while (dR.Read())
                {
                    try
                    {
                        // �`���V�ʖ����\���X�e�[�^�X
                        STATUS_MAISU = false;

                        // �O���b�h�Ƀf�[�^�\��
                        dataGridView2[1, e.RowIndex].Value = dR["�`���V��"].ToString();
                        dataGridView2[2, e.RowIndex].Value = dR["�z�z����"].ToString();
                        dataGridView2[3, e.RowIndex].Value = dR["�z�z�`��"].ToString();
                        dataGridView2[4, e.RowIndex].Value = dR["����"].ToString();
                        dataGridView2[5, e.RowIndex].Value = dR["�}�ԋL��"].ToString();
                        dataGridView2[6, e.RowIndex].Value = dR["����ID"].ToString();
                        dataGridView2[7, e.RowIndex].Value = Double.Parse(dR["�z�z�P��"].ToString(),System.Globalization.NumberStyles.Any).ToString("#,##0.00");
                        dataGridView2[8, e.RowIndex].Value = int.Parse(dR["�\�薇��"].ToString());
                        dataGridView2[10, e.RowIndex].Value = int.Parse(dR["���z�z����"].ToString());
                        dataGridView2[12, e.RowIndex].Value = int.Parse(dR["�����敪"].ToString());    // 2015/07/15
                        dataGridView2[13, e.RowIndex].Value = int.Parse(dR["�񍐖���"].ToString());    // 2015/07/15
                        dataGridView2[14, e.RowIndex].Value = "0"; //�폜�t���O    // 2015/07/15
                        dataGridView2[15, e.RowIndex].Value = "0"; //�ǉ��t���O    // 2015/07/15
                        dataGridView2[16, e.RowIndex].Value = int.Parse(dR["���z�z���L��"].ToString());    // 2015/07/15
                        dataGridView2[17, e.RowIndex].Value = int.Parse(dR["�}�ԗL��"].ToString());    // 2015/07/15

                        //dataGridView2[18, e.RowIndex].Value = dR["�Ԓn��"].ToString();
                        //dataGridView2[19, e.RowIndex].Value = dR["�}���V������"].ToString();
                        //dataGridView2[20, e.RowIndex].Value = int.Parse(dR["���R"].ToString(),System.Globalization.NumberStyles.Any);
                        //dataGridView2[21, e.RowIndex].Value = dR["���̑����e"].ToString();

                        // �`���V�ʖ����\���X�e�[�^�X
                        STATUS_MAISU = true;

                        // �}�ԃZ���ҏW�F2017/10/03
                        if (Utility.strToInt(dR["�}�ԗL��"].ToString()) != 0)
                        {
                            dataGridView2[5, e.RowIndex].ReadOnly = false;
                        }
                        else
                        {
                            dataGridView2[5, e.RowIndex].ReadOnly = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }

                    //�����f�[�^�͐ԕ\��
                    GridviewSet.KanryoColorShow(dataGridView2, e.RowIndex);
                }

                dR.Close();
                cArea.Close();
            }
            else if (e.ColumnIndex == 12)
            {
                //�����f�[�^�͐ԕ\��
                GridviewSet.KanryoColorShow(dataGridView2, e.RowIndex);
            }
            else if (e.ColumnIndex == 10)
            {
                // �`���V�ʖ����\��
                if (STATUS_MAISU)
                {
                    GridviewSet.MaisuSubTotal(dataGridView2);
                }
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     �z�z���׍s�ǉ� </summary>
        /// <param name="tempID">
        ///     �z�z�G���AID </param>
        ///----------------------------------------------------------------
        private void HaifuItemAdd(int tempID)
        {
            int iX;
            dataGridView2.Rows.Add();
            iX = dataGridView2.Rows.Count;

            dataGridView2[0, iX - 1].Value = tempID.ToString();
            dataGridView2[15, iX - 1].Value = "1";  //�ǉ��t���O    // 2015/07/15

        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     �z�z�G���A�f�[�^�̃X�e�[�^�X��"�o�^��"(1)�ɂ��� </summary>
        /// <param name="tempID">
        ///     �z�z�G���AID</param>
        ///----------------------------------------------------------------
        private void HaihuStatusUpdate(int tempID)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update �z�z�G���A ";
            sqlStr += "set �X�e�[�^�X = 1, ";
            sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
            sqlStr += "where �z�z�G���A.ID = " + tempID.ToString();
            
            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("�z�z�G���A�f�[�^�̃X�e�[�^�X�X�V�Ɏ��s���܂���(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     �z�z�w���f�[�^���X�V���� </summary>
        /// <param name="tempID">
        ///     ID</param>
        /// <param name="tempShijiID">
        ///     �z�z�w��ID</param>
        /// <param name="tempTanka">
        ///     �P��</param>
        /// <param name="temphMaisu">
        ///     �z�z����</param>
        /// <param name="temphMaisu2">
        ///     �񍐖���</param>
        /// <param name="tempKanryo">
        ///     �����敪</param>
        /// <param name="tempStatus">
        ///     �X�e�[�^�X</param>
        ///-----------------------------------------------------------------
        private void HaihuAreaUpdate(int tempID,int tempShijiID,double tempTanka,int temphMaisu,
            �@�@�@�@�@�@�@�@�@�@�@�@ int temphMaisu2,int tempKanryo,int tempStatus,string tempEdaban)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update �z�z�G���A ";
            sqlStr += "set ";
            sqlStr += "�z�z�w��ID = " + tempShijiID + ",";
            //sqlStr += "�z�z�P�� = " + tempTanka + ",";
            sqlStr += "���z�z���� = " + temphMaisu + ",";
            sqlStr += "���c�� = �\�薇�� - " + temphMaisu + ",";
            sqlStr += "�񍐖��� = " + temphMaisu2 + ",";
            sqlStr += "�񍐎c�� = �\�薇�� - " + temphMaisu2 + ",";
            sqlStr += "�����敪 = " + tempKanryo + ",";
            sqlStr += "�X�e�[�^�X = " + tempStatus + ",";
            sqlStr += "�}�ԋL�� = '" + tempEdaban + "',";

            //sqlStr += "�Ԓn�� = '',";
            //sqlStr += "�}���V������ = '',";
            //sqlStr += "���R = 0,";
            //sqlStr += "���̑����e = '',";

            sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
            sqlStr += "where (�z�z�G���A.ID = " + tempID.ToString() + ") and ";
            sqlStr += "(�X�e�[�^�X <> 0)";

            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("�z�z�G���A�f�[�^�̍X�V�Ɏ��s���܂���(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     �z�z�G���A�f�[�^�𖢑I����Ԃɖ߂� </summary>
        /// <param name="tempID">
        ///     �z�z�w��ID</param>
        ///----------------------------------------------------------------
        private void HaihuAreaClear(int tempID)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update �z�z�G���A ";
            sqlStr += "set ";
            sqlStr += "�z�z�w��ID = 0,";
            sqlStr += "�z�z�P�� = 0,";
            sqlStr += "���z�z���� = 0,";
            sqlStr += "���c�� = �\�薇��,";
            sqlStr += "�񍐖��� = 0,";
            sqlStr += "�񍐎c�� = �\�薇��,";
            sqlStr += "�����敪 = 0,";
            sqlStr += "�X�e�[�^�X = 0,";
            sqlStr += "�}�ԋL�� = '',";

            //sqlStr += "�Ԓn�� = '',";
            //sqlStr += "�}���V������ = '',";
            //sqlStr += "���R = 0,";
            //sqlStr += "���̑����e = '',";

            sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
            sqlStr += "where (�z�z�G���A.�z�z�w��ID = " + tempID.ToString() + ") ";

            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("�z�z�G���A�f�[�^�̏������Ɏ��s���܂���(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter_Haifu(dataGridView2.SelectedRows[0].Index);
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView2.Rows.Count == 0) return;
            if (dataGridView2.SelectedRows.Count == 0) return;

            //EnterKey������̍s�ړ����֎~����
            e.Handled = true;

            GridEnter_Haifu(dataGridView2.SelectedRows[0].Index);

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     �z�z�G���A�폜 </summary>
        /// <param name="eRow">
        ///     �s�C���f�b�N�X</param>
        ///--------------------------------------------------------------------
        private void areaDel(int eRow)
        {
            //�z�z�G���A���폜����
            //if (dataGridView2.SelectedRows.Count == 0) return;

            if (MessageBox.Show("�I�𒆂̃`���V�z�z�f�[�^��z�z�w�������珜�O���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            dataGridView2[14, eRow].Value = "1";    // 2015/07/15
            dataGridView2.Rows[eRow].DefaultCellStyle.ForeColor = Color.LightGray;

            //�f�[�^�x�[�X�X�V
            HaihuAreaUpdate(Int32.Parse(dataGridView2[0, eRow].Value.ToString()),
             (int)(0), (int)(0), (int)(0), (int)(0), (int)(0), (int)(0), "");

            //���z�z���폜
            string sqlStr;
            Control.FreeSql fCon = new Control.FreeSql();
            sqlStr = "";
            sqlStr += "delete from ���z�z��� ";
            sqlStr += "where �z�z�G���AID = " + Int32.Parse(dataGridView2[0, eRow].Value.ToString());

            if (fCon.Execute(sqlStr) == false)
            {
                MessageBox.Show("���z�z���̍폜�Ɏ��s���܂���", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            fCon.Close();

            dataGridView2.CurrentCell = null;

            //�z�z�w��������{�^��
            button4.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            const int G_COUNT = 15; //�z�z�w�����̖��׍s��

            int iX = 0;
    
            if (fMode.Mode == 0)     //�O���b�h����I�������z�z�w�������������
            {
                //�����̔z�z�w������A���������
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("�z�z�w���f�[�^���I������Ă��܂���", "�f�[�^���I��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (MessageBox.Show("�I������Ă���z�z�w�����𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                //���ӎ�������������
                foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                {
                    ShijiMemoUpdate(dataGridView1[0, r.Index].Value.ToString(), txtChuui.Text);
                }

                //�I������Ă���s���擾
                iX = 0;
                foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                {
                    GridEnter(r.Index);  //�f�[�^��ʕ\��

                    if (dataGridView2.Rows.Count > 0)
                    {
                        ShijiReportSet(G_COUNT);    //�������
                    }

                    StatusBack();   //���o�^�`���V�f�[�^�𖢑I����Ԃɖ߂�

                    DispClear();    //��ʕ\����������

                    iX++;
                }
            }
            else
            {
                //��ʕ\������Ă���z�z�w�������������
                if (MessageBox.Show("�z�z�w�����𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                ShijiReportSet(G_COUNT);    //�������
            }

        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     �z�z�w�������s�f�[�^�̒��ӏ��������������� </summary>
        /// <param name="tempID">
        ///     �z�z�w��ID</param>
        /// <param name="tempMemo">
        ///     ���ӏ���������</param>
        ///----------------------------------------------------------------
        private void ShijiMemoUpdate(string tempID,string tempMemo)
        {
            string sqlStr;
            Control.FreeSql fCon = new Control.FreeSql();

            sqlStr = "";
            sqlStr += "update �z�z�w�� set ";
            sqlStr += "���ӎ��� = '" + tempMemo + "' ";
            sqlStr += " where �z�z�w��.ID = " + tempID;
            
            fCon.Execute(sqlStr);

            fCon.Close();
        }

        private void ShijiReportSet(int G_COUNT)
        {
            int pCnt;

            //�y�[�W�J�E���g
            //pCnt = dataGridView2.Rows.Count / G_COUNT + 1;

            // 2015/11/18
            pCnt = dataGridView2.Rows.Count / G_COUNT;

            if ((dataGridView2.Rows.Count % G_COUNT) > 0)
            {
                pCnt++;
            }

            for (int i = 1; i <= pCnt; i++)
            {
                ShijiReport(pCnt, i, G_COUNT);
            }
        }

        private void ShijiReport(int tempPage, int tempCurrentPage, int tempFixRows)
        {

            const int S_GYO = 5;    //�G�N�Z���t�@�C�����ׂ�5�s�ڂ����
            int dgvIndex;
            int i;

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���z�z�w���񍐏� , Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    oxlsSheet.Cells[1, 1] = "�w���ԍ��F" + String.Format("{0:0000000}", int.Parse(txthID.Text, System.Globalization.NumberStyles.Any));         //�w���ԍ�
                    oxlsSheet.Cells[1, 17] = "���s���F" + DateTime.Today.ToLongDateString() + "�@P." + tempCurrentPage.ToString() + "/" + tempPage.ToString();  //���s��

                    //�z�z�G���A����
                    i = 0;
                    while (true)
                    {
                        dgvIndex = tempFixRows * (tempCurrentPage - 1) + i; //�f�[�^�O���b�h�r���[�̍s�C���f�b�N�X�����߂�

                        oxlsSheet.Cells[i + S_GYO, 1] = dataGridView2[1, dgvIndex].Value.ToString();   //�`���V��
                        oxlsSheet.Cells[i + S_GYO, 4] = dataGridView2[2, dgvIndex].Value.ToString();   //�z�z�敪
                        oxlsSheet.Cells[i + S_GYO, 6] = dataGridView2[3, dgvIndex].Value.ToString();   //�z�z�`��
                        oxlsSheet.Cells[i + S_GYO, 7] = dataGridView2[4, dgvIndex].Value.ToString();   //�z�z��Z��
                        oxlsSheet.Cells[i + S_GYO, 14] = dataGridView2[6, dgvIndex].Value.ToString();   //�G���AID
                        oxlsSheet.Cells[i + S_GYO, 15] = dataGridView2[7, dgvIndex].Value.ToString();   //�P��
                        oxlsSheet.Cells[i + S_GYO, 16] = dataGridView2[8, dgvIndex].Value.ToString();   //�\�薇��

                        if (dataGridView2[9, dgvIndex].Value != null)
                        {
                            oxlsSheet.Cells[i + S_GYO, 17] = dataGridView2[9, dgvIndex].Value.ToString();   //�\�薇���v
                        }
                        
                        //�{���̒��ӎ���
                        oxlsSheet.Cells[25, 1] = txtChuui.Text;

                        //�O���b�h�ŏI�s�̂Ƃ��I��
                        if (dgvIndex == (dataGridView2.Rows.Count - 1)) break;

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

                    //////DialogResult ret;

                    ////////�_�C�A���O�{�b�N�X�̏����ݒ�
                    //////saveFileDialog1.Title = "�z�z�w�����ۑ�";
                    //////saveFileDialog1.OverwritePrompt = true;
                    //////saveFileDialog1.RestoreDirectory = true;
                    //////saveFileDialog1.FileName = "�z�z�w����_" + String.Format("{0:0000000}", int.Parse(txthID.Text, System.Globalization.NumberStyles.Any));

                    ////////�����y�[�W�̂Ƃ��A�y�[�W�����t�^
                    //////if (tempPage > 1)
                    //////{
                    //////    saveFileDialog1.FileName += "_" + tempCurrentPage.ToString();
                    //////}

                    //////saveFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xls)|*.xls|�S�Ẵt�@�C��(*.*)|*.*";

                    ////////�_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
                    //////string fileName;
                    //////ret = saveFileDialog1.ShowDialog();

                    //////if (ret == System.Windows.Forms.DialogResult.OK)
                    //////{
                    //////    fileName = saveFileDialog1.FileName;
                    //////    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                    //////                    Type.Missing, Type.Missing, Type.Missing,
                    //////                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    //////                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //////}

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�z�z�w����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                MessageBox.Show(e.Message, "�z�z�w����", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            
            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void frmHaifuShiji_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("�I�����܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }
                
            //���o�^�̃`���V�f�[�^��߂�
            StatusBack();

        }

        private void txtStaffID_Validating(object sender, CancelEventArgs e)
        {
            int d;
            string str;

            // �����͂܂��̓X�y�[�X�݂͉̂�
            if ((this.txtStaffID.Text).Trim().Length < 1)
            {
                label13.Text = "";
                return;
            }

            // �������H
            if (txtStaffID.Text == null)
            {
                MessageBox.Show("�����œ��͂��Ă�������", "�z�z���R�[�h", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtStaffID.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", "�z�z���R�[�h", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            //�R�[�h����
            string sqlStr;
            Control.�z�z�� cStaff = new Control.�z�z��();
            OleDbDataReader dR;

            sqlStr = " where ID = " + txtStaffID.Text.ToString();
            dR = cStaff.FillBy(sqlStr);

            if (dR.HasRows == false)
            {
                MessageBox.Show("���o�^�R�[�h�ł�", "�z�z���R�[�h", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                label13.Text = "";
                dR.Close();
                cStaff.Close();
                return;
            }
            else
            {
                while (dR.Read())
                {
                    label13.Text = dR["����"].ToString();
                }
            }

            dR.Close();
            cStaff.Close();
        }

        /// <summary>
        /// �z�z���̓V���\������
        /// </summary>
        private void tenkouUpdate()
        {

            //�V��\��
            string sqlStr;
            OleDbDataReader dRt;
            Control.�V�� cTenkou = new Control.�V��();

            sqlStr = "";
            sqlStr += "where ���t = '" + jDate.Value.ToShortDateString() + "'";

            dRt = cTenkou.FillBy(sqlStr);

            while (dRt.Read())
            {
                txtTenkou.Text = dRt["�V��"].ToString();
            }

            dRt.Close();
            cTenkou.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //�V��o�^
            frmTenkou frm = new frmTenkou();
            frm.ShowDialog();

            //�V��\��
            tenkouUpdate();
        }

        private void jDate_ValueChanged(object sender, EventArgs e)
        {
            //�V��\��
            tenkouUpdate();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if ((txtStaffID.Text != null) && (txtStaffID.Text != "") && (txtStaffID.Text != "0"))
            {
                frmTesuuryouMeisai frm = new frmTesuuryouMeisai(1);

                frm.�z�z�� = jDate.Value;
                frm.�z�z��ID = int.Parse(txtStaffID.Text, System.Globalization.NumberStyles.Any);
                frm.�z�z���� = label13.Text;

                frm.ShowDialog();
            }
            else
            {
                MessageBox.Show("�z�z������͂��Ă�������",MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //�|�X�e�B���O�G���A�o�^
            frmPosting frm = new frmPosting();
            frm.ShowDialog();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            button4.Enabled = true;
        }

        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t' && e.KeyChar != '.')
                e.Handled = true;
        }

        private void dataGridView2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            string cName = dataGridView2.CurrentCell.OwningColumn.Name;
            if (cName == "col8" || cName == "col11")
            {
                //�C�x���g�n���h����������ǉ�����Ă��܂��̂ōŏ��ɍ폜����
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);

                //�C�x���g�n���h����ǉ�����
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
            else
            {
                //�C�x���g�n���h�����폜����
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void dataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentCellAddress.X == 12)
            {
                if (dataGridView2.IsCurrentCellDirty)
                {
                    dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 18)    // �ҏW��ʕ\��
            {
                GridEnter_Haifu(e.RowIndex);
            }
            else if (e.ColumnIndex == 19)   // �z�z�G���A�폜
            {
                areaDel(e.RowIndex);
            }
        }

        private void dataGridView2_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //dataGridView2.CurrentCell = null;
        }

        private void nDate_Enter(object sender, EventArgs e)
        {
            dataGridView2.CurrentCell = null;   // 2017/10/03
        }

        private void txtStaffID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
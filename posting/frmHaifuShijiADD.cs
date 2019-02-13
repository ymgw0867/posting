using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;

namespace posting
{
    public partial class frmHaifuShijiADD : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�z�z�G���A cArea = new Entity.�z�z�G���A();
        Entity.�z�z�w�� cMaster = new Entity.�z�z�w��();

        const string MESSAGE_CAPTION = "�z�z�w���f�[�^�o�^";
        const int STATUS_RES = 1;
        const int STATUS_DEF = 0;
        const string UPDATE_OK = "0";
        const string UPDATE_NO = "1";
        const int FLG_ON = 1;
        const int FLG_OFF = 0;

        public frmHaifuShijiADD()
        {
            InitializeComponent();

            // �f�[�^�ǂݍ���
            jAdp.Fill(dts.��1);
            aAdp.Fill(dts.�z�z�G���A);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.��1TableAdapter jAdp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.�z�z�G���ATableAdapter aAdp = new darwinDataSetTableAdapters.�z�z�G���ATableAdapter();

        DateTime sDay;
        DateTime eDay;

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�z�z�t���O��ON�ɂ���
            Utility.FlgOnOff(FLG_ON);

            // ���t�J�����_�[�p�̊J�n������I�������擾
            sDay = dts.�z�z�G���A.Where(a => a.�z�z�w��ID == 0 && a.�X�e�[�^�X == 0).Min(a => a.��1Row.�z�z�J�n��);
            eDay = dts.�z�z�G���A.Where(a => a.�z�z�w��ID == 0 && a.�X�e�[�^�X == 0).Max(a => a.��1Row.�z�z�I����);

            TimeSpan ts = eDay - sDay;
            double days = ts.TotalDays;

            // �Œ��P�O�O���Ƃ��܂�
            if (days > 100)
            {
                days = 100;
                eDay = sDay.AddDays(days);
            }

            //��ʐݒ�
            GridviewSet.Setting(dataGridView2, dts, jAdp, aAdp, sDay, days);
            GridviewSet.Setting2(dataGridView1);
            GridviewSet.ShowData(dataGridView2, sDay, eDay);
            dataGridView2.CurrentCell = null; //�I����Ԃ��������
        }

        // �f�[�^�O���b�h�r���[�N���X
        private class GridviewSet
        {
            /// <summary>
            /// �f�[�^�O���b�h�r���[�̒�`���s���܂�
            /// </summary>
            /// <param name="tempDGV">�f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            public static void Setting(DataGridView tempDGV, darwinDataSet dts, darwinDataSetTableAdapters.��1TableAdapter jAdp, darwinDataSetTableAdapters.�z�z�G���ATableAdapter aAdp, DateTime sDay, double days)
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
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 559;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�z�z�G���AID");
                    tempDGV.Columns.Add("col2", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col3", "�`���V��");
                    tempDGV.Columns.Add("col4", "�رID");
                    tempDGV.Columns.Add("col5", "�z�z��Z��");
                    tempDGV.Columns.Add("col6", "�\�薇��");

                    tempDGV.Columns.Add("col7", "�z�z�J�n");
                    tempDGV.Columns.Add("col8", "�z�z�I��");
                    tempDGV.Columns.Add("col9", "���z");
                    tempDGV.Columns.Add("col10", "�z�z����");
                    tempDGV.Columns.Add("col11", "�z�z�`��");
                    tempDGV.Columns.Add("col12", "�P��");
                    tempDGV.Columns.Add("col13", "����");
                    tempDGV.Columns.Add("col14", "���z���O");

                    for (int i = 0; i <= days; i++)
                    {
                        string gDay = string.Empty;

                        if (sDay.AddDays(i).Day == 1)
                        {
                            gDay = sDay.AddDays(i).Month.ToString() + "/" + sDay.AddDays(i).Day.ToString();
                        }
                        else
                        {
                            gDay = sDay.AddDays(i).Day.ToString();
                        }

                        tempDGV.Columns.Add("day" + i.ToString(), gDay);
                        tempDGV.Columns["day" + i.ToString()].Width = 20;
                        tempDGV.Columns["day" + i.ToString()].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    tempDGV.Columns[4].Frozen = true;

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 100;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 230;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    //tempDGV.Columns[8].Width = 366;

                    //tempDGV.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[8].Width = 200;
                    tempDGV.Columns[13].Width = 80;

                    tempDGV.Columns[0].Visible = false;
                    tempDGV.Columns[9].Visible = false;
                    tempDGV.Columns[10].Visible = false;
                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;
                    //tempDGV.Columns[13].Visible = false; // 2014/11/26

                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

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
                    tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", (float)9.5, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 559;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col0", "�O���[�v");
                    tempDGV.Columns.Add("col1", "�z�z�G���AID");
                    tempDGV.Columns.Add("col2", "�󒍔ԍ�");
                    tempDGV.Columns.Add("col3", "�`���V��");
                    tempDGV.Columns.Add("col4", "�رID");
                    tempDGV.Columns.Add("col5", "�z�z��Z��");
                    tempDGV.Columns.Add("col6", "�\�薇��");

                    tempDGV.Columns.Add("col7", "�z�z�J�n");
                    tempDGV.Columns.Add("col8", "�z�z�I��");
                    tempDGV.Columns.Add("col9", "���z");
                    tempDGV.Columns.Add("col10", "�z�z����");
                    tempDGV.Columns.Add("col11", "�z�z�`��");
                    tempDGV.Columns.Add("col12", "�P��");
                    tempDGV.Columns.Add("col13", "����");
                    tempDGV.Columns.Add("col14", "�X�V����");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 70;
                    tempDGV.Columns[5].Width = 230;
                    tempDGV.Columns[6].Width = 80;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 286;

                    tempDGV.Columns[1].Visible = false;
                    tempDGV.Columns[10].Visible = false;
                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;
                    tempDGV.Columns[13].Visible = false;
                    tempDGV.Columns[14].Visible = false;

                    //tempDGV.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[8].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //tempDGV.Columns[9].SortMode = DataGridViewColumnSortMode.NotSortable;

                    tempDGV.Columns[6].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //tempDGV.Columns[0].SortMode = DataGridViewColumnSortMode.Automatic;

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

            ///----------------------------------------------------------------
            /// <summary>
            ///     �O���b�h�Ɏ󒍃f�[�^��\������ </summary>
            /// <param name="tempDGV">
            ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g��</param>
            ///----------------------------------------------------------------
            public static void ShowData(DataGridView tempDGV, DateTime sDay, DateTime eDay)
            {
                string sqlSTRING = "";
                string strDate;
                int iX;

                try
                {
                    tempDGV.RowCount = 0;
                    
                    //�f�[�^���[�_�[���擾����
        
                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select �z�z�G���A.ID,�z�z�G���A.��ID,��.�`���V��,";
                    sqlSTRING += "�z�z�G���A.����ID,����.���� as ����,��.�z�z�J�n��,��.�z�z�I����,";
                    sqlSTRING += "�z�z�G���A.���z�敪,�z�z�G���A.�\�薇��,";
                    sqlSTRING += "��.�z�z����,�z�z�`��.���� as �z�z�`��,��.���z���O ";
                    sqlSTRING += "from ((�z�z�G���A left join �� on �z�z�G���A.��ID = ��.ID) ";
                    sqlSTRING += "left join ���� on �z�z�G���A.����ID = ����.ID) ";
                    sqlSTRING += "left join �z�z�`�� on ��.�z�z�`�� = �z�z�`��.ID ";
                    sqlSTRING += "where (�z�z�G���A.�z�z�w��ID = 0) and (�z�z�G���A.�X�e�[�^�X = 0) ";
                    sqlSTRING += "order by �z�z�G���A.��ID,�z�z�G���A.����ID";
                                        
                    //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����
                    Control.FreeSql cArea = new Control.FreeSql();
                    dR = cArea.free_dsReader(sqlSTRING);

                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {
                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = int.Parse(dR["ID"].ToString());
                            tempDGV[1, iX].Value = long.Parse(dR["��ID"].ToString());
                            tempDGV[2, iX].Value = dR["�`���V��"].ToString();
                            tempDGV[3, iX].Value = int.Parse(dR["����ID"].ToString());
                            tempDGV[4, iX].Value = dR["����"].ToString();
                            tempDGV[5, iX].Value = int.Parse(dR["�\�薇��"].ToString());

                            strDate = dR["�z�z�J�n��"].ToString() + "";
                            if (strDate == "")
                            {
                                tempDGV[6, iX].Value = "";
                            }
                            else
                            {
                                tempDGV[6, iX].Value = Convert.ToDateTime(dR["�z�z�J�n��"].ToString() + "");
                            }

                            strDate = dR["�z�z�I����"].ToString() + "";
                            if (strDate == "")
                            {
                                tempDGV[7, iX].Value = "";
                            }
                            else
                            {
                                tempDGV[7, iX].Value = Convert.ToDateTime(dR["�z�z�I����"].ToString() + "");
                            }

                            if (dR["���z�敪"].ToString() == "1")
                            {
                                tempDGV[8, iX].Value = "��";
                            }
                            else
                            {
                                tempDGV[8, iX].Value = "";
                            }

                            //tempDGV[8, iX].Value = dR["�z�z����"].ToString();
                            //tempDGV[9, iX].Value = dR["�z�z�`��"].ToString() + "";
                            //tempDGV[10, iX].Value = dR["�P��"].ToString();
                            //tempDGV[11, iX].Value = dR["����"].ToString();

                            // ���z���O 2014/11/26
                            tempDGV["col14", iX].Value = Utility.nullToInt(dR["���z���O"]).ToString();

                            // �z�z���}�[�L���O
                            for (int i = 0; sDay.AddDays(i) <= eDay; i++)
                            {
                                if (sDay.AddDays(i) >= DateTime.Parse(dR["�z�z�J�n��"].ToString()) &&
                                    sDay.AddDays(i) <= DateTime.Parse(dR["�z�z�I����"].ToString()))
                                {
                                    tempDGV["day" + i.ToString(), iX].Value = "��";
                                    tempDGV["day" + i.ToString(), iX].Style.BackColor = Color.LightPink;
                                    tempDGV["day" + i.ToString(), iX].Style.ForeColor = Color.LightPink;
                                }
                                else
                                {
                                    tempDGV["day" + i.ToString(), iX].Value = "";
                                    tempDGV["day" + i.ToString(), iX].Style.BackColor = Color.White;
                                    tempDGV["day" + i.ToString(), iX].Style.ForeColor = Color.White;
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
                    cArea.Close();

                    //if (tempDGV.RowCount <= 27)
                    //{
                    //    tempDGV.Columns[8].Width = 77;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[8].Width = 60;
                    //}

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //�z�z�G���A���ꊇ�I������
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("���ׂ��I������Ă��܂���", "���ז��I��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            } 

            if (MessageBox.Show("�I�𒆂̃`���V�z�z�f�[�^��z�z�w�����f�[�^�֒ǉ����܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            int iY = 0;
 
            //�I���f�[�^���w����TAB�ֈړ�����
            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {

                dataGridView1.Rows.Add();
                iY = dataGridView1.Rows.Count;

                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    // 2014/11/26 �u���z���O�v�ȍ~��̓X�L�b�v
                    if (i < 13)
                    {
                        dataGridView1[i + 1, iY - 1].Value = dataGridView2[i, r.Index].Value;
                    }
                }

                dataGridView1[14, iY - 1].Value = UPDATE_OK; //�X�V����
 
                //�I���f�[�^�̃X�e�[�^�X��(1)�ɏ���������
                HaihuStatusUpdate(int.Parse(dataGridView2[0, r.Index].Value.ToString()),STATUS_RES);

                iY++;
            }

            ////�z�z���w���f�[�^�ĕ\��
            //GridviewSet.ShowData(dataGridView2);

            //TAB2��\��
            tabControl1.SelectedIndex = 1;

            dataGridView1.CurrentCell = null; //�I����Ԃ��������

        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�\�����̃`���V�f�[�^�S�Ă�I����Ԃɂ��܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.Rows[i].Selected = true;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("���݂̑I����Ԃ��������܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                dataGridView2.Rows[r.Index].Selected = false;
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Ending();
        }

        private void Ending()
        {
            if (MessageBox.Show("�I�����܂��B���݁A�z�z�w���ݒ�^�u�ɕ\�����̃f�[�^�͓o�^����܂���B" + Environment.NewLine + "��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            //�\��f�[�^�̃X�e�[�^�X��0�ɖ߂�
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                HaihuStatusUpdate(int.Parse(dataGridView1[1, i].Value.ToString()), STATUS_DEF);
            }

            //�z�z�t���O��OFF�ɂ���
            Utility.FlgOnOff(FLG_OFF);

            this.Close();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("���z�\�����s���܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            int toID;
            DateTime sDate;
            DateTime eDate;

            int targetID;
            DateTime tsDate;
            DateTime teDate;

            string sqlStr, msgStr;

            // �����̃��[�v
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                // �z�z���Ԃ��o�^����Ă�����̂�ΏۂƂ���
                // ���z���O�Č��͑ΏۊO�Ƃ��܂� 2014/11/26
                if (dataGridView2[6, i].Value.ToString().Trim() != "" &&
                    dataGridView2[7, i].Value.ToString().Trim() != "" && 
                    dataGridView2["col14", i].Value.ToString() != "1")
                {
                    // ID
                    toID = Int32.Parse(dataGridView2[3, i].Value.ToString());

                    //�z�z�J�n��
                    sDate = Convert.ToDateTime(dataGridView2[6, i].Value.ToString());

                    //�z�z�I����
                    eDate = Convert.ToDateTime(dataGridView2[7, i].Value.ToString());

                    //����̃��[�v
                    for (int iX = i + 1; iX < dataGridView2.Rows.Count; iX++)
                    {
                        // �z�z���Ԃ��o�^����Ă�����̂�ΏۂƂ���
                        // ���z���O�Č��͑ΏۊO�Ƃ��܂� 2014/11/26
                        if (dataGridView2[6, iX].Value.ToString().Trim() != "" &&
                            dataGridView2[7, iX].Value.ToString().Trim() != "" &&
                            dataGridView2["col14", iX].Value.ToString() != "1")
                        {
                            //�����ID
                            targetID = Int32.Parse(dataGridView2[3, iX].Value.ToString());

                            //����̔z�z�J�n��
                            tsDate = Convert.ToDateTime(dataGridView2[6, iX].Value.ToString());

                            //����̔z�z�I����
                            teDate = Convert.ToDateTime(dataGridView2[7, iX].Value.ToString());

                            if (toID == targetID)
                            {
                                //�z�z���Ԃ��d������ꍇ�A���z����
                                //�p�^�[���@�@���Ԃ̈ꕔ�d��
                                if (((tsDate >= sDate) && (tsDate <= eDate)) || ((teDate >= sDate) && (teDate <= eDate)))
                                {
                                    dataGridView2[8, i].Value = "��";
                                    dataGridView2[8, iX].Value = "��";
                                    dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                                    dataGridView2.Rows[iX].DefaultCellStyle.ForeColor = Color.Red;
                                }

                                //�p�^�[���A ����̊��Ԓ����̂܂�
                                if ((tsDate <= sDate) && (eDate <= teDate))
                                {
                                    dataGridView2[8, i].Value = "��";
                                    dataGridView2[8, iX].Value = "��";
                                    dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                                    dataGridView2.Rows[iX].DefaultCellStyle.ForeColor = Color.Red;
                                }
                            }
                        }
                    }

                    //�z�z�w���ς݂Ŕz�z�������̃f�[�^��ΏۂɌ���
                    // ���> ��r����z�z���F�󒍃f�[�^�̔z�z�J�n���`�z�z�I���� 2009/10/13

                    sqlStr = "";

                    //////sqlStr += "select �z�z�G���A.�z�z�w��ID,�z�z�G���A.����ID,�z�z�w��.�z�z�� ";
                    //////sqlStr += "from �z�z�G���A inner join �z�z�w�� ";
                    //////sqlStr += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                    //////sqlStr += "where ";
                    //////sqlStr += "(�z�z�G���A.�z�z�w��ID <> 0) and ";
                    //////sqlStr += "(�z�z�G���A.�X�e�[�^�X = 2) and ";
                    //////sqlStr += "(�z�z�G���A.�����敪 = 0) and ";
                    //////sqlStr += "(�z�z�G���A.����ID = " + toID.ToString() + ") ";
                    //////sqlStr += "order by �z�z�G���A.�z�z�w��ID";

                    sqlStr += "select �z�z�G���A.�z�z�w��ID,�z�z�G���A.����ID,��.�z�z�J�n��,��.�z�z�I���� ";
                    sqlStr += "from �z�z�G���A inner join �� ";
                    sqlStr += "on �z�z�G���A.��ID = ��.ID ";
                    sqlStr += "where ";
                    sqlStr += "(�z�z�G���A.�z�z�w��ID <> 0) and ";
                    sqlStr += "(�z�z�G���A.�X�e�[�^�X = 2) and ";
                    sqlStr += "(�z�z�G���A.�����敪 = 0) and ";
                    sqlStr += "(�z�z�G���A.����ID = " + toID.ToString() + ") ";
                    sqlStr += "order by �z�z�G���A.��ID";

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlStr);

                    while (dR.Read())
                    {
                        //////�z�z�w�����̔z�z��
                        ////tsDate = Convert.ToDateTime(dR["�z�z��"].ToString());

                        //////�z�z�w�����̔z�z�����z�z���ԂɊY�����邩
                        ////if ((tsDate >= sDate) && (tsDate <= eDate))
                        ////{
                        ////    msgStr = dataGridView2[8, i].Value.ToString();

                        ////    if (msgStr != "")
                        ////    {
                        ////        msgStr += "�A";
                        ////    }

                        ////    msgStr += "�w����:" + dR["�z�z�w��ID"].ToString();
                        ////    dataGridView2[8, i].Value = msgStr;
                        ////    dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                        ////}

                        //����̔z�z�J�n��(�󒍃f�[�^���)
                        tsDate = Convert.ToDateTime(dR["�z�z�J�n��"].ToString());

                        //����̔z�z�I����(�󒍃f�[�^���)
                        teDate = Convert.ToDateTime(dR["�z�z�I����"].ToString());

                        //�z�z���Ԃ��d������ꍇ�A���z����
                        //�p�^�[���@�@���Ԃ̈ꕔ�d��
                        if (((tsDate >= sDate) && (tsDate <= eDate)) || ((teDate >= sDate) && (teDate <= eDate)))
                        {
                            msgStr = dataGridView2[8, i].Value.ToString();

                            if (msgStr != "")
                            {
                                msgStr += "�A";
                            }

                            msgStr += "�w����:" + dR["�z�z�w��ID"].ToString();
                            dataGridView2[8, i].Value = msgStr;
                            dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                        }
                        else
                        {
                            //�p�^�[���A ����̊��Ԓ����̂܂�
                            if ((tsDate <= sDate) && (eDate <= teDate))
                            {
                                msgStr = dataGridView2[8, i].Value.ToString();

                                if (msgStr != "")
                                {
                                    msgStr += "�A";
                                }

                                msgStr += "�w����:" + dR["�z�z�w��ID"].ToString();
                                dataGridView2[8, i].Value = msgStr;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                            }
                        }
                    }

                    dR.Close();
                    fCon.Close();
                }
            }

            MessageBox.Show("�I�����܂���", "���z�\��", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// �z�z�G���A�f�[�^�̃X�e�[�^�X��������Ԃɖ߂�
        /// </summary>
        /// <param name="tempID">�z�z�G���AID</param>
        private void HaihuStatusUpdate(int tempID,int tempStatus)
        {
            Control.FreeSql cUp = new Control.FreeSql();

            string sqlStr = "";

            sqlStr += "update �z�z�G���A ";
            sqlStr += "set ";
            sqlStr += "�z�z�w��ID = 0,";
            sqlStr += "�X�e�[�^�X = " + tempStatus.ToString() + ",";
            sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
            sqlStr += "where �z�z�G���A.ID = " + tempID.ToString();

            if (cUp.Execute(sqlStr) == false)
            {
                MessageBox.Show("�z�z�G���A�f�[�^�̃X�e�[�^�X�X�V�Ɏ��s���܂���(" + tempID.ToString() + ")", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            cUp.Close();
        }

        private void frmHaifuShijiADD_FormClosing(object sender, FormClosingEventArgs e)
        {
            Dispose();
        }

        private void button6_Click(object sender, EventArgs e)
        {

            const int DIALOG_OK = 1;
            //const int DIALOG_NO = 0;

            //�I�𒆂̃f�[�^���O���[�v������
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("���ׂ�I�����Ă�������", "���ז��I��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            haihuGroup fSub = new haihuGroup();

            if (fSub.ShowDialog(this) == DialogResult.OK)
            {
                if (fSub.sStatus == DIALOG_OK)
                {
                    //�I���f�[�^���O���[�s���O����
                    foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                    {
                        dataGridView1[0, r.Index].Value = fSub.sGroup;
                    }
                }

                //���ёւ��������߂�
                DataGridViewColumn sortColumn = dataGridView1.Columns[0];

                //���ёւ��̕����i�������~�����j�����߂�
                ListSortDirection sortDirection = ListSortDirection.Ascending;

                //���ёւ����s��
                dataGridView1.Sort(sortColumn, sortDirection);

                //�蓮�\�[�g�֎~�Ƃ���
                //dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;

            }

            fSub.Dispose();

            dataGridView1.CurrentCell = null;


        }

        private void button7_Click(object sender, EventArgs e)
        {
            //�z�z�G���A��I����������
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("���ׂ��I������Ă��܂���", "���ז��I��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�I�𒆂̃`���V�z�z�f�[�^�𖢎w���f�[�^�ɖ߂��܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�I���f�[�^�̃X�e�[�^�X��0�ɖ߂�
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                HaihuStatusUpdate(int.Parse(dataGridView1[1, r.Index].Value.ToString()), STATUS_DEF);
            }
            
            //�I���f�[�^���O���b�h����폜����
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                if (!r.IsNewRow)
                {
                    dataGridView1.Rows.Remove(r);
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                GridviewSet.ShowData(dataGridView2, sDay, eDay);
            }

            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    GridviewSet.ShowData(dataGridView2, sDay, eDay);
                    dataGridView2.CurrentCell = null;
                    break;

                case 1:
                    dataGridView1.CurrentCell = null;
                    button6.Enabled = false;
                    button7.Enabled = false;

                    if (dataGridView1.RowCount > 0)
                    {
                        button8.Enabled = true;
                    }
                    else
                    {
                        button8.Enabled = false;
                    }

                    break;

                default:
                    break;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //if (dataGridView1.SelectedRows.Count == 0)
            //{
            //    MessageBox.Show("���ׂ��I������Ă��܂���", "���ז��I��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[0, i].Value == null)
                {
                    MessageBox.Show("�O���[�v������Ă��Ȃ��z�z�f�[�^������܂��B�S�Ă̔z�z�f�[�^���O���[�v�����čēx���s���Ă��������B", "���O���[�v��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[0, i].Value.ToString() == "")
                {
                    MessageBox.Show("�O���[�v������Ă��Ȃ��z�z�f�[�^������܂��B�S�Ă̔z�z�f�[�^���O���[�v�����čēx���s���Ă��������B", "���O���[�v��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            //���ёւ��������߂�
            DataGridViewColumn sortColumn = dataGridView1.Columns[0];

            //���ёւ��̕����i�������~�����j�����߂�
            ListSortDirection sortDirection = ListSortDirection.Ascending;

            //���ёւ����s��
            dataGridView1.Sort(sortColumn, sortDirection);

            if (MessageBox.Show("�z�z�w���f�[�^�ɓo�^���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�z�z�w���f�[�^�o�^����
            haifuAdd();

            //�X�V����OK�̃f�[�^�̓O���b�h����폜����

            //�Z���N�g
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[14, i].Value.ToString() == UPDATE_OK)
                {
                    dataGridView1.Rows[i].Selected = true;
                }
            }

            //�폜
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                if (!r.IsNewRow)
                {
                    dataGridView1.Rows.Remove(r);
                }
            }

            //�I������
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("�S�Ă̔z�z�f�[�^�̓o�^�������I�����܂���", "�m�F", MessageBoxButtons.OK, MessageBoxIcon.Information);
                tabControl1.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("�ꕔ�̔z�z�f�[�^�̓o�^�������I�����܂���ł���", "�m�F", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void haifuAdd()
        {
            string sGroup = "";
            int sRow = 0, eRow = 0;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if ((sGroup != "") && (sGroup != dataGridView1[0, i].Value.ToString()))
                {
                    HaifuUpdate(sRow,eRow);
                    sRow = i;
                }

                sGroup = dataGridView1[0, i].Value.ToString();
                eRow = i;
            }

            HaifuUpdate(sRow, eRow);
        }

        private void HaifuUpdate(int sRow,int eRow)
        {
            try
            {
                HaifuShijiData();   //�z�z�w���N���X�Ƀf�[�^�Z�b�g

                Control.DataControl Con;
                OleDbConnection cn;
                OleDbTransaction tran;
                OleDbCommand SCom;

                //ID���̔�
                string sqlStr = "";
                int gID = (int)(1);

                sqlStr = "select max(ID) as ID from �z�z�w�� ";
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
                    //�z�z�w���f�[�^�o�^����
                    sqlStr = "";
                    sqlStr += "insert into �z�z�w�� ";
                    sqlStr += "(ID,�z�z��,���͓�,�z�z��ID,��ʔ�,��ʋ�ԊJ�n,��ʋ�ԏI��,";
                    sqlStr += "�z�z�J�n����,�z�z�I������,�I�����|�[�g,���z�z�敪,���z�z���R,";
                    sqlStr += "�o�^�N����,�ύX�N����) ";
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
                    sqlStr += "'" + cMaster.�ύX�N���� + "')";

                    SCom.CommandText = sqlStr;

                    SCom.ExecuteNonQuery();

                    //�z�z�G���A�f�[�^�X�V
                    string sID;
                    const string sSTATUS = "2"; 

                    for (int i = sRow; i <= eRow; i++)
                    {
                        sID = dataGridView1[1, i].Value.ToString();

                        sqlStr = "";
                        sqlStr += "update �z�z�G���A ";
                        sqlStr += "set ";
                        sqlStr += "�z�z�w��ID = " + gID.ToString() + ",";
                        sqlStr += "���c�� = �\�薇��,";
                        sqlStr += "�񍐎c�� = �\�薇��,";
                        sqlStr += "�X�e�[�^�X = " + sSTATUS + ",";
                        sqlStr += "�ύX�N���� = '" + DateTime.Today + "' ";
                        sqlStr += "where (�z�z�G���A.ID = " + sID + ") and ";
                        sqlStr += "(�X�e�[�^�X <> 0)";

                        //if (dataGridView1[0, i].Value.ToString() == "2")
                        //{
                        //    sqlStr += "(�X�e�[�^�X <> 0)";
                        //}


                        SCom.CommandText = sqlStr;

                        SCom.ExecuteNonQuery();
                    }

                    tran.Commit();

                    //�X�V���ʏ�������
                    UpdateFlg(sRow, eRow, UPDATE_OK);

                    //MessageBox.Show("�V�K�o�^����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (Exception ex)
                {
                    tran.Rollback();

                    //�X�V���ʏ�������
                    UpdateFlg(sRow, eRow, UPDATE_NO);

                    MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    MessageBox.Show("�o�^�Ɏ��s���܂����B���[���o�b�N���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //���o�^�z�z�f�[�^�̃X�e�[�^�X��߂�
                    //StatusBack();
                }

                cn.Close();

                Con.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "�X�V����", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void UpdateFlg(int sRow,int eRow, string sFlg)
        {
            //�X�V���ʏ�������
            for (int i = sRow; i <= eRow; i++)
            {
                dataGridView1[14, i].Value = sFlg;
            }
        }

        private void HaifuShijiData()
        {
            //�z�z�w���N���X�Ƀf�[�^�Z�b�g
            cMaster.�z�z�� = DateTime.Today;
            cMaster.���͓� = DateTime.Today;
            cMaster.�z�z��ID = 0;

            cMaster.��ʔ� = 0;
            cMaster.��ʋ�ԊJ�n = "";
            cMaster.��ʋ�ԏI�� = "";

            cMaster.�z�z�J�n���� = "";
            cMaster.�z�z�I������ = "";

            cMaster.�I�����|�[�g = "";

            cMaster.���z�z�敪 = "";
            cMaster.���z�z���R = "";

            cMaster.���ӎ��� = "";

            cMaster.�o�^�N���� = DateTime.Today;
            cMaster.�ύX�N���� = DateTime.Today;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                button6.Enabled = true;
                button7.Enabled = true;
            }
            else
            {
                button6.Enabled = false;
                button7.Enabled = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Ending();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //�I�𒆂̃f�[�^�̍��v�z�z������\������
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("���ׂ�I�����Ă�������", "���ז��I��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int sMai = 0;

            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                if ((dataGridView1[6, r.Index].Value != null) && (Utility.NumericCheck(dataGridView1[6, r.Index].Value.ToString()) != false))
                {
                    sMai += int.Parse(dataGridView1[6, r.Index].Value.ToString(), System.Globalization.NumberStyles.Any);
                }
            }

            MessageBox.Show(sMai.ToString("#,##0") + " ���ł�", "���v�z�z����", MessageBoxButtons.OK, MessageBoxIcon.Information);

            dataGridView1.CurrentCell = null;
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, "���w���f�[�^");
        }

    }
}
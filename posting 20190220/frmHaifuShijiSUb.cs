using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace posting
{
    public partial class frmHaifuShijiSUb : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�z�z�G���A cMaster = new Entity.�z�z�G���A();

        const string MESSAGE_CAPTION = "�z�z�w���f�[�^�o�^";
        const int FLG_ON = 1;
        const int FLG_OFF = 0;

        public frmHaifuShijiSUb()
        {
            
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�z�z�t���OON
            Utility.FlgOnOff(FLG_ON);

            //��ʐݒ�
            GridviewSet.Setting(dataGridView2);
            GridviewSet.ShowData(dataGridView2);
            dataGridView2.CurrentCell = null; //�I����Ԃ��������
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
                    tempDGV.DefaultCellStyle.Font = new Font("�l�r �o�S�V�b�N", (float)9.5, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 505;

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

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 100;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 70;
                    tempDGV.Columns[4].Width = 230;
                    tempDGV.Columns[5].Width = 80;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 349;

                    tempDGV.Columns[0].Visible = false;
                    tempDGV.Columns[9].Visible = false;
                    tempDGV.Columns[10].Visible = false;
                    tempDGV.Columns[11].Visible = false;
                    tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[5].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[6].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[7].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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

            public static void ShowData(DataGridView tempDGV)
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
                    sqlSTRING += "��.�z�z����,�z�z�`��.���� as �z�z�`�� ";
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

            //�z�z�G���A���ꊇ�o�^����
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("���ׂ�I�����Ă�������", "���ז��I��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            } 

            if (MessageBox.Show("�I�𒆂̃`���V�z�z�f�[�^��z�z�w�����֒ǉ��o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            int iX;
            int iY = 0;

            iX = dataGridView2.SelectedRows.Count;

            _Count = dataGridView2.SelectedRows.Count;
            F_ID = new int[iX];
 
            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                F_ID[iY] = Int32.Parse(dataGridView2[0, r.Index].Value.ToString());
        
                iY++;
            }
        }

        //�I�����ꂽ�z�z�G���A�f�[�^�̍s��
        private int _Count;
        public int Count
        {
            get 
            {
                return this._Count; 
            }
            set
            {
                this._Count = value;
            }
        }

        //�I�����ꂽ�z�z�G���A�f�[�^ID�̃C���f�N�T
        private int[] F_ID;
        public int this[int iX]
        {
            set
            {
                this.F_ID[iX] = value;
            }
            get
            {
                return this.F_ID[iX];
            }
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
            if (MessageBox.Show("�I�����܂��B���݁A�I����Ԃ̃f�[�^�͓o�^����܂���B" + Environment.NewLine + "��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            button4.DialogResult = DialogResult.No;
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

            string sqlStr,msgStr;

            //�����̃��[�v
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                //�z�z���Ԃ��o�^����Ă�����̂�ΏۂƂ���
                if ((dataGridView2[6, i].Value.ToString().Trim() != "") &&
                    (dataGridView2[7, i].Value.ToString().Trim() != ""))
                {

                    //ID
                    toID = Int32.Parse(dataGridView2[3, i].Value.ToString());

                    //�z�z�J�n��
                    sDate = Convert.ToDateTime(dataGridView2[6, i].Value.ToString());

                    //�z�z�I����
                    eDate = Convert.ToDateTime(dataGridView2[7, i].Value.ToString());

                    //����̃��[�v
                    for (int iX = i + 1; iX < dataGridView2.Rows.Count; iX++)
                    {
                        //�z�z���Ԃ��o�^����Ă�����̂�ΏۂƂ���
                        if ((dataGridView2[6, iX].Value.ToString().Trim() != "") &&
                            (dataGridView2[7, iX].Value.ToString().Trim() != ""))
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
                    sqlStr = "";
                    sqlStr += "select �z�z�G���A.�z�z�w��ID,�z�z�G���A.����ID,�z�z�w��.�z�z�� ";
                    sqlStr += "from �z�z�G���A inner join �z�z�w�� ";
                    sqlStr += "on �z�z�G���A.�z�z�w��ID = �z�z�w��.ID ";
                    sqlStr += "where ";
                    sqlStr += "(�z�z�G���A.�z�z�w��ID <> 0) and ";
                    sqlStr += "(�z�z�G���A.�X�e�[�^�X = 2) and ";
                    sqlStr += "(�z�z�G���A.�����敪 = 0) and ";
                    sqlStr += "(�z�z�G���A.����ID = " + toID.ToString() + ") ";
                    sqlStr += "order by �z�z�G���A.�z�z�w��ID";

                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlStr);

                    while(dR.Read())
                    {
                        //�z�z�w�����̔z�z��
                        tsDate = Convert.ToDateTime(dR["�z�z��"].ToString());

                        //�z�z�w�����̔z�z�����z�z���ԂɊY�����邩
                        if ((tsDate >= sDate) && (tsDate <= eDate))
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

                    dR.Close();
                    fCon.Close();

                }
            }

            MessageBox.Show("�I�����܂���", "���z�\��", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dataGridView2, "���w���f�[�^");
        }

        private void frmHaifuShijiSUb_FormClosing(object sender, FormClosingEventArgs e)
        {
            //�z�z�t���OOFF
            Utility.FlgOnOff(FLG_OFF);
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;

namespace posting
{
    public partial class frmKouza : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�U������ cMaster = new Entity.�U������();

        const string MESSAGE_CAPTION = "�U�������}�X�^�[";

        public frmKouza()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.�U������' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            GridviewSet.Setting(dataGridView1);
            this.�U������TableAdapter.Fill(this.darwinDataSet.�U������);

            //������ʃZ�b�g
            Utility.ComboKouza.load(comboBox1);

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
                    tempDGV.Height = 180;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    //tempDGV.Columns.Add("col1", "����");
                    //tempDGV.Columns.Add("col2", "����");
                    //tempDGV.Columns.Add("col3", "���l");

                    tempDGV.Columns[0].Width = 70;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 200;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;

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

            /// <summary>
            /// �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾����
            /// </summary>
            /// <param name="dgv">�ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            public static Boolean GetData(DataGridView dgv,ref Entity.�U������ tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.�U������ Kouza = new Control.�U������();
                OleDbDataReader dr;

                sqlStr = " where �U������.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Kouza.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.���Z�@�֖� = dr["���Z�@�֖�"].ToString() + "";
                        tempC.�x�X�� = dr["�x�X��"].ToString();
                        tempC.������� = Convert.ToInt32(dr["�������"].ToString());
                        tempC.�����ԍ� = dr["�����ԍ�"].ToString();
                        tempC.�������` = dr["�������`"].ToString();
                    }
                }
                else
                {
                    dr.Close();
                    Kouza.Close();
                    return false;
                }

                dr.Close();
                Kouza.Close();
                return true;
            }

            //////public static void ShowData(DataGridView tempDGV)
            //////{
            //////    string sqlSTRING = "";

            //////    try
            //////    {
            //////        tempDGV.RowCount = 0;

            //////        //�������}�X�^�[�̃f�[�^���[�_�[���擾����
            //////        Control.DataControl dCon = new Control.DataControl();

            //////        sqlSTRING = "select * from m_Costname " +
            //////                    "order by ID";

            //////        dR = dCon.FreeReader(sqlSTRING);

            //////        iX = 0;

            //////        while (dR.Read())
            //////        {
            //////            tempDGV.Rows.Add();

            //////            tempDGV[0, iX].Value = dR["ID"];
            //////            tempDGV[1, iX].Value = NullConvert.Noth(dR["������"]);
            //////            tempDGV[2, iX].Value = NullConvert.Noth(dR["���l"]);
            //////            //tempDGV[1, iX].Value = dR["������"];
            //////            //tempDGV[2, iX].Value = dR["���l"];
            //////            iX++;
            //////        }

            //////        dR.Close();

            //////        dCon.Close();

            //////    }
            //////    catch (Exception e)
            //////    {
            //////        MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
            //////    }

            //////}

        }

        //�O���b�h����f�[�^��I��
        private void GridEnter()
        {

            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[1, dataGridView1.SelectedRows[iX].Index].Value + "���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�U�������I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {

                    //�f�[�^���擾����
                    if (GridviewSet.GetData(dataGridView1,ref cMaster) == false)
                    {
                        MessageBox.Show("�Y������f�[�^���}�X�^�[�ɓo�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'�f�[�^�l���擾
                    //txtCode.Text = cMaster.ID.ToString();
                    txtName1.Text = cMaster.���Z�@�֖�;
                    txtName2.Text = cMaster.�x�X��;

                    Utility.ComboKouza.selectedIndex(comboBox1, Int32.Parse(cMaster.�������.ToString()));

                    txtName2.Text = cMaster.�x�X��;
                    txtNumber.Text = cMaster.�����ԍ�;
                    txtMeigi.Text = cMaster.�������`;

                    //ID�e�L�X�g�{�b�N�X�͕ҏW�s�Ƃ���
                    //txtCode.Enabled = false;

                    //�{�^�����
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //�t�H�[�����[�h�X�e�[�^�X:�ύX�폜

                    txtName1.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "��ʃN���A", MessageBoxButtons.OK);
                }
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            // Enterkey�ȊO�͑ΏۊO
            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            GridEnter();
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
                fMode.Mode = 0;

                //txtCode.Enabled = true;
                //txtCode.Text = "";
                txtName1.Text = "";
                txtName2.Text = "";
                comboBox1.SelectedIndex = -1;
                txtNumber.Text = ""; ;
                txtMeigi.Text = "";

                btnDel.Enabled = false;
                btnClr.Enabled = false;

                if (this.dataGridView1.RowCount > 0)
                {
                    btnCsv.Enabled = true;
                }
                else
                {
                    btnCsv.Enabled = false;
                }

                //txtCode.Focus();
                txtName1.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�I������Ă���f�[�^��j�����܂��B��낵���ł����H","�m�F",MessageBoxButtons.YesNo,MessageBoxIcon.Question)== DialogResult.No )
                return;
                
            DispClear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {

                if (fDataCheck() == true)
                {
                    Control.�U������ Kouza = new Control.�U������();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Kouza.Close();
                                return;
                            }

                            if (Kouza.DataInsert(cMaster) == true)
                            {
                                MessageBox.Show("�V�K�o�^����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�V�K�o�^�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //�X�V
                            if (MessageBox.Show("�X�V���܂��B��낵���ł����H", "�X�V�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Kouza.Close();
                                return;
                            }

                            if (Kouza.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Kouza.Close();

                    DispClear();

                    //�f�[�^�� 'darwinDataSet.�U������' �e�[�u���ɓǂݍ��݂܂��B
                    this.�U������TableAdapter.Fill(this.darwinDataSet.�U������);

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

            string str;
            double d;

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
                    //Control.�U������ Kouza = new Control.�U������();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Kouza.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Kouza.Close();
                    //    throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    //}

                    //dr.Close();
                    //Kouza.Close();

                }

                //���Z�@�֖��`�F�b�N
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("���Z�@�֖�����͂��Ă�������");
                }

                //�x�X���`�F�b�N
                if (txtName2.Text.Trim().Length < 1)
                {
                    txtName2.Focus();
                    throw new Exception("�x�X������͂��Ă�������");
                }

                //������ʃ`�F�b�N
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("������ʂ�I�����Ă�������");
                }

                //�����ԍ��F�������H
                if (txtNumber.Text == null)
                {
                    this.txtNumber.Focus();
                    throw new Exception("�����ԍ��͐����œ��͂��Ă�������");
                }

                str = this.txtNumber.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtNumber.Focus();
                    throw new Exception("�����ԍ��͐����œ��͂��Ă�������");
                }

                // �����͂܂��̓X�y�[�X�݂͕̂s��
                if ((this.txtNumber.Text).Trim().Length < 1)
                {
                    this.txtNumber.Focus();
                    throw new Exception("�����ԍ�����͂��Ă�������");
                }

                //�������`�`�F�b�N
                if (txtMeigi.Text.Trim().Length < 1)
                {
                    txtMeigi.Focus();
                    throw new Exception("�������`����͂��Ă�������");
                }

                ////�[���͕s��
                //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                //{
                //    this.txtCode.Focus();
                //    throw new Exception("�[���͓o�^�ł��܂���");
                //}

                //�N���X�Ƀf�[�^�Z�b�g
                //cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.���Z�@�֖� = txtName1.Text.ToString();
                cMaster.�x�X�� = txtName2.Text.ToString();

                Utility.ComboKouza cmb1 = new Utility.ComboKouza();
                cmb1 = (Utility.ComboKouza)comboBox1.SelectedItem;
                cMaster.������� = cmb1.ID;

                cMaster.�����ԍ� = txtNumber.Text.ToString();
                cMaster.�������` = txtMeigi.Text.ToString(); 

                if (fMode.Mode == 0) cMaster.�o�^�N���� = DateTime.Today;
                cMaster.�ύX�N���� = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "�ێ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void txtEnter(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;
            }

            if (sender == txtNumber)
            {
                objtxt = txtNumber;
            }

            if (sender == txtMeigi)
            {
                objtxt = txtMeigi;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {

            TextBox objtxt = new TextBox();

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;
            }

            if (sender == txtNumber)
            {
                objtxt = txtNumber;
            }

            if (sender == txtMeigi)
            {
                objtxt = txtMeigi;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //���ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���@
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(��.�U������ID = " + cMaster.ID.ToString() + ")  ";

            OleDbDataReader dr;
            Control.�� Order = new Control.��();
            dr = Order.FillBy(SqlStr);

            //�Y���U�������̎󒍃f�[�^���o�^����Ă���Ƃ��͍폜�s�Ƃ���
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "�̎󒍃f�[�^���o�^����Ă��܂�", txtName1.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Order.Close();
                return;
            }

            dr.Close();
            Order.Close();

            //���ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���A
            SqlStr = " where ";
            SqlStr += "(������.�U������ID1 = " + cMaster.ID.ToString() + ") or ";
            SqlStr += "(������.�U������ID2 = " + cMaster.ID.ToString() + ") ";

            Control.������ Seikyu = new Control.������();
            dr = Seikyu.FillBy(SqlStr);

            //�Y���U�������̎󒍃f�[�^���o�^����Ă���Ƃ��͍폜�s�Ƃ���
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "�̐����f�[�^���o�^����Ă��܂�", txtName1.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Seikyu.Close();
                return;
            }

            dr.Close();
            Seikyu.Close();

            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�폜
            Control.�U������ Kouza = new Control.�U������();
            if (Kouza.DataDelete(cMaster.ID)==true)
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Kouza.Close();

            DispClear();

            //�f�[�^�� 'darwinDataSet.�U������' �e�[�u���ɓǂݍ��݂܂��B
            this.�U������TableAdapter.Fill(this.darwinDataSet.�U������);

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

            GridEnter();
        }

    }
}
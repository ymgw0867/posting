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
    public partial class frmTax : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�ŗ� cMaster = new Entity.�ŗ�();

        const string MESSAGE_CAPTION = "����ŗ�";

        public frmTax()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView1);

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.�ŗ�' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            this.�ŗ�TableAdapter.Fill(this.darwinDataSet.�ŗ�);

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

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].Width = 60;
                    tempDGV.Columns[2].Width = 120;
                    tempDGV.Columns[3].Width = 260;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.�ŗ� tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.�ŗ� sTax = new Control.�ŗ�();
                OleDbDataReader dr;

                sqlStr = " where �ŗ�.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = sTax.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.�ݒ�ŗ� = Convert.ToInt32(dr["�ŗ�"].ToString());
                        tempC.�J�n�N���� = Convert.ToDateTime(dr["�J�n�N����"].ToString());
                        tempC.���l = dr["���l"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    sTax.Close();
                    return false;
                }

                dr.Close();
                sTax.Close();
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

            if (MessageBox.Show(msgStr, "�ŗ��I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    txtName1.Text = cMaster.�ݒ�ŗ�.ToString();
                    tDate.Value = cMaster.�J�n�N����;
                    txtMemo.Text = cMaster.���l;

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
            GridEnter();
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
                tDate.Checked = false;
                txtMemo.Text = "";

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
            if (MessageBox.Show("�I������Ă���f�[�^��ύX���Ȃ��ŏI�����܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
                
            DispClear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {

                if (fDataCheck() == true)
                {
                    Control.�ŗ� sTax = new Control.�ŗ�();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                sTax.Close();
                                return;
                            }

                            if (sTax.DataInsert(cMaster) == true)
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
                                sTax.Close();
                                return;
                            }

                            if (sTax.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", "�����}�X�^�[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    sTax.Close();

                    DispClear();

                    //�f�[�^�� 'darwinDataSet.�ŗ�' �e�[�u���ɓǂݍ��݂܂��B
                    this.�ŗ�TableAdapter.Fill(this.darwinDataSet.�ŗ�);

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

                    // �������H
                    if (txtName1.Text == null)
                    {
                        this.txtName1.Focus();
                        throw new Exception("�ŗ��͐����œ��͂��Ă�������");
                    }

                    str = this.txtName1.Text;

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtName1.Focus();
                        throw new Exception("�ŗ��͐����œ��͂��Ă�������");
                    }

                    // �����͂܂��̓X�y�[�X�݂͕̂s��
                    if ((this.txtName1.Text).Trim().Length < 1)
                    {
                        this.txtName1.Focus();
                        throw new Exception("�ŗ����͂��Ă�������");
                    }

                    ////�[���͕s��
                    //if (Convert.ToInt32(this.txtName1.Text.ToString()) == 0)
                    //{
                    //    this.txtName1.Focus();
                    //    throw new Exception("�[���͓o�^�ł��܂���");
                    //}

                //    //�o�^�ς݃R�[�h�����ׂ�
                //    string sqlStr;
                //    Control.�󒍎�� sTax = new Control.�󒍎��();
                //    OleDbDataReader dr;

                //    sqlStr = " where ID = " + txtCode.Text.ToString();
                //    dr = sTax.FillBy(sqlStr);

                //    if (dr.HasRows == true)
                //    {
                //        txtCode.Focus();
                //        dr.Close();
                //        sTax.Close();
                //        throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                //    }

                //    dr.Close();
                //    sTax.Close();

                }

                //�J�n�N�����`�F�b�N
                if (tDate.Checked == false)
                {
                    tDate.Focus();
                    throw new Exception("�E�v������͂��Ă�������");
                }

                //�N���X�Ƀf�[�^���Z�b�g
                cMaster.�ݒ�ŗ� = Convert.ToInt32(txtName1.Text.ToString());
                cMaster.�J�n�N���� = Convert.ToDateTime(tDate.Text.ToString());
                cMaster.���l = txtMemo.Text.ToString();

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

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            ////���ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���
            //string SqlStr;
            //SqlStr = " where ";
            //SqlStr += "(��.�󒍎��ID = " + txtCode.Text.ToString() + ")  ";

            //OleDbDataReader dr;
            //Control.�� Order = new Control.��();
            //dr = Order.FillBy(SqlStr);

            ////�Y����ʂ̎󒍃f�[�^���o�^����Ă���Ƃ��͍폜�s�Ƃ���
            //if (dr.HasRows == true)
            //{
            //    MessageBox.Show(txtName1.Text.ToString() + "�̎󒍃f�[�^���o�^����Ă��܂�", txtName1.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    dr.Close();
            //    Order.Close();
            //    return;
            //}

            //dr.Close();
            //Order.Close();
            
            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�폜
            Control.�ŗ� sTax = new Control.�ŗ�();
            if (sTax.DataDelete(cMaster.ID)==true)
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            sTax.Close();

            DispClear();

            //�f�[�^�� 'darwinDataSet.����' �e�[�u���ɓǂݍ��݂܂��B
            this.�ŗ�TableAdapter.Fill(this.darwinDataSet.�ŗ�);

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
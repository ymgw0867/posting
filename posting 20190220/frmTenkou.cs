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
    public partial class frmTenkou : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�V�� cMaster = new Entity.�V��();

        const string MESSAGE_CAPTION = "�V�����";

        public frmTenkou()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            GridviewSet.Setting(dataGridView1);
            Utility.ComboTenkou.load(comboBox1);
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
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 310;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "���t");
                    tempDGV.Columns.Add("col2", "�V��");
                    tempDGV.Columns.Add("col3", "���t");
                    tempDGV.Columns.Add("col4", "�V��");

                    tempDGV.Columns[0].Width = 110;
                    tempDGV.Columns[1].Width = 110;
                    tempDGV.Columns[2].Width = 110;
                    tempDGV.Columns[3].Width = 110;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.�V�� tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.�V�� Tenkou = new Control.�V��();
                OleDbDataReader dR;

                sqlStr = " where �V��.���t = '" + dgv[0, dgv.SelectedRows[iX].Index].Value.ToString() + "'";
                dR = Tenkou.FillBy(sqlStr);

                if (dR.HasRows == true)
                {
                    while (dR.Read() == true)
                    {
                        tempC.���t = DateTime.Parse(dR["���t"].ToString());
                        tempC.�V�� = dR["�V��"].ToString() + "";
                    }
                }
                else
                {
                    dR.Close();
                    Tenkou.Close();
                    return false;
                }

                dR.Close();
                Tenkou.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV,int tempYear,int tempMonth)
            {
                string sqlSTRING = "";
                string rDate;
                int iDay,iX,c1,c2;
                DateTime sDate;

                try
                {
                    //�V��f�[�^�̃f�[�^���[�_�[���擾����
                    OleDbDataReader dR;
                    Control.�V�� dCon = new Control.�V��();

                    tempDGV.RowCount = 0;

                    iDay = 0;
                    iX = 0;

                    while (true)
                    {
                        iDay++;

                        rDate = tempYear.ToString() + "/" + tempMonth.ToString() + "/" + iDay.ToString();

                        if (DateTime.TryParse(rDate, out sDate) == true)
                        {
                            if (iX < 16)
                            {
                                tempDGV.Rows.Add();
                                c1 = 0;
                                c2 = 1;
                            }
                            else
                            {
                                c1 = 2;
                                c2 = 3;
                            }

                            
                            tempDGV[c1, iX%16].Value = rDate + "(" + ("�����ΐ��؋��y").Substring(int.Parse(sDate.DayOfWeek.ToString("d")), 1) + ")";
                            tempDGV[c2, iX%16].Value = "";

                            sqlSTRING = "where ���t = '" + rDate + "'";

                            dR = dCon.FillBy(sqlSTRING);

                            while (dR.Read())
                            {
                                tempDGV[c2, iX%16].Value = dR["�V��"].ToString();
                            }

                            dR.Close();
                        }
                        else
                        {
                            break;
                        }

                        iX++;
                    }

                    dCon.Close();

                    tempDGV.CurrentCell = null;

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }

            }

        }

        //�O���b�h����f�[�^��I��
        private void GridEnter()
        {

            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value + "���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    //txtDATE.Text = cMaster.���t.ToShortDateString();

                    //�{�^�����
                    fMode.Mode = 1;     //�t�H�[�����[�h�X�e�[�^�X:�ύX�폜

                    comboBox1.Focus();
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

                txtYear.Text = "";
                txtMonth.Text = "";
                dateTimePicker1.Text = "";
                dateTimePicker1.Enabled = false;
                comboBox1.Text = "";
                comboBox1.Enabled = false;

                btnUpdate.Enabled = false;

                dateTimePicker1.Focus();
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
                    Control.�V�� cTenkou = new Control.�V��();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                cTenkou.Close();
                                return;
                            }

                            if (cTenkou.DataInsert(cMaster) == true)
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
                                cTenkou.Close();
                                return;
                            }

                            if (cTenkou.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    cTenkou.Close();

                    //���t�ʓV��\��
                    GridviewSet.ShowData(dataGridView1, int.Parse(txtYear.Text), int.Parse(txtMonth.Text));

                    //�V��R���{�����[�h
                    Utility.ComboTenkou.load(comboBox1);

                    comboBox1.Text = "";

                    //DispClear();

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
                    //if (txtDATE.Text == null)
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("�R�[�h�͐����œ��͂��Ă�������");
                    //}

                    //str = this.txtDATE.Text;

                    //if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    //{
                    //}
                    //else
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("�R�[�h�͐����œ��͂��Ă�������");
                    //}

                    //// �����͂܂��̓X�y�[�X�݂͕̂s��
                    //if ((this.txtDATE.Text).Trim().Length < 1)
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("�R�[�h����͂��Ă�������");
                    //}

                    ////�[���͕s��
                    //if (Convert.ToInt32(this.txtDATE.Text.ToString()) == 0)
                    //{
                    //    this.txtDATE.Focus();
                    //    throw new Exception("�[���͓o�^�ł��܂���");
                    //}

                    ////�o�^�ς݃R�[�h�����ׂ�
                    //string sqlStr;
                    //Control.���� Shozoku = new Control.����();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtDATE.Text.ToString();
                    //dr = Shozoku.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtDATE.Focus();
                    //    dr.Close();
                    //    Shozoku.Close();
                    //    throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    //}

                    //dr.Close();
                    //Shozoku.Close();

                }

                //���̃`�F�b�N
                if (comboBox1.Text.Trim().Length < 1)
                {
                    comboBox1.Focus();
                    throw new Exception("�V�����͂��Ă�������");
                }

                //�����N���X�Ƀf�[�^�Z�b�g
                cMaster.���t = dateTimePicker1.Value;
                cMaster.�V��  = comboBox1.Text.ToString();

                if (fMode.Mode == 0) cMaster.�o�^�N���� = DateTime.Today;
                cMaster.�ύX�N���� = DateTime.Today;

                //�o�^�ς݂��H
                OleDbDataReader dR;
                Control.�V�� cTenkou = new Control.�V��();
                dR = cTenkou.FillBy("where ���t = '" + dateTimePicker1.Text + "'");

                if (dR.HasRows == true)
                {
                    fMode.Mode = 1;
                }
                else
                {
                    fMode.Mode = 0;
                }

                dR.Close();
                cTenkou.Close();

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;

            }

        }

        private void txtEnter(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtYear)
            {
                objtxt = txtYear;
            }

            if (sender == txtMonth)
            {
                objtxt = txtMonth;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtYear)
            {
                objtxt = txtYear;
            }

            if (sender == txtMonth)
            {
                objtxt = txtMonth;
            }

            objtxt.BackColor = Color.White;

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

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (Utility.NumericCheck(txtYear.Text) == false)
                {
                    txtYear.Focus();
                    throw new Exception("�N�͐����œ��͂��Ă�������");
                }

                if (Utility.NumericCheck(txtMonth.Text) == false)
                {
                    txtMonth.Focus();
                    throw new Exception("���͐����œ��͂��Ă�������");
                }

                if ((int.Parse(txtMonth.Text) < 1) || (int.Parse(txtMonth.Text) > 12))
                {
                    txtMonth.Focus();
                    throw new Exception("��������������܂���");
                }

                //�J�����_�[�Ώی��\��
                dateTimePicker1.Enabled = true;
                dateTimePicker1.Value = DateTime.Parse(txtYear.Text + "/" + txtMonth.Text + "/01");

                //���t�ʓV��\��
                GridviewSet.ShowData(dataGridView1, int.Parse(txtYear.Text), int.Parse(txtMonth.Text));

                //���͊J��
                dateTimePicker1.Enabled = true;
                comboBox1.Enabled = true;
                btnUpdate.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void dateTimePicker1_Validating(object sender, CancelEventArgs e)
        {
            if ((dateTimePicker1.Value).Year.ToString() != txtYear.Text || (dateTimePicker1.Value).Month.ToString() != txtMonth.Text)
            {
                MessageBox.Show("�Ώ۔N���ƈقȂ�܂�", "�N��", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Cancel = true;
                dateTimePicker1.Focus();
            }
        }

    }
}
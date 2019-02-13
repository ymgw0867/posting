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
    public partial class frmTesuuryouMeisai : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�x���T�� cMaster = new Entity.�x���T��();
        int ShowStatus;

        const string MESSAGE_CAPTION = "�z�z�萔���x���E�T������";

        public frmTesuuryouMeisai(int tempStatus)
        {
            InitializeComponent();

            ShowStatus = tempStatus;
        }

        private void form_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView1);
            GridviewSet.ShowData(dataGridView1); 

            DispClear();

            //�z�z�w����ʂ���̌Ăяo���̂Ƃ�
            if (ShowStatus == 1)
            {
                dateTimePicker1.Value = F_hDate;
                txthCode.Text = F_�z�z��ID.ToString();
                lblName.Text = F_�z�z����;
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
                    tempDGV.Columns.Add("col1", "����");
                    tempDGV.Columns.Add("col2", "���t");
                    tempDGV.Columns.Add("col3", "�z�z��ID");
                    tempDGV.Columns.Add("col4", "����");
                    tempDGV.Columns.Add("col5", "�E�v");
                    tempDGV.Columns.Add("col6", "�P��");
                    tempDGV.Columns.Add("col7", "����");
                    tempDGV.Columns.Add("col8", "���z");
                    tempDGV.Columns.Add("col9", "�x���T��");
                    tempDGV.Columns.Add("col10", "�o�^�N����");
                    tempDGV.Columns.Add("col11", "�ύX�N����");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 80;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 200;
                    tempDGV.Columns[5].Width = 70;
                    tempDGV.Columns[6].Width = 70;
                    tempDGV.Columns[7].Width = 70;
                    tempDGV.Columns[8].Width = 80;
                    tempDGV.Columns[9].Width = 90;
                    tempDGV.Columns[10].Width = 90;

                    tempDGV.Columns[1].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "yyyy/M/dd";
                    tempDGV.Columns[10].DefaultCellStyle.Format = "yyyy/M/dd";

                    tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                    tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
            public static Boolean GetData(DataGridView dgv,ref Entity.�x���T�� tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.�x���T�� sCon = new Control.�x���T��();
                OleDbDataReader dr;

                sqlStr = " where �x���T��.ID = " + int.Parse(dgv[0, dgv.SelectedRows[iX].Index].Value.ToString());
                dr = sCon.FillBy(sqlStr);

                if (dr.HasRows == false)
                {

                    dr.Close();
                    sCon.Close();
                    return false;
                }

                while (dr.Read())
                {
                    tempC.ID = int.Parse(dr["ID"].ToString());
                    tempC.���t = DateTime.Parse(dr["���t"].ToString());
                    tempC.�z�z��ID = int.Parse(dr["�z�z��ID"].ToString());
                    tempC.�z�z���� = dr["����"].ToString();
                    tempC.�E�v = dr["�E�v"].ToString() + "";
                    tempC.�P�� = double.Parse(dr["�P��"].ToString(),System.Globalization.NumberStyles.Any);
                    tempC.���� = int.Parse(dr["����"].ToString(),System.Globalization.NumberStyles.Any);
                    tempC.���z = double.Parse(dr["���z"].ToString(),System.Globalization.NumberStyles.Any);
                    tempC.�x���T���敪 = int.Parse(dr["�x���T���敪"].ToString());
                }

                dr.Close();
                sCon.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV)
            {
                int iX;

                try
                {
                    tempDGV.RowCount = 0;

                    //�x���T���}�X�^�[�̃f�[�^���[�_�[���擾����
                    OleDbDataReader dR;
                    Control.�x���T�� sCon = new Control.�x���T��();
                    dR = sCon.FillBy("order by ID desc");
                    iX = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = dR["ID"].ToString();
                        tempDGV[1, iX].Value = DateTime.Parse(dR["���t"].ToString());
                        tempDGV[2, iX].Value = dR["�z�z��ID"].ToString();
                        tempDGV[3, iX].Value = dR["����"].ToString();
                        tempDGV[4, iX].Value = dR["�E�v"].ToString();
                        tempDGV[5, iX].Value = double.Parse(dR["�P��"].ToString()).ToString("#,##0.0");
                        tempDGV[6, iX].Value = int.Parse(dR["����"].ToString()).ToString("#,##0");
                        tempDGV[7, iX].Value = double.Parse(dR["���z"].ToString()).ToString("#,##0.0");

                        switch (dR["�x���T���敪"].ToString())
                        {
                            case "0":
                                tempDGV[8, iX].Value = "�x��";
                                break;

                            case "1":
                                tempDGV[8, iX].Value = "�T��";
                                break;
                        }

                        tempDGV[9, iX].Value = DateTime.Parse(dR["�o�^�N����"].ToString());
                        tempDGV[10, iX].Value = DateTime.Parse(dR["�ύX�N����"].ToString());

                        iX++;
                    }

                    dR.Close();

                    sCon.Close();

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
            msgStr += dataGridView1[1, dataGridView1.SelectedRows[iX].Index].Value + "���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�f�[�^�I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    dateTimePicker1.Value = cMaster.���t;
                    txthCode.Text = cMaster.�z�z��ID.ToString();
                    lblName.Text = cMaster.�z�z����;
                    txtName2.Text = cMaster.�E�v;
                    txtTanka.Text = cMaster.�P��.ToString();
                    txtSuuryou.Text = cMaster.����.ToString();
                    txtkingaku.Text = cMaster.���z.ToString();

                    comboBox1.SelectedIndex = cMaster.�x���T���敪;

                    //�{�^�����
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

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

                comboBox1.SelectedIndex = -1;
                txthCode.Text = "";
                lblName.Text = "";
                txtName2.Text = "";
                txtTanka.Text = "";
                txtSuuryou.Text = "";
                txtkingaku.Text = "0";

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

                comboBox1.Focus();
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
                    Control.�x���T�� sCon = new Control.�x���T��();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                sCon.Close();
                                return;
                            }

                            if (sCon.DataInsert(cMaster) == true)
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
                                sCon.Close();
                                return;
                            }

                            if (sCon.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", "�x���T��", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    sCon.Close();

                    DispClear();

                    GridviewSet.ShowData(dataGridView1);

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
                    //Control.���� Shozoku = new Control.����();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Shozoku.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Shozoku.Close();
                    //    throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    //}

                    //dr.Close();
                    //Shozoku.Close();

                }

                //���̃`�F�b�N
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("�x���T����I�����Ă�������");
                }

                //�z�z���`�F�b�N
                if (txthCode.Text == "")
                {
                    txthCode.Focus();
                    throw new Exception("�z�z��ID��I�����Ă�������");
                }

                //�N���X�Ƀf�[�^�Z�b�g
                cMaster.���t = DateTime.Parse(dateTimePicker1.Text);
                cMaster.�z�z��ID = int.Parse(txthCode.Text);
                cMaster.�E�v = txtName2.Text.ToString();
                cMaster.�P�� = double.Parse(txtTanka.Text,System.Globalization.NumberStyles.Any);
                cMaster.���� = int.Parse(txtSuuryou.Text,System.Globalization.NumberStyles.Any);
                cMaster.���z = double.Parse(txtkingaku.Text, System.Globalization.NumberStyles.Any);
                cMaster.�x���T���敪 = comboBox1.SelectedIndex;

                if (fMode.Mode == 0) cMaster.�o�^�N���� = DateTime.Today;
                cMaster.�ύX�N���� = DateTime.Today;

                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
                
            }

        }

        private void btnDel_Click(object sender, EventArgs e)
        {   
            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�폜
            Control.�x���T�� sCon = new Control.�x���T��();
            if (sCon.DataDelete(cMaster.ID)==true)
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            sCon.Close();

            DispClear();

            GridviewSet.ShowData(dataGridView1);
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

        private void txthCode_Validating(object sender, CancelEventArgs e)
        {
            if (txthCode.Text == "") return;

            if (Utility.NumericCheck(txthCode.Text) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            OleDbDataReader dR;
            Control.�z�z�� cStaff = new Control.�z�z��();
            dR = cStaff.FillBy("where ID = " + txthCode.Text);

            if (dR.HasRows == false)
            {
                lblName.Text = "";

                MessageBox.Show("�o�^����Ă��Ȃ��z�z��ID�ł�", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            else
            {
                while (dR.Read())
                {
                    lblName.Text = dR["����"].ToString();                    
                }
            }

            dR.Close();
            cStaff.Close();

            return;
        }

        private void txthCode_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txthCode) txtObj = txthCode;
            if (sender == txtName2) txtObj = txtName2;
            if (sender == txtTanka) txtObj = txtTanka;
            if (sender == txtSuuryou) txtObj = txtSuuryou;
            if (sender == txtkingaku) txtObj = txtkingaku;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void txthCode_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txthCode) txtObj = txthCode;
            if (sender == txtName2) txtObj = txtName2;
            if (sender == txtTanka) txtObj = txtTanka;
            if (sender == txtSuuryou) txtObj = txtSuuryou;
            if (sender == txtkingaku) txtObj = txtkingaku;

            txtObj.BackColor = Color.White;
        }

        private void txtTanka_Validating(object sender, CancelEventArgs e)
        {
            double kin;

            if (txtTanka.Text == "")
            {
                txtTanka.Text = "0";
                return;
            }

            if (Utility.NumericCheck(txtTanka.Text) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (Utility.NumericCheck(txtSuuryou.Text) == true)
            {
                kin = double.Parse(txtTanka.Text) * int.Parse(txtSuuryou.Text);
                txtkingaku.Text = kin.ToString("#,##0");
                return;
            }
            else
            {
                txtkingaku.Text = "0";
            }

        }

        private void txtSuuryou_Validating(object sender, CancelEventArgs e)
        {
            try
            {

                double kin;

                if (txtSuuryou.Text == "")
                {
                    txtSuuryou.Text = "0";
                    return;
                }

                if (Utility.NumericCheck(txtSuuryou.Text) == false)
                {
                    MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }

                if (Utility.NumericCheck(txtTanka.Text) == true)
                {
                    kin = double.Parse(txtTanka.Text) * int.Parse(txtSuuryou.Text);
                    txtkingaku.Text = kin.ToString("#,##0");
                    return;
                }
                else
                {
                    txtkingaku.Text = "0";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "�P������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSuuryou.Focus();
                return;
            }

        }

        //�v���p�e�B
        private DateTime F_hDate;

        public DateTime �z�z��
        {
            set { F_hDate = value; }
        }

        private int F_�z�z��ID;

        public int �z�z��ID
        {
            set { F_�z�z��ID = value; }
        }

        private string F_�z�z����;

        public string �z�z����
        {
            get { return F_�z�z����; }
            set { F_�z�z���� = value; }
        }

    }
}
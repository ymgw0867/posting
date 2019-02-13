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
    public partial class frmSystem : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.��Џ�� cMaster = new Entity.��Џ��();

        const string MESSAGE_CAPTION = "��Џ��}�X�^�[";

        public frmSystem()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.��Џ��' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            //GridviewSet.Setting(dataGridView1);
            //this.��Џ��TableAdapter.Fill(this.darwinDataSet.��Џ��);

            //������ʃZ�b�g
            Utility.ComboKouza.load(comboBox1);

            DispClear();

            GridEnter();

        }

        private Boolean GetData(ref Entity.��Џ�� tempC)
        {
            Control.��Џ�� Kaisha = new Control.��Џ��();
            OleDbDataReader dr;

            dr = Kaisha.Fill();

            if (dr.HasRows == true)
            {
                while (dr.Read() == true)
                {
                    tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                    tempC.��Ж� = dr["��Ж�"].ToString() + "";
                    tempC.��\�Ҏ��� = dr["��\�Ҏ���"].ToString() + "";
                    tempC.��E�� = dr["��E��"].ToString() + "";
                    tempC.�d�b�ԍ� = dr["�d�b�ԍ�"].ToString() + "";
                    tempC.FAX�ԍ� = dr["FAX�ԍ�"].ToString() + "";
                    tempC.�Z��1 = dr["�Z��1"].ToString() + "";
                    tempC.�Z��2 = dr["�Z��2"].ToString() + "";
                    tempC.�X�֔ԍ� = dr["�X�֔ԍ�"].ToString() + "";
                    tempC.���[���A�h���X = dr["���[���A�h���X"].ToString() + "";
                    tempC.������ = dr["������"].ToString() + "";
                    tempC.�S���Җ� = dr["�S���Җ�"].ToString() + "";
                    tempC.���L����1 = dr["���L����1"].ToString() + "";
                    tempC.���L����2 = dr["���L����2"].ToString() + "";
                    tempC.�˗��l�R�[�h = dr["�˗��l�R�[�h"].ToString() + "";
                    tempC.�˗��l�� = dr["�˗��l��"].ToString() + "";
                    tempC.���Z�@�փR�[�h = dr["���Z�@�փR�[�h"].ToString() + "";
                    tempC.���Z�@�֖� = dr["���Z�@�֖�"].ToString() + "";
                    tempC.�x�X�R�[�h = dr["�x�X�R�[�h"].ToString() + "";
                    tempC.�x�X�� = dr["�x�X��"].ToString() + "";
                    tempC.������� = Int32.Parse(dr["�������"].ToString());
                    tempC.�����ԍ� = dr["�����ԍ�"].ToString() + "";
                    tempC.�z�z�t���O = int.Parse(dr["�z�z�t���O"].ToString());
                    tempC.�X�֔ԍ�CSV�p�X = dr["�X�֔ԍ�CSV�p�X"].ToString();
                }
            }
            else
            {
                dr.Close();
                Kaisha.Close();
                return false;
            }

            dr.Close();
            Kaisha.Close();
            return true;
        }

        //�O���b�h����f�[�^��I��
        private void GridEnter()
        {
            try
            {
                //�f�[�^���擾����
                if (GetData(ref cMaster) == true)
                {
                    //'�f�[�^�l���擾
                    //txtCode.Text = cMaster.ID.ToString();
                    txtName.Text = cMaster.��Ж�;
                    txtDaihyo.Text = cMaster.��\�Ҏ���;
                    txtYaku.Text = cMaster.��E��;
                    txtTel.Text = cMaster.�d�b�ԍ�;
                    txtFax.Text = cMaster.FAX�ԍ�;
                    txtAddress1.Text = cMaster.�Z��1;
                    txtAddress2.Text = cMaster.�Z��2;
                    mtxtZipCode.Text = cMaster.�X�֔ԍ�;
                    txtEmail.Text = cMaster.���[���A�h���X;
                    txtBusho.Text = cMaster.������;
                    txtTantou.Text = cMaster.�S���Җ�;
                    txtMemo1.Text = cMaster.���L����1;
                    txtMemo2.Text = cMaster.���L����2;
                    txtIraiCode.Text = cMaster.�˗��l�R�[�h;
                    txtIraiName.Text = cMaster.�˗��l��;
                    txtBankCode.Text = cMaster.���Z�@�փR�[�h;
                    txtBankName.Text = cMaster.���Z�@�֖�;
                    txtShitenCode.Text = cMaster.�x�X�R�[�h;
                    txtShitenName.Text = cMaster.�x�X��;
                    Utility.ComboKouza.selectedIndex(comboBox1, Int32.Parse(cMaster.�������.ToString()));
                    txtNumber.Text = cMaster.�����ԍ�;

                    txtFlg.Text = cMaster.�z�z�t���O.ToString();
                    txtZipPath.Text = cMaster.�X�֔ԍ�CSV�p�X;   // 2015/10/04

                    //ID�e�L�X�g�{�b�N�X�͕ҏW�s�Ƃ���
                    //txtCode.Enabled = false;

                    //�{�^�����

                    fMode.Mode = 1;     //�t�H�[�����[�h�X�e�[�^�X:�ύX�폜

                    txtName.Focus();
                }
                else
                {
                    fMode.Mode = 0;
                    txtName.Focus();
                }


            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }

        /// <summary>
        /// ��ʂ��N���A����
        /// </summary>
        private void DispClear()
        {

            try
            {
                fMode.Mode = 0;

                txtName.Text = "";
                txtDaihyo.Text = "";
                txtYaku.Text = "";
                txtTel.Text = "";
                txtFax.Text = "";
                txtAddress1.Text = "";
                txtAddress2.Text = "";
                mtxtZipCode.Text = "";
                txtEmail.Text = "";
                txtBusho.Text = "";
                txtTantou.Text = "";
                txtMemo1.Text = "";
                txtMemo2.Text = "";
                txtIraiCode.Text = "";
                txtIraiName.Text = "";
                txtBankCode.Text = "";
                txtBankName.Text = "";
                txtShitenCode.Text = "";
                txtShitenName.Text = "";
                comboBox1.SelectedIndex = -1;
                txtNumber.Text = "";

                //�z�z�t���O
                label22.Visible = false;
                txtFlg.Visible = false;

                txtZipPath.Text = string.Empty;     // 2015/10/04

                txtName.Focus();
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
                    Control.��Џ�� Kaisha = new Control.��Џ��();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Kaisha.Close();
                                return;
                            }

                            if (Kaisha.DataInsert(cMaster) == true)
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
                                Kaisha.Close();
                                return;
                            }

                            if (Kaisha.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Kaisha.Close();
                    this.Close();
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
                    //Control.��Џ�� Kaisha = new Control.��Џ��();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Kaisha.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Kaisha.Close();
                    //    throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    //}

                    //dr.Close();
                    //Kaisha.Close();

                }

                //��Ж��`�F�b�N
                if (txtName.Text.Trim().Length < 1)
                {
                    txtName.Focus();
                    throw new Exception("��Ж�����͂��Ă�������");
                }

                //���Z�@�փR�[�h
                if (txtBankCode.Text.ToString().Trim() != "")
                {
                    str = this.txtBankCode.Text.ToString();

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtBankCode.Focus();
                        throw new Exception("���Z�@�փR�[�h�͐����œ��͂��Ă�������");
                    }
                }

                //�x�X�R�[�h
                if (txtShitenCode.Text.ToString().Trim() != "")
                {
                    str = this.txtShitenCode.Text.ToString();

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtShitenCode.Focus();
                        throw new Exception("�x�X�R�[�h�͐����œ��͂��Ă�������");
                    }
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

                //�N���X�Ƀf�[�^�Z�b�g
                //cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.��Ж� = txtName.Text.ToString();
                cMaster.��\�Ҏ��� = txtDaihyo.Text.ToString();
                cMaster.��E�� = txtYaku.Text.ToString();
                cMaster.�d�b�ԍ� = txtTel.Text.ToString();
                cMaster.FAX�ԍ� = txtFax.Text.ToString();
                cMaster.�Z��1 = txtAddress1.Text.ToString();
                cMaster.�Z��2 = txtAddress2.Text.ToString();

                if (mtxtZipCode.Text.ToString().Replace("-", "").Trim() == "")
                {
                    cMaster.�X�֔ԍ� = "";
                }
                else
                {
                    cMaster.�X�֔ԍ� = mtxtZipCode.Text.ToString();
                }

                cMaster.���[���A�h���X = txtEmail.Text.ToString();
                cMaster.������ = txtBusho.Text.ToString();
                cMaster.�S���Җ� = txtTantou.Text.ToString();
                cMaster.���L����1 = txtMemo1.Text.ToString();
                cMaster.���L����2 = txtMemo2.Text.ToString();
                cMaster.�˗��l�R�[�h = txtIraiCode.Text.ToString();
                cMaster.�˗��l�� = txtIraiName.Text.ToString();
                cMaster.���Z�@�փR�[�h = txtBankCode.Text.ToString();
                cMaster.���Z�@�֖� = txtBankName.Text.ToString();
                cMaster.�x�X�R�[�h = txtShitenCode.Text.ToString();
                cMaster.�x�X�� = txtShitenName.Text.ToString();

                Utility.ComboKouza cmb1 = new Utility.ComboKouza();
                cmb1 = (Utility.ComboKouza)comboBox1.SelectedItem;
                cMaster.������� = cmb1.ID;
                cMaster.�����ԍ� = txtNumber.Text.ToString();

                if (txtFlg.Visible == true)
                {
                    cMaster.�z�z�t���O = int.Parse(txtFlg.Text);
                }

                if (fMode.Mode == 0) cMaster.�o�^�N���� = DateTime.Today;
                cMaster.�ύX�N���� = DateTime.Today;
                cMaster.�X�֔ԍ�CSV�p�X = txtZipPath.Text;

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
            MaskedTextBox objMtxt = new MaskedTextBox();

            if (sender == txtName)
            {
                objtxt = txtName;
            }

            if (sender == txtDaihyo)
            {
                objtxt = txtDaihyo;
            }

            if (sender == txtYaku)
            {
                objtxt = txtYaku;
            }

            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;
            }
            
            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtMemo1)
            {
                objtxt = txtMemo1;
            }

            if (sender == txtMemo2)
            {
                objtxt = txtMemo2;
            }

            if (sender == txtIraiCode)
            {
                objtxt = txtIraiCode;
            }

            if (sender == txtIraiName)
            {
                objtxt = txtIraiName;
            }
            
            if (sender == txtBankCode)
            {
                objtxt = txtBankCode;
            }
            
            if (sender == txtBankName)
            {
                objtxt = txtBankName;
            }

            if (sender == txtShitenName)
            {
                objtxt = txtShitenName;
            }
            
            if (sender == txtShitenCode)
            {
                objtxt = txtShitenCode;
            }

            if (sender == txtNumber)
            {
                objtxt = txtNumber;
            }

            if (sender == txtFlg)
            {
                objtxt = txtFlg;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;
            
            objMtxt.SelectAll();
            objMtxt.BackColor = Color.LightGray;
        }

        private void txtLeave(object sender, EventArgs e)
       {

           TextBox objtxt = new TextBox();
           MaskedTextBox objMtxt = new MaskedTextBox();

           if (sender == txtName)
           {
               objtxt = txtName;
           }

           if (sender == txtDaihyo)
           {
               objtxt = txtDaihyo;
           }

           if (sender == txtYaku)
           {
               objtxt = txtYaku;
           }

           if (sender == txtTel)
           {
               objtxt = txtTel;
           }

           if (sender == txtFax)
           {
               objtxt = txtFax;
           }

           if (sender == txtAddress1)
           {
               objtxt = txtAddress1;
           }

           if (sender == txtAddress2)
           {
               objtxt = txtAddress2;
           }

           if (sender == mtxtZipCode)
           {
               objMtxt = mtxtZipCode;
           }

           if (sender == txtEmail)
           {
               objtxt = txtEmail;
           }

           if (sender == txtBusho)
           {
               objtxt = txtBusho;
           }

           if (sender == txtTantou)
           {
               objtxt = txtTantou;
           }

           if (sender == txtMemo1)
           {
               objtxt = txtMemo1;
           }

           if (sender == txtMemo2)
           {
               objtxt = txtMemo2;
           }

           if (sender == txtIraiCode)
           {
               objtxt = txtIraiCode;
           }

           if (sender == txtIraiName)
           {
               objtxt = txtIraiName;
           }

           if (sender == txtBankCode)
           {
               objtxt = txtBankCode;
           }

           if (sender == txtBankName)
           {
               objtxt = txtBankName;
           }

           if (sender == txtShitenName)
           {
               objtxt = txtShitenName;
           }

           if (sender == txtShitenCode)
           {
               objtxt = txtShitenCode;
           }

           if (sender == txtNumber)
           {
               objtxt = txtNumber;
           }

           if (sender == txtFlg)
           {
               objtxt = txtFlg;
           }

           objtxt.BackColor = Color.White;
           objMtxt.BackColor = Color.White;

        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void txtFlg_Validating(object sender, CancelEventArgs e)
        {
            if (Utility.NumericCheck(txtFlg.Text) == false)
            {
                MessageBox.Show("������0,�܂���1�̂ݗL���ł�","�z�z�t���O");
                e.Cancel = true;
                return;
            }

            if ((txtFlg.Text != "0") && (txtFlg.Text != "1"))
            {
                MessageBox.Show("������0,�܂���1�̂ݗL���ł�", "�z�z�t���O");
                e.Cancel = true;
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�z�z�t���O�̓V�X�e���Ǘ��҂̊m�F�̂����ύX���Ă��������B��낵���ł���","�z�z�t���O�ݒ�",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            label22.Visible = true;
            txtFlg.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sPath = getZipPath();
            if (sPath != string.Empty)
            {
                txtZipPath.Text = sPath;
            }
        }

        private string getZipPath()
        {
            DialogResult ret;

            // �_�C�A���O�{�b�N�X�̏����ݒ�
            openFileDialog1.Title = "�X�֔ԍ�CSV�t�@�C���I��";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "CSV�t�@�C��(*.csv)|*.csv|�S�Ẵt�@�C��(*.*)|*.*";

            // �_�C�A���O�{�b�N�X�̕\��
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

    }
}
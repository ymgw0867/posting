using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;
using MyLibrary;
using MyTextKana;


namespace posting
{
    public partial class frmStaff : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�z�z�� cMaster = new Entity.�z�z��();

        const string MESSAGE_CAPTION = "�z�z���}�X�^�[";

        public frmStaff()
        {
            InitializeComponent();
        }

        string[] zipArray = null;

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);
            
            this.darwinDataSet.Clear();
            this.darwinDataSet.EnforceConstraints = false;
            //this.�z�z��TableAdapter.Fill(this.darwinDataSet.�z�z��);
            this.���O�C�����[�U�[TableAdapter.Fill(this.darwinDataSet.���O�C�����[�U�[);

            // �f�[�^�O���b�h�r���[��`
            gridSetting(dataGridView1);

            // �O���b�h�r���[�f�[�^�\��
            //gridSerach(dataGridView1);

            //�����R���{�f�[�^���[�h
            Utility.ComboOffice.load(cmbShozoku);

            //�x���`�ԃR���{
            cmbShiharaiSet();

            //������ʃR���{�f�[�^���[�h
            Utility.ComboKouza.load(cmbShubetsu);

            // �X�֔ԍ�CSV�ǂݍ���
            Utility.zipCsvLoad(ref zipArray);

            // ��ʏ�����
            DispClear();
        }

        #region �O���b�h�r���[�J������`
        string colID = "col1";
        string colName = "col2";
        string colFuri = "col3";
        string colZip = "col4";
        string colAdd = "col5";
        string colMobile = "col6";
        string colTel = "col7";
        string colPcMail = "col8";
        string colMbMail = "col9";
        string colTourokuDt = "col10";
        string colKinmuKbn = "col11";
        string colHaihuKbn = "col12";
        string colHaifuMemo = "col13";
        string colShiharaiKbn = "col14";
        string colOfficeCode = "col15";
        string colBankCode1 = "col16";
        string colBankName = "col17";
        string colBankFuri = "col18";
        string colShitenCode = "col19";
        string colShitenName2 = "col20";
        string colShitenFuri = "col21";
        string colKouzaKbn = "col22";
        string colKouzaNum = "col23";
        string colKouzaName = "col24";
        string colMemo = "col25";
        string colAddDt = "col26";
        string colUpDt = "col27";
        string colMyNum = "col28";
        string colUser = "col29";
        #endregion

        /// <summary>
        /// �f�[�^�O���b�h�r���[�̒�`���s���܂�
        /// </summary>
        /// <param name="tempDGV">�f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        private void gridSetting(DataGridView tempDGV)
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
                tempDGV.Height = 184;

                // ��s�̐F
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add(colID, "����");
                tempDGV.Columns.Add(colName, "����");
                tempDGV.Columns.Add(colFuri, "�t���K�i");
                tempDGV.Columns.Add(colZip, "�X�֔ԍ�");
                tempDGV.Columns.Add(colAdd, "�Z��");
                tempDGV.Columns.Add(colMobile, "�g�єԍ�");
                tempDGV.Columns.Add(colTel, "����TEL");
                tempDGV.Columns.Add(colPcMail, "PCMail");
                tempDGV.Columns.Add(colMbMail, "�g��MAil");
                tempDGV.Columns.Add(colTourokuDt, "�o�^��");
                tempDGV.Columns.Add(colKinmuKbn, "�Ζ��敪");
                tempDGV.Columns.Add(colHaihuKbn, "�X���z�z");
                tempDGV.Columns.Add(colHaifuMemo, "�X�����l");
                tempDGV.Columns.Add(colShiharaiKbn, "�x���敪");
                tempDGV.Columns.Add(colOfficeCode, "������");
                tempDGV.Columns.Add(colBankCode1, "��s�R�[�h");
                tempDGV.Columns.Add(colBankName, "���Z�@��");
                tempDGV.Columns.Add(colBankFuri, "�J�i");
                tempDGV.Columns.Add(colShitenCode, "�x�X�R�[�h");
                tempDGV.Columns.Add(colShitenName2, "�x�X��");
                tempDGV.Columns.Add(colShitenFuri, "�J�i");
                tempDGV.Columns.Add(colKouzaKbn, "�������");
                tempDGV.Columns.Add(colKouzaNum, "�����ԍ�");
                tempDGV.Columns.Add(colKouzaName, "���`");
                tempDGV.Columns.Add(colMemo, "���l");
                tempDGV.Columns.Add(colAddDt, "�o�^�N����");
                tempDGV.Columns.Add(colUpDt, "�X�V�N����");
                tempDGV.Columns.Add(colMyNum, "�}�C�i���o�[");
                tempDGV.Columns.Add(colUser, "�X�V���[�U�[");

                tempDGV.Columns[0].Width = 60;
                tempDGV.Columns[1].Width = 180;
                tempDGV.Columns[2].Width = 160;
                tempDGV.Columns[3].Width = 80;
                tempDGV.Columns[4].Width = 300;
                tempDGV.Columns[5].Width = 120;
                tempDGV.Columns[6].Width = 120;
                tempDGV.Columns[7].Width = 140;
                tempDGV.Columns[8].Width = 140;
                tempDGV.Columns[9].Width = 100;
                tempDGV.Columns[10].Width = 80;
                tempDGV.Columns[11].Width = 100;
                tempDGV.Columns[12].Width = 160;
                tempDGV.Columns[13].Width = 80;
                tempDGV.Columns[14].Width = 80;
                tempDGV.Columns[15].Width = 100;
                tempDGV.Columns[16].Width = 160;
                tempDGV.Columns[17].Width = 80;
                tempDGV.Columns[25].Width = 130;    // �o�^�N���� 2015/07/16
                tempDGV.Columns[26].Width = 130;    // �X�V�N���� 2015/07/16
                tempDGV.Columns[27].Width = 130;    // �}�C�i���o�[
                tempDGV.Columns[28].Width = 130;    // �X�V���[�U�[

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

        /// --------------------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾���� </summary>
        /// <param name="dgv">
        ///     �ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        /// --------------------------------------------------------------------------
        private Boolean GetData(DataGridView dgv,ref Entity.�z�z�� tempC)
        {
            int iX = 0;
            string sqlStr;
            Control.�z�z�� Staff = new Control.�z�z��();
            OleDbDataReader dr;

            sqlStr = " where �z�z��.ID = " + Utility.strToInt(dgv[0, dgv.SelectedRows[iX].Index].Value.ToString());
            dr = Staff.FillBy(sqlStr);

            if (dr.HasRows == true)
            {
                while (dr.Read() == true)
                {
                    tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                    tempC.���� = dr["����"].ToString() + "";
                    tempC.�t���K�i = dr["�t���K�i"].ToString() + "";
                    tempC.�X�֔ԍ� = dr["�X�֔ԍ�"].ToString() + "";
                    tempC.�Z�� = dr["�Z��"].ToString() + "";
                    tempC.�g�ѓd�b�ԍ� = dr["�g�ѓd�b�ԍ�"].ToString() + "";
                    tempC.����d�b�ԍ� = dr["����d�b�ԍ�"].ToString() + "";
                    tempC.PC�A�h���X = dr["PC�A�h���X"].ToString() + "";
                    tempC.�g�уA�h���X = dr["�g�уA�h���X"].ToString() + "";
                    tempC.�o�^�� = dr["�o�^��"].ToString() + "";
                    tempC.�Ζ��敪 = Int32.Parse(dr["�Ζ��敪"].ToString());
                    tempC.�X���z�z�敪 = Int32.Parse(dr["�X���z�z�敪"].ToString());
                    tempC.�X���z�z���l = dr["�X���z�z���l"].ToString() + "";
                    tempC.���Ə��R�[�h = Int32.Parse(dr["���Ə��R�[�h"].ToString());
                    tempC.�x���敪 = dr["�x���敪"].ToString() + "";
                    tempC.���Z�@�փR�[�h = dr["���Z�@�փR�[�h"].ToString();
                    tempC.���Z�@�֖� = dr["���Z�@�֖�"].ToString() + "";
                    tempC.���Z�@�֖��J�i = dr["���Z�@�֖��J�i"].ToString() + "";
                    tempC.�x�X�R�[�h = dr["�x�X�R�[�h"].ToString();
                    tempC.�x�X�� = dr["�x�X��"].ToString() + "";
                    tempC.�x�X���J�i = dr["�x�X���J�i"].ToString() + "";
                    tempC.������� = Int32.Parse(dr["�������"].ToString());
                    tempC.�����ԍ� = dr["�����ԍ�"].ToString() + "";
                    tempC.�������`�J�i = dr["�������`�J�i"].ToString() + "";
                    tempC.���l = dr["���l"].ToString() + "";
                    tempC.�}�C�i���o�[ = dr["�}�C�i���o�["].ToString() + "";
                    tempC.���[�U�[ID = Utility.strToInt(dr["���[�U�[ID"].ToString());
                }
            }
            else
            {
                dr.Close();
                Staff.Close();
                return false;
            }

            dr.Close();
            Staff.Close();
            return true;
        }


        //�O���b�h����f�[�^��I��
        private void GridEnter()
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[colName, dataGridView1.SelectedRows[iX].Index].Value + "���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�z�z���I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    //�f�[�^���擾����
                    if (GetData(dataGridView1,ref cMaster) == false)
                    {
                        MessageBox.Show("�Y������f�[�^���}�X�^�[�ɓo�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'�f�[�^�l���擾
                    txtCode.Text = cMaster.ID.ToString();
                    txtName1.Text = cMaster.����;
                    txtFuri.Text = cMaster.�t���K�i;
                    mtxtZipCode.Text = cMaster.�X�֔ԍ�;
                    txtAddress.Text = cMaster.�Z��;
                    txtMobile.Text = cMaster.�g�ѓd�b�ԍ�;
                    txtTel.Text = cMaster.����d�b�ԍ�;
                    txteMail.Text = cMaster.PC�A�h���X;
                    txteMailm.Text = cMaster.�g�уA�h���X;

                    if (cMaster.�o�^�� == "")
                    {
                        iDate.Checked = false;
                    }
                    else
                    {
                        iDate.Checked = true;
                        iDate.Value = Convert.ToDateTime(cMaster.�o�^��);
                    }

                    if (cMaster.�Ζ��敪 == 1)
                    {
                        chkKinmu.Checked = true;
                    }
                    else
                    {
                        chkKinmu.Checked = false;
                    }

                    if (cMaster.�X���z�z�敪 == 1)
                    {
                        chkGaitou.Checked = true;
                    }
                    else
                    {
                        chkGaitou.Checked = false;
                    }

                    txtGaitouMemo.Text = cMaster.�X���z�z���l;
                    cmbShiharai.Text = cMaster.�x���敪;

                    Utility.ComboOffice.selectedIndex(cmbShozoku, cMaster.���Ə��R�[�h);

                    txtBankCode.Text = cMaster.���Z�@�փR�[�h;
                    txtBank.Text = cMaster.���Z�@�֖�;
                    txtBankFuri.Text = cMaster.���Z�@�֖��J�i;

                    txtShitenCode.Text = cMaster.�x�X�R�[�h;
                    txtShiten.Text = cMaster.�x�X��;
                    txtShitenFuri.Text = cMaster.�x�X���J�i;

                    Utility.ComboKouza.selectedIndex(cmbShubetsu, cMaster.�������);
                    txtKouza.Text = cMaster.�����ԍ�;
                    txtMeigi.Text = cMaster.�������`�J�i;
                    txtMyNumber.Text = cMaster.�}�C�i���o�[;      // 2015/07/16
                    txtMemo.Text = cMaster.���l.ToString();

                    //ID�e�L�X�g�{�b�N�X�͕ҏW�s�Ƃ���
                    txtCode.Enabled = false;

                    //�{�^�����
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //�t�H�[�����[�h�X�e�[�^�X:�ύX�폜

                    txtName1.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "��ʕ\��", MessageBoxButtons.OK);
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

                txtCode.Enabled = true;
                txtCode.Text = "";
                txtName1.Text = "";
                txtFuri.Text = "";
                mtxtZipCode.Text = "";
                txtAddress.Text = "";
                txtMobile.Text = "";
                txtTel.Text = "";
                txteMail.Text = "";
                txteMailm.Text = "";
                iDate.Checked = false;
                chkKinmu.Checked = false;
                chkGaitou.Checked = false;
                txtGaitouMemo.Text = "";
                cmbShiharai.SelectedIndex = -1;
                cmbShiharai.SelectedIndex = 0;  // 2018/02/22 �T���f�t�H���g�Ƃ���
                //cmbShozoku.SelectedIndex = -1;
                cmbShozoku.SelectedIndex = 1;   // 2018/02/22 �V�h���f�t�H���g�Ƃ��� 
                txtBankCode.Text = "";
                txtBank.Text = "";
                txtBankFuri.Text = "";
                txtShitenCode.Text = "";
                txtShiten.Text = "";
                txtShitenFuri.Text = "";
                cmbShubetsu.SelectedIndex = -1;
                txtKouza.Text = "";
                txtMeigi.Text = "";
                txtMemo.Text = "";
                txtMyNumber.Text = "";  // 2015/07/16

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

                txtCode.Focus();
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
                    Control.�z�z�� Staff = new Control.�z�z��();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Staff.Close();
                                return;
                            }

                            if (Staff.DataInsert(cMaster) == true)
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
                                Staff.Close();
                                return;
                            }

                            if (Staff.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", "�����}�X�^�[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Staff.Close();

                    DispClear();

                    //�f�[�^�� 'darwinDataSet.�z�z��' �e�[�u���ɓǂݍ��݂܂��B
                    this.�z�z��TableAdapter.Fill(this.darwinDataSet.�z�z��);
                    //this.�z�z��gridviewTableAdapter.Fill(this.darwinDataSet.�z�z��gridview);

                    // �O���b�h�ĕ\��
                    gridSerach(dataGridView1);
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
                    if (txtCode.Text == null)
                    {
                        this.txtCode.Focus();
                        throw new Exception("�R�[�h�͐����œ��͂��Ă�������");
                    }

                    str = this.txtCode.Text;

                    if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                    {
                    }
                    else
                    {
                        this.txtCode.Focus();
                        throw new Exception("�R�[�h�͐����œ��͂��Ă�������");
                    }

                    // �����͂܂��̓X�y�[�X�݂͕̂s��
                    if ((this.txtCode.Text).Trim().Length < 1)
                    {
                        this.txtCode.Focus();
                        throw new Exception("�R�[�h����͂��Ă�������");
                    }

                    //�[���͕s��
                    if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                    {
                        this.txtCode.Focus();
                        throw new Exception("�[���͓o�^�ł��܂���");
                    }

                    //�o�^�ς݃R�[�h�����ׂ�
                    string sqlStr;
                    Control.�z�z�� Staff = new Control.�z�z��();
                    OleDbDataReader dr;

                    sqlStr = " where ID = " + txtCode.Text.ToString();
                    dr = Staff.FillBy(sqlStr);

                    if (dr.HasRows == true)
                    {
                        txtCode.Focus();
                        dr.Close();
                        Staff.Close();
                        throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    }

                    dr.Close();
                    Staff.Close();
                }

                //���̃`�F�b�N
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("��������͂��Ă�������");
                }

                ////�x���`�ԃ`�F�b�N
                //if (cmbShiharai.SelectedIndex == -1)
                //{
                //    cmbShiharai.Focus();
                //    throw new Exception("�x���`�Ԃ�I�����Ă�������");
                //}

                ////���Ə��`�F�b�N
                //if (cmbShozoku.SelectedIndex == -1)
                //{
                //    cmbShozoku.Focus();
                //    throw new Exception("���Ə���I�����Ă�������");
                //}

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

                // �x�X�R�[�h
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

                // �}�C�i���o�[
                if (txtMyNumber.Text != string.Empty)
                {
                    if (Utility.strToLong(txtMyNumber.Text) == 0)
                    {
                        this.txtShitenCode.Focus();
                        throw new Exception("�}�C�i���o�[�͐����œ��͂��Ă�������");
                    }

                    if (txtMyNumber.Text.Length != 12)
                    {
                        this.txtShitenCode.Focus();
                        throw new Exception("�}�C�i���o�[�͐�����12�����͂��Ă�������");
                    }
                }

                //�N���X�Ƀf�[�^�Z�b�g
                cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.���� = txtName1.Text.ToString();
                cMaster.�t���K�i = txtFuri.Text.ToString();

                if (mtxtZipCode.Text.ToString().Replace("-", "").Trim() == "")
                {
                    cMaster.�X�֔ԍ� = "";
                }
                else
                {
                    cMaster.�X�֔ԍ� = mtxtZipCode.Text.ToString();
                }

                cMaster.�Z�� = txtAddress.Text.ToString();
                cMaster.�g�ѓd�b�ԍ� = txtMobile.Text.ToString();
                cMaster.����d�b�ԍ� = txtTel.Text.ToString();
                cMaster.PC�A�h���X = txteMail.Text.ToString();
                cMaster.�g�уA�h���X = txteMailm.Text.ToString();

                if (iDate.Checked == false)
                {
                    cMaster.�o�^�� = "";
                }
                else
                {
                    cMaster.�o�^�� = iDate.Value.ToShortDateString();
                }

                if (chkKinmu.Checked == true)
                {
                    cMaster.�Ζ��敪 = 1;
                }
                else
                {
                    cMaster.�Ζ��敪 = 0;
                }

                if (chkGaitou.Checked == true)
                {
                    cMaster.�X���z�z�敪 = 1;
                }
                else
                {
                    cMaster.�X���z�z�敪 = 0;
                }

                cMaster.�X���z�z���l = txtGaitouMemo.Text.ToString();
                cMaster.�x���敪 = cmbShiharai.Text;

                if (cmbShozoku.SelectedIndex == -1)
                {
                    cMaster.���Ə��R�[�h = 0;
                }
                else
                {
                    Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                    cmb1 = (Utility.ComboOffice)cmbShozoku.SelectedItem;
                    cMaster.���Ə��R�[�h = cmb1.ID;
                }

                cMaster.���Z�@�փR�[�h = txtBankCode.Text.ToString().Trim();
                cMaster.���Z�@�֖� = txtBank.Text.ToString();

                if (txtBankFuri.Text.ToString().Length > 15)
                {
                    cMaster.���Z�@�֖��J�i = txtBankFuri.Text.ToString().Substring(0, 15);
                }
                else
                {
                    cMaster.���Z�@�֖��J�i = txtBankFuri.Text.ToString();
                }

                cMaster.�x�X�R�[�h = txtShitenCode.Text.ToString().Trim();
                cMaster.�x�X�� = txtShiten.Text.ToString();

                if (txtShitenFuri.Text.ToString().Length > 15)
                {
                    cMaster.�x�X���J�i = txtShitenFuri.Text.ToString().Substring(0, 15);
                }
                else
                {
                    cMaster.�x�X���J�i = txtShitenFuri.Text.ToString();
                }

                if (cmbShubetsu.SelectedIndex == -1)
                {
                    cMaster.������� = 0;
                }
                else
                {
                    Utility.ComboKouza cmb2 = new Utility.ComboKouza();
                    cmb2 = (Utility.ComboKouza)cmbShubetsu.SelectedItem;
                    cMaster.������� = cmb2.ID;
                }

                cMaster.�����ԍ� = txtKouza.Text.ToString();
                cMaster.�������`�J�i = txtMeigi.Text.ToString();
                cMaster.���l = txtMemo.Text.ToString();

                if (fMode.Mode == 0)
                {
                    cMaster.�o�^�N���� = DateTime.Now;
                    cMaster.�ύX�N���� = DateTime.Now;
                }
                else
                {
                    cMaster.�ύX�N���� = DateTime.Now;
                }

                cMaster.�}�C�i���o�[ = txtMyNumber.Text;      // 2015/07/16
                cMaster.���[�U�[ID = global.loginUserID;      // 2015/07/16

                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "�ێ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        private void txtMEnter(object sender, EventArgs e)
        {
            MaskedTextBox objMtxt = (MaskedTextBox)sender;
            
            objMtxt.SelectAll();
            objMtxt.BackColor = Color.LightGray;
        }

        private void txtEnter(object sender, EventArgs e)
        {
            TextBox objtxt = (TextBox)sender;

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

            // 2019/03/16
            if (objtxt == txtName1 || objtxt == txtBank || objtxt == txtShiten)
            {
                MyTextKana.TextKana.textEnter(objtxt);
            }
        }

        private void txtMLeave(object sender, EventArgs e)
        {
            MaskedTextBox objMtxt = (MaskedTextBox)sender;
            objMtxt.BackColor = Color.White;
        }

        private void txtLeave(object sender, EventArgs e)
        {
            TextBox objtxt = (TextBox)sender;
            objtxt.BackColor = Color.White;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            // ���ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(�z�z�w��.�z�z��ID = " + txtCode.Text.ToString() + ")  ";

            OleDbDataReader dr;
            Control.�z�z�w�� Shiji = new Control.�z�z�w��();
            dr = Shiji.FillBy(SqlStr);

            // �Y���z�z�����o�^����Ă���Ƃ��͍폜�s�Ƃ���
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "���z�z�w���f�[�^�ɓo�^����Ă��܂�", txtName1.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Shiji.Close();
                return;
            }

            dr.Close();
            Shiji.Close();
            
            // �폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // �f�[�^�폜
            Control.�z�z�� Staff = new Control.�z�z��();
            if (Staff.DataDelete(Convert.ToInt32(txtCode.Text.ToString()))==true)
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Staff.Close();

            DispClear();

            // �f�[�^�� 'darwinDataSet.�z�z��' �e�[�u���ɓǂݍ��݂܂��B
            this.�z�z��TableAdapter.Fill(this.darwinDataSet.�z�z��);
            //this.�z�z��TableAdapter.Fill(this.darwinDataSet.�z�z��gridview);

            // �O���b�h�ĕ\��
            gridSerach(dataGridView1);
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

        private void cmbShiharaiSet()
        {
            cmbShiharai.Items.Clear();
            cmbShiharai.Items.Add("�T");
            cmbShiharai.Items.Add("��");
        }

        private void label14_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Form frm = new frmOffice();
            frm.ShowDialog();

            //���Ə��R���{�{�b�N�X�f�[�^���[�h
            Utility.ComboOffice.load(this.cmbShozoku);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // �O���b�h�r���[�ɍi�荞�ݕ\��
            gridSerach(dataGridView1);
        }

        private void gridSerach(DataGridView d)
        {
            ////�f�[�^�� 'darwinDataSet.�z�z��' �e�[�u���ɓǂݍ��݂܂��B
            //darwinDataSet ds = new darwinDataSet();
            //ds.Clear();
            //ds.EnforceConstraints = false;

            //const int MINID = 0;
            //const int MAXID = 999999;

            //int ID1, ID2;

            ////�z�z��ID�̎w��L��
            //if (textBoxID.Text.Length > 0)
            //{
            //    ID1 = int.Parse(textBoxID.Text);
            //    ID2 = int.Parse(textBoxID.Text);
            //}
            //else
            //{
            //    ID1 = MINID;
            //    ID2 = MAXID;
            //}

            //this.�z�z��TableAdapter.FillByName(ds.�z�z��, ID1, ID2, "%" + textBox1.Text.ToString() + "%");
            //dataGridView1.DataSource = ds.�z�z��;


            // 2015/07/17
            this.Cursor = Cursors.WaitCursor;   // �ҋ@�J�[�\����\��
            this.�z�z��TableAdapter.Fill(this.darwinDataSet.�z�z��);
            //this.�z�z��TableAdapter.Fill(this.darwinDataSet.�z�z��gridview);

            // 2015/07/16
            var s = this.darwinDataSet.�z�z��.Where(a => a.ID > 0);

            if (textBoxID.Text != string.Empty)
            {
                s = s.Where(a => a.ID == Utility.strToInt(textBoxID.Text.Trim()));
            }

            if (textBox1.Text != string.Empty)
            {
                s = s.Where(a => a.����.Contains(textBox1.Text));
            }

            d.Rows.Clear();
            int r = 0;

            foreach (var t in s)
            {
                d.Rows.Add();

                d[colID, r].Value = t.ID.ToString();
                d[colName, r].Value = t.����;
                d[colFuri, r].Value = t.�t���K�i;
                d[colZip, r].Value = t.�X�֔ԍ�;
                d[colAdd, r].Value = t.�Z��;
                d[colMobile, r].Value = t.�g�ѓd�b�ԍ�;
                d[colTel, r].Value = t.����d�b�ԍ�;
                d[colPcMail, r].Value = t.PC�A�h���X;
                d[colMbMail, r].Value = t.�g�уA�h���X;

                if (t.Is�o�^��Null())
                {
                    d[colTourokuDt, r].Value = string.Empty;
                }
                else
                {
                    d[colTourokuDt, r].Value = t.�o�^��.ToShortDateString();
                }

                d[colKinmuKbn, r].Value = t.�Ζ��敪.ToString();
                d[colHaihuKbn, r].Value = t.�X���z�z�敪.ToString();
                d[colHaifuMemo, r].Value = t.�X���z�z���l;
                d[colShiharaiKbn, r].Value = t.�x���敪;
                d[colOfficeCode, r].Value = t.���Ə��R�[�h.ToString();
                d[colBankCode1, r].Value = t.���Z�@�փR�[�h;
                d[colBankName, r].Value = t.���Z�@�֖�;
                d[colBankFuri, r].Value = t.���Z�@�֖��J�i;
                d[colShitenCode, r].Value = t.�x�X�R�[�h;
                d[colShitenName2, r].Value = t.�x�X��;
                d[colShitenFuri, r].Value = t.�x�X���J�i;
                d[colKouzaKbn, r].Value = t.�������.ToString();
                d[colKouzaNum, r].Value = t.�����ԍ�;
                d[colKouzaName, r].Value = t.�������`�J�i;
                d[colMemo, r].Value = t.���l;
                d[colAddDt, r].Value = t.�o�^�N����;
                d[colUpDt, r].Value = t.�ύX�N����;
                d[colMyNum, r].Value = t.�}�C�i���o�[;

                if (t.���O�C�����[�U�[Row == null)
                {
                    d[colUser, r].Value = string.Empty;
                }
                else
                {
                    d[colUser, r].Value = t.���O�C�����[�U�[Row.���O�C�����[�U�[;
                }

                r++;
            }

            d.CurrentCell = null;
            this.Cursor = Cursors.Default;      // �J�[�\����߂�
        }
        
        private void txtName1_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtName1.Text) == false)
            {
                txtFuri.Text = "";
                txtMeigi.Text = "";
            }           
        }

        private void txtName1_KeyDown(object sender, KeyEventArgs e)
        {
            string furiName;
            furiName  = MyTextKana.TextKana.textBox_KeyDown(txtName1, sender, e);

            //txtFuri.Text = furiName;      // 2019/03/16 �R�����g��
            //txtMeigi.Text = furiName;     // 2019/03/16 �R�����g��

            // 2019/03/16
            if (furiName != "")
            {
                txtFuri.Text += furiName; 
                txtMeigi.Text += furiName;
            }
        }

        private void txtBank_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtBank.Text) == false)
            {
                txtBankFuri.Text = "";
            }
        }

        private void txtBank_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/16
            string furi = MyTextKana.TextKana.textBox_KeyDown(txtBank, sender, e);

            // 2019/03/16
            txtBankFuri.Text += furi;
        }

        private void txtShiten_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtShiten.Text) == false)
            {
                txtShiten.Text = "";
            }
        }

        private void txtShiten_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/16
            string furi = MyTextKana.TextKana.textBox_KeyDown(txtShiten, sender, e);

            // 2019/03/16
            txtShitenFuri.Text += furi;
        }

        private void textBoxID_Validated(object sender, EventArgs e)
        {
            if (textBoxID.Text.Length > 0)
            {
                if (Utility.NumericCheck(textBoxID.Text) == false)
                {
                    MessageBox.Show("�z�z��ID�͐����œ��͂��Ă�������","����ID");
                    textBoxID.Focus();
                }
            }
        }

        private void txtMyNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void txtMobile_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        private void fillByAllToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.�z�z��TableAdapter.FillByAll(this.darwinDataSet.�z�z��);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillByAllToolStripButton_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.�z�z��TableAdapter.FillByAll(this.darwinDataSet1.�z�z��);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillByAllToolStripButton_Click_2(object sender, EventArgs e)
        {
            try
            {
                this.�z�z��TableAdapter.FillByAll(this.darwinDataSet1.�z�z��);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void mtxtZipCode_TextChanged(object sender, EventArgs e)
        {
            if (zipArray == null)
            {
                return;
            }

            MaskedTextBox mTxt = null;
            TextBox txtAdd = null;
            bool mc = false;

            mTxt = mtxtZipCode;
            txtAdd = txtAddress;

            string zipText = mTxt.Text.Replace("-", "");

            if (zipText.Length == 7)
            {
                foreach (var t in zipArray)
                {
                    string[] r = t.Split(',');

                    if (zipText == r[2].Replace("\"", ""))
                    {
                        // �Z��
                        string ad = r[6].Replace("\"", "") + r[7].Replace("\"", "") + r[8].Replace("\"", "");

                        // �����Z���Ȃ�΍ĕ\�����Ȃ�
                        if (!txtAdd.Text.Contains(ad))
                        {
                            txtAdd.Text = ad;
                        }
                        txtAdd.Focus();
                        mc = true;
                        break;
                    }
                }
            }
        }
    }
}
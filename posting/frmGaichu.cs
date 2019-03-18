using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;
using System.Linq;

namespace posting
{
    public partial class frmGaichu : Form
    {
        public frmGaichu()
        {
            InitializeComponent();

            // �f�[�^��ǂݍ���
            adp.Fill(dts.�O����);
            lAdp.Fill(dts.���O�C�����[�U�[);
        }

        string[] zipArray = null;

        Utility.formMode fMode = new Utility.formMode();
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.�O����TableAdapter adp = new darwinDataSetTableAdapters.�O����TableAdapter();
        darwinDataSetTableAdapters.��1TableAdapter jAdp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.���O�C�����[�U�[TableAdapter lAdp = new darwinDataSetTableAdapters.���O�C�����[�U�[TableAdapter();

        const string MESSAGE_CAPTION = "�O����}�X�^�[";

        // �f�[�^�O���b�h�J������`
        string colID = "col1";
        string colName = "col2";
        string colZip = "col3";
        string colAdd1 = "col4";
        string colAdd2 = "col5";
        string colAddDate = "col6";
        string colUpDate = "col7";
        string colUserID = "col8";
        string colSday = "col9";    // �x���� 2018/01/03

        bool jStatus = false;

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // �O���b�h�r���[��`
            gridViewSetting(dataGridView1);

            // �O���b�h�r���[�f�[�^�\��
            gridViewDataShow(dataGridView1, textBox1.Text);
            
            //�S���Ј��R���{
            Utility.ComboShain.load(cmbShain);

            // ��ʏ�����
            DispClear();
            textBox1.Text = string.Empty;

            // �X�֔ԍ�CSV�ǂݍ���
            Utility.zipCsvLoad(ref zipArray);
        }

        /// <summary>
        /// �f�[�^�O���b�h�r���[�̒�`���s���܂�
        /// </summary>
        /// <param name="tempDGV">�f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        private void gridViewSetting(DataGridView tempDGV)
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
                tempDGV.Height = 217;

                // ��s�̐F
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add(colID, "�R�[�h");
                tempDGV.Columns.Add(colName, "����");
                tempDGV.Columns.Add(colZip, "�X�֔ԍ�");
                tempDGV.Columns.Add(colAdd1, "�Z���P");
                tempDGV.Columns.Add(colAdd2, "�Z���Q");
                tempDGV.Columns.Add(colAddDate, "�o�^�N����");
                tempDGV.Columns.Add(colUpDate, "�X�V�N����");
                tempDGV.Columns.Add(colUserID, "���[�U�[ID");
                tempDGV.Columns.Add(colSday, "�x����");

                tempDGV.Columns[colID].Width = 90;
                tempDGV.Columns[colName].Width = 200;
                tempDGV.Columns[colZip].Width = 110;
                tempDGV.Columns[colAdd1].Width = 200;
                tempDGV.Columns[colAdd2].Width = 200;
                tempDGV.Columns[colAddDate].Width = 160;
                tempDGV.Columns[colUpDate].Width = 160;
                tempDGV.Columns[colUserID].Width = 110;
                tempDGV.Columns[colSday].Width = 90;

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

        private void gridViewDataShow(DataGridView dg, string sName)
        {
            int r = 0;
            dg.Rows.Clear();

            foreach (var t in dts.�O����.OrderBy(a => a.ID))
            {
                if (sName ==string.Empty || (sName != string.Empty && t.����.Contains(sName)))
                {
                    dg.Rows.Add();
                    dg[colID, r].Value = t.ID.ToString();
                    dg[colName, r].Value = t.����;
                    dg[colZip, r].Value = t.�X�֔ԍ�;
                    dg[colAdd1, r].Value = t.�Z��1;
                    dg[colAdd2, r].Value = t.�Z��2;
                    dg[colAddDate, r].Value = t.�o�^�N����.ToString();
                    dg[colUpDate, r].Value = t.�X�V�N����.ToString();

                    if (t.���O�C�����[�U�[Row == null)
                    {
                        dg[colUserID, r].Value = string.Empty;
                    }
                    else
                    {
                        dg[colUserID, r].Value = t.���O�C�����[�U�[Row.���O�C�����[�U�[;
                    }

                    dg[colSday, r].Value = t.�x����.ToString();

                    r++;
                }
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     �O���b�h����f�[�^��I�� </summary>
        ///-----------------------------------------------------------
        private void GridEnter()
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[colName, dataGridView1.SelectedRows[iX].Index].Value + "���I������܂���" + Environment.NewLine;
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�O����I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    //�f�[�^���擾����
                    int sID = int.Parse(dataGridView1[colID, dataGridView1.SelectedRows[0].Index].Value.ToString());

                    darwinDataSet.�O����Row r = (darwinDataSet.�O����Row)dts.�O����.Single(a => a.ID == sID);

                    // �f�[�^�l���擾
                    txtFuri.Text = r.�t���K�i;
                    txtName2.Text = r.����;
                    txtTantou.Text = r.�S���Җ�;
                    txtBusho.Text = r.�S������;
                    mtxtZipCode.Text = r.�X�֔ԍ�;
                    txtAddress1.Text = r.�Z��1;
                    txtAddress2.Text = r.�Z��2;
                    txtTel.Text = r.�d�b�ԍ�;
                    txtFax.Text = r.FAX�ԍ�;
                    txtEmail.Text = r.eMail;
                    cmbShain.SelectedValue = r.�S���c��;
                    txtMemo.Text = r.���l;
                    txtShiharaibi.Text = r.�x����.ToString();

                    //ID�e�L�X�g�{�b�N�X�͕ҏW�s�Ƃ���
                    //txtCode.Enabled = false;

                    //�{�^�����
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     // �t�H�[�����[�h�X�e�[�^�X:�ύX�폜
                    fMode.ID = sID;

                    txtName2.Focus();
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

                txtFuri.Text = "";
                txtName2.Text = "";
                txtTantou.Text = "";
                txtBusho.Text = "";
                mtxtZipCode.Text = "";
                txtAddress1.Text = "";
                txtAddress2.Text = "";
                txtTel.Text = "";
                txtFax.Text = "";
                txtEmail.Text = "";
                cmbShain.SelectedIndex = -1;
                txtMemo.Text = "";
                txtShiharaibi.Text = "";

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

                txtName2.Focus();
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
                if (fDataCheck())
                {
                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                return;
                            }

                            if (dataAdd())
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
                                return;
                            }

                            dataUpDate();

                            break;
                    }

                    // �f�[�^�x�[�X�X�V
                    adp.Update(dts.�O����);

                    //�f�[�^�� 'darwinDataSet.�O����' �e�[�u���ɓǂݍ��݂܂��B
                    adp.Fill(dts.�O����);

                    // ��ʏ�����
                    DispClear();

                    // �O���b�h�ĕ\��
                    gridViewDataShow(dataGridView1, textBox1.Text);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"�X�V����",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }


        /// -----------------------------------------------------------------
        /// <summary>
        ///     �O����f�[�^�Z�b�g�X�V </summary>
        /// <returns>
        ///     true:�s�X�V�����Afalse:�s�X�V���s</returns>
        /// -----------------------------------------------------------------
        private bool dataUpDate()
        {
            try
            {
                darwinDataSet.�O����Row s = dts.�O����.Single(a => a.ID == fMode.ID);
                dataRowSet(s);
                
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
        
        /// -----------------------------------------------------------------
        /// <summary>
        ///     �O����V�K�o�^ </summary>
        /// <returns>
        ///     true:�s�ǉ������Afalse:�s�ǉ����s</returns>
        /// -----------------------------------------------------------------
        private bool dataAdd()
        {
            try
            {
                darwinDataSet.�O����Row r = dts.�O����.New�O����Row();
                dataRowSet(r);

                // �f�[�^�Z�b�g�ɍs�ǉ�
                dts.�O����.Add�O����Row(r);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     �O����f�[�^Row�Ƀf�[�^�Z�b�g </summary>
        /// <param name="r">
        ///     darwinDataSet.�O����Row </param>
        /// --------------------------------------------------------------------
        private darwinDataSet.�O����Row dataRowSet(darwinDataSet.�O����Row r)
        {
            r.���� = txtName2.Text;
            r.�t���K�i = txtFuri.Text;
            r.�S���Җ� = txtTantou.Text;
            r.�S������ = txtBusho.Text;

            if (((mtxtZipCode.Text).Replace("-", "")).Trim() == "")
            {
                r.�X�֔ԍ� = "";
            }
            else
            {
                r.�X�֔ԍ� = mtxtZipCode.Text;
            }

            r.�Z��1 = txtAddress1.Text;
            r.�Z��2 = txtAddress2.Text;
            r.�d�b�ԍ� = txtTel.Text;
            r.FAX�ԍ� = txtFax.Text;
            r.eMail = txtEmail.Text;

            Utility.ComboShain cmb1 = new Utility.ComboShain();

            if (cmbShain.SelectedIndex == -1)
            {
                r.�S���c�� = 0;
            }
            else
            {
                cmb1 = (Utility.ComboShain)cmbShain.SelectedItem;
                r.�S���c�� = cmb1.ID;
            }

            r.���l = txtMemo.Text;

            if (fMode.Mode == 0)
            {
                r.�o�^�N���� = DateTime.Now;
            }

            r.�X�V�N���� = DateTime.Now;
            r.���[�U�[ID = global.loginUserID;
            r.�x���� = Utility.strToInt(txtShiharaibi.Text);

            return r;
        }


        ///--------------------------------------------------------------
        /// <summary>
        ///     �o�^�f�[�^�`�F�b�N </summary>
        /// <returns>
        ///     true:�G���[�Ȃ�, false:�G���[�L��</returns>
        ///--------------------------------------------------------------
        private Boolean fDataCheck()
        {
            try
            {
                // ���̃`�F�b�N
                if (txtName2.Text.Trim().Length < 1)
                {
                    txtName2.Focus();
                    throw new Exception("���̂���͂��Ă�������");
                }

                // �x�����`�F�b�N
                if (Utility.strToInt(txtShiharaibi.Text) > 31)
                {
                    txtShiharaibi.Focus();
                    throw new Exception("�x����������������܂���");
                }

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

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}
            
            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;

                // 2019/03/16
                MyTextKana.TextKana.textEnter(txtName2);
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }
            
            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == txtShiharaibi)
            {
                objtxt = txtShiharaibi;
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

            //if (sender == txtCode)
            //{
            //    objtxt = txtCode;
            //}
            
            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;
            }
            
            if (sender == txtTel)
            {
                objtxt = txtTel;
            }

            if (sender == txtFax)
            {
                objtxt = txtFax;
            }

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;
            }

            if (sender == txtShiharaibi)
            {
                objtxt = txtShiharaibi;
            }

            objtxt.BackColor = Color.White;
            objMtxt.BackColor = Color.White;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            dataDel();
        }

        private void dataDel()
        {
            // �폜�m�F�F2018/01/04
            if (MessageBox.Show("�폜�������I������܂����B���s���Ă�낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // �f�[�^�ǂݍ��݁F2018/01/04
            if (!jStatus)
            {
                Cursor = Cursors.WaitCursor;
                jAdp.Fill(dts.��1);
                jStatus = true;
                Cursor = Cursors.Default;
            }

            // �󒍃f�[�^�ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���
            if (dts.��1.Any(a => a.�O����ID�c�� == fMode.ID || a.�O����ID�x�� == fMode.ID))
            {
                MessageBox.Show(txtName2.Text.ToString() + " �͎󒍃f�[�^�ɓo�^����Ă��܂�", txtName2.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // �폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // �f�[�^�폜
            if (dts.�O����.Any(a => a.ID == fMode.ID))
            {
                darwinDataSet.�O����Row r = dts.�O����.Single(a => a.ID == fMode.ID);
                r.Delete();

                // �f�[�^�x�[�X�X�V
                adp.Update(dts.�O����);

                // ���b�Z�[�W
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // ���b�Z�[�W
                MessageBox.Show("�Y������O����f�[�^������܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            // ��ʏ�����
            DispClear();

            //�f�[�^�� 'darwinDataSet.�O����' �e�[�u���ɓǂݍ��݂܂��B
            adp.Fill(dts.�O����);

            // �O���b�h�ĕ\��
            gridViewDataShow(dataGridView1, textBox1.Text);
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

        private void button1_Click(object sender, EventArgs e)
        {
            gridViewDataShow(dataGridView1, textBox1.Text);
        }
        
        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmShain();

            frm.ShowDialog();
            Utility.ComboShain.load(cmbShain);
        }
        
        private void txtName2_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/16
            string furi = MyTextKana.TextKana.textBox_KeyDown(txtName2, sender, e);

            // 2019/03/16
            txtFuri.Text += furi;
        }

        private void txtName2_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtName2.Text) == false)
            {
                txtFuri.Text = "";
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
            txtAdd = txtAddress1;

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

        private void txtShiharaibi_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
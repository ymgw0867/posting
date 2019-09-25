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
    public partial class frmPosting : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Utility.areaMode aMode = new Utility.areaMode();
        Entity.�� cMaster = new Entity.��();
        Entity.�z�z�G���A cArea = new Entity.�z�z�G���A();

        const string MESSAGE_CAPTION = "�|�X�e�B���O�G���A";

        public frmPosting()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.��' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            GridviewSet.Setting(dataGridView1);
            this.darwinDataSet.Clear();
            this.darwinDataSet.EnforceConstraints = false;
            this.��TableAdapter.FillByPosting(darwinDataSet.��);

            //if (dataGridView1.RowCount <= 8)
            //{
            //    dataGridView1.Columns[3].Width = 288;
            //}
            //else
            //{
            //    dataGridView1.Columns[3].Width = 271;
            //}

           

            //�|�X�e�B���O�G���A
            GridviewSet.AriaSetting(dataGridView2);

            //�����}�X�^
            GridviewSet.TownSetting(dataGridView3);

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
                    tempDGV.Height = 163;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    //tempDGV.Columns.Add("col1", "����");
                    //tempDGV.Columns.Add("col2", "����");
                    //tempDGV.Columns.Add("col3", "���l");

                    tempDGV.Columns[0].Width = 90;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 280;
                    tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static void AriaSetting(DataGridView tempDGV)
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
                    //tempDGV.Height = 230;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "ID");
                    tempDGV.Columns.Add("col2", "�G���AID");
                    tempDGV.Columns.Add("col3", "�z�z�G���A");
                    tempDGV.Columns.Add("col4", "�\�薇��");
                    tempDGV.Columns.Add("col5", "�w����");

                    tempDGV.Columns[0].Visible = false;

                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 70;

                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[3].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = true;

                    // �ҏW����
                    tempDGV.ReadOnly = true;
                    //tempDGV.Columns[0].ReadOnly = true;
                    //tempDGV.Columns[1].ReadOnly = true;
                    //tempDGV.Columns[2].ReadOnly = true;

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

            public static void TownSetting(DataGridView tempDGV)
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
                    //tempDGV.Height = 140;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�G���AID");
                    tempDGV.Columns.Add("col2", "����");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

            public static Boolean GetPostingData(int tempID, ref Entity.�z�z�G���A tempC)
            {
                string sqlStr;

                Control.�z�z�G���A dArea = new Control.�z�z�G���A();
                OleDbDataReader dr;

                sqlStr = " where �z�z�G���A.ID = " + tempID.ToString();
                dr = dArea.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.����ID = Int32.Parse(dr["����ID"].ToString());
                        tempC.�\�薇�� = Int32.Parse(dr["�\�薇��"].ToString());
                        tempC.��ID = long.Parse(dr["��ID"].ToString());
                        tempC.�z�z�w��ID = Int32.Parse(dr["�z�z�w��ID"].ToString());
                        tempC.�z�z�P�� =double.Parse(dr["�z�z�P��"].ToString(), System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.�z�z�� = dr["�z�z��"].ToString();
                        tempC.���z�z���� = Int32.Parse(dr["���z�z����"].ToString());
                        tempC.���c�� = Int32.Parse(dr["���c��"].ToString());
                        tempC.�񍐖��� = Int32.Parse(dr["�񍐖���"].ToString());
                        tempC.�񍐎c�� = Int32.Parse(dr["�񍐎c��"].ToString());
                        tempC.���z�敪 = Int32.Parse(dr["���z�敪"].ToString());
                        tempC.�}�ԋL�� = dr["�}�ԋL��"].ToString();
                        tempC.�����敪 = Int32.Parse(dr["�����敪"].ToString());
                        tempC.�X�e�[�^�X = Int32.Parse(dr["�X�e�[�^�X"].ToString());
                    }
                }
                else
                {
                    dr.Close();
                    dArea.Close();
                    return false;
                }

                dr.Close();
                dArea.Close();
                return true;
            }


            public static void AreaShowData(DataGridView tempDGV,long tempID)
            {
                int iX;

                try
                {
                    tempDGV.RowCount = 0;

                    //�z�z�G���A�f�[�^���擾����
                    OleDbDataReader dr;
                    Control.�z�z�G���A dArea = new Control.�z�z�G���A();
                    dr = dArea.FillBy("where ��ID = " + tempID.ToString() + " order by ����ID");

                    iX = 0;

                    while (dr.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = long.Parse(dr["ID"].ToString());
                        tempDGV[1, iX].Value = dr["����ID"].ToString();
                        tempDGV[3, iX].Value = Int32.Parse(dr["�\�薇��"].ToString());
                        tempDGV[4, iX].Value = Int32.Parse(dr["�z�z�w��ID"].ToString());
                        iX++;
                    }

                    tempDGV.CurrentCell = null;

                    //if (tempDGV.RowCount <= 18)
                    //{
                    //    tempDGV.Columns[2].Width = 218;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[2].Width = 201;
                    //}

                    dr.Close();
                    dArea.Close();

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

            try
            {

                if (MessageBox.Show(dataGridView1[3, dataGridView1.SelectedRows[0].Index].Value.ToString() + " ���I������܂����B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                DispClear();

                //�f�[�^���擾����
                OleDbDataReader dr;
                Control.�� cOrder = new Control.��();
                dr = cOrder.FillBy("where ID = " + dataGridView1[0,dataGridView1.SelectedRows[0].Index].Value.ToString());
                
                if (dr.HasRows == false)
                {
                    MessageBox.Show("�Y������f�[�^���o�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                //'�f�[�^�l���擾
                while (dr.Read())
                {
                    txtID.Text = dr["ID"].ToString();
                    txtCName.Text = "";
                    label8.Text = int.Parse(dr["����"].ToString()).ToString("#,##0");
                    label11.Text = double.Parse(dr["�z�z�P��"].ToString(),System.Globalization.NumberStyles.Any).ToString("#,##0.00");

                    //���Ӑ於
                    OleDbDataReader drt;
                    Control.���Ӑ� Client = new Control.���Ӑ�();
                    drt = Client.FillBy("where ID = " + dr["���Ӑ�ID"].ToString());

                    while (drt.Read())
                    {
                        txtCName.Text = drt["����"].ToString();
                    }

                    drt.Close();

                    txtChirashi.Text = dr["�`���V��"].ToString();   
                }

                dr.Close();

                cOrder.Close();

                //�z�z�G���A�f�[�^�\��
                GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
                MaisuSum();

                //txtTotal.Text = GetMaisuTotal().ToString("#,##0");
                //int Zan;
                //Zan = Int32.Parse(textBox2.Text, System.Globalization.NumberStyles.Any) - Int32.Parse(txtTotal.Text, System.Globalization.NumberStyles.Any);
                //textBox3.Text = Zan.ToString("#,##0");
      
                //�{�^���\��
                txtAdd.Enabled = true;
                txtAdel.Enabled = true;
                txtAclear.Enabled = true;

                button2.Enabled = true;
                button4.Enabled = true;

                tabPage3.Text = txtChirashi.Text + " �F �|�X�e�B���O�G���A�\";

                txtAreaID.Enabled = true;
                txtAreaName.Enabled = true;
                txtHaihuMaisu.Enabled = true;
                textBox5.Enabled = true;

                txtAreaID.Focus();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "�f�[�^�\��", MessageBoxButtons.OK);
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            if (dataGridView1.Rows.Count == 0) return;

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
                aMode.Mode = 0;

                txtID.Text = "";
                txtCName.Text = "";
                txtChirashi.Text = "";

                txtAreaID.Text = "";
                txtAreaName.Text = "";
                txtHaihuMaisu.Text = "";

                txtAreaID.Enabled = false;
                txtAreaName.Enabled = false;
                txtHaihuMaisu.Enabled = false;

                dataGridView2.Rows.Clear();
                dataGridView3.Rows.Clear();

                label11.Text = "";
                label8.Text = "";
                label7.Text = "";
                label10.Text = "";
                textBox5.Text = "";
                textBox5.Enabled = false;

                txtAdel.Enabled = false;

                button2.Enabled = false;
                button5.Enabled = false;
                button4.Enabled = false;

                txtAdd.Enabled = false;
                txtAdel.Enabled = false;
                txtAclear.Enabled = false;

                tabPage3.Text = "�|�X�e�B���O�G���A�\";

                dataGridView1.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }
        
        /// <summary>
        /// ��ʂ��N���A����
        /// </summary>
        private void DispClear2()
        {

            try
            {
                aMode.Mode = 0;

                txtAreaID.Text = "";
                txtAreaName.Text = "";
                txtHaihuMaisu.Text = "";

                textBox5.Text = "";
                txtAdel.Enabled = false;

                txtAreaID.Focus();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }

        }

        //�o�^�f�[�^�`�F�b�N
        private Boolean fDataCheck()
        {
            string str;
            int d;

            try
            {

                //�G���AID�`�F�b�N
                if (txtAreaID.Text == null)
                {
                    this.txtAreaID.Focus();
                    throw new Exception("�G���AID�͐����œ��͂��Ă�������");
                }

                str = txtAreaID.Text;

                if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                {
                    this.txtAreaID.Focus();
                    throw new Exception("�G���AID������������܂���");
                }

                if (Int32.Parse(txtAreaID.Text.ToString()) < 0)
                {
                    this.txtAreaID.Focus();
                    throw new Exception("�G���AID������������܂���");
                }

                //�z�z�����`�F�b�N
                if (txtHaihuMaisu.Text == null)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("�z�z�����͐����œ��͂��Ă�������");
                }

                str = txtHaihuMaisu.Text;

                if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("�z�z����������������܂���");
                }

                if (Int32.Parse(txtHaihuMaisu.Text.ToString(),System.Globalization.NumberStyles.AllowThousands) < 0)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("�z�z����������������܂���");
                }

                //�N���X�Ƀf�[�^�Z�b�g
                cArea.����ID = Int32.Parse(txtAreaID.Text.ToString());
                cArea.�\�薇�� = Int32.Parse(txtHaihuMaisu.Text.ToString(),System.Globalization.NumberStyles.AllowThousands);
                cArea.��ID = long.Parse(txtID.Text.ToString());

                if (aMode.Mode == 0)
                {
                    cArea.�z�z�w��ID = 0;

                    if (Utility.NumericCheck(label11.Text) == false)
                    {
                        cArea.�z�z�P�� = 0;
                    }
                    else
                    {
                        cArea.�z�z�P�� = double.Parse(label11.Text, System.Globalization.NumberStyles.Any);
                    }
                    
                    cArea.�z�z�� = "";
                    cArea.���z�z���� = 0;
                    cArea.���c�� = cArea.�\�薇��;
                    cArea.�񍐖��� = 0;
                    cArea.�񍐎c�� = cArea.�\�薇��;
                    cArea.���z�敪 = 0;
                    cArea.�����敪 = 0;
                    cArea.�X�e�[�^�X = 0;
                    cArea.�o�^�N���� = DateTime.Today;
                }
                else
                {
                    cArea.���c�� = cArea.�\�薇�� - cArea.���z�z����;
                    cArea.�񍐎c�� = cArea.�\�薇�� - cArea.�񍐖���;
                }

                cArea.�ύX�N���� = DateTime.Today;

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

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == txtAreaID)
            {
                objtxt = txtAreaID;
            }

            if (sender == txtHaihuMaisu)
            {
                objtxt = txtHaihuMaisu;
            }

            if (sender == textBox5)
            {
                objtxt = textBox5;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {

           TextBox objtxt = new TextBox();

           string str;
           double d;

           try
           {
               if (sender == textBox1)
               {
                   objtxt = textBox1;
               }

               //�z�z�G���AID
               if (sender == txtAreaID)
               {
                   objtxt = txtAreaID;

                   if (txtAreaID.Text == null) txtAreaID.Text = "0";

                   str = txtAreaID.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtAreaID.Text = "0";

                   txtAreaName.Text = GetTownName(txtAreaID.Text.ToString());
               }

               if (sender == txtHaihuMaisu)
               {
                   objtxt = txtHaihuMaisu;
               }

               if (sender == textBox5)
               {
                   objtxt = textBox5;
               }

               objtxt.BackColor = Color.White;
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message,"�G���[���b�Z�[�W");
           }
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
            darwinDataSet ds = new darwinDataSet();
            ds.Clear();
            ds.EnforceConstraints = false;
            this.��TableAdapter.FillByPostingName(ds.��, "%" + textBox1.Text.ToString() + "%");
            dataGridView1.DataSource = ds.��;

            if (dataGridView1.RowCount <= 8)
            {
                dataGridView1.Columns[3].Width = 288;
            }
            else
            {
                dataGridView1.Columns[3].Width = 271;
            }

            DispClear();
        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void txtAdd_Click(object sender, EventArgs e)
        {
            //�z�z�G���A�o�^
            try
            {
                if (fDataCheck() == true)
                {

                    Control.�z�z�G���A dArea = new Control.�z�z�G���A();

                    switch (aMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                break;

                            if (dArea.DataInsert(cArea) == false)
                            {
                                MessageBox.Show("�V�K�o�^�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                        case 1: //�X�V
                            if (MessageBox.Show("�X�V���܂��B��낵���ł����H", "�X�V�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                break;

                            if (dArea.DataUpdate(cArea) == false)
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    dArea.Close();

                    DispClear2();
                    
                    txtAreaID.Focus();

                    //�z�z�G���A�ĕ\��
                    GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
                    MaisuSum();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "�ێ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex != 1) return;

            dataGridView2[2, e.RowIndex].Value = GetTownName(dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString());
        }

        private string GetTownName(string tempID)
        {
            //�z�z�G���A��������
            string strName = ""; 
            OleDbDataReader dr;

            Control.���� cTown = new Control.����();
            dr = cTown.FillBy("where ID = " + tempID);

            while (dr.Read())
            {
                strName = dr["����"].ToString();
            }

            dr.Close();
            return strName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                //�z�z�G���A��������
                OleDbDataReader dr;
                int iX = 0;

                Control.���� cTown = new Control.����();
                dr = cTown.FillBy("where ���� like '%" + textBox5.Text.ToString() + "%' order by ID");
                dataGridView3.Rows.Clear();

                while (dr.Read())
                {
                    dataGridView3.Rows.Add();
                    dataGridView3[0, iX].Value = dr["ID"];
                    dataGridView3[1, iX].Value = dr["����"];
                    iX++;
                }

                //if (dataGridView3.RowCount <= 11)
                //{
                //    dataGridView3.Columns[1].Width = 280;
                //}
                //else
                //{
                //    dataGridView3.Columns[1].Width = 263;
                //}

                dr.Close();
                cTown.Close();

                dataGridView3.Focus();
                dataGridView3.CurrentCell = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        //�z�z�G���A�̒������ꊇ�o�^����
        {
            if (MessageBox.Show("�I������Ă��钬�����ꊇ�ǉ��o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            foreach (DataGridViewRow r in dataGridView3.SelectedRows)
            {
                txtAreaID.Text = dataGridView3[0, r.Index].Value.ToString();
                txtHaihuMaisu.Text = "0";

                if (fDataCheck() == true)
                {
                    Control.�z�z�G���A dArea = new Control.�z�z�G���A();
                    if (dArea.DataInsert(cArea) == false)
                    {
                        MessageBox.Show(dataGridView3[1, r.Index].Value.ToString() + "�̐V�K�o�^�Ɏ��s���܂���", "�����ꊇ�o�^", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    dArea.Close();
                }

            }

            DispClear2();

            txtAreaID.Focus();

            //�z�z�G���A�ĕ\��
            GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
            MaisuSum();
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            PostingGridEnter();
        }

        private void PostingGridEnter()
        {
            if (MessageBox.Show(dataGridView2[2, dataGridView2.SelectedRows[0].Index].Value.ToString() + " ���I������܂����B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            int iX = 0;

            txtAreaID.Text = dataGridView2[1, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtAreaName.Text = dataGridView2[2, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtHaihuMaisu.Text = dataGridView2[3, dataGridView2.SelectedRows[iX].Index].Value.ToString();

            aMode.Mode = 1;
            aMode.RowIndex = dataGridView2.SelectedRows[iX].Index;
            txtAdel.Enabled = true;

            //�z�z�G���A�f�[�^���擾
            GridviewSet.GetPostingData(Int32.Parse(dataGridView2[0, dataGridView2.SelectedRows[iX].Index].Value.ToString()), ref cArea);

            //�G���AID�R�[�h
            txtAreaID.Focus();
        }

        private void TownGridEnter()
        {
            if (MessageBox.Show(dataGridView3[1, dataGridView3.SelectedRows[0].Index].Value.ToString() + " ���I������܂����B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            txtAreaID.Text = dataGridView3[0, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtAreaName.Text = dataGridView3[1, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtHaihuMaisu.Focus();
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            TownGridEnter();
        }

        private int GetMaisuTotal()
        {
            int Total = 0;

            for (int i = 0; i <= dataGridView2.Rows.Count - 1; i++)
            {
                Total += Int32.Parse(dataGridView2[3, i].Value.ToString(), System.Globalization.NumberStyles.AllowThousands);
            }

            return Total;
        }

        private void txtAdel_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("�|�X�e�B���O�G���A�f�[�^���I������Ă��܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�I�����ꂽ " + dataGridView2.SelectedRows.Count.ToString() + "���̃|�X�e�B���O�G���A�f�[�^���폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                dataGridView2.CurrentCell = null;
                return;
            }

            foreach (DataGridViewRow r in dataGridView2.SelectedRows)
            {
                int aID;
                aID = int.Parse(dataGridView2[0, r.Index].Value.ToString(),System.Globalization.NumberStyles.Any);

                //���R�[�h�폜
                Control.�z�z�G���A dArea = new Control.�z�z�G���A();

                if (dArea.DataDelete(aID) == false)
                {
                    MessageBox.Show("�폜�Ɏ��s���܂����B�G���AID�F" + aID.ToString(), MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                dArea.Close();
            }

            DispClear2();

            //�z�z�G���A�ĕ\��
            GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
            MaisuSum();




            ////�폜
            //if (MessageBox.Show(txtAreaName.Text + " ���폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            //    return;

            ////�f�[�^�폜
            //Control.�z�z�G���A dArea = new Control.�z�z�G���A();
            //if (dArea.DataDelete(cArea.ID) == true)
            //{
            //    MessageBox.Show("�폜����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //else
            //{
            //    MessageBox.Show("�폜�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}

            //dArea.Close();

            //DispClear2();

            ////�z�z�G���A�ĕ\��
            //GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));
            //MaisuSum();

        }

        private void txtAclear_Click(object sender, EventArgs e)
        {
            DispClear();
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (dataGridView2.Rows.Count == 0) return;
            if (dataGridView2.SelectedRows.Count == 0) return;

            if (e.KeyCode.ToString() == "Return")
            {
                PostingGridEnter();
            }
        }

        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (dataGridView3.Rows.Count == 0) return;
            if (dataGridView3.SelectedRows.Count == 0) return;

            if (e.KeyCode.ToString() == "Return")
            {
                TownGridEnter();
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            button5.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GetExcelPos();

        }

        private void GetExcelPos()
        {

            DialogResult ret;

            //�_�C�A���O�{�b�N�X�̏����ݒ�
            openFileDialog1.Title = "�|�X�e�B���O�G���A�\�̑I��";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xls)|*.xls|�S�Ẵt�@�C��(*.*)|*.*";

            //�_�C�A���O�{�b�N�X�̕\��
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel) return;

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " ���I������܂����B��낵���ł���?", "�|�X�e�B���O�G���AExcel�V�[�g��荞��", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int S_GYO = 2;    //�G�N�Z���t�@�C�����o���s�i���ׂ�2�s�ڂ���j

            //�}�E�X�|�C���^��ҋ@�ɂ���
            this.Cursor = Cursors.WaitCursor;

            //string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            Excel.Range dRng;
            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            int iX = S_GYO;
            int Cnt = 0;
            int err;
            string cellID;
            string cellMaisu;
            int d;

            try
            {

                while (true) 
                {
                    err = 0;

                    //�G���AID
                    dRng = (Excel.Range)oxlsSheet.Cells[iX,1];

                    //�󔒂Ȃ珈���I��
                    if ((dRng.Text.ToString().Trim() + "") == "")
                        break;

                    cellID = dRng.Text.ToString().Trim();

                    //ID�`�F�b�N
                    if (cellID == null)
                    {
                        err = 1;
                        MessageBox.Show("�G���AID������������܂���B���̍s�̓X�L�b�v����܂��B�@" + Environment.NewLine + iX.ToString() + "�s�ځ@�F�@" + cellID, "��荞�݃G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (int.TryParse(cellID, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                    {
                        err = 1;
                        MessageBox.Show("�G���AID������������܂���B���̍s�̓X�L�b�v����܂��B�@" + Environment.NewLine + iX.ToString() + "�s�ځ@�F�@" + cellID, "��荞�݃G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (Int32.Parse(cellID, System.Globalization.NumberStyles.Any) < 0)
                    {
                        err = 1;
                        MessageBox.Show("�G���AID������������܂���B���̍s�̓X�L�b�v����܂��B�@" + Environment.NewLine + iX.ToString() + "�s�ځ@�F�@" + cellID, "��荞�݃G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    //�z�z����
                    dRng = (Excel.Range)oxlsSheet.Cells[iX,3];
                    cellMaisu = dRng.Text.ToString().Trim();

                    //�`�F�b�N
                    if (cellMaisu == null)
                    {
                        err = 1;
                        MessageBox.Show("�z�z����������������܂���B���̍s�̓X�L�b�v����܂��B�@" + Environment.NewLine + iX.ToString() + "�s�ځ@�F�@" + cellMaisu, "��荞�݃G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (int.TryParse(cellMaisu, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                    {
                        err = 1;
                        MessageBox.Show("�z�z����������������܂���B���̍s�̓X�L�b�v����܂��B�@" + Environment.NewLine + iX.ToString() + "�s�ځ@�F�@" + cellMaisu, "��荞�݃G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (Int32.Parse(cellMaisu,System.Globalization.NumberStyles.Any) < 0)
                    {
                        err = 1;
                        MessageBox.Show("�z�z����������������܂���B���̍s�̓X�L�b�v����܂��B�@" + Environment.NewLine + iX.ToString() + "�s�ځ@�F�@" + cellMaisu, "��荞�݃G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    //�G���[�̂Ƃ��͓ǂݔ�΂�
                    if (err == 0)
                    {

                        txtAreaID.Text = cellID.ToString();
                        txtHaihuMaisu.Text = cellMaisu.ToString();

                        if (fDataCheck() == true)
                        {
                            Control.�z�z�G���A dArea = new Control.�z�z�G���A();
                            if (dArea.DataInsert(cArea) == false)
                            {
                                MessageBox.Show("�V�K�o�^�Ɏ��s���܂���", "Excel�V�[�g��荞��", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                Cnt++;
                            }

                            dArea.Close();
                        }

                    }

                    iX++;
                }

                MessageBox.Show(Cnt.ToString() + " ���̔z�z�G���A����荞�݂܂���", "��荞�ݏI��", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //�}�E�X�|�C���^�����ɖ߂�
                this.Cursor = Cursors.Default;

                // �m�F�̂���Excel�̃E�B���h�E��\������
                //oXls.Visible = true;

                //���
                //oxlsSheet.PrintPreview(true);

                //�ۑ�����
                oXls.DisplayAlerts = false;

                //Book���N���[�Y
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                //Excel���I��
                oXls.Quit();
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message, "���", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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


            DispClear2();

            txtAreaID.Focus();

            //�z�z�G���A�ĕ\��
            GridviewSet.AreaShowData(dataGridView2, long.Parse(txtID.Text.ToString()));

            MaisuSum();
        }

        private void MaisuSum()
        {
            label10.Text = GetMaisuTotal().ToString("#,##0");
            int Zan;
            Zan = Int32.Parse(label8.Text, System.Globalization.NumberStyles.Any) - Int32.Parse(label10.Text, System.Globalization.NumberStyles.Any);
            label7.Text = Zan.ToString("#,##0");

            if (Zan < 0)
            {
                label7.ForeColor = Color.Red;
                MessageBox.Show("�|�X�e�B���O�G���A�������z�z�����𒴂��Ă��܂�", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                label7.ForeColor = Color.Black;
            }

        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            txtAdel.Enabled = true;
        }

    }
}
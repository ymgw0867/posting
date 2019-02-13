using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MyLibrary;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmTown : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.���� cMaster = new Entity.����();

        const string MESSAGE_CAPTION = "�����}�X�^�[";

        public frmTown()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {

            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView1);

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.����' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            this.����TableAdapter.Fill(this.darwinDataSet.����);

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
                    tempDGV.Height = 181;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    //tempDGV.Columns.Add("col1", "����");
                    //tempDGV.Columns.Add("col2", "����");
                    //tempDGV.Columns.Add("col3", "���l");

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 100;
                    //tempDGV.Columns[3].Width = 160;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;

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

            /// <summary>
            /// �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾����
            /// </summary>
            /// <param name="dgv">�ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            public static Boolean GetData(DataGridView dgv,ref Entity.���� tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.���� Town = new Control.����();
                OleDbDataReader dr;

                sqlStr = " where ����.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Town.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.���� = dr["����"].ToString() + "";
                        tempC.�s�撬���R�[�h = int.Parse(dr["�s�撬���R�[�h"].ToString(),System.Globalization.NumberStyles.Any);
                        tempC.���l = dr["���l"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    Town.Close();
                    return false;
                }

                dr.Close();
                Town.Close();
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

            if (MessageBox.Show(msgStr, "�����I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                    txtCode.Text = cMaster.ID.ToString();
                    txtName1.Text = cMaster.����.ToString();
                    txtCityCode.Text = cMaster.�s�撬���R�[�h.ToString();
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
                txtCityCode.Text = "0";
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
                    Control.���� Town = new Control.����();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Town.Close();
                                return;
                            }

                            if (Town.DataInsert(cMaster) == true)
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
                                Town.Close();
                                return;
                            }

                            if (Town.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", "�����}�X�^�[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Town.Close();

                    DispClear();

                    //�f�[�^�� 'darwinDataSet.����' �e�[�u���ɓǂݍ��݂܂��B
                    this.����TableAdapter.Fill(this.darwinDataSet.����);
                    dataGridView1.DataSource = this.darwinDataSet.����;

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
                    Control.���� Town = new Control.����();
                    OleDbDataReader dr;

                    sqlStr = " where ID = " + txtCode.Text.ToString();
                    dr = Town.FillBy(sqlStr);

                    if (dr.HasRows == true)
                    {
                        txtCode.Focus();
                        dr.Close();
                        Town.Close();
                        throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    }

                    dr.Close();
                    Town.Close();

                }

                //���̃`�F�b�N
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("���̂���͂��Ă�������");
                }

                //�s�撬���R�[�h
                if (Utility.NumericCheck(txtCityCode.Text) == false)
                {
                    txtCityCode.Focus();
                    throw new Exception("�s�撬���R�[�h�͐����œ��͂��Ă�������");
                }

                //�}�X�^�[�`�F�b�N
                if (txtCityCode.Text != "0")
                {
                    string sqlSTR;
                    OleDbDataReader dR;
                    Control.FreeSql fCon = new Control.FreeSql();

                    sqlSTR = "";
                    sqlSTR += "select * from �s�撬�� where ID = " + txtCityCode.Text;
                    dR = fCon.free_dsReader(sqlSTR);

                    if (dR.HasRows == false)
                    {
                        txtCityCode.Focus();
                        dR.Close();
                        fCon.Close();
                        throw new Exception("�Y������s�撬���R�[�h������܂���");
                    }

                    dR.Close();
                    fCon.Close();
                }


                //�N���X�Ƀf�[�^�Z�b�g
                cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());
                cMaster.���� = txtName1.Text.ToString();
                cMaster.�s�撬���R�[�h = int.Parse(txtCityCode.Text.ToString(), System.Globalization.NumberStyles.Any);
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

            if (sender == txtCode)
            {
                objtxt = txtCode;
            }

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtCityCode)
            {
                objtxt = txtCityCode;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
        {

            TextBox objtxt = new TextBox();

            if (sender == txtCode)
            {
                objtxt = txtCode;
            }

            if (sender == txtName1)
            {
                objtxt = txtName1;
            }

            if (sender == txtCityCode)
            {
                objtxt = txtCityCode;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            objtxt.BackColor = Color.White;

        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //���ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���
            string SqlStr;
            SqlStr = " where ";
            SqlStr += "(�z�z�G���A.����ID = " + txtCode.Text.ToString() + ")  ";

            OleDbDataReader dr;
            Control.�z�z�G���A Area = new Control.�z�z�G���A();
            dr = Area.FillBy(SqlStr);

            //�Y���������o�^����Ă���Ƃ��͍폜�s�Ƃ���
            if (dr.HasRows == true)
            {
                MessageBox.Show(txtName1.Text.ToString() + "���z�z�G���A�f�[�^�ɓo�^����Ă��܂�", txtName1.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dr.Close();
                Area.Close();
                return;
            }

            dr.Close();
            Area.Close();
            
            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�폜
            Control.���� Town = new Control.����();
            if (Town.DataDelete(Convert.ToInt32(txtCode.Text.ToString()))==true)
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION,  MessageBoxButtons.OK, MessageBoxIcon.Information);
            Town.Close();

            DispClear();

            //�f�[�^�� 'darwinDataSet.����' �e�[�u���ɓǂݍ��݂܂��B
            this.����TableAdapter.Fill(this.darwinDataSet.����);
            dataGridView1.DataSource = this.darwinDataSet.����;

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
            this.����TableAdapter.FillByName(ds.����,"%" + textBox1.Text.ToString() + "%");
            dataGridView1.DataSource = ds.����;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GetExcelPos();
        }

        private void GetExcelPos()
        {

            DialogResult ret;

            //�_�C�A���O�{�b�N�X�̏����ݒ�
            openFileDialog1.Title = "�s�撬���R�[�h�\�̑I��";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xls)|*.xls|�S�Ẵt�@�C��(*.*)|*.*";

            //�_�C�A���O�{�b�N�X�̕\��
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel) return;

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " ���I������܂����B��낵���ł���?", "�s�撬���R�[�h�\��荞��", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            const int S_GYO = 1;    //�G�N�Z���t�@�C�����o���s�i���ׂ�1�s�ڂ���j

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
            int err;
            string cellID, cellItem1, cellItem2;
            string sqlSTR;

            try
            {

                while (true)
                {
                    err = 0;

                    //�s�撬���R�[�h
                    dRng = (Excel.Range)oxlsSheet.Cells[iX, 2];

                    //�󔒂Ȃ珈���I��
                    if ((dRng.Text.ToString().Trim() + "") == "")
                        break;

                    cellID = dRng.Text.ToString();

                    //�s���{��
                    dRng = (Excel.Range)oxlsSheet.Cells[iX, 3];
                    cellItem1 = dRng.Text.ToString();

                    //�s�撬����
                    dRng = (Excel.Range)oxlsSheet.Cells[iX, 4];
                    cellItem2 = dRng.Text.ToString();

                    //�R�[�h�`�F�b�N
                    if (Utility.NumericCheck(cellID) == false)
                    {
                        err = 1;
                    }

                    //�G���[�̂Ƃ��͓ǂݔ�΂�
                    if (err == 0)
                    {
                        Control.FreeSql fCon = new Control.FreeSql();

                        sqlSTR = "";
                        sqlSTR += "insert into �s�撬�� ";
                        sqlSTR += "(ID,�s���{��,�s�撬��,�敪1,�敪2,���l,�o�^�N����,�ύX�N����) ";
                        sqlSTR += "values (" + cellID + ",";
                        sqlSTR += "'" + cellItem1 + "',";
                        sqlSTR += "'" + cellItem2 + "',";
                        sqlSTR += "0,0,'',";
                        sqlSTR += "'" + DateTime.Today.ToShortDateString() + "',";
                        sqlSTR += "'" + DateTime.Today.ToShortDateString() + "')";

                        if (fCon.Execute(sqlSTR) == false)
                        {
                            MessageBox.Show("�s�撬���}�X�^�[�̓o�^�Ɏ��s���܂����B" + cellID, "�s�撬���R�[�h�\�ǂݍ���", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            fCon.Close();
                            break;
                        }

                        fCon.Close();

                    }

                    iX++;
                }

                MessageBox.Show("�I�����܂���", "�s�撬���R�[�h�\�ǂݍ���", MessageBoxButtons.OK, MessageBoxIcon.Information);

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

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string cName;
            string sqlSTR;
            OleDbDataReader dR;
            Control.FreeSql fCon = new Control.FreeSql();

            dR = fCon.free_dsReader("select * from �s�撬�� order by ID");

            while(dR.Read())
            {
                cName = dR["�s�撬��"].ToString();
                Control.FreeSql fCon2 = new Control.FreeSql();

                sqlSTR = "";
                sqlSTR += "update ���� ";
                sqlSTR += "set ����.�s�撬���R�[�h = " + dR["ID"].ToString();
                sqlSTR += "where (����.���� like '" + cName + "%') and ";
                sqlSTR += "(����.�s�撬���R�[�h = 0)";

                fCon2.Execute(sqlSTR);

                fCon2.Close();

            }

            dR.Close();

            fCon.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmTownSub fSub = new frmTownSub();

            if (fSub.ShowDialog(this) == DialogResult.OK)
            {
                txtCityCode.Text = fSub.�s�撬���R�[�h.ToString(); ;
            }
            else
            {
            }

            fSub.Dispose();
        }

    }
}
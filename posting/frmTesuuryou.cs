using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Linq;

namespace posting
{
    public partial class frmTesuuryou : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.��Џ�� sMaster = new Entity.��Џ��();
        Entity.�S��g���[���[���R�[�h sZtr = new Entity.�S��g���[���[���R�[�h();

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.�z�z��TableAdapter adp = new darwinDataSetTableAdapters.�z�z��TableAdapter();

        const string MESSAGE_CAPTION = "�z�z�萔������";

        int[] fIDArray;

        public frmTesuuryou()
        {
            InitializeComponent();

            // �z�z���f�[�^�ǂݍ���
            adp.Fill(dts.�z�z��);
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //�O���b�h��`
            GridviewSet.Setting(dataGridView2);
            button1.Enabled = false;

            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            DispClear();

            // �z�z���w��R���{�{�b�N�X�l�Z�b�g 2016/02/01
            addCmbHaihuItems();
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     �z�z���w��R���{�{�b�N�X�l�Z�b�g </summary>
        ///---------------------------------------------------------------
        private void addCmbHaihuItems()
        {
            comboBox2.Items.Add("�S��");
            comboBox2.Items.Add("�z�z���w��");

            // ���Ə��}�X�^�[�̎��Ə�����ǉ�
            darwinDataSetTableAdapters.���Ə�TableAdapter jAdp = new darwinDataSetTableAdapters.���Ə�TableAdapter();
            jAdp.Fill(dts.���Ə�);

            int iX =0;

            foreach (var t in dts.���Ə�.OrderBy(a => a.ID))
            {
                comboBox2.Items.Add(t.����);

                iX++;
                Array.Resize(ref fIDArray, iX);
                fIDArray[iX - 1] = t.ID;
            }

            comboBox2.SelectedIndex = 0;
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
                    tempDGV.Height = 505;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "���Ə�");
                    tempDGV.Columns.Add("col2", "�z�z��ID");
                    tempDGV.Columns.Add("col3", "�z�z����");
                    tempDGV.Columns.Add("col4", "�w���ԍ�");
                    tempDGV.Columns.Add("col5", "���t");
                    tempDGV.Columns.Add("col6", "�`���V��");
                    tempDGV.Columns.Add("col7", "�z�z�G���A");
                    tempDGV.Columns.Add("col8", "�P��");
                    tempDGV.Columns.Add("col9", "�z�z��");
                    tempDGV.Columns.Add("col10", "���v");

                    tempDGV.Columns[0].Width = 70;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 100;
                    tempDGV.Columns[3].Width = 80;
                    tempDGV.Columns[4].Width = 90;
                    //tempDGV.Columns[5].Width = 200;
                    //tempDGV.Columns[6].Width = 200;
                    tempDGV.Columns[7].Width = 70;
                    tempDGV.Columns[8].Width = 80;
                    tempDGV.Columns[9].Width = 80;

                    tempDGV.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    tempDGV.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                    //tempDGV.Columns[0].Visible = false;
                    //tempDGV.Columns[9].Visible = false;
                    //tempDGV.Columns[10].Visible = false;
                    //tempDGV.Columns[11].Visible = false;
                    //tempDGV.Columns[12].Visible = false;

                    tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0.0";
                    tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";
                    tempDGV.Columns[9].DefaultCellStyle.Format = "#,##0";

                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

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

            public static void ShowData(DataGridView tempDGV, int tempYear, int tempMonth, string sDate, string eDate, string tempSKbn,int tempHIndex,int tempHID)
            {

                string sqlSTRING = "";
                int iX;

                try
                {
                    //�J�[�\����ҋ@�\��
                    Cursor.Current = Cursors.WaitCursor;

                    tempDGV.RowCount = 0;
                    
                    //�f�[�^���[�_�[���擾����
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();
                    cn = Con.GetConnection();

                    OleDbDataReader dR;

                    sqlSTRING = "";
                    sqlSTRING += "select * from ";
                    sqlSTRING += "(";

                    //�z�z���і���
                    sqlSTRING += "select �z�z�w��.ID as �z�z�w��ID,�z�z�w��.�z�z��,���Ə�.ID as ���Ə�ID,";
                    sqlSTRING += "���Ə�.���� as ���Ə�����, �z�z��.ID as �z�z��ID,";
                    sqlSTRING += "�z�z��.����, ���Ӑ�.ID as ���Ӑ�ID, ���Ӑ�.���� as ���Ӑ旪��,";
                    sqlSTRING += "��.�`���V��, ����.����, �z�z�G���A.�z�z�P�� as �P��,";
                    sqlSTRING += "�z�z�G���A.���z�z����,";
                    sqlSTRING += "�z�z�G���A.�z�z�P�� * �z�z�G���A.���z�z���� as ���z,��.ID,'0' as sortID  ";
                    sqlSTRING += "from �z�z�w�� inner join �z�z�G���A ";
                    sqlSTRING += "on �z�z�w��.ID = �z�z�G���A.�z�z�w��ID inner join �� ";
                    sqlSTRING += "on �z�z�G���A.��ID = ��.ID left join �z�z�� ";
                    sqlSTRING += "on �z�z�w��.�z�z��ID = �z�z��.ID left join ���� ";
                    sqlSTRING += "on �z�z�G���A.����ID = ����.ID left join ���Ə� ";
                    sqlSTRING += "on �z�z��.���Ə��R�[�h = ���Ə�.ID left join ���Ӑ� ";
                    sqlSTRING += "on ��.���Ӑ�ID = ���Ӑ�.ID ";

                    //�T���͌�
                    switch (tempSKbn)
                    {
                        case "��":
                            sqlSTRING += "where (year(�z�z�w��.���͓�) = ?) and (month(�z�z�w��.���͓�) = ?) ";
                            break;

                        case "�T":
                            sqlSTRING += "where ((�z�z�w��.���͓�) >= ?) and ((�z�z�w��.���͓�) <= ?) ";
                            break;
                    }

                    sqlSTRING += "and (�z�z��.�x���敪 = ?) ";

                    //�z�z���w��
                    if (tempHIndex == 1)
                    {
                            sqlSTRING += "and (�z�z��.ID = ?) ";
                    }
                    else if (tempHIndex >= 2)
                    {
                        // 2016/02/01
                        sqlSTRING += "and (���Ə�.ID = ?) ";
                    }


                    //��ʔ�
                    sqlSTRING += "union ";

                    sqlSTRING += "select �z�z�w��.ID AS �z�z�w��ID,�z�z�w��.�z�z��,���Ə�.ID AS ���Ə�ID,";
                    sqlSTRING += "���Ə�.���� AS ���Ə�����,�z�z��.ID AS �z�z��ID,�z�z��.����,";
                    sqlSTRING += "'0' AS ���Ӑ�ID,'' AS ���Ӑ旪��,'��ʔ�' AS �`���V��,'' AS ����,";
                    sqlSTRING += "�z�z�w��.��ʔ� AS �P��,'1' AS ���z�z����,�z�z�w��.��ʔ� AS ���z,";
                    sqlSTRING += "'999999999999' AS ID, '0' AS sortID ";

                    sqlSTRING += "FROM �z�z�w�� LEFT OUTER JOIN ";
                    sqlSTRING += "�z�z�� ON �z�z�w��.�z�z��ID = �z�z��.ID LEFT OUTER JOIN ";
                    sqlSTRING += "���Ə� ON �z�z��.���Ə��R�[�h = ���Ə�.ID ";
                    sqlSTRING += "WHERE (�z�z�w��.��ʔ� > 0) ";

                    //�T���͌�
                    switch (tempSKbn)
                    {
                        case "��":
                            sqlSTRING += "and  (year(�z�z�w��.���͓�) = ?) and (month(�z�z�w��.���͓�) = ?) ";
                            break;

                        case "�T":
                            sqlSTRING += "and  ((�z�z�w��.���͓�) >= ?) and ((�z�z�w��.���͓�) <= ?) ";
                            break;
                    }

                    sqlSTRING += "and (�z�z��.�x���敪 = ?) ";

                    //�z�z���w��
                    if (tempHIndex == 1)
                    {
                        sqlSTRING += "and (�z�z��.ID = ?) ";
                    }
                    else if (tempHIndex >= 2)
                    {
                        // 2016/02/01
                        sqlSTRING += "and (���Ə�.ID = ?) ";
                    }

                    //�x���E�T������
                    sqlSTRING += "union ";

                    sqlSTRING += "select '0' as �z�z�w��ID,�x���T��.���t as �z�z��,���Ə�.ID as ���Ə�ID,";
                    sqlSTRING += "���Ə�.���� as ���Ə�����,�z�z��.ID as �z�z��ID,";
                    sqlSTRING += "�z�z��.����,0 as ���Ӑ�ID,'' as ���Ӑ旪��,";
                    sqlSTRING += "�x���T��.�E�v as �`���V��,'' as ����,";

                    //�P�� : �T�����ڂ̂Ƃ���(-1)
                    sqlSTRING += "case �x���T��.�x���T���敪 when 0 ";
                    sqlSTRING += "then �x���T��.�P�� ";
                    sqlSTRING += "else �x���T��.�P�� * (-1) ";
                    sqlSTRING += "end,";

                    sqlSTRING += "�x���T��.���� as ���z�z����,";

                    //���z : �T�����ڂ̂Ƃ���(-1)
                    sqlSTRING += "case �x���T��.�x���T���敪 when 0 ";
                    sqlSTRING += "then �x���T��.���z ";
                    sqlSTRING += "else �x���T��.���z * (-1) ";
                    sqlSTRING += "end,";
                    
                    sqlSTRING += "999999999999 as ID,'2' as sortID ";

                    sqlSTRING += "from �x���T�� inner join �z�z�� ";
                    sqlSTRING += "on �x���T��.�z�z��ID = �z�z��.ID left join ���Ə� ";
                    sqlSTRING += "on �z�z��.���Ə��R�[�h = ���Ə�.ID ";

                    //�T���͌�
                    switch (tempSKbn)
                    {
                        case "��":
                            sqlSTRING += "where (year(�x���T��.���t) = ?) and (month(�x���T��.���t) = ?) ";
                            break;

                        case "�T":
                            sqlSTRING += "where ((�x���T��.���t) >= ?) and ((�x���T��.���t) <= ?) ";
                            break;
                    }

                    sqlSTRING += "and (�z�z��.�x���敪 = ?) ";

                    //�z�z���w��
                    if (tempHIndex == 1)
                    {
                        sqlSTRING += "and (�z�z��.ID = ?) ";
                    }
                    else if (tempHIndex >= 2)
                    {
                        // 2016/02/01
                        sqlSTRING += "and (���Ə�.ID = ?) ";
                    }

                    sqlSTRING += ") as U ";
                    sqlSTRING += "order by ���Ə�ID,U.�z�z��ID,sortID,U.�z�z�w��ID,U.ID";
                                        
                    //�z�z�w���f�[�^�̃f�[�^���[�_�[���擾����
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.CommandText = sqlSTRING;

                    switch (tempSKbn)
                    {
                        case "��":
                            SCom.Parameters.AddWithValue("@P1", tempYear);
                            SCom.Parameters.AddWithValue("@P2", tempMonth);
                            SCom.Parameters.AddWithValue("@P3", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 �z�z���܂��͎��Ə��w��
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH1", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P4", tempYear);
                            SCom.Parameters.AddWithValue("@P5", tempMonth);
                            SCom.Parameters.AddWithValue("@P6", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH2", tempHID);
                            //}

                            // 2016/02/01 �z�z���܂��͎��Ə��w��
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH2", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P7", tempYear);
                            SCom.Parameters.AddWithValue("@P8", tempMonth);
                            SCom.Parameters.AddWithValue("@P9", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH3", tempHID);
                            //}

                            // 2016/02/01 �z�z���܂��͎��Ə��w��
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH3", tempHID);
                            }

                            break;

                        case "�T":
                            SCom.Parameters.AddWithValue("@P1", sDate);
                            SCom.Parameters.AddWithValue("@P2", eDate);
                            SCom.Parameters.AddWithValue("@P3", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 �z�z���܂��͎��Ə��w��
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH1", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P4", sDate);
                            SCom.Parameters.AddWithValue("@P5", eDate);
                            SCom.Parameters.AddWithValue("@P6", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 �z�z���܂��͎��Ə��w��
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH2", tempHID);
                            }

                            SCom.Parameters.AddWithValue("@P7", sDate);
                            SCom.Parameters.AddWithValue("@P8", eDate);
                            SCom.Parameters.AddWithValue("@P9", tempSKbn);

                            //if (tempHIndex == 1)
                            //{
                            //    SCom.Parameters.AddWithValue("@PH1", tempHID);
                            //}

                            // 2016/02/01 �z�z���܂��͎��Ə��w��
                            if (tempHIndex >= 1)
                            {
                                SCom.Parameters.AddWithValue("@PH3", tempHID);
                            }

                            break;
                    }

                    SCom.Connection = cn;

                    dR = SCom.ExecuteReader();
                    
                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        try
                        {

                            tempDGV.Rows.Add();

                            tempDGV[0, iX].Value = dR["���Ə�����"].ToString();
                            tempDGV[1, iX].Value = int.Parse(dR["�z�z��ID"].ToString());
                            tempDGV[2, iX].Value = dR["����"].ToString() + "";
                            tempDGV[3, iX].Value = dR["�z�z�w��ID"].ToString();
                            tempDGV[4, iX].Value = DateTime.Parse(dR["�z�z��"].ToString()).ToShortDateString();
                            tempDGV[5, iX].Value = dR["�`���V��"].ToString();
                            tempDGV[6, iX].Value = dR["����"].ToString() + "";
                            tempDGV[7, iX].Value = double.Parse(dR["�P��"].ToString()).ToString("#,##0.00");
                            tempDGV[8, iX].Value = int.Parse(dR["���z�z����"].ToString());
                            tempDGV[9, iX].Value = System.Math.Floor(double.Parse(dR["���z"].ToString()) + 0.5);

                            iX++;
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }

                    dR.Close();

                    cn.Close();

                    Con.Close();

                    //�J�[�\���\����߂�
                    Cursor.Current = Cursors.Default;
                }
                catch (Exception e)
                {
                    //�J�[�\���\����߂�
                    Cursor.Current = Cursors.Default;

                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);

                }

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�I�����܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int tempHID = 0;

            if (DataCheck() == false) return;

            if (MessageBox.Show("�z�z�萔�����ׂ�\�����܂��B��낵���ł����H", MESSAGE_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                return;

            //���ו\��
            //switch (comboBox2.SelectedIndex)
            //{
            //    case 0:
            //        tempHID = 0;
            //        break;

            //    case 1:
            //        tempHID = int.Parse(txtHID.Text);
            //        break;
            //}

            // 2016/02/01
            if (comboBox2.SelectedIndex == 0)
            {
                // �S��
                tempHID = 0;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                // �z�z���w��
                tempHID = int.Parse(txtHID.Text);
            }
            else
            {
                // ���Ə��w��
                tempHID = fIDArray[comboBox2.SelectedIndex - 2];
            }
            
            GridviewSet.ShowData(dataGridView2, int.Parse(txtYear.Text), int.Parse(txtMonth.Text),dateTimePicker1.Value.ToShortDateString(),dateTimePicker2.Value.ToShortDateString(), comboBox1.Text,comboBox2.SelectedIndex,tempHID);
            
            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("�Y���҂͂���܂���ł���", MESSAGE_CAPTION);

                //����{�^��
                button1.Enabled = false;

                //�U���f�[�^�쐬�֘A�I�u�W�F�N�g
                label12.Enabled = false;
                label13.Enabled = false;
                label14.Enabled = false;
                label15.Enabled = false;
                txtfMonth.Enabled = false;
                txtfDay.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
            }
            else
            {
                //����{�^��
                button1.Enabled = true;

                //�U���f�[�^�쐬�֘A�I�u�W�F�N�g
                label12.Enabled = true;
                label13.Enabled = true;
                label14.Enabled = true;
                label15.Enabled = true;
                txtfMonth.Enabled = true;
                txtfDay.Enabled = true;
                button2.Enabled = true;

                //�z�z���ʎ萔���f�[�^�o�̓{�^��
                button3.Enabled = true;
            }

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ShowPosting(dataGridView1, Convert.ToDateTime(dateTimePicker1.Value.ToShortDateString()), long.Parse(dataGridView2[0, dataGridView2.SelectedRows[0].Index].Value.ToString()));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("�z�z����������̏؂𔭍s���܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�O���b�h�ĕ\��
            int tempHID = 0;

            if (DataCheck() == false) return;

            // 2016/02/01
            if (comboBox2.SelectedIndex == 0)
            {
                // �S��
                tempHID = 0;
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                // �z�z���w��
                tempHID = int.Parse(txtHID.Text);
            }
            else
            {
                // ���Ə��w��
                tempHID = fIDArray[comboBox2.SelectedIndex - 2];
            }

            GridviewSet.ShowData(dataGridView2, int.Parse(txtYear.Text), int.Parse(txtMonth.Text), dateTimePicker1.Value.ToShortDateString(), dateTimePicker2.Value.ToShortDateString(), comboBox1.Text, comboBox2.SelectedIndex, tempHID);

            dataGridView2.CurrentCell = null;

            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("�Y���҂͂���܂���ł���", MESSAGE_CAPTION);
                return;
            }

            // ���s 2016/09/19
            newReport();
        }

        private void newReport()
        {
            // �v�����g�_�C�A���O
            PrintDialog pd = new PrintDialog();
            if (pd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            //�������
            int staffID = 0;
            int dCnt = 0;
            double Gur = 0;
            string staffName = "";
            string sKikan = "";

            string[] sCell01 = new string[1];
            string[] sCell02 = new string[1];
            string[] sCell03 = new string[1];
            string[] sCell04 = new string[1];
            string[] sCell05 = new string[1];
            string[] sCell06 = new string[1];
            string[] sCell07 = new string[1];

            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���z�z����������̏�, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[2];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    for (int iX = 0; iX <= dataGridView2.RowCount - 1; iX++)
                    {
                        // ���
                        if ((staffID != 0) && (staffID != int.Parse(dataGridView2[1, iX].Value.ToString())))
                        {
                            // �G�N�Z���V�[�g�Ƀf�[�^���Z�b�g
                            xlsDataSet(ref oXlsBook, ref oxlsSheet, Gur, staffID, staffName, sKikan, sCell01, sCell02, sCell03, sCell04, sCell05, sCell06, sCell07);                            
                        }

                        if (staffID != int.Parse(dataGridView2[1, iX].Value.ToString()))
                        {
                            // �Ώۊ���
                            switch (comboBox1.Text)
                            {
                                case "��":
                                    sKikan = "�Ώۊ��� : " + txtYear.Text + label1.Text + txtMonth.Text + label2.Text;
                                    break;

                                case "�T":
                                    sKikan = "�Ώۊ��� : " + DateTime.Parse(dateTimePicker1.Value.ToString()).AddDays(-1).ToShortDateString() + label4.Text + DateTime.Parse(dateTimePicker2.Value.ToString()).AddDays(-1).ToShortDateString();
                                    break;
                            }

                            // �z�z����
                            staffName = dataGridView2[2, iX].Value.ToString();

                            // �萔�����v
                            Gur = 0;

                            // �z�z���ʂ̖��א�
                            dCnt = 0;
                        }

                        //����
                        Array.Resize(ref sCell01, dCnt + 1);
                        Array.Resize(ref sCell02, dCnt + 1);
                        Array.Resize(ref sCell03, dCnt + 1);
                        Array.Resize(ref sCell04, dCnt + 1);
                        Array.Resize(ref sCell05, dCnt + 1);
                        Array.Resize(ref sCell06, dCnt + 1);
                        Array.Resize(ref sCell07, dCnt + 1);

                        sCell01[dCnt] = dataGridView2[3, iX].Value.ToString();
                        sCell02[dCnt] = dataGridView2[4, iX].Value.ToString();
                        sCell03[dCnt] = dataGridView2[5, iX].Value.ToString();
                        sCell04[dCnt] = dataGridView2[6, iX].Value.ToString();
                        sCell05[dCnt] = dataGridView2[7, iX].Value.ToString();
                        sCell06[dCnt] = dataGridView2[8, iX].Value.ToString();
                        sCell07[dCnt] = dataGridView2[9, iX].Value.ToString();

                        // �萔���v�Z
                        Gur += double.Parse(dataGridView2[9, iX].Value.ToString());

                        // �z�z��ID
                        staffID = int.Parse(dataGridView2[1, iX].Value.ToString());

                        dCnt++;
                    }

                    // �G�N�Z���V�[�g�Ƀf�[�^���Z�b�g
                    xlsDataSet(ref oXlsBook, ref oxlsSheet, Gur, staffID, staffName, sKikan, sCell01, sCell02, sCell03, sCell04, sCell05, sCell06, sCell07);

                    // �ۑ�����
                    oXls.DisplayAlerts = false;

                    // �e���v���[�g�p�̃V�[�g���폜����
                    ((Excel.Worksheet)oXlsBook.Sheets["��̏�"]).Delete();    // �z�z����������̏��V�[�g
                    ((Excel.Worksheet)oXlsBook.Sheets["������"]).Delete();    // �z�z���������V�[�g

                    // Excel�̃E�B���h�E��\������
                    //oXls.Visible = true;

                    // BOOK�ň��
                    oXlsBook.PrintOut(1, Type.Missing, 1, false, pd.PrinterSettings.PrinterName, Type.Missing, Type.Missing, Type.Missing);

                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    // Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excel���I��
                    oXls.Quit();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    // Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excel���I��
                    oXls.Quit();
                }
                finally
                {
                    // COM �I�u�W�F�N�g�̎Q�ƃJ�E���g��������� 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //�}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;

                    MessageBox.Show("�������I�����܂���", "�������",MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "�������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        ///---------------------------------------------------------------------------------------
        /// <summary>
        ///     �z�z����������̏��V�[�g�A�z�z���������V�[�g�Ƀf�[�^���Z�b�g���� </summary>
        /// <param name="oXlsBook">
        ///     ExcelBook�I�u�W�F�N�g </param>
        /// <param name="oxlsSheet">
        ///     ExcelSheet�I�u�W�F�N�g </param>
        /// <param name="Gur">
        ///     �萔�����v </param>
        /// <param name="staffID">
        ///     �X�^�b�t�R�[�h </param>
        /// <param name="staffName">
        ///     �X�^�b�t�� </param>
        /// <param name="sKikan">
        ///     �Ώۊ��� </param>
        /// <param name="sCell01">
        ///     ���׍��ڂP </param>
        /// <param name="sCell02">
        ///     ���׍��ڂQ </param>
        /// <param name="sCell03">
        ///     ���׍��ڂR </param>
        /// <param name="sCell04">
        ///     ���׍��ڂS </param>
        /// <param name="sCell05">
        ///     ���׍��ڂT </param>
        /// <param name="sCell06">
        ///     ���׍��ڂU </param>
        /// <param name="sCell07">
        ///     ���׍��ڂV </param>
        ///---------------------------------------------------------------------------------------
        private void xlsDataSet(ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, double Gur, int staffID, string staffName, string sKikan,
                                  string[] sCell01, string[] sCell02, string[] sCell03, string[] sCell04,
                                  string[] sCell05, string[] sCell06, string[] sCell07)
        {
            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            Excel.Worksheet jyuryoSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet seikyuSheet = (Excel.Worksheet)oXlsBook.Sheets[2];

            const int S_GYO = 14;    //�G�N�Z���t�@�C�����׊J�n�s
            const int S_ROWSMAX = 7; //�G�N�Z���t�@�C����ő�l

            // �z�z����������̏��V�[�g��ǉ�����
            jyuryoSheet.Copy(Type.Missing, oxlsSheet);
            oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[oXlsBook.Sheets.Count];

            // �Ώۊ���
            oxlsSheet.Cells[12, 3] = sKikan;

            // �z�z��ID
            oxlsSheet.Cells[4, 1] = staffID.ToString();

            // �z�z����
            oxlsSheet.Cells[4, 2] = staffName + "�@�a";

            // ���z
            oxlsSheet.Cells[10, 6] = Gur;

            for (int iz = 0; iz < sCell01.Length; iz++)
            {
                // ����
                oxlsSheet.Cells[iz + S_GYO, 1] = sCell01[iz];
                oxlsSheet.Cells[iz + S_GYO, 2] = sCell02[iz];
                oxlsSheet.Cells[iz + S_GYO, 3] = sCell03[iz];
                oxlsSheet.Cells[iz + S_GYO, 4] = sCell04[iz];
                oxlsSheet.Cells[iz + S_GYO, 5] = sCell05[iz];
                oxlsSheet.Cells[iz + S_GYO, 6] = sCell06[iz];
                oxlsSheet.Cells[iz + S_GYO, 7] = sCell07[iz];

                // �Z�������֎������R�r��������
                rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }

            // �\�S�̂Ɏ����c�r��������
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // �\�S�̂̍��[�c�r��
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // �\�S�̂̉E�[�c�r��
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


            // �z�z�������� ////////////////////////////////////////////////////////////////////////////////////
            
            // �z�z���������V�[�g��ǉ�����
            seikyuSheet.Copy(Type.Missing, oxlsSheet);
            oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[oXlsBook.Sheets.Count];

            // �Ώۊ���
            oxlsSheet.Cells[12, 3] = sKikan;

            // �z�z����
            oxlsSheet.Cells[4, 2] = staffID.ToString() + "  " + staffName;

            // �z�z���Z��
            OleDbDataReader dR;
            Control.�z�z�� sCon = new Control.�z�z��();
            dR = sCon.FillBy("where �z�z��.ID = " + staffID.ToString());
            while (dR.Read())
            {
                oxlsSheet.Cells[9, 2] = dR["�Z��"];
            }

            dR.Close();
            sCon.Close();

            // ���z
            oxlsSheet.Cells[10, 6] = Gur;

            for (int iz = 0; iz < sCell01.Length; iz++)
            {
                // ����
                oxlsSheet.Cells[iz + S_GYO, 1] = sCell01[iz];
                oxlsSheet.Cells[iz + S_GYO, 2] = sCell02[iz];
                oxlsSheet.Cells[iz + S_GYO, 3] = sCell03[iz];
                oxlsSheet.Cells[iz + S_GYO, 4] = sCell04[iz];
                oxlsSheet.Cells[iz + S_GYO, 5] = sCell05[iz];
                oxlsSheet.Cells[iz + S_GYO, 6] = sCell06[iz];
                oxlsSheet.Cells[iz + S_GYO, 7] = sCell07[iz];

                // �Z�������֎������R�r��������
                rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }

            // �\�S�̂Ɏ����c�r��������
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // �\�S�̂̍��[�c�r��
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // �\�S�̂̉E�[�c�r��
            rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
            rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
            oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        }

        private void KanryoReport(double Gur, string staffID, string staffName, string sKikan,
                                  string[] sCell01,string[] sCell02,string[] sCell03,string[] sCell04,
                                  string[] sCell05,string[] sCell06,string[] sCell07, string prnName)
        {
            const int S_GYO = 14;    //�G�N�Z���t�@�C�����׊J�n�s
            const int S_ROWSMAX = 7; //�G�N�Z���t�@�C����ő�l

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���z�z����������̏�, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {

                    //�Ώۊ���
                    oxlsSheet.Cells[12, 3] = sKikan;

                    //�z�z��ID
                    oxlsSheet.Cells[4, 1] = staffID;

                    //�z�z����
                    oxlsSheet.Cells[4, 2] = staffName + "�@�a";

                    //���z
                    oxlsSheet.Cells[10, 6] = Gur;

                    for (int iX = 0; iX < sCell01.Length; iX++)
                    {

                        //����
                        oxlsSheet.Cells[iX + S_GYO, 1] = sCell01[iX];
                        oxlsSheet.Cells[iX + S_GYO, 2] = sCell02[iX];
                        oxlsSheet.Cells[iX + S_GYO, 3] = sCell03[iX];
                        oxlsSheet.Cells[iX + S_GYO, 4] = sCell04[iX];
                        oxlsSheet.Cells[iX + S_GYO, 5] = sCell05[iX];
                        oxlsSheet.Cells[iX + S_GYO, 6] = sCell06[iX];
                        oxlsSheet.Cells[iX + S_GYO, 7] = sCell07[iX];

                        ////�Z���㕔�֎������R�r��������
                        //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //�Z�������֎������R�r��������
                        rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }

                    //�\�S�̂Ɏ����c�r��������
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̍��[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̉E�[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;

                    // �m�F�̂���Excel�̃E�B���h�E��\������
                    //oXls.Visible = true;

                    //���
                    //oxlsSheet.PrintPreview(false);
                    //oxlsSheet.PrintOut(1, Type.Missing, 1, false, oXls.ActivePrinter, Type.Missing, Type.Missing, Type.Missing);
                    oxlsSheet.PrintOut(1, Type.Missing, 1, false, prnName, Type.Missing, Type.Missing, Type.Missing);


                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    //�ۑ�����
                    oXls.DisplayAlerts = false;

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();

                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();
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
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "�������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void KanryoReportS(double Gur, int staffID, string staffName, string sKikan,
                                  string[] sCell01, string[] sCell02, string[] sCell03, string[] sCell04,
                                  string[] sCell05, string[] sCell06, string[] sCell07, string prnName)
        {

            const int S_GYO = 14;    //�G�N�Z���t�@�C�����׊J�n�s
            const int S_ROWSMAX = 7; //�G�N�Z���t�@�C����ő�l

            try
            {

                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���z�z��������, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {

                    //�Ώۊ���
                    oxlsSheet.Cells[12, 3] = sKikan;

                    //�z�z����
                    oxlsSheet.Cells[4, 2] = staffID.ToString() + "  " + staffName;

                    //�z�z���Z��
                    OleDbDataReader dR;
                    Control.�z�z�� sCon = new Control.�z�z��();
                    dR = sCon.FillBy("where �z�z��.ID = " + staffID.ToString());
                    while (dR.Read())
                    {
                        oxlsSheet.Cells[9, 2] = dR["�Z��"];
                    }

                    dR.Close();
                    sCon.Close();

                    //���z
                    oxlsSheet.Cells[10, 6] = Gur;

                    for (int iX = 0; iX < sCell01.Length; iX++)
                    {

                        //����
                        oxlsSheet.Cells[iX + S_GYO, 1] = sCell01[iX];
                        oxlsSheet.Cells[iX + S_GYO, 2] = sCell02[iX];
                        oxlsSheet.Cells[iX + S_GYO, 3] = sCell03[iX];
                        oxlsSheet.Cells[iX + S_GYO, 4] = sCell04[iX];
                        oxlsSheet.Cells[iX + S_GYO, 5] = sCell05[iX];
                        oxlsSheet.Cells[iX + S_GYO, 6] = sCell06[iX];
                        oxlsSheet.Cells[iX + S_GYO, 7] = sCell07[iX];

                        ////�Z���㕔�֎������R�r��������
                        //rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        //rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        //oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //�Z�������֎������R�r��������
                        rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                        rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }

                    //�\�S�̂Ɏ����c�r��������
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̍��[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�\�S�̂̉E�[�c�r��
                    rng[0] = (Excel.Range)oxlsSheet.Cells[S_GYO, S_ROWSMAX];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;

                    // �m�F�̂���Excel�̃E�B���h�E��\������
                    //oXls.Visible = true;

                    //���
                    //oxlsSheet.PrintPreview(false);
                    oxlsSheet.PrintOut(1, Type.Missing, 1, false, prnName, Type.Missing, Type.Missing, Type.Missing);


                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    //�ۑ�����
                    oXls.DisplayAlerts = false;

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();

                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();
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
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "�������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void txtYear_Validating(object sender, CancelEventArgs e)
        {

            string str;
            int d;

            if (txtYear.Text == null)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtYear.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

        }

        private void txtMonth_Validating(object sender, CancelEventArgs e)
        {

            string str;
            int d;

            if (txtMonth.Text == null)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtMonth.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("1�`12�œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtHID) txtObj = txtHID;
            if (sender == txtfMonth) txtObj = txtfMonth;
            if (sender == txtfDay) txtObj = txtfDay;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtHID) txtObj = txtHID;
            if (sender == txtfMonth) txtObj = txtfMonth;
            if (sender == txtfDay) txtObj = txtfDay;

            txtObj.BackColor = Color.White;
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)

            {
		        case 0:
                    
                    //�T����
                    label3.Enabled = false;
                    dateTimePicker1.Enabled = false;
                    label4.Enabled = false;
                    dateTimePicker2.Enabled = false;

                    //�Ώی�
                    label5.Enabled = true;
                    txtYear.Enabled = true;
                    label1.Enabled = true;
                    txtMonth.Enabled = true;
                    label2.Enabled = true;

                    break;

                case 1:

                    //�T����
                    label3.Enabled = true;
                    dateTimePicker1.Enabled = true;
                    label4.Enabled = true;
                    dateTimePicker2.Enabled = true;

                    //�Ώی�
                    label5.Enabled = false;
                    txtYear.Enabled = false;
                    label1.Enabled = false;
                    txtMonth.Enabled = false;
                    label2.Enabled = false;

                     break;
            }

        }

        private void DispClear()
        {

            //�T����
            label3.Enabled = false;
            dateTimePicker1.Enabled = false;
            label4.Enabled = false;
            dateTimePicker2.Enabled = false;

            //�Ώی�
            label5.Enabled = false;
            txtYear.Enabled = false;
            label1.Enabled = false;
            txtMonth.Enabled = false;
            label2.Enabled = false;

            //�z�z���w��
            label6.Enabled = false;
            txtHID.Enabled = false;
            label7.Enabled = false;

            //�U���f�[�^�쐬�֘A�I�u�W�F�N�g
            label12.Enabled = false;
            label13.Enabled = false;
            label14.Enabled = false;
            label15.Enabled = false;
            txtfMonth.Enabled = false;
            txtfDay.Enabled = false;

            button2.Enabled = false;
            button3.Enabled = false;
        }

        private bool DataCheck()
        {
            try
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("���܂��͏T��I�����Ă�������");
                }

                if (comboBox2.SelectedIndex == -1)
                {
                    comboBox2.Focus();
                    throw new Exception("�z�z���w���I�����Ă�������");
                }

                if (comboBox2.SelectedIndex == 1 && txtHID.Text == "")
                {
                    txtHID.Focus();
                    throw new Exception("�z�z��ID����͂��Ă�������");
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "�ێ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            switch (comboBox2.SelectedIndex)
            {
                case 0:
                    label6.Enabled = false;
                    txtHID.Enabled = false;
                    label7.Enabled = false;
                    break;

                case 1:
                    label6.Enabled = true;
                    txtHID.Enabled = true;
                    label7.Enabled = true;
                    break;
            }
        }

        private void txtHID_Validating(object sender, CancelEventArgs e)
        {
            int d;
            string str;

            // �����͂܂��̓X�y�[�X�݂͉̂�
            if ((this.txtHID.Text).Trim().Length < 1)
            {
                label7.Text = "";
                return;
            }

            // �������H
            if (txtHID.Text == null)
            {
                MessageBox.Show("�����œ��͂��Ă�������", "�z�z���R�[�h", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtHID.Text;

            if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", "�z�z���R�[�h", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            //�R�[�h����
            string sqlStr;
            Control.�z�z�� cStaff = new Control.�z�z��();
            OleDbDataReader dR;

            sqlStr = " where ID = " + txtHID.Text.ToString();
            dR = cStaff.FillBy(sqlStr);

            if (dR.HasRows == false)
            {
                MessageBox.Show("���o�^�R�[�h�ł�", "�z�z���R�[�h", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                label7.Text = "";
                dR.Close();
                cStaff.Close();
                return;
            }
            else
            {
                while (dR.Read())
                {

                    ////�z�z���̎x���P�ʂƏƍ�
                    //if (comboBox1.Text != dR["�x���敪"].ToString())
                    //{
                    //    MessageBox.Show("�w�肵���z�z���̎萔���x���P�ʂ́u" + comboBox1.Text + "�v�ł͂���܂���", "�z�z��", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    e.Cancel = true;
                    //    label7.Text = "";
                    //    dR.Close();
                    //    cStaff.Close();
                    //    return;
                    //}

                    label7.Text = dR["����"].ToString();
                }
            }

            dR.Close();
            cStaff.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (DataCheckZengin() == false) return;
                if (MessageBox.Show("�S��t�H�[�}�b�g�̐U���f�[�^���o�͂��܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                Cursor.Current = Cursors.WaitCursor;

                DialogResult ret;

                //�_�C�A���O�{�b�N�X�̏����ݒ�
                saveFileDialog1.Title = "�U���f�[�^�ۑ�";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = "�U���f�[�^_" + DateTime.Now.ToString("yyyyMMddhhmmss");

                saveFileDialog1.Filter = "÷��̧��(*.txt)|*.txt";

                //�_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret != System.Windows.Forms.DialogResult.OK) return;

                fileName = saveFileDialog1.FileName;

                //�w�b�_�N���X
                MakeZenginHead(fileName);

                //�f�[�^�N���X
                MakeZenginData(fileName);

                //�g���[���N���X
                MakeZenginTr(fileName);

                //�G���h�N���X
                MakeZenginEnd(fileName);

                //�I�����b�Z�[�W
                MessageBox.Show("�����I��", "�U���f�[�^�쐬");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private bool DataCheckZengin()
        {
            try
            {
                if (Utility.NumericCheck(txtfMonth.Text) == false)
                {
                    txtfMonth.Focus();
                    throw new Exception("�U�����͐����œ��͂��Ă�������");
                }

                if (int.Parse(txtfMonth.Text, System.Globalization.NumberStyles.Any) < 1 || int.Parse(txtfMonth.Text, System.Globalization.NumberStyles.Any) > 12)
                {
                    txtfMonth.Focus();
                    throw new Exception("�U����������������܂���");
                }

                if (Utility.NumericCheck(txtfDay.Text) == false)
                {
                    txtfDay.Focus();
                    throw new Exception("�U�����͐����œ��͂��Ă�������");
                }

                if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) < 1)
                {
                    txtfDay.Focus();
                    throw new Exception("�U����������������܂���");
                }

                switch (int.Parse(txtfMonth.Text,System.Globalization.NumberStyles.Any))
                {
                    case 1:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("�U����������������܂���");break;
                    case 2:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 29) throw new Exception("�U����������������܂���"); break;
                    case 3:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("�U����������������܂���"); break;
                    case 4:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("�U����������������܂���"); break;
                    case 5:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("�U����������������܂���"); break;
                    case 6:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("�U����������������܂���"); break;
                    case 7:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("�U����������������܂���"); break;
                    case 8:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("�U����������������܂���"); break;
                    case 9:  if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("�U����������������܂���"); break;
                    case 10: if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("�U����������������܂���"); break;
                    case 11: if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 30) throw new Exception("�U����������������܂���"); break;
                    case 12: if (int.Parse(txtfDay.Text, System.Globalization.NumberStyles.Any) > 31) throw new Exception("�U����������������܂���"); break;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "�ێ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

        }
        private void MakeZenginHead(string tempPath)
        {
            const string DATAKBN = "1";
            const string SHUBETSU = "21";
            const string BANKCODE = "0005";
            const string BANKNAME = "���޼ĳ�ֳUFJ";

            string stTarget;
            string sBuff;
            
            Entity.�S��w�b�_ cHead = new Entity.�S��w�b�_();

            //��Џ��擾
            OleDbDataReader dR;
            Control.��Џ�� cSys = new Control.��Џ��();
            dR = cSys.Fill();

            while (dR.Read())
            {
                sMaster.ID = int.Parse(dR["ID"].ToString(),System.Globalization.NumberStyles.Any);
                sMaster.�˗��l�R�[�h = dR["�˗��l�R�[�h"].ToString();
                sMaster.�˗��l�� = dR["�˗��l��"].ToString();
                sMaster.���Z�@�փR�[�h = dR["���Z�@�փR�[�h"].ToString();
                sMaster.���Z�@�֖� = dR["���Z�@�֖�"].ToString();
                sMaster.�x�X�R�[�h = dR["�x�X�R�[�h"].ToString();
                sMaster.�x�X�� = dR["�x�X��"].ToString();
                sMaster.������� = int.Parse(dR["�������"].ToString(),System.Globalization.NumberStyles.Any);
                sMaster.�����ԍ� = dR["�����ԍ�"].ToString();
            }

            dR.Close();
            cSys.Close();

            //�w�b�_�N���X�Ƀf�[�^���Z�b�g
            cHead.�f�[�^�敪 = DATAKBN;
            cHead.��ʃR�[�h = SHUBETSU;
            cHead.�R�[�h�敪 = " ";
            stTarget = "";
            cHead.�U���˗��l�R�[�h = stTarget.PadLeft(10);
            cHead.�U���˗��l�� = sMaster.�˗��l��.PadRight(40);
            cHead.��g�� = txtfMonth.Text.PadLeft(2, '0') + txtfDay.Text.PadLeft(2, '0');
            cHead.�d����s�ԍ� = BANKCODE;
            cHead.�d����s�� = BANKNAME.PadRight(15);
            cHead.�d���x�X�ԍ� = sMaster.�x�X�R�[�h.Trim().PadLeft(3, '0');
            cHead.�d���x�X�� = sMaster.�x�X��.PadRight(15);
            cHead.�a����� = sMaster.�������.ToString();
            cHead.�����ԍ� = sMaster.�����ԍ�.Trim().PadLeft(7, '0');
            stTarget = "";
            cHead.�_�~�[ = stTarget.PadLeft(17);
 
            //�o�̓��C��
            sBuff = "";
            sBuff += cHead.�f�[�^�敪;
            sBuff += cHead.��ʃR�[�h;
            sBuff += cHead.�R�[�h�敪;
            sBuff += cHead.�U���˗��l�R�[�h;
            sBuff += cHead.�U���˗��l��;
            sBuff += cHead.��g��;
            sBuff += cHead.�d����s�ԍ�;
            sBuff += cHead.�d����s��;
            sBuff += cHead.�d���x�X�ԍ�;
            sBuff += cHead.�d���x�X��;
            sBuff += cHead.�a�����;
            sBuff += cHead.�����ԍ�;
            sBuff += cHead.�_�~�[;

            //�e�L�X�g�f�[�^�o��
            System.IO.StreamWriter tZengin = new System.IO.StreamWriter(tempPath, false,System.Text.Encoding.GetEncoding(932));
            tZengin.WriteLine(sBuff);
            tZengin.Close();

        }

        private void MakeZenginData(string tempPath)
        {
            int staffID = 0;
            double Gur = 0;
            double GurTotal = 0;
            int dCnt = 0;

            //�t�@�C���I�[�v��
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true,System.Text.Encoding.GetEncoding(932));

            for (int iX = 0; iX <= dataGridView2.RowCount - 1; iX++)
            {

                //�f�[�^�쐬
                if ((staffID != 0) && (staffID != int.Parse(dataGridView2[1, iX].Value.ToString())))
                {
                    MakeZenginData_meisai(staffID, Gur, tZengin);
                    dCnt ++;
                }

                if (staffID != int.Parse(dataGridView2[1, iX].Value.ToString()))
                {
                    //�萔�����v
                    Gur = 0;
                }

                //�萔���v�Z
                Gur += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                GurTotal += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);

                //�z�z��ID
                staffID = int.Parse(dataGridView2[1, iX].Value.ToString());
            }

            MakeZenginData_meisai(staffID, Gur, tZengin);
            dCnt++;

            //�t�@�C���N���[�Y
            tZengin.Close();

            //�g���[���[�N���X�ɍ��v�����ƍ��v���z���Z�b�g
            sZtr.���v���� = dCnt.ToString().PadLeft(6, '0');
            sZtr.���v���z = GurTotal.ToString().PadLeft(12, '0');
            
        }

        private void MakeZenginData_meisai(int tempID, double tempKin,System.IO.StreamWriter tempZengin)
        {
            const string DATAKUBUN = "2";
            const string SHINKICODE = "0";
            const string FURIKOMISHITEIKUBUN = "7";

            string stTarget;
            string sBuff;
            string zKouza;
            string hKouza;

            Entity.�S��f�[�^���R�[�h cData = new Entity.�S��f�[�^���R�[�h();

            //�z�z�������擾
            OleDbDataReader dR;
            Control.�z�z�� cStaff = new Control.�z�z��();
            dR = cStaff.FillBy("where ID = " + tempID.ToString());

            while (dR.Read())
            {
                cData.��d����s�ԍ� = dR["���Z�@�փR�[�h"].ToString().Trim().PadLeft(4, '0');
                cData.��d���x�X�ԍ� = dR["�x�X�R�[�h"].ToString().Trim().PadLeft(3, '0');
                cData.�a����� = dR["�������"].ToString();

                //�����ԍ��S�p�����p�ɕϊ�
                zKouza = dR["�����ԍ�"].ToString().Trim() + "";
                hKouza = Strings.StrConv(zKouza, VbStrConv.Narrow, 0);
                cData.�����ԍ� = hKouza.PadLeft(7, '0');

                cData.���l�� = dR["�������`�J�i"].ToString().Trim().PadRight(30);
            }

            dR.Close();
            cStaff.Close();

            //�N���X�Ƀf�[�^���Z�b�g
            stTarget = "";
            cData.�f�[�^�敪 = DATAKUBUN;
            cData.��d����s�� = stTarget.PadLeft(15);
            cData.��d���x�X�� = stTarget.PadLeft(15);
            cData.��`�������ԍ� = stTarget.PadRight(4);
            cData.�U�����z = tempKin.ToString().PadLeft(10, '0');
            cData.�V�K�R�[�h = SHINKICODE;
            cData.�ڋq�R�[�h1 = tempID.ToString().PadLeft(10, '0');
            cData.�ڋq�R�[�h2 = tempID.ToString().PadLeft(10, '0');
            cData.�U���w��敪 = FURIKOMISHITEIKUBUN;
            cData.���ʕ\��  = stTarget.PadLeft(1);
            cData.�_�~�[ = stTarget.PadLeft(7);

            //�o�̓��C��
            sBuff = "";
            sBuff += cData.�f�[�^�敪;
            sBuff += cData.��d����s�ԍ�;
            sBuff += cData.��d����s��;
            sBuff += cData.��d���x�X�ԍ�;
            sBuff += cData.��d���x�X��;
            sBuff += cData.��`�������ԍ�;
            sBuff += cData.�a�����;
            sBuff += cData.�����ԍ�;
            sBuff += cData.���l��;
            sBuff += cData.�U�����z;
            sBuff += cData.�V�K�R�[�h;
            sBuff += cData.�ڋq�R�[�h1;
            sBuff += cData.�ڋq�R�[�h2;
            sBuff += cData.�U���w��敪;
            sBuff += cData.���ʕ\��;
            sBuff += cData.�_�~�[;

            //�e�L�X�g�f�[�^�o��
            tempZengin.WriteLine(sBuff);

        }

        private void MakeZenginTr(string tempPath)
        {
            
            const string DATAKUBUN = "8";
            string stTarget = "";
            string sbuff;

            //�t�@�C���I�[�v��
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true, System.Text.Encoding.GetEncoding(932));

            //�N���X�Ƀf�[�^�Z�b�g
            sZtr.�f�[�^�敪 = DATAKUBUN;
            sZtr.�_�~�[ = stTarget.PadLeft(101);

            //���C���o��
            sbuff = "";
            sbuff += sZtr.�f�[�^�敪;
            sbuff += sZtr.���v����;
            sbuff += sZtr.���v���z;
            sbuff += sZtr.�_�~�[;

            //�e�L�X�g�f�[�^�o��
            tZengin.WriteLine(sbuff);

            //�t�@�C���N���[�Y
            tZengin.Close();

        }

        private void MakeZenginEnd(string tempPath)
        {

            const string DATAKUBUN = "9";
            string stTarget = "";
            string sbuff;

            Entity.�S��G���h���R�[�h cEnd = new Entity.�S��G���h���R�[�h();

            //�t�@�C���I�[�v��
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true, System.Text.Encoding.GetEncoding(932));

            //�N���X�Ƀf�[�^�Z�b�g
            cEnd.�f�[�^�敪 = DATAKUBUN;
            cEnd.�_�~�[ = stTarget.PadLeft(119);

            //���C���o��
            sbuff = "";
            sbuff += cEnd.�f�[�^�敪;
            sbuff += cEnd.�_�~�[;

            //�e�L�X�g�f�[�^�o��
            tZengin.WriteLine(sbuff);

            //�t�@�C���N���[�Y
            tZengin.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {            
            try
            {
                if (MessageBox.Show("�z�z���ʎ萔���ꗗ���o�͂��܂��B��낵���ł����H", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

                Cursor.Current = Cursors.WaitCursor;

                DialogResult ret;
                string fNameHD;

                // �t�@�C�������o��
                switch (comboBox1.Text)
                {
                    case "�T":
                        fNameHD = dateTimePicker1.Value.ToLongDateString() + "����" + dateTimePicker2.Value.ToLongDateString() + "�܂�_"; 
                        break;

                    case "��":
                        fNameHD = txtYear.Text + "�N" + txtMonth.Text + "��_";
                        break;

                    default:
                        fNameHD = "";
                        break;
                }

                // �_�C�A���O�{�b�N�X�̏����ݒ�
                saveFileDialog1.Title = "�z�z���ʎ萔���ꗗ";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = fNameHD + "�z�z���ʎ萔���ꗗ_" + DateTime.Now.ToString("yyyyMMddhhmmss");

                saveFileDialog1.Filter = "csv̧��(*.csv)|*.csv";

                // �_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret != System.Windows.Forms.DialogResult.OK) return;

                fileName = saveFileDialog1.FileName;

                // �f�[�^�N���X
                MakeIchiranData(fileName);

                // �I�����b�Z�[�W
                MessageBox.Show("�����I��", "�z�z���ʎ萔���ꗗ");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private void MakeIchiranData(string tempPath)
        {
            int staffID = 0;
            string staffName = "";
            string officeName = "";
            double Gur = 0;
            double GurTotal = 0;
            int dCnt = 0;

            //�t�@�C���I�[�v��
            System.IO.StreamWriter tZengin;
            tZengin = new System.IO.StreamWriter(tempPath, true,System.Text.Encoding.GetEncoding(932));

            for (int iX = 0; iX <= dataGridView2.RowCount - 1; iX++)
            {
                //���o���o��
                if (iX == 0)
                {
                    MakeIchiranData_Header(tZengin);
                }

                //�f�[�^�쐬
                if ((staffID != 0) && (staffID != int.Parse(dataGridView2[1, iX].Value.ToString())))
                {
                    MakeIchiranData_meisai(officeName, staffID, staffName, Gur, tZengin);
                    dCnt++;
                }

                if (staffID != int.Parse(dataGridView2[1, iX].Value.ToString()))
                {
                    //�萔�����v
                    Gur = 0;
                }

                //�萔���v�Z
                Gur += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);
                GurTotal += double.Parse(dataGridView2[9, iX].Value.ToString(), System.Globalization.NumberStyles.Any);

                //�z�z��ID,����,���Ə�
                officeName = dataGridView2[0, iX].Value.ToString();
                staffID = int.Parse(dataGridView2[1, iX].Value.ToString());
                staffName = dataGridView2[2, iX].Value.ToString();
            }

            MakeIchiranData_meisai(officeName, staffID, staffName, Gur, tZengin);
            dCnt++;

            //�t�@�C���N���[�Y
            tZengin.Close();

        }
        private void MakeIchiranData_Header(System.IO.StreamWriter tempZengin)
        {
            string sBuff;

            //�o�̓��C�� : 2015/07/16 �}�C�i���o�[�ǉ�
            sBuff = "���Ə�,�z�z��ID,�z�z����,�萔��,�}�C�i���o�[";

            //�e�L�X�g�f�[�^�o��
            tempZengin.WriteLine(sBuff);

        }
        private void MakeIchiranData_meisai(string tempOffice,int tempID, string tempName,double tempKin, System.IO.StreamWriter tempZengin)
        {
            string sBuff;

            //�o�̓��C��
            sBuff = "";
            sBuff += tempOffice  + ",";
            sBuff += tempID.ToString() + ",";
            sBuff += tempName + ",";
            sBuff += tempKin.ToString() + ",";

            // �}�C�i���o�[ : 2015/07/16
            foreach (var t in dts.�z�z��.Where(a => a.ID == tempID))
            {
                sBuff += t.�}�C�i���o�[;
            }

            //�e�L�X�g�f�[�^�o��
            tempZengin.WriteLine(sBuff);
        }
    }
}
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
    public partial class frmOrder_0308 : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Utility.areaMode aMode = new Utility.areaMode();
        Utility.����ŗ� cTax = new Utility.����ŗ�();
        Entity.�� cMaster = new Entity.��();

        const string MESSAGE_CAPTION = "�󒍊m�菑";

        public frmOrder_0308()
        {
            InitializeComponent();
        }

        private void form_Load(object sender, EventArgs e)
        {
            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.��' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            GridviewSet.Setting(dataGridView1);
            this.darwinDataSet.Clear();
            this.darwinDataSet.EnforceConstraints = false;
            this.��TableAdapter.Fill(this.darwinDataSet.��);

            //�|�X�e�B���O�G���A
            GridviewSet.AriaSetting(dataGridView2);

            //�����}�X�^
            GridviewSet.TownSetting(dataGridView3);

            //�󒍋敪�R���{
            cmbJkbnSet();

            //���Ə��R���{
            Utility.ComboOffice.load(cmbOffice);

            //���Ӑ�R���{
            Utility.ComboClient.load(cmbClient);

            //�󒍎�ʃR���{
            Utility.ComboJshubetsu.load(cmbNaiyou);

            //�z�z�`�ԃR���{
            Utility.ComboFkeitai.load(cmbFkeitai);

            //�z�z�����R���{
            cmbFjyoukenSet();

            //���^�R���{
            Utility.ComboSize.load(cmbSize);

            //�z�z�����R���{
            cmbFyuyoSet();

            //�[�i�`�ԃR���{
            cmbNkeitaiSet();

            //�������@�R���{
            Utility.ComboShimebi.load(cmbNyukin);

            //�U�������R���{
            Utility.ComboFuri.load(cmbFuri);

            //�񍐎����R���{
            cmbHjikiSet();

            //�񍐐��x�R���{
            cmbHseidoSet();

            //�񍐕��@�R���{
            cmbHhouhouSet();

            //�ŗ��擾
            cTax.Ritsu = GetTaxRT(DateTime.Today);

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
                    tempDGV.ColumnHeadersHeight = 16;
                    tempDGV.RowTemplate.Height = 16;

                    // �S�̂̍���
                    tempDGV.Height = 164;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    //tempDGV.Columns.Add("col1", "����");
                    //tempDGV.Columns.Add("col2", "����");
                    //tempDGV.Columns.Add("col3", "���l");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 90;
                    tempDGV.Columns[11].Width = 90;

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
                    tempDGV.ColumnHeadersHeight = 16;
                    tempDGV.RowTemplate.Height = 16;

                    // �S�̂̍���
                    tempDGV.Height = 230;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�G���AID");
                    tempDGV.Columns.Add("col2", "�z�z�G���A");
                    tempDGV.Columns.Add("col3", "�z�z����");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 200;
                    tempDGV.Columns[2].Width = 80;

                    tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[2].DefaultCellStyle.Format = "#,##0";

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    //tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // �ҏW����
                    tempDGV.Columns[0].ReadOnly = true;
                    tempDGV.Columns[1].ReadOnly = true;
                    tempDGV.Columns[2].ReadOnly = true;

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
                    tempDGV.ColumnHeadersHeight = 16;
                    tempDGV.RowTemplate.Height = 16;

                    // �S�̂̍���
                    tempDGV.Height = 140;

                    // ��s�̐F
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�G���AID");
                    tempDGV.Columns.Add("col2", "����");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 200;

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

            /// <summary>
            /// �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾����
            /// </summary>
            /// <param name="dgv">�ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            public static Boolean GetData(DataGridView dgv,ref Entity.�� tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.�� Order = new Control.��();
                OleDbDataReader dr;

                sqlStr = " where ��.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Order.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.���Ə�ID = Int32.Parse(dr["���Ə�ID"].ToString());
                        tempC.�󒍓� = Convert.ToDateTime(dr["�󒍓�"].ToString());
                        tempC.�󒍋敪 = dr["�󒍋敪"].ToString() + "";
                        tempC.���Ӑ�ID = Int32.Parse(dr["���Ӑ�ID"].ToString());

                        //tempC.�Ј�ID = Int32.Parse(dr["�Ј�ID"].ToString());

                        tempC.�`���V�� = dr["�`���V��"].ToString() + "";
                        tempC.�󒍎��ID = Int32.Parse(dr["�󒍎��ID"].ToString());
                        tempC.�P�� = Convert.ToDouble(dr["�P��"].ToString());
                        tempC.���� = Int32.Parse(dr["����"].ToString());
                        tempC.���z = Int32.Parse(dr["���z"].ToString(),System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.����� = Int32.Parse(dr["�����"].ToString(),System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.�ō����z = Int32.Parse(dr["�ō����z"].ToString(),System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.�l���z = Int32.Parse(dr["�l���z"].ToString(), System.Globalization.NumberStyles.AllowDecimalPoint);
                        tempC.������z = Int32.Parse(dr["������z"].ToString(), System.Globalization.NumberStyles.AllowDecimalPoint);                        
                        tempC.�ŗ� = Int32.Parse(dr["�ŗ�"].ToString());
                        tempC.���^ = Int32.Parse(dr["���^"].ToString());
                        tempC.�˗��� = dr["�˗���"].ToString();
                        tempC.���� = Convert.ToDouble(dr["����"].ToString());
                        tempC.�z�z�`�� = Int32.Parse(dr["�z�z�`��"].ToString());
                        tempC.�z�z���� = dr["�z�z����"].ToString() + "";
                        tempC.�z�z�J�n�� = dr["�z�z�J�n��"].ToString();
                        tempC.�z�z�I���� = dr["�z�z�I����"].ToString();
                        tempC.�z�z�P�\ = dr["�z�z�P�\"].ToString() + "";
                        tempC.�[�i�\��� = dr["�[�i�\���"].ToString();
                        tempC.�[�i�`�� = dr["�[�i�`��"].ToString() + "";
                        tempC.������ = Int32.Parse(dr["������"].ToString());
                        tempC.�������@ = dr["�������@"].ToString() + "";
                        tempC.�����\��� = dr["�����\���"].ToString();
                        tempC.�񍐎��� = dr["�񍐎���"].ToString() + "";
                        tempC.�񍐐��x = dr["�񍐐��x"].ToString() + "";
                        tempC.�񍐕��@ = dr["�񍐕��@"].ToString() + "";
                        tempC.���[���A�h���X = dr["���[���A�h���X"].ToString() + "";
                        tempC.�U������ID = Int32.Parse(dr["�U������ID"].ToString());
                        tempC.���z�z���L�� = Int32.Parse(dr["���z�z���L��"].ToString());
                        tempC.�}�ԗL�� = Int32.Parse(dr["�}�ԗL��"].ToString());
                        tempC.���L���� = dr["���L����"].ToString() + "";
                        tempC.�G���A���l = dr["�G���A���l"].ToString() + "";
                    }
                }
                else
                {
                    dr.Close();
                    Order.Close();
                    return false;
                }

                dr.Close();
                Order.Close();
                return true;
            }

            //public static void ShowData(DataGridView tempDGV)
            //{
            //    string sqlSTRING = "";

            //    try
            //    {
            //        tempDGV.RowCount = 0;

            //        //�󒍃f�[�^�̃f�[�^���[�_�[���擾����
            //        Control.FreeSql cOrder = new Control.FreeSql();
            //        cOrder.Execute("");
            //        Control.DataControl dCon = new Control.DataControl();

            //        sqlSTRING = "select * from m_Costname " +
            //                    "order by ID";

            //        dR = dCon.FreeReader(sqlSTRING);

            //        iX = 0;

            //        while (dR.Read())
            //        {
            //            tempDGV.Rows.Add();

            //            tempDGV[0, iX].Value = dR["ID"];
            //            tempDGV[1, iX].Value = NullConvert.Noth(dR["������"]);
            //            tempDGV[2, iX].Value = NullConvert.Noth(dR["���l"]);
            //            //tempDGV[1, iX].Value = dR["������"];
            //            //tempDGV[2, iX].Value = dR["���l"];
            //            iX++;
            //        }

            //        dR.Close();

            //        dCon.Close();

            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
            //    }

            //}

        }

        //�O���b�h����f�[�^��I��
        private void GridEnter()
        {

            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[2, dataGridView1.SelectedRows[iX].Index].Value + " " + dataGridView1[3, dataGridView1.SelectedRows[iX].Index].Value +"���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�󒍊m�菑�I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
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
                        MessageBox.Show("�Y������f�[�^���o�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    //'�f�[�^�l���擾
                    jDate.Value = Convert.ToDateTime(cMaster.�󒍓�.ToString());
                    comboBox1.Text = cMaster.�󒍋敪;

                    Utility.ComboOffice.selectedIndex(cmbOffice, Int32.Parse(cMaster.���Ə�ID.ToString()));
                    
                    //�N���C�A���g���\��
                    Utility.ComboClient.selectedIndex(cmbClient, Int32.Parse(cMaster.���Ӑ�ID.ToString()));

                    txtCzipcode.Text = "";
                    txtName2.Text = "";
                    txtCbusho.Text = "";
                    txtTantou.Text = "";
                    txtCtel.Text = "";
                    txtCfax.Text  = "";
                    txtCbusho.Text = "";
                    txtCtantou.Text = "";

                    //�N���C�A���g���\��
                    ClientShow(cMaster.���Ӑ�ID);

                    txtChirashi.Text = cMaster.�`���V��;
                    Utility.ComboJshubetsu.selectedIndex(cmbNaiyou, Int32.Parse(cMaster.�󒍎��ID.ToString()));

                    txtTanka.Text = cMaster.�P��.ToString("#,##0.0");
                    txtMai.Text  = cMaster.����.ToString("#,##0");
                    txtUri.Text = cMaster.���z.ToString("#,##0");
                    txtTax.Text = cMaster.�����.ToString("#,##0");
                    txtZeikomi.Text = cMaster.�ō����z.ToString("#,##0");
                    txtNebiki.Text = cMaster.�l���z.ToString("#,##0");
                    txtUriTL.Text = cMaster.������z.ToString("#,##0");

                    Utility.ComboFkeitai.selectedIndex(cmbFkeitai, Int32.Parse(cMaster.�z�z�`��.ToString()));
                    cmbFjyouken.Text = cMaster.�z�z����;
                    Utility.ComboSize.selectedIndex(cmbSize, Int32.Parse(cMaster.���^.ToString()));

                    txtIraisaki.Text = cMaster.�˗���;
                    txtGenka.Text = cMaster.����.ToString("#,##0.0");

                    if (cMaster.�z�z�J�n��.ToString() == "")
                    {
                        StartDate.Checked = false;
                    }
                    else
                    {
                        StartDate.Value = Convert.ToDateTime(cMaster.�z�z�J�n��.ToString());
                    }

                    if (cMaster.�z�z�I����.ToString() == "")
                    {
                        EndDate.Checked = false;
                    }
                    else
                    {
                        EndDate.Value = Convert.ToDateTime(cMaster.�z�z�I����.ToString());
                    }

                    cmbFyuyo.Text = cMaster.�z�z�P�\;

                    if (Int32.Parse(cMaster.������.ToString()) == 1)
                    {
                        checkBox1.Checked = true;
                    }
                    else
                    {
                        checkBox1.Checked = false;
                    }

                    if (cMaster.�[�i�\���.ToString() == "")
                    {
                        NouhinDate.Checked = false;
                    }
                    else
                    {
                        NouhinDate.Value = Convert.ToDateTime(cMaster.�[�i�\���.ToString());
                    }

                    cmbNkeitai.Text = cMaster.�[�i�`��;
                    cmbNyukin.Text = cMaster.�������@;
                    
                    if (cMaster.�����\���.ToString() == "")
                    {
                        NyukinDate.Checked = false;
                    }
                    else
                    {
                        NyukinDate.Value = Convert.ToDateTime(cMaster.�����\���.ToString());
                    }

                    Utility.ComboFuri.selectedIndex(cmbFuri, cMaster.�U������ID);
                    cmbHjiki.Text = cMaster.�񍐎���;
                    cmbHseido.Text = cMaster.�񍐐��x;
                    cmbHhouhou.Text = cMaster.�񍐕��@;

                    txtEmail.Text = cMaster.���[���A�h���X;
                    txtMemo.Text = cMaster.���L����;
                    txtMemo2.Text = cMaster.�G���A���l;

                    if (Int32.Parse(cMaster.���z�z���L��.ToString()) == 1)
                    {
                        checkBox2.Checked = true;
                    }
                    else
                    {
                        checkBox2.Checked = false;
                    }

                    if (Int32.Parse(cMaster.�}�ԗL��.ToString()) == 1)
                    {
                        checkBox3.Checked = true;
                    }
                    else
                    {
                        checkBox3.Checked = false;
                    }

                    //ID�e�L�X�g�{�b�N�X�͕ҏW�s�Ƃ���
                    //txtCode.Enabled = false;

                    //�{�^�����
                    btnDel.Enabled = true;
                    btnClr.Enabled = true;

                    fMode.Mode = 1;     //�t�H�[�����[�h�X�e�[�^�X:�ύX�폜

                    jDate.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "�f�[�^�\��", MessageBoxButtons.OK);
                }
            }

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            // Enterkey�ȊO�͑ΏۊO
            if (e.KeyCode.ToString() != "Return")
            {
                return;
            }

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

                jDate.Value = DateTime.Today;
                comboBox1.SelectedIndex = -1;
                cmbOffice.SelectedIndex = -1;
                cmbClient.SelectedIndex = -1;
                txtCzipcode.Text = "";
                txtName2.Text = "";
                txtCbusho.Text = "";
                txtTantou.Text = "";
                txtCtel.Text = "";
                txtCfax.Text = "";
                txtCbusho.Text = "";
                txtCtantou.Text = "";

                txtChirashi.Text = "";
                cmbNaiyou.SelectedIndex=-1;
                txtTanka.Text = "0";
                txtMai.Text = "0";
                txtUri.Text = "0";
                txtTax.Text = "0";
                txtZeikomi.Text = "0";
                txtNebiki.Text = "0";
                txtUriTL.Text = "0";

                cmbFkeitai.SelectedIndex=-1;
                cmbFjyouken.SelectedIndex = -1;
                cmbSize.SelectedIndex = -1;

                txtIraisaki.Text = "";
                txtGenka.Text = "0";

                StartDate.Checked = false;
                EndDate.Checked = false;
                cmbFyuyo.SelectedIndex = -1;
                checkBox1.Checked = false;
                NouhinDate.Checked = false;
                cmbNkeitai.SelectedIndex = -1;
                cmbNyukin.SelectedIndex = -1;
                NyukinDate.Checked = false;
                cmbFuri.SelectedIndex = -1;
                cmbHjiki.SelectedIndex = -1;
                cmbHseido.SelectedIndex = -1;
                cmbHhouhou.SelectedIndex = -1;

                txtEmail.Text = "";
                txtMemo.Text = "";
                txtMemo2.Text = "";
                checkBox2.Checked = false;
                checkBox3.Checked = false;

                btnDel.Enabled = false;
                btnClr.Enabled = false;

                txtAreaID.Text = "";
                txtAreaName.Text = "";
                txtHaihuMaisu.Text = "";

                dataGridView2.Rows.Clear();
                dataGridView3.Rows.Clear();

                textBox5.Text = "";
                txtAdel.Enabled = false;

                //txtCode.Focus();
                jDate.Focus();
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
                    Control.�� Order = new Control.��();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                break;

                            if (Order.DataInsert(cMaster) == true)
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
                                break;

                            if (Order.DataUpdate(cMaster) == true)
                            {
                                MessageBox.Show("�X�V����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("�X�V�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            break;

                    }

                    Order.Close();

                    DispClear();

                    //�f�[�^�� 'darwinDataSet.��' �e�[�u���ɓǂݍ��݂܂��B
                    this.��TableAdapter.Fill(this.darwinDataSet.��);
                    dataGridView1.DataSource = this.darwinDataSet.��;

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
                    //Control.�� Order = new Control.��();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Order.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Order.Close();
                    //    throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    //}

                    //dr.Close();
                    //Order.Close();

                }
                //�󒍋敪�`�F�b�N
                if (comboBox1.SelectedIndex == -1)
                {
                    comboBox1.Focus();
                    throw new Exception("�󒍋敪��I�����Ă�������");
                }

                //���Ə��h�c�`�F�b�N
                if (cmbOffice.SelectedIndex == -1)
                {
                    cmbOffice.Focus();
                    throw new Exception("���Ə���I�����Ă�������");
                }

                //�N���C�A���g�`�F�b�N
                if (cmbClient.SelectedIndex == -1)
                {
                    cmbClient.Focus();
                    throw new Exception("�N���C�A���g��I�����Ă�������");
                }

                //�`���V���`�F�b�N
                if (txtChirashi.Text.Trim().Length < 1)
                {
                    txtChirashi.Focus();
                    throw new Exception("�`���V������͂��Ă�������");
                }

                //�󒍓��e�`�F�b�N
                if (cmbNaiyou.SelectedIndex == -1)
                {
                    cmbNaiyou.Focus();
                    throw new Exception("�󒍓��e��I�����Ă�������");
                }

                //�P���F�������H
                if (txtTanka.Text == null)
                {
                    this.txtTanka.Focus();
                    throw new Exception("�P���͐����œ��͂��Ă�������");
                }

                str = this.txtTanka.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtTanka.Focus();
                    throw new Exception("�P���͐����œ��͂��Ă�������");
                }

                //�����F�������H
                if (txtMai.Text == null)
                {
                    this.txtMai.Focus();
                    throw new Exception("�����͐����œ��͂��Ă�������");
                }

                str = this.txtMai.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtMai.Focus();
                    throw new Exception("�����͐����œ��͂��Ă�������");
                }

                //�z�z�`�ԃ`�F�b�N
                if (cmbFkeitai.SelectedIndex == -1)
                {
                    cmbFkeitai.Focus();
                    throw new Exception("�z�z�`�Ԃ�I�����Ă�������");
                }

                //���^�`�F�b�N
                if (cmbSize.SelectedIndex == -1)
                {
                    cmbSize.Focus();
                    throw new Exception("�T�C�Y��I�����Ă�������");
                }

                //�����F�������H
                if (txtGenka.Text == null)
                {
                    this.txtGenka.Focus();
                    throw new Exception("�����͐����œ��͂��Ă�������");
                }

                str = this.txtGenka.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtGenka.Focus();
                    throw new Exception("�����͐����œ��͂��Ă�������");
                }

                if (cmbNaiyou.Text == "�|�X�e�B���O")
                {
                    //�z�z�����`�F�b�N
                    if (cmbFjyouken.SelectedIndex == -1)
                    {
                        cmbFjyouken.Focus();
                        throw new Exception("�z�z������I�����Ă�������");
                    }

                    //�z�z����
                    if (StartDate.Value > EndDate.Value)
                    {
                        StartDate.Focus();
                        throw new Exception("�z�z���Ԃ�����������܂���");
                    }

                    //�z�z�P�\�`�F�b�N
                    if (cmbFyuyo.SelectedIndex == -1)
                    {
                        cmbFyuyo.Focus();
                        throw new Exception("�z�z�P�\��I�����Ă�������");
                    }

                    //�񍐎����`�F�b�N
                    if (cmbHjiki.SelectedIndex == -1)
                    {
                        cmbHjiki.Focus();
                        throw new Exception("�񍐎�����I�����Ă�������");
                    }

                    //�񍐐��x�`�F�b�N
                    if (cmbHseido.SelectedIndex == -1)
                    {
                        cmbHseido.Focus();
                        throw new Exception("�񍐐��x��I�����Ă�������");
                    }
                    //�񍐕��@�`�F�b�N
                    if (cmbHhouhou.SelectedIndex == -1)
                    {
                        cmbHhouhou.Focus();
                        throw new Exception("�񍐕��@��I�����Ă�������");
                    }
                }

                //�N���X�Ƀf�[�^�Z�b�g
                Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                cmb1 = (Utility.ComboOffice)cmbOffice.SelectedItem;
                cMaster.���Ə�ID = cmb1.ID;

                cMaster.�󒍓� = jDate.Value;
                cMaster.�󒍋敪 = comboBox1.Text;

                Utility.ComboClient cmb2 = new Utility.ComboClient();
                cmb2 = (Utility.ComboClient)cmbClient.SelectedItem;
                cMaster.���Ӑ�ID = cmb2.ID;

                cMaster.�`���V�� = txtChirashi.Text.ToString();

                Utility.ComboJshubetsu cmb3 = new Utility.ComboJshubetsu();
                cmb3 = (Utility.ComboJshubetsu)cmbNaiyou.SelectedItem;
                cMaster.�󒍎��ID = cmb3.ID;

                cMaster.�P�� = Convert.ToDouble(txtTanka.Text.ToString());
                cMaster.���� = Int32.Parse(txtMai.Text.ToString(),System.Globalization.NumberStyles.AllowThousands);
                cMaster.���z = Int32.Parse(txtUri.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.����� = Int32.Parse(txtTax.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.�ō����z = Int32.Parse(txtZeikomi.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.�l���z = Int32.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.������z = Int32.Parse(txtUriTL.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);
                cMaster.�ŗ� = cTax.Ritsu;

                Utility.ComboSize cmb4 = new Utility.ComboSize();
                cmb4 = (Utility.ComboSize)cmbSize.SelectedItem;
                cMaster.���^ = cmb4.ID;

                cMaster.�˗��� = txtIraisaki.Text.ToString();
                cMaster.���� = Convert.ToDouble(txtGenka.Text.ToString());

                Utility.ComboFkeitai cmb5 = new Utility.ComboFkeitai();
                cmb5 = (Utility.ComboFkeitai)cmbFkeitai.SelectedItem;
                cMaster.�z�z�`�� = cmb5.ID;

                cMaster.�z�z���� = cmbFjyouken.Text.ToString();

                if (StartDate.Checked == true)
                {
                    cMaster.�z�z�J�n�� = StartDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�z�z�J�n�� = "";
                }

                if (EndDate.Checked == true)
                {
                    cMaster.�z�z�I���� = EndDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�z�z�I���� = "";
                }

                cMaster.�z�z�P�\ = cmbFyuyo.Text.ToString();

                if (NouhinDate.Checked == true)
                {
                    cMaster.�[�i�\��� = NouhinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�[�i�\��� = "";
                }

                cMaster.�[�i�`�� = cmbNkeitai.Text.ToString();

                if (checkBox1.Checked == true)
                {
                    cMaster.������ = 1;
                }
                else
                {
                    cMaster.������ = 0;
                }

                if (cmbNyukin.SelectedIndex == -1)
                {
                    cMaster.�������@ = "";
                }
                else
                {
                    Utility.ComboShimebi cmb6 = new Utility.ComboShimebi();
                    cmb6 = (Utility.ComboShimebi)cmbNyukin.SelectedItem;
                    cMaster.�������@ = cmb6.Name;
                }

                if (NyukinDate.Checked == true)
                {
                    cMaster.�����\��� = NyukinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�����\��� = "";
                }

                cMaster.�񍐎��� = cmbHjiki.Text;
                cMaster.�񍐐��x = cmbHseido.Text;
                cMaster.�񍐕��@ = cmbHhouhou.Text;
                cMaster.���[���A�h���X = txtEmail.Text;

                if (cmbFuri.SelectedIndex == -1)
                {
                    cMaster.�U������ID = 0;
                }
                else
                {
                    Utility.ComboFuri cmb7 = new Utility.ComboFuri();
                    cmb7 = (Utility.ComboFuri)cmbFuri.SelectedItem;
                    cMaster.�U������ID = cmb7.ID;
                }

                if (checkBox2.Checked == true)
                {
                    cMaster.���z�z���L�� = 1;
                }
                else
                {
                    cMaster.���z�z���L�� = 0;
                }

                if (checkBox3.Checked == true)
                {
                    cMaster.�}�ԗL�� = 1;
                }
                else
                {
                    cMaster.�}�ԗL�� = 0;
                }

                cMaster.���L���� = txtMemo.Text.ToString();
                cMaster.�G���A���l = txtMemo2.Text.ToString();

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

            if (sender == txtChirashi)
            {
                objtxt = txtChirashi;
            }

            if (sender == txtTanka)
            {
                objtxt = txtTanka;
            }

            if (sender == txtMai)
            {
                objtxt = txtMai;
            }

            if (sender == txtNebiki)
            {
                objtxt = txtNebiki;
            }

            if (sender == txtIraisaki)
            {
                objtxt = txtIraisaki;
            }

            if (sender == txtGenka)
            {
                objtxt = txtGenka;
            }

            if (sender == txtEmail)
            {
                objtxt = txtEmail;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == txtMemo2)
            {
                objtxt = txtMemo2;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == cmbClient)
            {
                cmbClient.BackColor = Color.LightGray;
            }

            if (sender == txtAreaID)
            {
                objtxt = txtAreaID;
            }

            if (sender == txtHaihuMaisu)
            {
                objtxt = txtHaihuMaisu;
            }

            objtxt.SelectAll();
            objtxt.BackColor = Color.LightGray;

        }

        private void txtLeave(object sender, EventArgs e)
       {

           TextBox objtxt = new TextBox();

           double Kingaku;
           int KingakuTax;
           int KingakuZeikomi;
           int KingakuTL;
           string str;
           double d;

           try
           {


               if (sender == txtChirashi)
               {
                   objtxt = txtChirashi;
               }

               if (sender == txtTanka)
               {
                   objtxt = txtTanka;

                   if (txtTanka.Text == null) txtTanka.Text = "0";

                   str = txtTanka.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtTanka.Text = "0";

                   UriageSum(Convert.ToDateTime(jDate.Value), Convert.ToDouble(txtTanka.Text),
                             int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             int.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             out Kingaku, out KingakuTax, out KingakuZeikomi,out KingakuTL);


                   //������z
                   txtUri.Text = Kingaku.ToString("#,##0");

                   //����Ŋz�v�Z               
                   txtTax.Text = KingakuTax.ToString("#,##0");

                   //�ō����z
                   txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                   //������z
                   txtUriTL.Text  = KingakuTL.ToString("#,##0");
               }

               if (sender == txtMai)
               {
                   objtxt = txtMai;

                   if (txtMai.Text == null) txtMai.Text = "0";

                   str = txtMai.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtMai.Text = "0";

                   UriageSum(Convert.ToDateTime(jDate.Value), Convert.ToDouble(txtTanka.Text),
                             int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             int.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             out Kingaku, out KingakuTax, out KingakuZeikomi,out KingakuTL);

                   //������z
                   txtUri.Text = Kingaku.ToString("#,##0");

                   //����Ŋz�v�Z               
                   txtTax.Text = KingakuTax.ToString("#,##0");

                   //�ō����z
                   txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                   //������z
                   txtUriTL.Text = KingakuTL.ToString("#,##0");
               }

               if (sender == txtNebiki)
               {
                   objtxt = txtNebiki;

                   if (txtNebiki.Text == null) txtNebiki.Text = "0";

                   str = txtNebiki.Text;

                   if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                       txtNebiki.Text = "0";

                   UriageSum(Convert.ToDateTime(jDate.Value), Convert.ToDouble(txtTanka.Text),
                             int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             int.Parse(txtNebiki.Text.ToString(), System.Globalization.NumberStyles.AllowThousands),
                             out Kingaku, out KingakuTax, out KingakuZeikomi, out KingakuTL);

                   //������z
                   txtUri.Text = Kingaku.ToString("#,##0");

                   //����Ŋz�v�Z               
                   txtTax.Text = KingakuTax.ToString("#,##0");

                   //�ō����z
                   txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                   //������z
                   txtUriTL.Text = KingakuTL.ToString("#,##0");

               }

               if (sender == txtIraisaki)
               {
                   objtxt = txtIraisaki;
               }

               if (sender == txtGenka)
               {
                   objtxt = txtGenka;
               }

               if (sender == txtEmail)
               {
                   objtxt = txtEmail;
               }

               if (sender == txtMemo)
               {
                   objtxt = txtMemo;
               }

               if (sender == txtMemo2)
               {
                   objtxt = txtMemo2;
               }

               if (sender == textBox1)
               {
                   objtxt = textBox1;
               }

               if (sender == cmbClient)
               {
                   cmbClient.BackColor = Color.White;

                   //�N���C�A���g���\��
                   Utility.ComboClient cmbC = new Utility.ComboClient();
                   cmbC = (Utility.ComboClient)cmbClient.SelectedItem;
                   ClientShow(cmbC.ID);

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

               objtxt.BackColor = Color.White;
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message,"�G���[���b�Z�[�W");
           }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {

            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�폜
            Control.�� Order = new Control.��();
            if (Order.DataDelete(cMaster.ID) == true)
            {
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("�󒍃f�[�^�̍폜�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            Order.Close();

            //�z�z�G���A�f�[�^�폜
            string strSql;
            strSql = "";
            strSql += "delete from �z�z�G���A ";
            strSql += "where �z�z�G���A.��ID = " + cMaster.ID.ToString();

            Control.FreeSql fsql = new Control.FreeSql();
            if (fsql.Execute(strSql) == true)
            {
                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("�z�z�G���A�f�[�^�̍폜�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            fsql.Close();

            DispClear();

            //�f�[�^�� 'darwinDataSet.��' �e�[�u���ɓǂݍ��݂܂��B
            this.��TableAdapter.Fill(this.darwinDataSet.��);
            dataGridView1.DataSource = this.darwinDataSet.��;

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

        private void cmbJkbnSet()
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("�V�K");
            comboBox1.Items.Add("���s�[�g");
        }

        private void cmbFjyoukenSet()
        {
            cmbFjyouken.Items.Clear();
            cmbFjyouken.Items.Add("�\�薇���ǂ���");
            cmbFjyouken.Items.Add("�l�߂Ĕz�z �n�j");
        }

        private void cmbFyuyoSet()
        {
            cmbFyuyo.Items.Clear();
            cmbFyuyo.Items.Add("����");
            cmbFyuyo.Items.Add("�O��n�j");
        }

        private void cmbNkeitaiSet()
        {
            cmbNkeitai.Items.Clear();
            cmbNkeitai.Items.Add("��}��");
            cmbNkeitai.Items.Add("����");
            cmbNkeitai.Items.Add("�W�ׁF�o�b�N");
            cmbNkeitai.Items.Add("�W�ׁF�c��");
        }

        private void cmbHjikiSet()
        {
            cmbHjiki.Items.Clear();
            cmbHjiki.Items.Add("�f�C���[");
            cmbHjiki.Items.Add("�T�P��");
            cmbHjiki.Items.Add("�I����");
        }

        private void cmbHseidoSet()
        {
            cmbHseido.Items.Clear();
            cmbHseido.Items.Add("���z�z��");
            cmbHseido.Items.Add("�\�薇��");
        }

        private void cmbHhouhouSet()
        {
            cmbHhouhou.Items.Clear();
            cmbHhouhou.Items.Add("FAX");
            cmbHhouhou.Items.Add("���[��");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            darwinDataSet ds = new darwinDataSet();
            ds.Clear();
            ds.EnforceConstraints = false;
            this.��TableAdapter.FillByName(ds.��, "%" + textBox1.Text.ToString() + "%");
            dataGridView1.DataSource = ds.��;
        }

        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmOffice();

            frm.ShowDialog();
            Utility.ComboOffice.load(cmbOffice);

        }

        /// <summary>
        /// ����ŗ��擾
        /// </summary>
        /// <param name="tempDate">����t</param>
        /// <returns>�ŗ�</returns>
        private int GetTaxRT(DateTime tempDate)
        {
            //�ŗ��擾
            string Sqlstr;
            int Ritsu = 0;

            OleDbDataReader drTax;
            Control.�ŗ� cR = new Control.�ŗ�();
            Sqlstr = "where �ŗ�.�J�n�N���� <= '" + jDate.Value.ToShortDateString() + "' order by �J�n�N����";
            drTax = cR.FillBy(Sqlstr);

            while (drTax.Read())
            {
                Ritsu = Int32.Parse(drTax["�ŗ�"].ToString());
            }

            drTax.Close();

            return Ritsu;
            
        }

        /// <summary>
        /// ����Ōv�Z
        /// </summary>
        /// <param name="tempKin">�Ώۋ��z</param>
        /// <param name="tempTax">�ŗ�</param>
        /// <returns>����Ŋz</returns>
        private int GetTax(double tempKin,int tempTax)
        {
            double GakuD;
            int GakuI;

            GakuD = tempKin * tempTax / 100;
            GakuD += 0.5;
            GakuI = (int)GakuD;

            return GakuI;
        }

        private void UriageSum(DateTime tempDate,double tempTanka,int tempMai,int tempNebiki,
                               out double Kingaku,out int KingakuTax,out int KingakuZeikomi,
                               out int KingakuTL)
        {

            //������z
            Kingaku = tempTanka * tempMai;

            //�ŗ��Ď擾
            cTax.Ritsu = GetTaxRT(tempDate);

            //����Ŋz�v�Z               
            KingakuTax = GetTax(Kingaku, cTax.Ritsu);

            //�ō����z
            KingakuZeikomi = (int)Kingaku + KingakuTax;

            //������z
            KingakuTL = KingakuZeikomi - tempNebiki;
        }

        private void ClientShow(int tempID)
        {
            OleDbDataReader drt;
            Control.���Ӑ� Client = new Control.���Ӑ�();
            drt = Client.FillBy("where ID = " + tempID.ToString());

            while (drt.Read())
            {
                txtCzipcode.Text = drt["�X�֔ԍ�"].ToString();
                txtName2.Text = drt["�Z��1"].ToString() + "" + " " + drt["�Z��2"].ToString() + "";
                txtCbusho.Text = drt["������"].ToString() + "";
                txtCtantou.Text = drt["�S���Җ�"].ToString() + "";
                txtCtel.Text = drt["�d�b�ԍ�"].ToString() + "";
                txtCfax.Text = drt["FAX�ԍ�"].ToString() + "";

                OleDbDataReader drS;
                Control.�Ј� Shain = new Control.�Ј�();
                drS = Shain.FillBy("where ID = " + drt["�S���Ј��R�[�h"].ToString());

                while (drS.Read())
                {
                    txtTantou.Text = drS["����"].ToString();
                }

                drS.Close();

            }

            drt.Close();

        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void cmbClient_Click(object sender, EventArgs e)
        {

        }

        private void ��BindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void txtZeikomi_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void cmbFkeitai_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtAdd_Click(object sender, EventArgs e)
        {
            //�z�z�G���A�o�^

            string str;
            double d;
            int iX;

            try
            {
                if (txtHaihuMaisu.Text == null)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("�z�z�����͐����œ��͂��Ă�������");
                }

                str = txtHaihuMaisu.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("�z�z�����͐����œ��͂��Ă�������");
                }

                if (Int32.Parse(txtHaihuMaisu.Text.ToString()) < 0)
                {
                    this.txtHaihuMaisu.Focus();
                    throw new Exception("�z�z����������������܂���");
                }

                switch (aMode.Mode)
                {
                    case 0: //�o�^
                        dataGridView2.Rows.Add();
                        iX = dataGridView2.Rows.Count - 1;
                        dataGridView2[0, iX].Value = Int32.Parse(txtAreaID.Text.ToString());
                        //dataGridView2[1, iX].Value = txtAreaName.Text;
                        dataGridView2[2, iX].Value = Int32.Parse(txtHaihuMaisu.Text.ToString());
                        break;

                    case 1: //�X�V

                        break;
                }

                textAreaClear();
                txtAreaID.Focus();

                txtTotal.Text = GetMaisuTotal().ToString("#,##0");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MESSAGE_CAPTION + "�ێ�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void textAreaClear()
        {
            aMode.Mode = 0;

            txtAreaID.Text = "";
            txtAreaName.Text = "";
            txtHaihuMaisu.Text = "";
            txtAdel.Enabled = false;

            //�\�[�g�g�p��
            dataGridView2.Columns[0].SortMode = DataGridViewColumnSortMode.Automatic;
            dataGridView2.Columns[1].SortMode = DataGridViewColumnSortMode.Automatic;
            dataGridView2.Columns[2].SortMode = DataGridViewColumnSortMode.Automatic;
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex != 0) return;

            dataGridView2[1, e.RowIndex].Value = GetTownName(dataGridView2[e.ColumnIndex, e.RowIndex].Value.ToString());
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

                dr.Close();
                cTown.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        //�z�z�G���A�̒������ꊇ�o�^����
        {
            int iX = 0;
            foreach (DataGridViewRow r in dataGridView3.SelectedRows)
            {
                dataGridView2.Rows.Add();
                iX = dataGridView2.Rows.Count - 1;
                dataGridView2[0, iX].Value = Int32.Parse(dataGridView3[0, r.Index].Value.ToString());
                dataGridView2[2, iX].Value = 0;
                iX++;
            }

        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int iX = 0;

            txtAreaID.Text = dataGridView2[0, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtAreaName.Text = dataGridView2[1, dataGridView2.SelectedRows[iX].Index].Value.ToString();
            txtHaihuMaisu.Text  = dataGridView2[2, dataGridView2.SelectedRows[iX].Index].Value.ToString();

            //�ҏW���̓\�[�g�֎~
            dataGridView2.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;

            aMode.Mode = 1;
            aMode.RowIndex = dataGridView2.SelectedRows[iX].Index;
            txtAdel.Enabled = true;
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            txtAreaID.Text = dataGridView3[0, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtAreaName.Text = dataGridView3[1, dataGridView3.SelectedRows[0].Index].Value.ToString();
            txtHaihuMaisu.Focus();
            aMode.Mode = 0;
        }

        private int GetMaisuTotal()
        {
            int Total = 0;

            for (int i = 0; i <= dataGridView2.Rows.Count - 1; i++)
            {
                Total += Int32.Parse(dataGridView2[2, i].Value.ToString(), System.Globalization.NumberStyles.AllowThousands);
            }

            return Total;
        }

        private void txtAdel_Click(object sender, EventArgs e)
        {
            //�s�폜
            if (MessageBox.Show(dataGridView2[1,aMode.RowIndex].Value.ToString() + " ���폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            dataGridView2.Rows.RemoveAt(aMode.RowIndex);
        }

        private void txtAclear_Click(object sender, EventArgs e)
        {
            textAreaClear();
        }

    }
}
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
using System.Linq;

namespace posting
{
    public partial class frmOrder: Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Utility.����ŗ� cTax = new Utility.����ŗ�();
        Entity.�� cMaster = new Entity.��();

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.��1TableAdapter jAdp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.���Ӑ�TableAdapter cAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
        darwinDataSetTableAdapters.�Ј�TableAdapter sAdp = new darwinDataSetTableAdapters.�Ј�TableAdapter();
        darwinDataSetTableAdapters.�V������TableAdapter rAdp = new darwinDataSetTableAdapters.�V������TableAdapter();
        darwinDataSetTableAdapters.�z�z�G���ATableAdapter hAdp = new darwinDataSetTableAdapters.�z�z�G���ATableAdapter();

        // 2016/07/07
        darwinDataSetTableAdapters.�󒍕ҏW����TableAdapter lAdp = new darwinDataSetTableAdapters.�󒍕ҏW����TableAdapter();

        const string MESSAGE_CAPTION = "�󒍊m�菑";
        const string J_NAIYOU_POSTING = "�|�X�e�B���O";

        Int64 _orderNum = 0;

        int sLoginTag = 0;                    // �󒍊m�菑�ҏW�ݒ�@�ҏW�\���O�C���^�O
        DateTime dtLock = DateTime.Today;     // ���������s������
        bool orderEditStatus = true;          // �󒍊m�菑�ҏW�ݒ�

        Utility.comboGaichuUser[] cg = null;

        public frmOrder(Int64 orderNum)
        {
            InitializeComponent();

            //// �󒍃f�[�^�ǂݍ���
            //jAdp.Fill(dts.��1);

            // ���Ӑ�f�[�^�ǂݍ���
            cAdp.Fill(dts.���Ӑ�);

            // �Ј��f�[�^�ǂݍ���
            sAdp.Fill(dts.�Ј�);

            // 2016/07/07
            lAdp.Fill(dts.�󒍕ҏW����);

            // 2016/07/19
            rAdp.Fill(dts.�V������);

            //// 2017/01/27
            //hAdp.Fill(dts.�z�z�G���A);

            // �w��󒍔ԍ�
            _orderNum = orderNum;
        }

        private void form_Load(object sender, EventArgs e)
        {
            // 2018/01/05
            txtTanka.AutoSize = false;
            txtTanka.Height = 25;

            txtMai.AutoSize = false;
            txtMai.Height = 25;

            txtNebikigo.AutoSize = false;
            txtNebikigo.Height = 24;

            txtZeikomi.AutoSize = false;
            txtZeikomi.Height = 24;

            txtTax.AutoSize = false;
            txtTax.Height = 24;

            cmbeGaichu.AutoSize = false;
            cmbeGaichu.Height = 24;
            
            // �E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.��' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă��������B
            GridviewSet.Setting(dataGridView1);
            //this.darwinDataSet.Clear();
            //this.darwinDataSet.EnforceConstraints = false;
            //this.��TableAdapter.Fill(this.darwinDataSet.��);

            //////�󒍋敪�R���{
            ////cmbJkbnSet();

            // �O����R���{�{�b�N�X
            Utility.comboGaichu.itemLoad(cmbeGaichu);   // �c�Ɨp
            Utility.comboGaichu.itemLoad(cmbpGaichu);   // �x���p
            Utility.comboGaichu.itemLoad(cmbpGaichu2);  // �x���p2 : 2016/10/14
            Utility.comboGaichu.itemLoad(cmbpGaichu3);  // �x���p3 : 2016/10/14

            cg = Utility.comboGaichu.getArrayGaichu();  // 2018/01/03

            // ���Ə��R���{
            Utility.ComboOffice.load(cmbOffice);

            // ���Ӑ�R���{
            Utility.ComboClient.itemsLoad(cmbClient);

            // �󒍎�ʃR���{
            Utility.ComboJshubetsu.load(cmbNaiyou);
            Utility.ComboJshubetsu.load(cmbsNaiyou);

            // �z�z�`�ԃR���{
            Utility.ComboFkeitai.load(cmbFkeitai);

            // �z�z�����R���{
            cmbFjyoukenSet();

            // ���^�R���{
            Utility.ComboSize.load(cmbSize);
            Utility.ComboSize.load(cmbsSize);   // �����T�C�Y 2019/02/14

            // �Č���ʃR���{
            Utility.ComboAnshu.load(cmbAnShu);

            //�z�z�P�\�R���{
            //cmbFyuyoSet();    // 2015/06/23

            //�[�i�`�ԃR���{
            //cmbNkeitaiSet();  // 2015/06/23

            ////////�������@�R���{
            //////Utility.ComboShimebi.load(cmbNyukin);

            //�U�������R���{
            //Utility.ComboFuri.load(cmbFuri);

            //�񍐎����R���{
            //cmbHjikiSet();    // 2015/06/23

            //�񍐐��x�R���{
            //cmbHseidoSet();   // 2015/06/23

            //�񍐕��@�R���{
            //cmbHhouhouSet();  // 2015/06/23

            // �ŗ��擾
            cTax.Ritsu = GetTaxRT(DateTime.Today);

            // ��ʏ�����
            DispClear();

            sDt.Value = DateTime.Today.AddYears(-1);
            eDt.Value = DateTime.Today;

            // �󒍔ԍ��w��̂Ƃ�
            if (_orderNum != 0)
            {
                // �t�H�[�����[�h�X�e�[�^�X/��ID:�ύX�폜
                fMode.Mode = 1;
                fMode.jID = _orderNum;
                cMaster.ID = fMode.jID;

                // �f�[�^�\��
                dataShow();
            }

            // �󒍊m�菑�ҏW�ݒ���擾
            sLoginTag = getOrderLock(out dtLock);            
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     �󒍊m�菑�ҏW�ݒ���擾 </summary>
        /// <param name="dt">
        ///     ���������s��</param>
        /// <returns>
        ///     ���O�C���^�O</returns>
        ///-----------------------------------------------------------
        private int getOrderLock(out DateTime dt)
        {
            int sTag = 0;
            dt = DateTime.Parse("2999/01/01");

            if (dts.�󒍕ҏW����.Any(a => a.ID == global.lockKey))
            {
                var s = dts.�󒍕ҏW����.Single(a => a.ID == global.lockKey);
                sTag = s.���O�C���O���[�v;
                dt = s.���������s��;
            }

            return sTag;
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
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    tempDGV.EnableHeadersVisualStyles = false;

                    // ��w�b�_�[�\���ʒu�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // ��w�b�_�[�t�H���g�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 144;

                    // ��s�̐F
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns[0].Width = 110;
                    tempDGV.Columns[1].Width = 90;
                    tempDGV.Columns[2].Width = 200;
                    tempDGV.Columns[3].Width = 200;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 140;
                    tempDGV.Columns[11].Width = 140;
                    tempDGV.Columns[12].Width = 140;

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
                    
                    // �r��
                    tempDGV.BorderStyle = BorderStyle.Fixed3D;
                    tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ///--------------------------------------------------------------------------
            /// <summary>
            ///     �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾���� </summary>
            /// <param name="dgv">
            ///     �ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            ///--------------------------------------------------------------------------
            public static Boolean GetData(DataGridView dgv, ref Entity.�� tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.�� Order = new Control.��();
                OleDbDataReader dr;

                sqlStr = " where ��.ID = " + (long)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = Order.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = long.Parse(dr["ID"].ToString());
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
                        tempC.�z�z�P�� = double.Parse(dr["�z�z�P��"].ToString());
                        tempC.�˗��� = dr["�˗���"].ToString();
                        tempC.���� = Convert.ToDouble(dr["����"].ToString());
                        tempC.�z�z�`�� = Int32.Parse(dr["�z�z�`��"].ToString());
                        tempC.�z�z���� = dr["�z�z����"].ToString() + "";
                        tempC.�z�z�J�n�� = dr["�z�z�J�n��"].ToString();
                        tempC.�z�z�I���� = dr["�z�z�I����"].ToString();
                        //tempC.�z�z�P�\ = dr["�z�z�P�\"].ToString() + "";  // 2015/06/23
                        tempC.�[�i�\��� = dr["�[�i�\���"].ToString();
                        //tempC.�[�i�`�� = dr["�[�i�`��"].ToString() + "";  // 2015/06/23
                        tempC.������ = Int32.Parse(dr["������"].ToString());
                        tempC.������ID = int.Parse(dr["������ID"].ToString());
                        tempC.���������s�� = dr["���������s��"].ToString();
                        tempC.�������@ = dr["�������@"].ToString() + "";
                        tempC.�����\��� = dr["�����\���"].ToString();
                        //tempC.�񍐎��� = dr["�񍐎���"].ToString() + "";  // 2015/06/23
                        //tempC.�񍐐��x = dr["�񍐐��x"].ToString() + "";  // 2015/06/23
                        //tempC.�񍐕��@ = dr["�񍐕��@"].ToString() + "";  // 2015/06/23
                        //tempC.���[���A�h���X = dr["���[���A�h���X"].ToString() + "";    // 2015/06/23
                        tempC.�U������ID = Int32.Parse(dr["�U������ID"].ToString());
                        tempC.���z�z���L�� = Int32.Parse(dr["���z�z���L��"].ToString());
                        tempC.�}�ԗL�� = Int32.Parse(dr["�}�ԗL��"].ToString());
                        tempC.���L���� = dr["���L����"].ToString() + "";
                        tempC.�G���A���l = dr["�G���A���l"].ToString() + "";
                        tempC.�����敪 = int.Parse(dr["�����敪"].ToString());
                        tempC.���z���O = Int32.Parse(dr["���z���O"].ToString());
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
            msgStr += dataGridView1[3, dataGridView1.SelectedRows[iX].Index].Value +"���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "�󒍊m�菑�I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    // �t�H�[�����[�h�X�e�[�^�X/��ID:�ύX�폜
                    fMode.Mode = 1;
                    fMode.jID = (long)dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value;
                    cMaster.ID = fMode.jID;

                    // ���͊J�n�\���t�F2017/01/27
                    jDate.MinDate = DateTime.Parse("1900/01/01");
                    dtSeikyu.MinDate = DateTime.Parse("1900/01/01");
                    NyukinDate.MinDate = DateTime.Parse("1900/01/01");

                    // �f�[�^�\��
                    dataShow();

                    //// �t�H�[�����[�h�X�e�[�^�X/��ID:�ύX�폜
                    //fMode.Mode = 1; 
                    //fMode.jID = (long)dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value;
                    //cMaster.ID = fMode.jID;

                    //// �f�[�^���擾����
                    //if (!GetData(fMode.jID))
                    //{
                    //    MessageBox.Show("�Y������f�[�^���o�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    //    return;
                    //}

                    //// �{�^�����
                    //btnDel.Enabled = true;
                    //btnClr.Enabled = true;
                    //button2.Enabled = true;
                    //button3.Enabled = true;

                    //// �󒍔ԍ� 2015/07/13
                    //label45.Visible = false;
                    //groupBox3.Visible = false;

                    //jDate.Focus();
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "�f�[�^�\��", MessageBoxButtons.OK);
                }
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�\�� </summary>
        ///------------------------------------------------------------
        private void dataShow()
        {
            int hCnt = 0;

            // �f�[�^���擾����
            if (!GetData(fMode.jID, ref hCnt))
            {
                MessageBox.Show("�Y������f�[�^���o�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            // �{�^�����
            if (orderEditStatus)
            {
                btnUpdate.Enabled = true;
                btnDel.Enabled = true;
                button2.Enabled = true;
                label52.Visible = false;
            }
            else
            {
                btnUpdate.Enabled = false;
                btnDel.Enabled = false;
                button2.Enabled = false;
                label52.Visible = true;
            }

            // �z�z�G���A�o�^�ςݎ󒍃f�[�^�̂Ƃ��폜�s�Ƃ���F2017/01/27
            if (hCnt > 0)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }

            btnClr.Enabled = true;
            button3.Enabled = true;

            // �󒍔ԍ� 2015/07/13
            label45.Visible = false;
            groupBox3.Visible = false;

            // �󒍃f�[�^�R�s�[�����N�{�^���F2018/01/05
            lnkCopy.Visible = true;
            
            jDate.Focus();
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾���� </summary>
        /// <param name="dgv">
        ///     �ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        /// -----------------------------------------------------------------------------
        private Boolean GetData(long sID, ref int hCnt)
        {
            // �󒍃f�[�^�ǂݍ���
            DateTime sd = DateTime.Parse(sDt.Value.ToShortDateString());
            DateTime ed = DateTime.Parse(eDt.Value.ToShortDateString());
            jAdp.FillByFromYMDToYMD(dts.��1, sd, ed);

            //jAdp.Fill(dts.��1);

            foreach (var t in dts.��1.Where(a => a.ID == sID))
            {
                cmbNaiyou.SelectedValue = t.�󒍎��ID;
                jDate.Value = DateTime.Parse(t.�󒍓�.ToShortDateString());
                cmbOffice.SelectedValue = t.���Ə�ID;

                cmbClient.SelectedValue = t.���Ӑ�ID;
                txtCzipcode.Text = "";
                txtName2.Text = "";
                txtCbusho.Text = "";
                txtTantou.Text = "";
                txtCtel.Text = "";
                txtCfax.Text = "";
                txtCbusho.Text = "";
                txtCtantou.Text = "";

                //�N���C�A���g���\��
                ClientShow(t.���Ӑ�ID);

                txtChirashi.Text = t.�`���V��;

                txtUri.Text = t.���z.ToString("#,##0");  // �P���A��������ɕ\�� 2019/07/30
                txtTanka.Text = t.�P��.ToString("#,##0.00");
                txtMai.Text = t.����.ToString("#,##0");
                txtNebiki.Text = t.�l���z.ToString("#,##0");

                // �l������z
                int nebikigoKin = Utility.strToInt(txtUri.Text) - Utility.strToInt(txtNebiki.Text);
                txtNebikigo.Text = nebikigoKin.ToString("#,##0");

                txtTax.Text = t.�����.ToString("#,##0");
                txtZeikomi.Text = t.�ō����z.ToString("#,##0");

                cmbFkeitai.SelectedValue = t.�z�z�`��;
                cmbFjyouken.Text = t.�z�z����;
                cmbSize.SelectedValue = t.���^;
                txtHTanka.Text = t.�z�z�P��.ToString("#,##0.00");

                if (t.Is�z�z�J�n��Null())
                {
                    StartDate.Checked = false;
                }
                else
                {
                    StartDate.Checked = true;
                    StartDate.Value = DateTime.Parse(t.�z�z�J�n��.ToShortDateString());
                }

                if (t.Is�z�z�I����Null())
                {
                    EndDate.Checked = false;
                }
                else
                {
                    EndDate.Checked = true;
                    EndDate.Value = DateTime.Parse(t.�z�z�I����.ToShortDateString());
                }

                if (t.Is�[�i�\���Null())
                {
                    NouhinDate.Checked = false;
                }
                else
                {
                    NouhinDate.Checked = true;
                    NouhinDate.Value = DateTime.Parse(t.�[�i�\���.ToShortDateString());
                }
                
                if (t.Is�����\���Null())
                {
                    NyukinDate.Checked = false;
                }
                else
                {
                    NyukinDate.Checked = true;
                    NyukinDate.Value = DateTime.Parse(t.�����\���.ToShortDateString());
                }

                txtMemo.Text = t.���L����;
                //txtMemo2.Text = t.�G���A���l;

                txtSalesMemo.Text = t.�c�Ɣ��l;     // 2019/03/01

                if (int.Parse(t.���z�z���L��.ToString()) == 1)
                {
                    checkBox2.Checked = true;
                }
                else
                {
                    checkBox2.Checked = false;
                }

                if (int.Parse(t.�}�ԗL��.ToString()) == 1)
                {
                    checkBox3.Checked = true;
                }
                else
                {
                    checkBox3.Checked = false;
                }

                // �������� 2010/05/30
                txtSeikyuNumber.Text = t.������ID.ToString();

                // ���z���O 2014/11/26
                if (Utility.nullToInt(t.���z���O) == 1)
                {
                    chkHeihai.Checked = true;
                }
                else
                {
                    chkHeihai.Checked = false;
                }

                //���������s�� 2015/06/30
                if (t.Is���������s��Null())
                {
                    dtSeikyu.Checked = false;
                }
                else
                {
                    dtSeikyu.Checked = true;
                    dtSeikyu.Value = DateTime.Parse(t.���������s��.ToShortDateString());
                }

                // �Č���� 2015/06/30
                cmbAnShu.SelectedValue = t.�Č����;

                // �O����E�c�Ɨp 2015/06/30
                cmbeGaichu.SelectedValue = t.�O����ID�c��;

                // �O���x�����E�c�� 2015/06/30
                if (t.Is�O���x�����c��Null())
                {
                    dteGaichuPay.Checked = false;
                }
                else
                {
                    dteGaichuPay.Checked = true;
                    dteGaichuPay.Value = DateTime.Parse(t.�O���x�����c��.ToShortDateString());
                }

                // �O�������E�c�� 2015/06/30
                txteGaichuGenka.Text = t.�O�������c��.ToString("#,##0");

                // �O����E�x���p 2015/06/30
                cmbpGaichu.SelectedValue = t.�O����ID�x��;
                

                // �O���x�����E�x�� 2015/06/30
                if (t.Is�O���x�����x��Null())
                {
                    dtpGaichuPay.Checked = false;
                }
                else
                {
                    dtpGaichuPay.Checked = true;
                    dtpGaichuPay.Value = DateTime.Parse(t.�O���x�����x��.ToShortDateString());
                }

                // �O�������E�x�� 2015/06/30
                txtpGaichuGenka.Text = t.�O�������x��.ToString("#,##0");

                // �O����E�x���p2 2016/10/14
                cmbpGaichu2.SelectedValue = t.�O����ID�x��2;

                // �O���x�����E�x��2 2016/10/14
                if (t.Is�O���x�����x��2Null())
                {
                    dtpGaichuPay2.Checked = false;
                }
                else
                {
                    dtpGaichuPay2.Checked = true;
                    dtpGaichuPay2.Value = DateTime.Parse(t.�O���x�����x��2.ToShortDateString());
                }

                // �O�������E�x��2 2016/10/14
                txtpGaichuGenka2.Text = t.�O�������x��2.ToString("#,##0");

                // �O����E�x���p3 2016/10/14
                cmbpGaichu3.SelectedValue = t.�O����ID�x��3;

                // �O���x�����E�x��3 2016/10/14
                if (t.Is�O���x�����x��3Null())
                {
                    dtpGaichuPay3.Checked = false;
                }
                else
                {
                    dtpGaichuPay3.Checked = true;
                    dtpGaichuPay3.Value = DateTime.Parse(t.�O���x�����x��3.ToShortDateString());
                }

                // �O�������E�x��3 2016/10/14
                txtpGaichuGenka3.Text = t.�O�������x��3.ToString("#,##0");

                // �O���˗��� 2015/08/11
                if (t.Is�O���˗����x��Null())
                {
                    dtGaichuIrai.Checked = false;
                }
                else
                {
                    dtGaichuIrai.Checked = true;
                    dtGaichuIrai.Value = DateTime.Parse(t.�O���˗����x��.ToShortDateString());
                }

                // 2016/10/15
                if (t.Is�O���˗����x��2Null())
                {
                    dtGaichuIrai2.Checked = false;
                }
                else
                {
                    dtGaichuIrai2.Checked = true;
                    dtGaichuIrai2.Value = DateTime.Parse(t.�O���˗����x��2.ToShortDateString());
                }

                if (t.Is�O���˗����x��3Null())
                {
                    dtGaichuIrai3.Checked = false;
                }
                else
                {
                    dtGaichuIrai3.Checked = true;
                    dtGaichuIrai3.Value = DateTime.Parse(t.�O���˗����x��3.ToShortDateString());
                }

                // �O���n���� 2015/08/11
                if (t.Is�O���n����Null())
                {
                    dtGaichuWatashi.Checked = false;
                }
                else
                {
                    dtGaichuWatashi.Checked = true;
                    dtGaichuWatashi.Value = DateTime.Parse(t.�O���n����.ToShortDateString());
                }

                // 2016/10/15
                if (t.Is�O���n����2Null())
                {
                    dtGaichuWatashi2.Checked = false;
                }
                else
                {
                    dtGaichuWatashi2.Checked = true;
                    dtGaichuWatashi2.Value = DateTime.Parse(t.�O���n����2.ToShortDateString());
                }

                if (t.Is�O���n����3Null())
                {
                    dtGaichuWatashi3.Checked = false;
                }
                else
                {
                    dtGaichuWatashi3.Checked = true;
                    dtGaichuWatashi3.Value = DateTime.Parse(t.�O���n����3.ToShortDateString());
                }

                // �O���󂯓n���S���� 2015/08/11
                txtGaichuUkeTan.Text = t.�O���󂯓n���S����;
                txtGaichuUkeTan2.Text = t.�O���󂯓n���S����2;    // 2016/10/15
                txtGaichuUkeTan3.Text = t.�O���󂯓n���S����3;    // 2016/10/15

                // �O���ϑ�����
                txtGaichuMaisu.Text = t.�O���ϑ�����.ToString();
                txtGaichuMaisu2.Text = t.�O���ϑ�����2.ToString();    // 2016/10/15
                txtGaichuMaisu3.Text = t.�O���ϑ�����3.ToString();    // 2016/10/15

                // �Ǝ�
                if (t.�Ǝ� != null)
                {
                    txtGyoushu.Text = t.�Ǝ�;
                }
                else
                {
                    txtGyoushu.Text = string.Empty;
                }

                // ���͊����̎󒍊m�菑�͕ҏW���b�N���� 2016/07/19
                if (t.�V������Row != null && t.�V������Row.�������� == global.FLGON)
                {
                    orderEditStatus = false;
                    label52.Text = "���ɓ������������Ă��邽�ߕҏW�͏o���܂���";
                }
                else if (sLoginTag != global.FLGOFF)
                {
                    // �󒍊m�菑�ҏW���������� 2016/07/07
                    if (global.loginType == sLoginTag)
                    {
                        // �ҏW�\�ȃ��O�C���^�C�v
                        orderEditStatus = true;
                    }
                    else
                    {
                        if (!t.Is���������s��Null() && (t.���������s�� <= dtLock))
                        {
                            // ���������s�����ҏW���b�N���Ԓ�
                            orderEditStatus = false;
                            label52.Text = "���������s���� " + dtLock.ToShortDateString() + " �ȑO�̂��߂��̎󒍊m�菑�̕ҏW�͐�������Ă��܂�";
                        }
                        else
                        {
                            // ���������s�����ҏW���b�N���Ԉȍ~
                            orderEditStatus = true;
                        }
                    }

                    //// �z�z�G���A��荞�ݍς݂��H�@2017/01/27
                    //hCnt = t.Get�z�z�G���ARows().Count();


                    // �z�z�G���A��荞�ݍς݂��H�@2018/01/27
                    Control.DataControl Con = new Control.DataControl();
                    OleDbConnection cn = new OleDbConnection();

                    cn = Con.GetConnection();
                    OleDbCommand SCom = new OleDbCommand();

                    SCom.Connection = cn;

                    //�z�z�G���A��荞�ݍς݂��H
                    StringBuilder sb = new StringBuilder();
                    sb.Append("Select count(*) as cnt from �z�z�G���A ");
                    sb.Append("where ��ID = ").Append(t.ID);

                    SCom.CommandText = sb.ToString();
                    OleDbDataReader dr = SCom.ExecuteReader();

                    string cc = string.Empty;
                    while (dr.Read())
                    {
                        cc = dr["cnt"].ToString();
                    }

                    dr.Close();
                    SCom.Connection.Close();

                    hCnt = Utility.strToInt(cc);

                }
                else
                {
                    orderEditStatus = true;
                }

                // �m���ȑe�����z�F2019/07/30
                txteGaichuArari.Text = (Utility.strToInt(txtUri.Text) - Utility.strToInt(txteGaichuGenka.Text)).ToString("#,##0");

            }

            return true;
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

        ///----------------------------------------------------
        /// <summary>
        ///     ��ʂ��N���A���� </summary>
        ///----------------------------------------------------
        private void DispClear()
        {
            try
            {
                // �������� : 2015/06/24
                //sDt.Checked = false;  2018/01/27
                //eDt.Checked = false;  2018/01/27
                textBox1.Text = string.Empty;
                cmbsNaiyou.SelectedIndex = -1;
                txtsClient.Text = string.Empty;
                txtsGaichu.Text = string.Empty;
                txtsOrderNum.Text = string.Empty;

                // �������ڒǉ� 2019/02/14
                txtsMaisu.Text = string.Empty;  // ����
                cmbsSize.SelectedIndex = -1;    // �T�C�Y
                seDt.Checked = false;           // ��������

                // �\������
                fMode.Mode = 0;

                // �V�K���͎��ɂ̂ݓ��͓��t�����F2017/08/28
                DateTime minDate = DateTime.Today.AddMonths(-1);
                if (_orderNum == global.FLGOFF)
                {
                    // ���͓��t���� 2016/01/27
                    minDate = DateTime.Parse(minDate.Year.ToString() + "/" + minDate.Month.ToString() + "/01");

                    // �󒍓����͉\�J�n���F2017/01/27
                    jDate.Value = DateTime.Today;
                    jDate.MinDate = minDate;
                }

                //////comboBox1.SelectedIndex = -1;
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
                textBox4.Text = "";

                txtChirashi.Text = "";
                cmbNaiyou.SelectedIndex = -1;
                txtTanka.Text = "0";
                txtMai.Text = "0";
                txtUri.Text = "0";
                txtTax.Text = "0";
                txtZeikomi.Text = "0";
                txtNebiki.Text = "0";
                txtNebikigo.Text = "0";

                cmbFkeitai.SelectedIndex = -1;
                cmbFjyouken.SelectedIndex = -1;
                cmbSize.SelectedIndex = -1;
                //cmbSize.Text = string.Empty;

                txtHTanka.Text = "0";

                //txtIraisaki.Text = "";
                //txtGenka.Text = "0";

                StartDate.Checked = false;
                EndDate.Checked = false;

                //cmbFyuyo.SelectedIndex = -1;      // 2015/06/23
                //checkBox1.Checked = false;

                NouhinDate.Checked = false;
                NouhinDate.Enabled = true;  // 2019/02/14

                //cmbNkeitai.SelectedIndex = -1;    // 2015/06/23

                //////cmbNyukin.SelectedIndex = -1;

                NyukinDate.Checked = false;

                // �V�K���͎��ɂ̂ݓ��͓��t�����F2017/08/28
                if (_orderNum == global.FLGOFF)
                {
                    NyukinDate.MinDate = minDate;   // �x���������͊J�n�\���F2017/01/27
                }

                //cmbFuri.SelectedIndex = -1;

                //cmbHjiki.SelectedIndex = -1;      // 2015/06/23
                //cmbHseido.SelectedIndex = -1;     // 2015/06/23
                //cmbHhouhou.SelectedIndex = -1;    // 2015/06/23

                //txtEmail.Text = "";               // 2015/06/23

                txtMemo.Text = "";
                txtSalesMemo.Text = "";             // 2019/03/01
                //txtMemo2.Text = "";
                checkBox2.Checked = false;
                checkBox3.Checked = false;

                txtSeikyuNumber.Text = "0";             // 2010/05/30

                chkHeihai.Checked = false;              // 2014/11/26 ���z�`�F�b�N

                dtSeikyu.Checked = false;               // ���������s�� 2015/06/30

                // �V�K���͎��ɂ̂ݓ��͓��t�����F2017/08/28
                if (_orderNum == global.FLGOFF)
                {
                    dtSeikyu.MinDate = minDate;         // ���������s�����͉\�J�n���F2017/01/27
                }

                cmbAnShu.SelectedIndex = -1;            // �Č���ʃR���{�{�b�N�X 2015/06/30

                cmbeGaichu.SelectedIndex = -1;          // �c�ƁE�O���� 2015/06/30
                dteGaichuPay.Checked = false;           // �c�ƁE�O���x���� 2015/06/30
                txteGaichuGenka.Text = string.Empty;    // �c�ƁE�O������ 2015/06/30

                cmbpGaichu.SelectedIndex = -1;          // �x���E�O���� 2015/06/30
                dtpGaichuPay.Checked = false;           // �x���E�O���x���� 2015/06/30
                txtpGaichuGenka.Text = string.Empty;    // �x���E�O������ 2015/06/30
                lblSD1.Text = string.Empty;             // �x���� 2018/01/04

                cmbpGaichu2.SelectedIndex = -1;         // �x���E�O����2 2016/10/14
                dtpGaichuPay2.Checked = false;          // �x���E�O���x����2 2016/10/14
                txtpGaichuGenka2.Text = string.Empty;   // �x���E�O������2 2016/10/14
                lblSD2.Text = string.Empty;             // �x���� 2018/01/04

                cmbpGaichu3.SelectedIndex = -1;         // �x���E�O����3 2016/10/14
                dtpGaichuPay3.Checked = false;          // �x���E�O���x����3 2016/10/14
                txtpGaichuGenka3.Text = string.Empty;   // �x���E�O������3 2016/10/14
                lblSD3.Text = string.Empty;             // �x���� 2018/01/04

                dtGaichuIrai.Checked = false;           // �O���˗��� 2015/08/11
                dtGaichuWatashi.Checked = false;        // �O���n���� 2015/08/11
                txtGaichuUkeTan.Text = string.Empty;    // �O���󂯓n���S���� 2015/08/11
                txtGaichuMaisu.Text = string.Empty;     // �O���ɓn�������� 2015/09/20

                dtGaichuIrai2.Checked = false;          // �O���˗��� 2016/10/15
                dtGaichuWatashi2.Checked = false;       // �O���n���� 2016/10/15
                txtGaichuUkeTan2.Text = string.Empty;   // �O���󂯓n���S���� 2016/10/15
                txtGaichuMaisu2.Text = string.Empty;    // �O���ɓn�������� 2016/10/15

                dtGaichuIrai3.Checked = false;          // �O���˗��� 2016/10/15
                dtGaichuWatashi3.Checked = false;       // �O���n���� 2016/10/15
                txtGaichuUkeTan3.Text = string.Empty;   // �O���󂯓n���S���� 2016/10/15
                txtGaichuMaisu3.Text = string.Empty;    // �O���ɓn�������� 2016/10/15

                txtGyoushu.Text = string.Empty;         // �Ǝ�            2015/09/20

                tabControl1.SelectTab(0);
                tabControl1.TabPages[3].Text = "�e��";    // 2019/02/21

                // 2019/02/21
                txtaUri.Text = string.Empty;
                txtaGai1.Text = string.Empty;
                txtaGai2.Text = string.Empty;
                txtaGai3.Text = string.Empty;
                txtaGaiGenka1.Text = string.Empty;
                txtaGaiGenka2.Text = string.Empty;
                txtaGaiGenka3.Text = string.Empty;
                txtaArari.Text = string.Empty;

                btnUpdate.Enabled = true;
                btnDel.Enabled = false;
                btnClr.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;

                // �󒍔ԍ� 2015/07/13
                label45.Visible = true;
                rBtnOrderNum1.Checked = true;
                groupBox3.Visible = true;
                lblOrderNum.Text = string.Empty;
                
                //txtCode.Focus();
                jDate.Focus();

                // 2016/07/07
                label52.Visible = false;

                lblClientShimebi.Text = string.Empty;   // 2018/01/04
                lnkCopy.Visible = false;    // 2018/01/05
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
                    Control.�� Order = new Control.��();

                    switch (fMode.Mode)
                    {
                        case 0: //�V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                Order.Close();
                                return;
                            }

                            // �ȉ��A2015/07/15 Utility.getOrderNum()�ɒu������

                            ////�󒍓��擾
                            //DateTime sJDate = jDate.Value;

                            //ID���̔�
                            //string sqlStr = "";
                            //long gID = 0;
                            //long sWk;
                            //DateTime sJDate;
                            
                            ////�`�[�ԍ��ŏ��l
                            //sWk = sJDate.Year;
                            //cMaster.IDTemplateS = sWk * 100000000;

                            //sWk = sJDate.Month;
                            //cMaster.IDTemplateS += sWk * 1000000;

                            //sWk = sJDate.Day;
                            //cMaster.IDTemplateS += sWk * 10000;

                            //cMaster.IDTemplateS++;

                            ////�`�[�ԍ��ő�l
                            //cMaster.IDTemplateE = cMaster.IDTemplateS + 9998;

                            ////�󒍓��̓`�[�����邩�H
                            //sqlStr += "select max(ID) as ID from �� ";
                            //sqlStr += "where (ID >= " + cMaster.IDTemplateS.ToString() + ") and ";
                            //sqlStr += "(ID <= " + cMaster.IDTemplateE.ToString() + ")";

                            //OleDbDataReader dR;
                            //Control.FreeSql fCon = new Control.FreeSql();
                            //dR = fCon.free_dsReader(sqlStr);

                            //while (dR.Read())
                            //{
                            //    if (dR["ID"] == DBNull.Value)
                            //    {
                            //        gID = cMaster.IDTemplateS;�@//�Ȃ���΃e���v���[�g�̓`�[�ԍ��ŏ��l���Z�b�g
                            //    }
                            //    else
                            //    {
                            //        gID = long.Parse(dR["ID"].ToString()) + 1;�@//�����1�C���N�������g
                            //    }
                            //}

                            //dR.Close();
                            //fCon.Close();

                            // �󒍔ԍ��̔� 2015/07/15
                            if (rBtnOrderNum1.Checked)
                            {
                                // �����̔ԂŎ󒍔ԍ����擾
                                cMaster.ID = Utility.getOrderNum(jDate.Value);
                            }
                            else if (rBtnOrderNum2.Checked)
                            {
                                // �擾�ςݔԍ��ƕR�t��
                                cMaster.ID = long.Parse(lblOrderNum.Text);
                            }

                            //������ID
                            //cMaster.������ID = 0;

                            if (Order.DataInsert(cMaster) == true)
                            {
                                if (rBtnOrderNum2.Checked)
                                {
                                    // �󒍔ԍ��̔ԃf�[�^�X�V
                                    orderNumUpdate(cMaster.ID, cMaster.���Ӑ�ID);
                                }

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
                                Order.Close();
                                return;
                            }

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
                    //this.��TableAdapter.Fill(this.darwinDataSet.��);
                    //dataGridView1.DataSource = this.darwinDataSet.��;
                    dataGridView1.DataSource = null;

                    // �󒍃f�[�^�Z�b�g�ǂݍ���
                    //jAdp.Fill(dts.��1);

                    DateTime sd = DateTime.Parse(sDt.Value.ToShortDateString()); // 2018/01/27
                    DateTime ed = DateTime.Parse(eDt.Value.ToShortDateString()); // 2018/01/27
                    jAdp.FillByFromYMDToYMD(dts.��1, sd, ed); // 2018/01/27
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"�X�V����",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     �󒍔ԍ��̔ԕR�t���X�V </summary>
        /// <param name="jID">
        ///     �󒍔ԍ�</param>
        /// -------------------------------------------------------------------
        private void orderNumUpdate(long jID, int clientID)
        {
            darwinDataSetTableAdapters.�󒍔ԍ��̔�TableAdapter adp = new darwinDataSetTableAdapters.�󒍔ԍ��̔�TableAdapter();
            adp.Fill(dts.�󒍔ԍ��̔�);

            // �󒍔ԍ��̔ԕR�t���X�V
            if (dts.�󒍔ԍ��̔�.Any(a => a.�󒍔ԍ� == jID))
            {
                var s = dts.�󒍔ԍ��̔�.Single(a => a.�󒍔ԍ� == jID);
                s.���Ӑ�ID = clientID;
                s.�m�菑���� = 1;
                s.�m�菑���͓��t = DateTime.Now;
                s.�m�菑���̓��[�U�[ID = global.loginUserID;
                s.�X�V�N���� = DateTime.Now;

                // �f�[�^�x�[�X�X�V
                adp.Update(dts.�󒍔ԍ��̔�);
            }
        }


        // �o�^�f�[�^�`�F�b�N
        private Boolean fDataCheck()
        {
            string str;
            double d;
            int dInt;

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

                    // �R�t���󒍔ԍ��̂Ƃ�
                    if (rBtnOrderNum2.Checked && lblOrderNum.Text == string.Empty)
                    {
                        btnOrderNum.Focus();
                        throw new Exception("�擾�ς݂̎󒍔ԍ���I�����Ă�������");
                    }

                }

                //�󒍋敪�`�F�b�N
                //////if (comboBox1.SelectedIndex == -1)
                //////{
                //////    comboBox1.Focus();
                //////    throw new Exception("�󒍋敪��I�����Ă�������");
                //////}

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
                else
                {
                    // �󒍓��e�ƈČ���� 2015/06/30
                    if ((cmbNaiyou.SelectedIndex == 0 && cmbAnShu.SelectedIndex == 2) || 
                        (cmbNaiyou.SelectedIndex > 0 && cmbAnShu.SelectedIndex < 2))
                    {
                        cmbNaiyou.Focus();
                        throw new Exception("�󒍓��e�ƈČ���ʂ���v���Ă��܂���");
                    }
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

                if (int.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out dInt))
                {
                }
                else
                {
                    this.txtMai.Focus();
                    throw new Exception("����������������܂���");
                }

                ////////�����F�������H
                //////if (txtGenka.Text == null)
                //////{
                //////    this.txtGenka.Focus();
                //////    throw new Exception("�����͐����œ��͂��Ă�������");
                //////}

                //////str = this.txtGenka.Text;

                //////if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                //////{
                //////}
                //////else
                //////{
                //////    this.txtGenka.Focus();
                //////    throw new Exception("�����͐����œ��͂��Ă�������");
                //////}

                //�|�X�e�B���O�̂Ƃ��̂݃`�F�b�N�ΏۂƂ��鍀��
                if (cmbNaiyou.Text == J_NAIYOU_POSTING)
                {
                    //�z�z�`�ԃ`�F�b�N
                    if (cmbFkeitai.SelectedIndex == -1)
                    {
                        cmbFkeitai.Focus();
                        throw new Exception("�z�z�`�Ԃ�I�����Ă�������");
                    }

                    //�z�z�����`�F�b�N
                    if (cmbFjyouken.SelectedIndex == -1)
                    {
                        cmbFjyouken.Focus();
                        throw new Exception("�z�z������I�����Ă�������");
                    }

                    //���^�`�F�b�N
                    if (cmbSize.SelectedIndex == -1)
                    {
                        cmbSize.Focus();
                        throw new Exception("�T�C�Y��I�����Ă�������");
                    }

                    //�z�z�P���F�������H
                    if (Utility.NumericCheck(txtHTanka.Text) == false)
                    {
                        this.txtHTanka.Focus();
                        throw new Exception("�z�z�P���͐����œ��͂��Ă�������");
                    }

                    ////�z�z�J�n��
                    //if (StartDate.Checked == false)
                    //{
                    //    StartDate.Focus();
                    //    throw new Exception("�z�z�J�n������͂��Ă�������");
                    //}

                    ////�z�z�I����
                    //if (EndDate.Checked == false)
                    //{
                    //    EndDate.Focus();
                    //    throw new Exception("�z�z�I��������͂��Ă�������");
                    //}

                    //�z�z����
                    if (StartDate.Value > EndDate.Value)
                    {
                        StartDate.Focus();
                        throw new Exception("�z�z���Ԃ�����������܂���");
                    }
                    
                    // �[�i���K�{�Ƃ��� 2016/01/17
                    if (!NouhinDate.Checked)
                    {
                        NouhinDate.Focus();
                        throw new Exception("�[�i���������͂ł�");
                    }

                    ////�z�z�P�\�`�F�b�N    // 2015/06/23
                    //if (cmbFyuyo.SelectedIndex == -1)
                    //{
                    //    cmbFyuyo.Focus();
                    //    throw new Exception("�z�z�P�\��I�����Ă�������");
                    //}

                    ////�񍐎����`�F�b�N    // 2015/06/23
                    //if (cmbHjiki.SelectedIndex == -1)
                    //{
                    //    cmbHjiki.Focus();
                    //    throw new Exception("�񍐎�����I�����Ă�������");
                    //}

                    ////�񍐐��x�`�F�b�N    // 2015/06/23
                    //if (cmbHseido.SelectedIndex == -1)
                    //{
                    //    cmbHseido.Focus();
                    //    throw new Exception("�񍐐��x��I�����Ă�������");
                    //}

                    ////�񍐕��@�`�F�b�N    // 2015/06/23
                    //if (cmbHhouhou.SelectedIndex == -1)
                    //{
                    //    cmbHhouhou.Focus();
                    //    throw new Exception("�񍐕��@��I�����Ă�������");
                    //}
                }
                
                // �z�z�J�n���F2016/08/03 �󒍓��e���|�X�e�B���O�Ɍ��炸�K�{���͍��ڂƂ���
                if (StartDate.Checked == false)
                {
                    StartDate.Focus();
                    throw new Exception("�z�z�J�n������͂��Ă�������");
                }

                //�z�z�I�����F2017/01/27 �󒍓��e���|�X�e�B���O�Ɍ��炸�K�{���͍��ڂƂ���
                if (EndDate.Checked == false)
                {
                    EndDate.Focus();
                    throw new Exception("�z�z�I��������͂��Ă�������");
                }

                // �O���˗����ƊO���n�������@2015/09/20
                if (!dtGaichuIrai.Checked && Utility.strToInt(txtGaichuMaisu.Text) > 0)
                {
                    tabControl1.SelectTab(0);   // 2016/10/15
                    dtGaichuIrai.Focus();
                    throw new Exception("�O���˗����������͂œn�����������͂���Ă��܂�");
                }

                // 2016/10/15
                if (!dtGaichuIrai2.Checked && Utility.strToInt(txtGaichuMaisu2.Text) > 0)
                {
                    tabControl1.SelectTab(1);
                    dtGaichuIrai2.Focus();
                    throw new Exception("�O���˗����������͂œn�����������͂���Ă��܂�");
                }

                // 2016/10/15
                if (!dtGaichuIrai3.Checked && Utility.strToInt(txtGaichuMaisu3.Text) > 0)
                {
                    tabControl1.SelectTab(2);
                    dtGaichuIrai3.Focus();
                    throw new Exception("�O���˗����������͂œn�����������͂���Ă��܂�");
                }

                // �c�ƌ����@2015/11/18
                // �O���悪�I���ς݂�
                if (cmbeGaichu.SelectedIndex != -1)
                {
                    // �x�����������͂̂Ƃ�
                    if (!dteGaichuPay.Checked)
                    {
                        dteGaichuPay.Focus();
                        throw new Exception("�c�ƌ����̎x�����������͂ł�");
                    }

                    // �����������͂̂Ƃ�
                    if (txteGaichuGenka.Text == string.Empty)
                    {
                        txteGaichuGenka.Focus();
                        throw new Exception("�c�ƌ����������͂ł�");
                    }
                }
                else
                {
                    // 2016/01/17
                    cmbeGaichu.Focus();
                    throw new Exception("�O���悪�I������Ă��܂���");
                }

                // �x���������͍ς݂�
                if (dteGaichuPay.Checked)
                {
                    // �O���悪���I���̂Ƃ�
                    if (cmbeGaichu.SelectedIndex == -1)
                    {
                        cmbeGaichu.Focus();
                        throw new Exception("�c�ƌ����̊O�����I�����Ă�������");
                    }

                    // �����������͂̂Ƃ�
                    if (txteGaichuGenka.Text == string.Empty)
                    {
                        txteGaichuGenka.Focus();
                        throw new Exception("�c�ƌ����������͂ł�");
                    }
                }

                // ���������͍ς݂�
                if (Utility.strToDouble(txteGaichuGenka.Text) != 0)
                {
                    // �O���悪���I���̂Ƃ�
                    if (cmbeGaichu.SelectedIndex == -1)
                    {
                        cmbeGaichu.Focus();
                        throw new Exception("�c�ƌ����̊O�����I�����Ă�������");
                    }

                    // �x�����������͂̂Ƃ�
                    if (!dteGaichuPay.Checked)
                    {
                        dteGaichuPay.Focus();
                        throw new Exception("�x�����������͂ł�");
                    }
                }

                // �O����x���p�@2015/11/18
                // �x���p�O���悪�I���ς݂�
                if (cmbpGaichu.SelectedIndex != -1)
                {
                    // �x�����������͂̂Ƃ�
                    if (!dtpGaichuPay.Checked)
                    {
                        tabControl1.SelectTab(0);
                        dtpGaichuPay.Focus();
                        throw new Exception("�O����x���p�̎x�����������͂ł�");
                    }

                    // �����������͂̂Ƃ�
                    if (txtpGaichuGenka.Text == string.Empty)
                    {
                        tabControl1.SelectTab(0);
                        txtpGaichuGenka.Focus();
                        throw new Exception("�O����x���p�����������͂ł�");
                    }
                }
                else
                {
                    tabControl1.SelectTab(0);

                    // 2016/01/17
                    cmbpGaichu.Focus();
                    throw new Exception("�O���悪�I������Ă��܂���");
                }

                // �O����x���������͍ς݂�
                if (dtpGaichuPay.Checked)
                {
                    // �x���p�O���悪���I���̂Ƃ�
                    if (cmbpGaichu.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(0);
                        cmbpGaichu.Focus();
                        throw new Exception("�O����x���p�O�����I�����Ă�������");
                    }

                    // �����������͂̂Ƃ�
                    if (txtpGaichuGenka.Text == string.Empty)
                    {
                        tabControl1.SelectTab(0);
                        txtpGaichuGenka.Focus();
                        throw new Exception("�O����x���p�����������͂ł�");
                    }
                }

                // ���������͍ς݂�
                if (Utility.strToDouble(txtpGaichuGenka.Text) != 0)
                {
                    // �O����x���p�O���悪���I���̂Ƃ�
                    if (cmbpGaichu.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(0);
                        cmbpGaichu.Focus();
                        throw new Exception("�O����x���p�O�����I�����Ă�������");
                    }

                    // �O����x�����������͂̂Ƃ�
                    if (!dtpGaichuPay.Checked)
                    {
                        tabControl1.SelectTab(0);
                        dtpGaichuPay.Focus();
                        throw new Exception("�O����x�����������͂ł�");
                    }
                }


                // �O����x���p�Q�@2016/10/14
                // �x���p�O����Q���I���ς݂�
                if (cmbpGaichu2.SelectedIndex != -1)
                {
                    // �x�����������͂̂Ƃ�
                    if (!dtpGaichuPay2.Checked)
                    {
                        tabControl1.SelectTab(1);
                        dtpGaichuPay2.Focus();
                        throw new Exception("�O����x���p�Q�̎x�����������͂ł�");
                    }

                    // �����������͂̂Ƃ�
                    if (txtpGaichuGenka2.Text == string.Empty)
                    {
                        tabControl1.SelectTab(1);
                        txtpGaichuGenka2.Focus();
                        throw new Exception("�O����x���p�����Q�������͂ł�");
                    }
                }

                // �O����x����2�����͍ς݂�
                if (dtpGaichuPay2.Checked)
                {
                    // �x���p�O����2�����I���̂Ƃ�
                    if (cmbpGaichu2.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(1);
                        cmbpGaichu2.Focus();
                        throw new Exception("�O����x���p�O����Q��I�����Ă�������");
                    }

                    // ����2�������͂̂Ƃ�
                    if (txtpGaichuGenka2.Text == string.Empty)
                    {
                        tabControl1.SelectTab(1);
                        txtpGaichuGenka2.Focus();
                        throw new Exception("�O����x���p�����Q�������͂ł�");
                    }
                }

                // �����Q�����͍ς݂�
                if (Utility.strToDouble(txtpGaichuGenka2.Text) != 0)
                {
                    // �O����x���p�O����Q�����I���̂Ƃ�
                    if (cmbpGaichu2.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(1);
                        cmbpGaichu2.Focus();
                        throw new Exception("�O����x���p�O����Q��I�����Ă�������");
                    }

                    // �O����x�����Q�������͂̂Ƃ�
                    if (!dtpGaichuPay2.Checked)
                    {
                        tabControl1.SelectTab(1);
                        dtpGaichuPay2.Focus();
                        throw new Exception("�O����x�����Q�������͂ł�");
                    }
                }


                // �O����x���p�R�@2016/10/14
                // �x���p�O����R���I���ς݂�
                if (cmbpGaichu3.SelectedIndex != -1)
                {
                    // �x�����������͂̂Ƃ�
                    if (!dtpGaichuPay3.Checked)
                    {
                        tabControl1.SelectTab(2);
                        dtpGaichuPay3.Focus();
                        throw new Exception("�O����x���p�R�̎x�����������͂ł�");
                    }

                    // �����������͂̂Ƃ�
                    if (txtpGaichuGenka3.Text == string.Empty)
                    {
                        tabControl1.SelectTab(2);
                        txtpGaichuGenka3.Focus();
                        throw new Exception("�O����x���p�����R�������͂ł�");
                    }
                }

                // �O����x�����R�����͍ς݂�
                if (dtpGaichuPay3.Checked)
                {
                    // �x���p�O����3�����I���̂Ƃ�
                    if (cmbpGaichu3.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(2);
                        cmbpGaichu3.Focus();
                        throw new Exception("�O����x���p�O����R��I�����Ă�������");
                    }

                    // �����R�������͂̂Ƃ�
                    if (txtpGaichuGenka3.Text == string.Empty)
                    {
                        tabControl1.SelectTab(2);
                        txtpGaichuGenka3.Focus();
                        throw new Exception("�O����x���p�����R�������͂ł�");
                    }
                }

                // �����R�����͍ς݂�
                if (Utility.strToDouble(txtpGaichuGenka3.Text) != 0)
                {
                    // �O����x���p�O����R�����I���̂Ƃ�
                    if (cmbpGaichu3.SelectedIndex == -1)
                    {
                        tabControl1.SelectTab(2);
                        cmbpGaichu3.Focus();
                        throw new Exception("�O����x���p�O����R��I�����Ă�������");
                    }

                    // �O����x�����R�������͂̂Ƃ�
                    if (!dtpGaichuPay3.Checked)
                    {
                        tabControl1.SelectTab(2);
                        dtpGaichuPay3.Focus();
                        throw new Exception("�O����x�����R�������͂ł�");
                    }
                }
                
                // ���������s���K�{�Ƃ��� 2016/01/17
                if (!dtSeikyu.Checked)
                {
                    dtSeikyu.Focus();
                    throw new Exception("���������s���������͂ł�");
                }

                // �x�������K�{�Ƃ��� 2016/01/17
                if (!NyukinDate.Checked)
                {
                    NyukinDate.Focus();
                    throw new Exception("�x�������������͂ł�");
                }

                //�N���X�Ƀf�[�^�Z�b�g
                if (fMode.Mode == 1)
                {
                    cMaster.ID = fMode.jID;
                }

                Utility.ComboOffice cmb1 = new Utility.ComboOffice();
                cmb1 = (Utility.ComboOffice)cmbOffice.SelectedItem;
                cMaster.���Ə�ID = cmb1.ID;

                cMaster.�󒍓� = jDate.Value;

                //cMaster.�󒍋敪 = comboBox1.Text;
                cMaster.�󒍋敪 = "";

                Utility.ComboClient cmb2 = new Utility.ComboClient();
                cmb2 = (Utility.ComboClient)cmbClient.SelectedItem;
                cMaster.���Ӑ�ID = cmb2.ID;

                cMaster.�`���V�� = txtChirashi.Text.ToString();

                Utility.ComboJshubetsu cmb3 = new Utility.ComboJshubetsu();
                cmb3 = (Utility.ComboJshubetsu)cmbNaiyou.SelectedItem;
                cMaster.�󒍎��ID = cmb3.ID;

                cMaster.�P�� = Utility.strToDouble(txtTanka.Text);
                cMaster.���� = Utility.strToInt(txtMai.Text);
                cMaster.���z = Utility.strToInt(txtUri.Text);
                cMaster.����� = Utility.strToInt(txtTax.Text);
                cMaster.�ō����z = Utility.strToInt(txtZeikomi.Text);
                cMaster.�l���z = Utility.strToInt(txtNebiki.Text);
                cMaster.������z = Utility.strToInt(txtZeikomi.Text);
                cMaster.�ŗ� = cTax.Ritsu;

                //�T�C�Y(���^)ID�擾
                if (cmbSize.SelectedIndex != -1)
                {
                    Utility.ComboSize cmb4 = new Utility.ComboSize();
                    cmb4 = (Utility.ComboSize)cmbSize.SelectedItem;
                    cMaster.���^ = cmb4.ID;
                }
                else
                {
                    cMaster.���^ = 0;
                }

                cMaster.�z�z�P�� = Utility.strToDouble(txtHTanka.Text);
                //cMaster.�˗��� = txtIraisaki.Text.ToString();
                //cMaster.���� = Convert.ToDouble(txtGenka.Text.ToString());
                cMaster.�˗��� = "";
                cMaster.���� = (double)0;

                //�z�z�`��ID�擾
                if (cmbFkeitai.SelectedIndex != -1)
                {
                    Utility.ComboFkeitai cmb5 = new Utility.ComboFkeitai();
                    cmb5 = (Utility.ComboFkeitai)cmbFkeitai.SelectedItem;
                    cMaster.�z�z�`�� = cmb5.ID;
                }
                else
                {
                    cMaster.�z�z�`�� = 0;
                }

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

                cMaster.�z�z�P�\ = string.Empty;    // 2015/07/01

                if (NouhinDate.Checked == true)
                {
                    cMaster.�[�i�\��� = NouhinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�[�i�\��� = "";
                }

                cMaster.�[�i�`�� = string.Empty;    // 2015/07/01
                
                cMaster.������ = 1;

                //if (checkBox1.Checked == true)
                //{
                //    cMaster.������ = 1;
                //}
                //else
                //{
                //    cMaster.������ = 0;
                //}

                //��������
                cMaster.������ID = Int32.Parse(txtSeikyuNumber.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);

                cMaster.�������@ = "";

                //////if (cmbNyukin.SelectedIndex == -1)
                //////{
                //////    cMaster.�������@ = "";
                //////}
                //////else
                //////{
                //////    Utility.ComboShimebi cmb6 = new Utility.ComboShimebi();
                //////    cmb6 = (Utility.ComboShimebi)cmbNyukin.SelectedItem;
                //////    cMaster.�������@ = cmb6.Name;
                //////}

                if (NyukinDate.Checked == true)
                {
                    cMaster.�����\��� = NyukinDate.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�����\��� = "";
                }

                cMaster.�񍐎��� = string.Empty;        // 2015/07/01
                cMaster.�񍐐��x = string.Empty;        // 2015/07/01
                cMaster.�񍐕��@ = string.Empty;        // 2015/07/01
                cMaster.���[���A�h���X = string.Empty;  // 2015/07/01

                cMaster.�U������ID = 0;

                //if (cmbFuri.SelectedIndex == -1)
                //{
                //    cMaster.�U������ID = 0;
                //}
                //else
                //{
                //    Utility.ComboFuri cmb7 = new Utility.ComboFuri();
                //    cmb7 = (Utility.ComboFuri)cmbFuri.SelectedItem;
                //    cMaster.�U������ID = cmb7.ID;
                //}

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
                cMaster.�G���A���l = string.Empty;       // 2015/11/11
                cMaster.�c�Ɣ��l = txtSalesMemo.Text;   // 2019/03/01
                cMaster.�����敪 = 0;

                // 2014/11/26 ���z���O
                if (chkHeihai.Checked == true)  
                {
                    cMaster.���z���O = 1;
                }
                else
                {
                    cMaster.���z���O = 0;
                }

                if (fMode.Mode == 0)
                {
                    cMaster.�o�^�N���� = DateTime.Now;
                }

                cMaster.�ύX�N���� = DateTime.Now;

                // 2015/06/30
                if (dtSeikyu.Checked == true)
                {
                    cMaster.���������s�� = dtSeikyu.Value.ToShortDateString();
                }
                else
                {
                    cMaster.���������s�� = "";
                }

                // 2015/06/30
                if (cmbeGaichu.SelectedIndex != -1)
                { 
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbeGaichu.SelectedItem;
                    cMaster.�O����ID�c�� = cmb.ID;
                }
                else
                {
                    cMaster.�O����ID�c�� = 0;
                }

                // 2015/06/30
                if (dteGaichuPay.Checked)
                {
                    cMaster.�O����x�����c�� = dteGaichuPay.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O����x�����c�� = string.Empty;
                }

                // 2015/06/30
                cMaster.�O���挴���c�� = Utility.strToDouble(txteGaichuGenka.Text);

                // 2015/06/30
                if (cmbpGaichu.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbpGaichu.SelectedItem;
                    cMaster.�O����ID�x�� = cmb.ID;
                }
                else
                {
                    cMaster.�O����ID�x�� = 0;
                }

                // 2015/06/30
                if (dtpGaichuPay.Checked)
                {
                    cMaster.�O����x�����x�� = dtpGaichuPay.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O����x�����x�� = string.Empty;
                }

                // 2015/06/30
                cMaster.�O���挴���x�� = Utility.strToDouble(txtpGaichuGenka.Text);

                // 2016/10/14
                if (cmbpGaichu2.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbpGaichu2.SelectedItem;
                    cMaster.�O����ID�x��2 = cmb.ID;
                }
                else
                {
                    cMaster.�O����ID�x��2 = 0;
                }

                // 2016/10/14
                if (dtpGaichuPay2.Checked)
                {
                    cMaster.�O����x�����x��2 = dtpGaichuPay2.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O����x�����x��2 = string.Empty;
                }

                // 2016/10/14
                cMaster.�O���挴���x��2 = Utility.strToDouble(txtpGaichuGenka2.Text);

                // 2016/10/14
                if (cmbpGaichu3.SelectedIndex != -1)
                {
                    Utility.comboGaichu cmb = (Utility.comboGaichu)cmbpGaichu3.SelectedItem;
                    cMaster.�O����ID�x��3 = cmb.ID;
                }
                else
                {
                    cMaster.�O����ID�x��3 = 0;
                }

                // 2016/10/14
                if (dtpGaichuPay3.Checked)
                {
                    cMaster.�O����x�����x��3 = dtpGaichuPay3.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O����x�����x��3 = string.Empty;
                }

                // 2016/10/14
                cMaster.�O���挴���x��3 = Utility.strToDouble(txtpGaichuGenka3.Text);

                // 2015/06/30
                if (cmbAnShu.SelectedIndex != -1)
                {
                    Utility.ComboAnshu cmb = (Utility.ComboAnshu)cmbAnShu.SelectedItem;
                    cMaster.�Č���� = cmb.ID;
                }
                else
                {
                    cMaster.�Č���� = 0;
                }

                // 2015/06/30, 2015/07/10
                cMaster.���[�U�[ID = global.loginUserID;

                // 2015/08/11
                cMaster.�O����˗����c�� = string.Empty;

                // 2015/08/11
                if (dtGaichuIrai.Checked)
                {
                    cMaster.�O����˗����x�� = dtGaichuIrai.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O����˗����x�� = string.Empty;
                }

                // 2016/10/15
                if (dtGaichuIrai2.Checked)
                {
                    cMaster.�O����˗����x��2 = dtGaichuIrai2.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O����˗����x��2 = string.Empty;
                }

                if (dtGaichuIrai3.Checked)
                {
                    cMaster.�O����˗����x��3 = dtGaichuIrai3.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O����˗����x��3 = string.Empty;
                }

                // 2015/08/11
                if (dtGaichuWatashi.Checked)
                {
                    cMaster.�O���n���� = dtGaichuWatashi.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O���n���� = string.Empty;
                }

                // 2016/10/15
                if (dtGaichuWatashi2.Checked)
                {
                    cMaster.�O���n����2 = dtGaichuWatashi2.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O���n����2 = string.Empty;
                }

                if (dtGaichuWatashi3.Checked)
                {
                    cMaster.�O���n����3 = dtGaichuWatashi3.Value.ToShortDateString();
                }
                else
                {
                    cMaster.�O���n����3 = string.Empty;
                }
                                
                cMaster.�O���󂯓n���S���� = txtGaichuUkeTan.Text;     // 2015/08/11
                cMaster.�O���󂯓n���S����2 = txtGaichuUkeTan2.Text;   // 2016/10/15
                cMaster.�O���󂯓n���S����3 = txtGaichuUkeTan3.Text;   // 2016/10/15

                // 2015/09/20
                cMaster.�O���ϑ����� = Utility.strToInt(txtGaichuMaisu.Text);
                cMaster.�O���ϑ�����2 = Utility.strToInt(txtGaichuMaisu2.Text);   // 2016/10/15
                cMaster.�O���ϑ�����3 = Utility.strToInt(txtGaichuMaisu3.Text);   // 2016/10/15

                cMaster.�Ǝ� = txtGyoushu.Text;

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

            //if (sender == txtIraisaki)
            //{
            //    objtxt = txtIraisaki;
            //}

            //if (sender == txtGenka)
            //{
            //    objtxt = txtGenka;
            //}

            //if (sender == txtEmail)   // 2015/06/23
            //{
            //    objtxt = txtEmail;
            //}

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            //if (sender == txtMemo2)
            //{
            //    objtxt = txtMemo2;
            //}

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            if (sender == txtHTanka)
            {
                objtxt = txtHTanka;
            }

            if (sender == cmbClient)
            {
                cmbClient.BackColor = Color.LightSteelBlue;
            }

            // 2015/07/01
            if (sender == cmbeGaichu)
            {
                cmbeGaichu.BackColor = Color.LightSteelBlue;
            }

            if (sender == txteGaichuGenka)
            {
                objtxt = txteGaichuGenka;
            }

            if (sender == cmbpGaichu)
            {
                cmbpGaichu.BackColor = Color.LightSteelBlue;
            }

            if (sender == txtpGaichuGenka)
            {
                objtxt = txtpGaichuGenka;
            }

            if (sender == txtsClient)
            {
                objtxt = txtsClient;
            }

            if (sender == txtsGaichu)
            {
                objtxt = txtsGaichu;
            }

            if (sender == cmbAnShu)
            {
                cmbAnShu.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbNaiyou)
            {
                cmbNaiyou.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbOffice)
            {
                cmbOffice.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbsNaiyou)
            {
                cmbsNaiyou.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbFkeitai)
            {
                cmbFkeitai.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbFjyouken)
            {
                cmbFjyouken.BackColor = Color.LightSteelBlue;
            }

            if (sender == cmbSize)
            {
                cmbSize.BackColor = Color.LightSteelBlue;
            }

            if (sender == txtsOrderNum)
            {
                txtsOrderNum.BackColor = Color.LightSteelBlue;
            }


            objtxt.SelectAll();
            objtxt.BackColor = Color.LightSteelBlue;
        }

        private void txtLeave(object sender, EventArgs e)
        {
            TextBox objtxt = new TextBox();

            decimal Kingaku = 0;
            decimal KingakuTax = 0;
            decimal KingakuZeikomi = 0;
            decimal KingakuTL = 0;
            DateTime dt = DateTime.Today;
            string str;
            double d;
            decimal dTanka = 0;
            decimal dHTanka = 0;
            int dMai = 0;
            int dNebiki = 0;

            try
            {
                // ���e�̂Ƃ�
                if (sender == txtChirashi)
                {
                    objtxt = txtChirashi;
                }

                // �P���܂��͖����̂Ƃ� 2015/07/01
                if (sender == txtTanka || sender == txtMai || sender == txtNebiki)
                {
                    // �P��
                    if (sender == txtTanka)
                    {
                        objtxt = txtTanka;
                    }

                    // ����
                    if (sender == txtMai)
                    {
                        objtxt = txtMai;
                    }
                    
                    // ���t�擾
                    //DateTime.TryParse(jDate.Value.ToShortDateString(), out dt);  // 2019/08/26 �R�����g��
                    DateTime.TryParse(dtSeikyu.Value.ToShortDateString(), out dt);  // �������� 2019/08/26

                    // �P���擾
                    str = Utility.strToDouble(txtTanka.Text).ToString();

                    if (!decimal.TryParse(str, out dTanka))
                    {
                        dTanka = 0m;
                    }

                    // �����擾
                    dMai = Utility.strToInt(txtMai.Text);

                    // �l���z�擾
                    dNebiki = Utility.strToInt(txtNebiki.Text);

                    // ������z
                    Kingaku = Math.Floor(dTanka * (decimal)dMai);

                    // �l������z
                    decimal nebikigoKin = Kingaku - dNebiki;
                    txtNebikigo.Text = nebikigoKin.ToString("#,##0");

                    if (dtSeikyu.Checked)
                    {
                        // ���z�v�Z
                        UriageSum(dt, nebikigoKin, out KingakuTax, out KingakuZeikomi, out KingakuTL);
                    }

                    // ������z
                    txtUri.Text = Kingaku.ToString("#,##0");

                    // �c�ƌ����E�e�����z�F2019/02/21
                    txteGaichuArari.Text = (Utility.strToInt(txtUri.Text) - Utility.strToInt(txteGaichuGenka.Text)).ToString("#,##0");

                    // �O���e���F2019/02/21
                    int gai = Utility.strToInt(txtpGaichuGenka.Text) + Utility.strToInt(txtpGaichuGenka2.Text) + Utility.strToInt(txtpGaichuGenka3.Text);
                    int arari = Utility.strToInt(txtUri.Text) - gai;
                    txtaUri.Text = Kingaku.ToString("#,##0");
                    txtaArari.Text = arari.ToString("#,##0");

                    tabControl1.TabPages[3].Text = "�e���F" + arari.ToString("#,##0");

                    // ����Ŋz�v�Z               
                    txtTax.Text = KingakuTax.ToString("#,##0");

                    // �ō����z
                    txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");

                    // �c�ƌ����F�|�X�e�B���O�̂Ƃ� 2018/01/03
                    if (cmbNaiyou.Text == J_NAIYOU_POSTING)
                    {
                        // �z�z�P���擾
                        str = Utility.strToDouble(txtHTanka.Text).ToString();

                        if (!decimal.TryParse(str, out dHTanka))
                        {
                            dHTanka = 0m;
                        }

                        decimal genka = Math.Floor(dHTanka * (decimal)dMai);
                        txteGaichuGenka.Text = genka.ToString("#,##0");
                    }
                }

                // �l���̂Ƃ� 2015/07/01
                if (sender == txtNebiki)
                {
                    objtxt = txtNebiki;

                    //������z
                    //txtNebikigo.Text = (Utility.strToInt(txtZeikomi.Text) - Utility.strToInt(txtNebiki.Text)).ToString("#,##0");
                }

                // �z�z�P���̂Ƃ�  2018/01/03
                if (sender == txtHTanka)
                {
                    objtxt = txtHTanka;

                    // �c�ƌ����F�|�X�e�B���O�̂Ƃ�
                    if (cmbNaiyou.Text == J_NAIYOU_POSTING)
                    {
                        // �z�z�P���擾
                        str = Utility.strToDouble(txtHTanka.Text).ToString();

                        if (!decimal.TryParse(str, out dHTanka))
                        {
                            dHTanka = 0m;
                        }

                        // �����擾
                        dMai = Utility.strToInt(txtMai.Text);

                        decimal genka = Math.Floor(dHTanka * (decimal)dMai);
                        txteGaichuGenka.Text = genka.ToString("#,##0");
                    }
                }

                //if (sender == txtIraisaki)
                //{
                //    objtxt = txtIraisaki;
                //}

                //if (sender == txtGenka)
                //{
                //    objtxt = txtGenka;
                //}

                //if (sender == txtEmail)    // 2015/06/23
                //{
                //    objtxt = txtEmail;
                //}

                if (sender == txtMemo)
                {
                    objtxt = txtMemo;
                }

                //if (sender == txtMemo2)
                //{
                //    objtxt = txtMemo2;
                //}

                if (sender == textBox1)
                {
                    objtxt = textBox1;
                }

                if (sender == txtHTanka)
                {
                    objtxt = txtHTanka;
                }

                if (sender == cmbClient)
                {
                    cmbClient.BackColor = Color.White;

                    if (cmbClient.SelectedIndex != -1)
                    {
                        //�N���C�A���g���\��
                        Utility.ComboClient cmbC = new Utility.ComboClient();
                        cmbC = (Utility.ComboClient)cmbClient.SelectedItem;
                        ClientShow(cmbC.ID);
                    }
                }

                // 2015/07/01
                if (sender == cmbsNaiyou)
                {
                    cmbsNaiyou.BackColor = Color.White;
                }

                if (sender == cmbNaiyou)
                {
                    cmbNaiyou.BackColor = Color.White;
                }

                if (sender == cmbOffice)
                {
                    cmbOffice.BackColor = Color.White;
                }

                if (sender == cmbFkeitai)
                {
                    cmbFkeitai.BackColor = Color.White;
                }

                if (sender == cmbFjyouken)
                {
                    cmbFjyouken.BackColor = Color.White;
                }

                if (sender == cmbSize)
                {
                    cmbSize.BackColor = Color.White;
                }

                if (sender == cmbAnShu)
                {
                    cmbAnShu.BackColor = Color.White;
                }

                if (sender == cmbeGaichu)
                {
                    cmbeGaichu.BackColor = Color.White;
                }

                if (sender == cmbpGaichu)
                {
                    cmbpGaichu.BackColor = Color.White;
                }

                if (sender == txteGaichuGenka)
                {
                    objtxt = txteGaichuGenka;
                }

                if (sender == txtpGaichuGenka)
                {
                    objtxt = txtpGaichuGenka;
                }

                if (sender == txtsClient)
                {
                    objtxt = txtsClient;
                }

                if (sender == txtsGaichu)
                {
                    objtxt = txtsGaichu;
                }

                if (sender == txtsOrderNum)
                {
                    objtxt = txtsOrderNum;
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
            // �V�������f�[�^�ƂȂ������߈ȉ��R�����g�� 2015/12/02
            //���������s�ς݂��H  2010/05/30
            //if (txtSeikyuNumber.Text.ToString() != "0")
            //{
            //    MessageBox.Show("���ɐ����������s�ς݂ł��B(��" + txtSeikyuNumber.Text + ")" + Environment.NewLine + "�󒍃f�[�^���폜����ɂ͐������f�[�^���폜������Ɏ��s���Ă�������", "���������s�ς�", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}

            // 2015/12/02 �������ԍ��擾
            int seikyuNum = Utility.strToInt(txtSeikyuNumber.Text);

            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            
            string strSql;

            //�f�[�^�폜
            Control.DataControl Con = new Control.DataControl();
            OleDbConnection cn = new OleDbConnection();
            OleDbTransaction tran;

            cn = Con.GetConnection();

            //�g�����U�N�V�����̊J�n
            tran = cn.BeginTransaction();

            OleDbCommand SCom = new OleDbCommand();

            SCom.Connection = cn;
            SCom.Transaction = tran;

            try
            {
                //�󒍃f�[�^�̍폜
                strSql = "";
                strSql += "delete from �� ";
                strSql += "where ID = " + cMaster.ID.ToString();

                SCom.CommandText = strSql;

                SCom.ExecuteNonQuery();

                //�z�z�G���A�f�[�^�̍폜
                strSql = "";
                strSql += "delete from �z�z�G���A ";
                strSql += "where �z�z�G���A.��ID = " + cMaster.ID.ToString();

                SCom.CommandText = strSql;

                SCom.ExecuteNonQuery();

                //�R�~�b�g
                tran.Commit();

                MessageBox.Show("�폜����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // �󒍃f�[�^�ǂݍ���
                jAdp.Fill(dts.��1);
                
                // �������f�[�^�擾
                if (dts.��1.Count(a => a.������ID == seikyuNum) > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("�폜�����󒍃f�[�^���܂ސ������f�[�^������܂��B").Append(Environment.NewLine);
                    sb.Append("�������ߏ������s���Đ������z���Čv�Z���Ă��������B");

                    MessageBox.Show(sb.ToString(), "�������f�[�^�̍č쐬���K�v�ł�", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                //���[���o�b�N
                tran.Rollback();
                MessageBox.Show(ex.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MessageBox.Show("�폜�Ɏ��s���܂����B���[���o�b�N���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            cn.Close();
            Con.Close();

            DispClear();

            //�f�[�^�� 'darwinDataSet.��' �e�[�u���ɓǂݍ��݂܂��B
            //this.��TableAdapter.Fill(this.darwinDataSet.��);
            //dataGridView1.DataSource = this.darwinDataSet.��;

            dataGridView1.DataSource = null;

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

        //////private void cmbJkbnSet()
        //////{
        //////    comboBox1.Items.Clear();
        //////    comboBox1.Items.Add("�V�K");
        //////    comboBox1.Items.Add("���s�[�g");
        //////}

        private void cmbFjyoukenSet()
        {
            cmbFjyouken.Items.Clear();
            cmbFjyouken.Items.Add("�\�薇���ǂ���");
            cmbFjyouken.Items.Add("�l�߂Ĕz�z �n�j");
        }

        //private void cmbFyuyoSet()    // 2015/06/23
        //{
        //    cmbFyuyo.Items.Clear();
        //    cmbFyuyo.Items.Add("����");
        //    cmbFyuyo.Items.Add("�O��n�j");
        //}

        //private void cmbNkeitaiSet()  // 2015/06/23
        //{
        //    cmbNkeitai.Items.Clear();
        //    cmbNkeitai.Items.Add("��}��");
        //    cmbNkeitai.Items.Add("����");
        //    cmbNkeitai.Items.Add("�W�ׁF�o�b�N");
        //    cmbNkeitai.Items.Add("�W�ׁF�c��");
        //}

        //private void cmbHjikiSet()    // 2015/06/23
        //{
        //    cmbHjiki.Items.Clear();
        //    cmbHjiki.Items.Add("�f�C���[");
        //    cmbHjiki.Items.Add("�T�P��");
        //    cmbHjiki.Items.Add("�I����");
        //}

        //private void cmbHseidoSet()   // 2015/06/23
        //{
        //    cmbHseido.Items.Clear();
        //    cmbHseido.Items.Add("���z�z��");
        //    cmbHseido.Items.Add("�\�薇��");
        //}

        //private void cmbHhouhouSet()  // 2015/06/23
        //{
        //    cmbHhouhou.Items.Clear();
        //    cmbHhouhou.Items.Add("FAX");
        //    cmbHhouhou.Items.Add("���[��");
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            //darwinDataSet ds = new darwinDataSet();
            //ds.Clear();
            //ds.EnforceConstraints = false;
            //this.��TableAdapter.FillByName(ds.��, "%" + textBox1.Text.ToString() + "%");
            //dataGridView1.DataSource = ds.��;

            dataSerach();
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     �󒍃f�[�^���� </summary>
        /// ------------------------------------------------------------------
        private void dataSerach()
        {
            this.Cursor = Cursors.WaitCursor;
            
            darwinDataSet dts = new darwinDataSet();
            dts.EnforceConstraints = false;

            darwinDataSetTableAdapters.��TableAdapter adp = new darwinDataSetTableAdapters.��TableAdapter();
            darwinDataSetTableAdapters.�O����TableAdapter gAdp = new darwinDataSetTableAdapters.�O����TableAdapter();
            darwinDataSetTableAdapters.���Ӑ�TableAdapter cAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
            darwinDataSetTableAdapters.�Ј�TableAdapter eAdp = new darwinDataSetTableAdapters.�Ј�TableAdapter();

            //adp.Fill(dts.��);

            DateTime sd = DateTime.Parse(sDt.Value.ToShortDateString()); // 2018/01/27
            DateTime ed = DateTime.Parse(eDt.Value.ToShortDateString()); // 2018/01/27
            adp.FillByFromYMDToYMD(dts.��, sd, ed); // 2018/01/27

            gAdp.Fill(dts.�O����);
            cAdp.Fill(dts.���Ӑ�);
            eAdp.Fill(dts.�Ј�);

            var s = dts.��.Where(a => a.ID >= 0);

            // ���O�C�����[�U�[�̎󒍃f�[�^�ێ琧��
            if (global.loginOrderMntType == 0)
            {
                s = s.Where(a => a.�o�^���[�U�[ID == global.loginUserID);
            }

            // �󒍓�
            if (sDt.Checked)
            {
                s = s.Where(a => a.�󒍓� >= DateTime.Parse(sDt.Value.ToShortDateString()));
            }

            if (eDt.Checked)
            {
                s = s.Where(a => a.�󒍓� <= DateTime.Parse(eDt.Value.ToShortDateString()));
            }

            // �`���V��
            if (textBox1.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.�`���V��.Contains(textBox1.Text.Trim()));
            }

            // �󒍔ԍ�
            if (txtsOrderNum.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.ID == Utility.strToLong(txtsOrderNum.Text));
            }

            // �󒍓��e�i�󒍎��ID�j
            if (cmbsNaiyou.SelectedIndex != -1)
            {
                Utility.ComboJshubetsu cmbJ = (Utility.ComboJshubetsu)cmbsNaiyou.SelectedItem;
                s = s.Where(a => a.�󒍎��ID == cmbJ.ID); 
            }

            // �O����
            if (txtsGaichu.Text != string.Empty)
            {
                s = s.Where(a => a.�O����Row != null && a.�O����Row.����.Contains(txtsGaichu.Text));
            }

            // ���Ӑ�
            if (txtsClient.Text != string.Empty)
            {
                s = s.Where(a => (!a.Is����Null() && a.����.Contains(txtsClient.Text)) || (!a.Is�t���K�iNull() && a.�t���K�i.Contains(txtsClient.Text)));
            }

            // �c�ƒS����
            if (txtsTantou.Text != string.Empty)
            {
                s = s.Where(a => (!a.Is���Ӑ�IDNull() && a.���Ӑ�Row.�Ј�Row != null && a.���Ӑ�Row.�Ј�Row.����.Contains(txtsTantou.Text)));
            }
            
            //�T�C�Y(���^)ID�擾 2019/02/14
            if (cmbsSize.SelectedIndex != -1)
            {
                Utility.ComboSize cmbs = new Utility.ComboSize();
                cmbs = (Utility.ComboSize)cmbsSize.SelectedItem;
                int sSize = cmbs.ID;

                s = s.Where(a => a.���^ == sSize);
            }


            // ���� 2019/02/14
            if (txtsMaisu.Text != string.Empty)
            {
                int sMai = Utility.strToInt(txtsMaisu.Text);
                if (sMai != global.FLGOFF)
                {
                    s = s.Where(a => a.���� == sMai);
                }
            }

            // �������� 2019/02/14
            if (seDt.Checked)
            {
                s = s.Where(a => a.���������s�� == DateTime.Parse(seDt.Value.ToShortDateString()));
            }


            // �f�[�^�O���b�h�r���[�Ƀf�[�^�\�[�X��ݒ�
            dataGridView1.DataSource = s.ToList();

            if (dataGridView1.RowCount > 0)
            {
                dataGridView1.CurrentCell = null;
            }

            this.Cursor = Cursors.Default;
        }
        
        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmOffice();

            frm.ShowDialog();
            Utility.ComboOffice.load(cmbOffice);
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     ����ŗ��擾 </summary>
        /// <param name="tempDate">
        ///     ����t</param>
        /// <returns>
        ///     �ŗ�</returns>
        ///  
        /// 2015/06/24
        ///------------------------------------------------------------------
        private int GetTaxRT(DateTime tempDate)
        {
            //�ŗ��擾
            int Ritsu = 0;

            darwinDataSet dts = new posting.darwinDataSet();
            darwinDataSetTableAdapters.�ŗ�TableAdapter adp = new darwinDataSetTableAdapters.�ŗ�TableAdapter();
            adp.Fill(dts.�ŗ�);

            foreach (var t in dts.�ŗ�.Where(a => a.�J�n�N���� <= tempDate).OrderByDescending(a => a.�J�n�N����))
            {
                Ritsu = t.�ŗ�;
                break;
            }

            return Ritsu;            
        }

        /// -------------------------------------------------------------
        /// <summary>
        ///     ����Ōv�Z </summary>
        /// <param name="tempKin">
        ///     �Ώۋ��z</param>
        /// <param name="tempTax">
        ///     �ŗ�</param>
        /// <returns>
        ///     ����Ŋz</returns>
        /// -------------------------------------------------------------
        private decimal GetTax(decimal tempKin, int tempTax)
        {
            decimal GakuD;
            //int GakuI;

            // �[���؎̂� 2015/07/01
            GakuD = Math.Floor(tempKin * tempTax / 100);

            //GakuD += (decimal)0.5;
            //GakuI = (int)GakuD;

            return GakuD;
        }

        /// ---------------------------------------------------------------------------------------
        /// <summary>
        ///     ������z���v�Z���� </summary>
        /// <param name="tempDate">
        ///     ������t�i���������Ƃ��� 2019/08/26�j</param>
        /// <param name="nebikigo">
        ///     �l������z </param>
        /// <param name="KingakuTax">
        ///     �����</param>
        /// <param name="KingakuZeikomi">
        ///     �ō�����</param>
        /// <param name="KingakuTL">
        ///     ������z</param>
        /// -----------------------------------------------------------------------------------------
        private void UriageSum(DateTime tempDate, decimal nebikigo, 
                                out decimal KingakuTax, out decimal KingakuZeikomi, out decimal KingakuTL)
        {
            //�ŗ��Ď擾
            cTax.Ritsu = Utility.GetTaxRT(tempDate);

            //����Ŋz�v�Z
            KingakuTax = Utility.GetTax(nebikigo, cTax.Ritsu);

            //�ō����z
            KingakuZeikomi = (int)nebikigo + KingakuTax;

            //������z
            KingakuTL = KingakuZeikomi;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     �N���C�A���g���\�� </summary>
        /// <param name="tempID">
        ///     ���Ӑ�ID</param>
        /// -------------------------------------------------------------------------
        private void ClientShow(int tempID)
        {
            foreach (var t in dts.���Ӑ�.Where(a => a.ID == tempID))
            {
                txtCzipcode.Text = t.�X�֔ԍ�;
                txtName2.Text = (t.�Z��1 + " " + t.�Z��2).Trim();

                // ��������
                txtCbusho.Text = t.������;
                //txtCtantou.Text = t.�S���Җ�;
                txtCtel.Text = t.�d�b�ԍ�; 
                txtCfax.Text = t.FAX�ԍ�;

                txtCtantou.Text = t.������S���Җ�;
                textBox4.Text = t.�����於��;

                if (t.�Ј�Row != null)
                {
                    txtTantou.Text = t.�Ј�Row.����;
                }
                else
                {
                    txtTantou.Text = string.Empty;
                }

                // ���� : 2018/01/03
                lblClientShimebi.Text = t.����.ToString() + "��";
            }
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

        private void button2_Click(object sender, EventArgs e)
        {
            OrderReport();
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     ��������� </summary>
        ///--------------------------------------------------------------------
        private void OrderReport()
        {
            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���������V�[�g��, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {

                    if (cmbClient.SelectedIndex != -1)
                    {
                        Utility.ComboClient cmbC = new Utility.ComboClient();
                        cmbC = (Utility.ComboClient)cmbClient.SelectedItem;
                        oxlsSheet.Cells[5, 2] = cmbC.NameShow; //�N���C�A���g��
                    }
                    else
                    {
                        oxlsSheet.Cells[5, 2] = "";
                    }

                    oxlsSheet.Cells[6, 2] = txtName2.Text;  //�Z��
                    oxlsSheet.Cells[9, 3] = txtCtel.Text;   //�d�b�ԍ�
                    oxlsSheet.Cells[9, 6] = txtCfax.Text;   //FAX�ԍ�

                    if (StartDate.Checked == true)
                    {
                        oxlsSheet.Cells[16, 1] = StartDate.Value.Month.ToString() + "��" + StartDate.Value.Day.ToString() + "���`"; //�z�z��
                    }
                    else
                    {
                        oxlsSheet.Cells[16, 1] = "";
                    }

                    oxlsSheet.Cells[16, 2] = cmbSize.Text;  //�T�C�Y
                    oxlsSheet.Cells[16, 3] = cmbFkeitai.Text;   //�z�z�`��
                    oxlsSheet.Cells[16, 5] = int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);   //�z�z����
                    oxlsSheet.Cells[16, 6] = double.Parse(txtTanka.Text.ToString(), System.Globalization.NumberStyles.Any);           //�P��
                    oxlsSheet.Cells[22, 6] = Utility.strToInt(txtNebiki.Text) * (-1);   // �l�� 
                    oxlsSheet.Cells[23, 6] = Utility.strToInt(txtTax.Text);   // ����� 

                    oxlsSheet.Cells[39, 6] = txtTantou.Text;    //�c�ƒS����

                    //�}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;

                    // �m�F�̂���Excel�̃E�B���h�E��\������
                    oXls.Visible = true;

                    //���
                    oxlsSheet.PrintPreview(true);

                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    //�ۑ�����
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //�_�C�A���O�{�b�N�X�̏����ݒ�
                    saveFileDialog1.Title = "�������ۑ�";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = "������_" + cmbClient.Text + " " + jDate.Value.ToLongDateString();
                    saveFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xls)|*.xls|�S�Ẵt�@�C��(*.*)|*.*";

                    //�_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName,Type.Missing,Type.Missing,
                                        Type.Missing,Type.Missing,Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange,Type.Missing,
                                        Type.Missing,Type.Missing,Type.Missing,Type.Missing);
                    }

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                MessageBox.Show(e.Message, "������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            NyukoReport();
        }

        private void NyukoReport()
        {
            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z�����ɊǗ��\�V�[�g��, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                try
                {
                    oxlsSheet.Cells[2, 2] = txtChirashi.Text; //���炵��
                    oxlsSheet.Cells[3, 2] = int.Parse(txtMai.Text.ToString(), System.Globalization.NumberStyles.AllowThousands);   //�z�z����

                    if (StartDate.Checked == true)
                    {
                        oxlsSheet.Cells[4, 2] = StartDate.Value.Month.ToString() + "/" + StartDate.Value.Day.ToString(); //�z�z�J�n��
                    }
                    else
                    {
                        oxlsSheet.Cells[4, 2] = "";
                    }

                    if (EndDate.Checked == true)
                    {
                        oxlsSheet.Cells[5, 2] = EndDate.Value.Month.ToString() + "/" + EndDate.Value.Day.ToString(); //�z�z�I����
                    }
                    else
                    {
                        oxlsSheet.Cells[5, 2] = "";
                    }

                    if (NouhinDate.Checked == true)
                    {
                        oxlsSheet.Cells[6, 2] = NouhinDate.Value.Month.ToString() + "/" + NouhinDate.Value.Day.ToString(); //�[�i�\���
                    }
                    else
                    {
                        oxlsSheet.Cells[6, 2] = "";
                    }

                    oxlsSheet.Cells[7, 2] = txtTantou.Text;         // �c�ƒS����
                    //oxlsSheet.Cells[8, 2] = cmbHseido.Text;       // �񍐐��x 2015/06/23 �P�p
                    oxlsSheet.Cells[8, 2] = cmbOffice.Text;         // ���Ə� 
                    oxlsSheet.Cells[9, 2] = cMaster.ID.ToString();  // ��ID 2015/06/23 

                    //�}�E�X�|�C���^�����ɖ߂�
                    this.Cursor = Cursors.Default;

                    // �m�F�̂���Excel�̃E�B���h�E��\������
                    oXls.Visible = true;

                    //���
                    oxlsSheet.PrintPreview(true);

                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    //�ۑ�����
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //�_�C�A���O�{�b�N�X�̏����ݒ�
                    saveFileDialog1.Title = "���ɊǗ��\�ۑ�";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = "���ɊǗ��\_" + txtChirashi.Text + " " + jDate.Value.ToLongDateString();
                    saveFileDialog1.Filter = "Microsoft Office Excel�t�@�C��(*.xls)|*.xls|�S�Ẵt�@�C��(*.*)|*.*";

                    //�_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "���ɊǗ��\", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                MessageBox.Show(e.Message, "���ɊǗ��\", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void label3_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            frmClient frm = new frmClient();
            frm.ShowDialog();

            //���Ӑ�R���{�ă��[�h : 2015/06/24
            //Utility.ComboClient.load(cmbClient);
            int cIdx = cmbClient.SelectedIndex;
            Utility.ComboClient.itemsLoad(cmbClient);
            cmbClient.SelectedIndex = cIdx;

            // ���Ӑ�f�[�^�ǂݍ���
            cAdp.Fill(dts.���Ӑ�);
        }

        private void cmbSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSize.SelectedIndex != -1)
            {
                Utility.ComboSize Cmbs = new Utility.ComboSize();
                Cmbs = (Utility.ComboSize)cmbSize.SelectedItem;

                txtHTanka.Text = Cmbs.Tanka.ToString("#,##0.0");
            }
        }

        private void cmbNaiyou_SelectedIndexChanged(object sender, EventArgs e)
        {
            //�|�X�e�B���O������ȊO���œ��͍��ڂ𐧌䂷�� 2010/1/27
            if (cmbNaiyou.SelectedIndex != -1)
            {
                // �|�X�e�B���O�ȊO�̂Ƃ�
                if (cmbNaiyou.SelectedIndex != 0)
                {
                    this.cmbFkeitai.SelectedIndex = -1;
                    this.cmbFkeitai.Enabled = false;

                    this.cmbFjyouken.SelectedIndex = -1;
                    this.cmbFjyouken.Enabled = false;

                    this.cmbSize.SelectedIndex = -1;

                    // 2015/11/10 ����A�f�U�C�����T�C�Y���͉\�Ƃ���
                    // 2016/01/19 �V���܍����T�C�Y���͉\�Ƃ���
                    if (cmbNaiyou.SelectedIndex == 1 || cmbNaiyou.SelectedIndex == 3 || cmbNaiyou.Text == "�V���܍�")
                    {
                        this.cmbSize.Enabled = true;
                    }
                    else
                    {
                        this.cmbSize.Enabled = false;
                    }

                    this.txtHTanka.Text = "0";
                    this.txtHTanka.Enabled = false;

                    //this.StartDate.Checked = false;
                    //this.StartDate.Enabled = false;

                    //this.EndDate.Checked = false;
                    //this.EndDate.Enabled = false;

                    // 2015/06/23
                    //this.cmbFyuyo.SelectedIndex = -1;
                    //this.cmbFyuyo.Enabled = false;

                    // 2019/02/14 �S�ē��͉\�ɕύX �R�����g��
                    //this.NouhinDate.Checked = false;
                    //this.NouhinDate.Enabled = false;

                    // 2015/06/23
                    //this.cmbNkeitai.SelectedIndex = -1;
                    //this.cmbNkeitai.Enabled = false;

                    // 2015/06/23
                    //this.cmbHjiki.SelectedIndex = -1;
                    //this.cmbHjiki.Enabled = false;

                    // 2015/06/23
                    //this.cmbHseido.SelectedIndex = -1;
                    //this.cmbHseido.Enabled = false;

                    // 2015/06/23
                    //this.cmbHhouhou.SelectedIndex = -1;
                    //this.cmbHhouhou.Enabled = false;

                    // 2015/06/23
                    //this.txtEmail.Text = "";
                    //this.txtEmail.Enabled = false;

                    //���z�z���v�`�F�b�N��
                    this.checkBox2.Checked = false;
                    this.checkBox2.Enabled = false;

                    //�}�ԗv�`�F�b�N��
                    this.checkBox3.Checked = false;
                    this.checkBox3.Enabled = false;

                    // �Č���� 2015/06/30
                    if (cmbAnShu.Items.Count > 0)
                    {
                        cmbAnShu.SelectedIndex = 2;
                    }

                    // ���z���O���͕s�Ƃ���@2015/11/27
                    chkHeihai.Enabled = false;
                }
                else
                {
                    this.cmbFkeitai.SelectedIndex = -1;
                    this.cmbFkeitai.Enabled = true;

                    this.cmbFjyouken.SelectedIndex = -1;
                    this.cmbFjyouken.Enabled = true;

                    this.cmbSize.SelectedIndex = -1;
                    this.cmbSize.Enabled = true;

                    this.txtHTanka.Text = "0";
                    this.txtHTanka.Enabled = true;

                    this.StartDate.Checked = false;
                    this.StartDate.Enabled = true;

                    this.EndDate.Checked = false;
                    this.EndDate.Enabled = true;

                    // 2015/06/23
                    //this.cmbFyuyo.SelectedIndex = -1;
                    //this.cmbFyuyo.Enabled = true;

                    // 2019/02/14 �S�ē��͉\�ɕύX �R�����g��
                    //this.NouhinDate.Checked = false;
                    //this.NouhinDate.Enabled = true;

                    // 2015/06/23
                    //this.cmbNkeitai.SelectedIndex = -1;
                    //this.cmbNkeitai.Enabled = true;

                    // 2015/06/23
                    //this.cmbHjiki.SelectedIndex = -1;
                    //this.cmbHjiki.Enabled = true;

                    // 2015/06/23
                    //this.cmbHseido.SelectedIndex = -1;
                    //this.cmbHseido.Enabled = true;

                    // 2015/06/23
                    //this.cmbHhouhou.SelectedIndex = -1;
                    //this.cmbHhouhou.Enabled = true;

                    // 2015/06/23
                    //this.txtEmail.Text = "";
                    //this.txtEmail.Enabled = true;

                    //���z�z���v�`�F�b�N��
                    this.checkBox2.Checked = false;
                    this.checkBox2.Enabled = true;

                    //�}�ԗv�`�F�b�N��
                    this.checkBox3.Checked = false;
                    this.checkBox3.Enabled = true;

                    // �Č���� 2015/06/30
                    if (cmbAnShu.Items.Count > 0)
                    {
                        cmbAnShu.SelectedIndex = 0;
                    }

                    // ���z���O���͉Ƃ���@2015/11/27
                    chkHeihai.Enabled = true;
                }
            }
        }

        private void cmbeGaichu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbpGaichu.Items.Count > 0)
            {
                if (fMode.Mode == 0)
                {
                    //cmbeGaichu.SelectedIndex = cmbeGaichu.SelectedIndex;
                    cmbpGaichu.SelectedIndex = cmbeGaichu.SelectedIndex;
                }
            }
        }

        private void dteGaichuPay_ValueChanged(object sender, EventArgs e)
        {
            if (fMode.Mode == 0)
            {
                dtpGaichuPay.Value = dteGaichuPay.Value;
            }
        }

        private void txteGaichuGenka_TextChanged(object sender, EventArgs e)
        {
            if (fMode.Mode == 0)
            {
                txtpGaichuGenka.Text = txteGaichuGenka.Text;
            }

            // �c�ƌ����ƊO���挴���̓C�R�[���Ƃ���i�O���挴���Q�A�R���Ȃ��Ƃ��j 2019/02/21
            if ((Utility.strToInt(txtpGaichuGenka2.Text) == global.FLGOFF) &&
                (Utility.strToInt(txtpGaichuGenka3.Text) == global.FLGOFF))
            {
                txtpGaichuGenka.Text = txteGaichuGenka.Text;
            }

            // �e�����z�F2019/02/21
            txteGaichuArari.Text = (Utility.strToInt(txtUri.Text) - Utility.strToInt(txteGaichuGenka.Text)).ToString("#,##0");

        }

        private void txtTanka_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void txtMai_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void rBtnOrderNum1_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnOrderNum1.Checked)
            {
                btnOrderNum.Enabled = false;
                lblOrderNum.Enabled = false;
            }
            else if (rBtnOrderNum2.Checked)
            {
                btnOrderNum.Enabled = true;
                lblOrderNum.Enabled = true;
            }
        }

        private void btnOrderNum_Click(object sender, EventArgs e)
        {
            frmGetOrderNumber frm = new frmGetOrderNumber();
            frm.ShowDialog();

            if (frm._orderNum != string.Empty)
            {
                lblOrderNum.Text = frm._orderNum;
            }
            frm.Dispose();
        }

        private void label13_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // �O����}�X�^�[�ێ�
            frmGaichu frm = new frmGaichu();
            frm.ShowDialog();

            // �O����R���{�{�b�N�X�A�C�e�����[�h
            int idx = cmbeGaichu.SelectedIndex;
            Utility.comboGaichu.itemLoad(cmbeGaichu);
            cmbeGaichu.SelectedIndex = idx;
        }

        private void label36_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // �O����}�X�^�[�ێ�
            frmGaichu frm = new frmGaichu();
            frm.ShowDialog();

            // �O����R���{�{�b�N�X�A�C�e�����[�h
            int idx = cmbpGaichu.SelectedIndex;
            Utility.comboGaichu.itemLoad(cmbpGaichu);
            cmbpGaichu.SelectedIndex = idx;

            //// �����p�O����R���{�{�b�N�X�A�C�e�����[�h
            //Utility.comboGaichu.itemLoad(cmbsGaichu);
        }

        private void label50_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void dtGaichuIrai_EnabledChanged(object sender, EventArgs e)
        {

        }

        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.��TableAdapter.FillBy(this.darwinDataSet.��);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillBy1ToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.��TableAdapter.FillBy1(this.darwinDataSet.��);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void cmbpGaichu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu.SelectedIndex >= 0)
                {
                    lblSD1.Text = cg[cmbpGaichu.SelectedIndex].shiharaibi.ToString() + "��";
                }
                else
                {
                    lblSD1.Text = string.Empty;
                }

                txtaGai1.Text = cmbpGaichu.Text;
            }
        }

        private void cmbpGaichu2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu2.SelectedIndex >= 0)
                {
                    lblSD2.Text = cg[cmbpGaichu2.SelectedIndex].shiharaibi.ToString() + "��";
                }
                else
                {
                    lblSD2.Text = string.Empty;
                }

                txtaGai2.Text = cmbpGaichu2.Text;
            }
        }

        private void cmbpGaichu3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu3.SelectedIndex >= 0)
                {
                    lblSD3.Text = cg[cmbpGaichu3.SelectedIndex].shiharaibi.ToString() + "��";
                }
                else
                {
                    lblSD3.Text = string.Empty;
                }

                txtaGai3.Text = cmbpGaichu3.Text;
            }
        }

        private void cmbpGaichu3_TextChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu3.Text == string.Empty)
                {
                    lblSD3.Text = string.Empty;
                }
            }
        }

        private void cmbpGaichu_TextChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu.Text == string.Empty)
                {
                    lblSD1.Text = string.Empty;
                }
            }
        }

        private void cmbpGaichu2_TextChanged(object sender, EventArgs e)
        {
            if (cg != null)
            {
                if (cmbpGaichu2.Text == string.Empty)
                {
                    lblSD2.Text = string.Empty;
                }
            }
        }

        private void lnkCopy_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmOrderCopy frm = new frmOrderCopy(dts, jAdp, fMode.jID);
            frm.ShowDialog();

            bool fs = frm.fStatus;
            frm.Dispose();

            DispClear();

            if (frm.fStatus)
            {
                Cursor = Cursors.WaitCursor;
                jAdp.Update(dts.��1);
                dataGridView1.DataSource = null;
                jAdp.Fill(dts.��1);
                Cursor = Cursors.Default;
            }
        }

        private void txtpGaichuGenka3_TextChanged(object sender, EventArgs e)
        {
            // ������z
            txtaUri.Text = txtUri.Text;

            // �O���e���F2019/02/21
            int gai = Utility.strToInt(txtpGaichuGenka.Text) + Utility.strToInt(txtpGaichuGenka2.Text) + Utility.strToInt(txtpGaichuGenka3.Text);
            int arari = Utility.strToInt(txtUri.Text) - gai;
            txtaArari.Text = arari.ToString("#,##0");

            tabControl1.TabPages[3].Text = "�e���F" + arari.ToString("#,##0");

            // �O������
            txtaGaiGenka1.Text = txtpGaichuGenka.Text;
            txtaGaiGenka2.Text = txtpGaichuGenka2.Text;
            txtaGaiGenka3.Text = txtpGaichuGenka3.Text;
        }

        private void txtMai_TextChanged(object sender, EventArgs e)
        {
            // �c�ƌ����F�|�X�e�B���O�̂Ƃ� 2019/03/16
            if (cmbNaiyou.Text == J_NAIYOU_POSTING)
            {
                txteGaichuGenka.Text = getGenka().ToString("#,##0");
            }
        }

        private void txtHTanka_TextChanged(object sender, EventArgs e)
        {
            // �c�ƌ����F�|�X�e�B���O�̂Ƃ� 2019/03/16
            if (cmbNaiyou.Text == J_NAIYOU_POSTING)
            {
                txteGaichuGenka.Text = getGenka().ToString("#,##0");
            }
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     �c�ƌ��������߂�F2019/03/16 </summary>
        /// <returns>
        ///     �c�ƌ���</returns>
        ///-------------------------------------------------------
        private decimal getGenka()
        {
            // �z�z�P���擾
            String str = Utility.strToDouble(txtHTanka.Text).ToString();

            decimal dHTanka = 0;

            if (!decimal.TryParse(str, out dHTanka))
            {
                dHTanka = 0m;
            }

            decimal genka = Math.Floor(dHTanka * Utility.strToDecimal(txtMai.Text.Replace(",","")));

            return genka;
        }

        private void dtSeikyu_ValueChanged(object sender, EventArgs e)
        {
            //
            //  ����ŎZ�����𐿋������Ƃ�������
            //      ���t�̕ύX�ɔ�������Ŋz�A�ō��݋��z�������ύX����
            //      2019-08-26
            //

            DateTime dt;

            if (dtSeikyu.Checked)
            {
                if (!DateTime.TryParse(dtSeikyu.Value.ToShortDateString(), out dt))
                {
                    return;
                }

                decimal nebikigoKin = Utility.strToDecimal(txtNebikigo.Text);
                decimal KingakuTax = 0;
                decimal KingakuZeikomi = 0;
                decimal KingakuTL = 0;

                // ���z�v�Z
                UriageSum(dt, nebikigoKin, out KingakuTax, out KingakuZeikomi, out KingakuTL);

                // ����Ŋz�v�Z               
                txtTax.Text = KingakuTax.ToString("#,##0");

                // �ō����z
                txtZeikomi.Text = KingakuZeikomi.ToString("#,##0");
            }
            else
            {
                // ����Ŋz�v�Z               
                txtTax.Text = global.FLGOFF.ToString();

                // �ō����z
                txtZeikomi.Text = Utility.strToDecimal(txtNebikigo.Text).ToString("#,##0");
            }
        }
    }
}
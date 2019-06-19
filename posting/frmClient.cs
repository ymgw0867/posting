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
    public partial class frmClient : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        //Entity.���Ӑ� cMaster = new Entity.���Ӑ�();

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.���Ӑ�TableAdapter cAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
        darwinDataSetTableAdapters.��1TableAdapter jAdp = new darwinDataSetTableAdapters.��1TableAdapter();
        darwinDataSetTableAdapters.������TableAdapter rAdp = new darwinDataSetTableAdapters.������TableAdapter();
        darwinDataSetTableAdapters.�Ј�TableAdapter sAdp = new darwinDataSetTableAdapters.�Ј�TableAdapter();
        //darwinDataSetTableAdapters.��Џ��TableAdapter kAdp = new darwinDataSetTableAdapters.��Џ��TableAdapter();

        const string MESSAGE_CAPTION = "���Ӑ�}�X�^�[";

        string[] zipArray = null;

        bool jStatus = false;
        bool sStatus = false;

        public frmClient()
        {
            InitializeComponent();

            // �f�[�^�ǂݍ��� 2015/07/05
            cAdp.Fill(dts.���Ӑ�);
            //jAdp.Fill(dts.��1);
            //rAdp.Fill(dts.������);
            sAdp.Fill(dts.�Ј�);
            //kAdp.Fill(dts.��Џ��);
        }

        private void form_Load(object sender, EventArgs e)
        {
            // �����於�̃Z�b�g
            setSeikyuName();

            // �E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // TODO: ���̃R�[�h�s�̓f�[�^�� 'darwinDataSet.���Ӑ�' �e�[�u���ɓǂݍ��݂܂��B�K�v�ɉ����Ĉړ��A�܂��͍폜�����Ă�������
            Setting(dataGridView1);
            gridSearch(dataGridView1);
            //this.���Ӑ�TableAdapter.Fill(this.darwinDataSet.���Ӑ�);

            // �h�̃R���{
            cmbKeishoSet();
            cmbKeishoSSet();    // ������h�� 2019/02/20

            // �Œʒm�R���{
            cmbTaxSet();

            // �s���{���R���{
            cmbCitySet(cmbCity);
            cmbCitySet(cmbCityS);

            // �S���Ј��R���{
            Utility.ComboShain.load(cmbShain);
            
            // �f�[�^������
            DispClear();

            // �X�֔ԍ�CSV�ǂݍ���
            Utility.zipCsvLoad(ref zipArray);

            // ������E�h�́A�������Z�b�g�c�[���{�^���\�� 2019/04/02
            DateTime dt = new DateTime(2019, 04, 30);
            if (DateTime.Today > dt)
            {
                button2.Visible = false;
            }
            else
            {
                button2.Visible = true;
            }
        }

        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     �����於�̃Z�b�g null�A�܂��͋󔒂̂Ƃ����̂��Z�b�g����</summary>
        /// -----------------------------------------------------------------------------------
        private void setSeikyuName()
        {
            int cnt = 0;

            foreach (var t in dts.���Ӑ�)
            {
                //if (t.Is�����於��Null())
                //{
                //    t.�����於�� = t.����;
                //    cnt++;
                //}
                //else if (t.�����於��.Trim() == string.Empty)
                //{
                //    t.�����於�� = t.����;
                //    cnt++;
                //}

                if (t.�����於��.Trim() == string.Empty)
                {
                    t.�����於�� = t.����;
                    cnt++;
                }
            }

            if (cnt > 0)
            {
                cAdp.Update(dts.���Ӑ�);
                cAdp.Fill(dts.���Ӑ�);
            }
        }


        #region �O���b�h�r���[�J������`
        string colID = "col1";
        string colRyaku = "col2";
        string colFuri = "col3";
        string colName = "col4";
        string colKeisho = "col5";
        string colTantou = "col6";
        string colBusho = "col7";
        string colZip = "col8";
        string colCity = "col9";
        string colAdd1 = "col10";
        string colAdd2 = "col11";
        string colTel = "col12";
        string colFax = "col13";
        string colEmail = "col14";
        string colEigyo = "col15";
        string colShimebi = "col16";
        string colZei = "col17";
        string colSeName = "col18";
        string colSeZip = "col19";
        string colSeCity = "col20";
        string colSeAdd1 = "col21";
        string colSeAdd2 = "col22";
        string colBiko = "col23";
        string colAddDt = "col24";
        string colUpDt = "col25";
        string colTantouS = "col26";
        #endregion

        /// -------------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̒�`���s���܂� </summary>
        /// <param name="tempDGV">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        /// -------------------------------------------------------------------
        private void Setting(DataGridView tempDGV)
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
                tempDGV.Height = 183;

                // ��s�̐F
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // �e�񕝎w��
                tempDGV.Columns.Add(colID, "����");
                tempDGV.Columns.Add(colRyaku, "����");
                tempDGV.Columns.Add(colFuri, "�t���K�i");
                tempDGV.Columns.Add(colName, "����");
                tempDGV.Columns.Add(colKeisho, "�h��");
                tempDGV.Columns.Add(colTantou, "�S���Җ�");
                tempDGV.Columns.Add(colBusho, "������");
                tempDGV.Columns.Add(colZip, "�X�֔ԍ�");
                tempDGV.Columns.Add(colCity, "�s���{��");
                tempDGV.Columns.Add(colAdd1, "�Z���P");
                tempDGV.Columns.Add(colAdd2, "�Z���Q");
                tempDGV.Columns.Add(colTel, "�d�b�ԍ�");
                tempDGV.Columns.Add(colFax, "FAX�ԍ�");
                tempDGV.Columns.Add(colEmail, "eMail");
                tempDGV.Columns.Add(colEigyo, "�S���c��");
                tempDGV.Columns.Add(colShimebi, "����");
                tempDGV.Columns.Add(colZei, "�Œʒm");
                tempDGV.Columns.Add(colSeName, "�����於��");
                tempDGV.Columns.Add(colSeZip, "�����恧");
                tempDGV.Columns.Add(colSeCity, "������s���{��");
                tempDGV.Columns.Add(colSeAdd1, "������P");
                tempDGV.Columns.Add(colSeAdd2, "������Q");
                tempDGV.Columns.Add(colTantouS, "������S����");
                tempDGV.Columns.Add(colBiko, "���l");
                tempDGV.Columns.Add(colAddDt, "�o�^�N����");
                tempDGV.Columns.Add(colUpDt, "�X�V�N����");

                tempDGV.Columns[0].Width = 70;
                tempDGV.Columns[1].Width = 200;
                tempDGV.Columns[2].Width = 200;
                tempDGV.Columns[3].Width = 200;
                tempDGV.Columns[4].Width = 60;
                tempDGV.Columns[5].Width = 100;
                tempDGV.Columns[6].Width = 160;
                tempDGV.Columns[7].Width = 100;
                tempDGV.Columns[8].Width = 120;
                tempDGV.Columns[9].Width = 200;
                tempDGV.Columns[10].Width = 200;

                tempDGV.Columns[11].Width = 120;
                tempDGV.Columns[12].Width = 120;
                tempDGV.Columns[13].Width = 160;
                tempDGV.Columns[14].Width = 100;
                tempDGV.Columns[15].Width = 60;
                tempDGV.Columns[16].Width = 80;
                tempDGV.Columns[17].Width = 120;
                tempDGV.Columns[18].Width = 120;
                tempDGV.Columns[19].Width = 200;
                tempDGV.Columns[20].Width = 200;
                tempDGV.Columns[21].Width = 200;
                tempDGV.Columns[22].Width = 110;
                tempDGV.Columns[23].Width = 120;
                tempDGV.Columns[24].Width = 140;
                tempDGV.Columns[25].Width = 140;

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

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾���� </summary>
        /// <param name="dgv">
        ///     �ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        /// -----------------------------------------------------------------------------
        private Boolean GetData(DataGridView dgv,ref Entity.���Ӑ� tempC, darwinDataSet dts, int sID)
        {
            //foreach (var t in dts.���Ӑ�.Where(a => a.ID == sID))
            //{

            //}

            int iX = 0;
            string sqlStr;
            Control.���Ӑ� Client = new Control.���Ӑ�();
            OleDbDataReader dr;

            sqlStr = " where ���Ӑ�.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
            dr = Client.FillBy(sqlStr);

            if (dr.HasRows == true)
            {
                while (dr.Read() == true)
                {
                    tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                    tempC.���� = dr["����"].ToString() + "";
                    tempC.�t���K�i = dr["�t���K�i"].ToString();
                    tempC.���� = dr["����"].ToString();
                    tempC.�h�� = dr["�h��"].ToString();
                    tempC.�S���Җ� = dr["�S���Җ�"].ToString();
                    tempC.������ = dr["������"].ToString();
                    tempC.�S���Җ� = dr["�S���Җ�"].ToString();
                    tempC.�X�֔ԍ� = dr["�X�֔ԍ�"].ToString();
                    tempC.�s���{�� = dr["�s���{��"].ToString();
                    tempC.�Z��1 = dr["�Z��1"].ToString();
                    tempC.�Z��2 = dr["�Z��2"].ToString();
                    tempC.�d�b�ԍ� = dr["�d�b�ԍ�"].ToString();
                    tempC.FAX�ԍ� = dr["FAX�ԍ�"].ToString();
                    tempC.���[���A�h���X = dr["���[���A�h���X"].ToString();
                    tempC.�S���Ј��R�[�h = Int32.Parse(dr["�S���Ј��R�[�h"].ToString());
                    tempC.���� = Int32.Parse(dr["����"].ToString());
                    tempC.�Œʒm = dr["�Œʒm"].ToString();
                    tempC.������X�֔ԍ� = dr["������X�֔ԍ�"].ToString();
                    tempC.������s���{�� = dr["������s���{��"].ToString();
                    tempC.������Z��1 = dr["������Z��1"].ToString();
                    tempC.������Z��2 = dr["������Z��2"].ToString();
                    tempC.���l = dr["���l"].ToString();
                    tempC.������S���Җ� = dr["������S���Җ�"].ToString();   // 2015/11/20
                }
            }
            else
            {
                dr.Close();
                Client.Close();
                return false;
            }

            dr.Close();
            Client.Close();
            return true;
        }

        //�O���b�h����f�[�^��I��
        private void GridEnter()
        {
            int iX = 0;
            string msgStr;

            msgStr = "";
            msgStr += dataGridView1[1, dataGridView1.SelectedRows[iX].Index].Value + "���I������܂���" + "\n";
            msgStr += "��낵���ł����H";

            if (MessageBox.Show(msgStr, "���Ӑ�I��", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No)
            {
                return;
            }
            else
            {
                try
                {
                    // �f�[�^���擾����
                    int sID = int.Parse(dataGridView1[0, dataGridView1.SelectedRows[iX].Index].Value.ToString());
                    if (dts.���Ӑ�.Any(a => a.ID == sID))
                    {
                        darwinDataSet.���Ӑ�Row r = dts.���Ӑ�.Single(a => a.ID == sID);

                        //'�f�[�^��\�����܂� 2015/07/05
                        //txtCode.Text = cMaster.ID.ToString();
                        txtName1.Text = r.����;
                        txtFuri.Text = r.�t���K�i;
                        txtName2.Text = r.����;
                        cmbKeisho.Text = r.�h��;
                        txtTantou.Text = r.�S���Җ�;
                        txtBusho.Text = r.������;
                        mtxtZipCode.Text = r.�X�֔ԍ�;
                        cmbCity.Text = r.�s���{��;
                        txtAddress1.Text = r.�Z��1;
                        txtAddress2.Text = r.�Z��2;
                        txtTel.Text = r.�d�b�ԍ�;
                        txtFax.Text = r.FAX�ԍ�;
                        txtEmail.Text = r.���[���A�h���X;

                        Utility.ComboShain.selectedIndex(cmbShain, Int32.Parse(r.�S���Ј��R�[�h.ToString()));

                        txtShimebi.Text = r.����.ToString();
                        cmbTax.Text = r.�Œʒm.ToString();

                        //if (r.Is�����於��Null())
                        //{
                        //    txtNameSeikyu.Text = string.Empty;
                        //}
                        //else
                        //{
                        //    txtNameSeikyu.Text = r.�����於��;
                        //}

                        txtNameSeikyu.Text = r.�����於��;

                        mtxtZipCodeS.Text = r.������X�֔ԍ�;
                        cmbCityS.Text = r.������s���{��;
                        txtAddress1S.Text = r.������Z��1;
                        txtAddress2S.Text = r.������Z��2;

                        // 2015/11/20
                        if (r.������S���Җ� != null)
                        {
                            txtTantouS.Text = r.������S���Җ�;
                        }
                        else
                        {
                            txtTantouS.Text = string.Empty;
                        }

                        txtMemo.Text = r.���l;

                        txtBushoS.Text = r.�����敔����;  // 2019/02/20
                        cmbKeishoS.Text = r.������h��;  // 2019/02/20
                        

                        //ID�e�L�X�g�{�b�N�X�͕ҏW�s�Ƃ���
                        //txtCode.Enabled = false;

                        //�{�^�����
                        btnDel.Enabled = true;
                        btnClr.Enabled = true;

                        fMode.Mode = 1;     // �t�H�[�����[�h�X�e�[�^�X:�ύX�폜
                        fMode.ID = r.ID;    // ���Ӑ�ID

                        txtName1.Focus();

                        button2.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("�Y������f�[�^���}�X�^�[�ɓo�^����Ă��܂���", "�����G���[", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                }
                catch (Exception e)
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

                //txtCode.Enabled = true;
                //txtCode.Text = "";

                txtName1.Text = "";
                txtFuri.Text = "";
                txtName2.Text = "";
                cmbKeisho.Text = "";
                txtTantou.Text = "";
                txtBusho.Text = "";
                mtxtZipCode.Text = "";
                cmbCity.Text = "";
                txtAddress1.Text = "";
                txtAddress2.Text = "";
                txtTel.Text = "";
                txtFax.Text = "";
                txtEmail.Text = "";
                cmbShain.SelectedIndex = -1;
                txtShimebi.Text = "0";
                cmbTax.Text = "";

                txtNameSeikyu.Text = "";    // 2015/07/05
                mtxtZipCodeS.Text = "";
                cmbCityS.Text = "";
                txtAddress1S.Text = "";
                txtAddress2S.Text = "";
                txtTantouS.Text = "";   // 2015/11/20
                txtMemo.Text = "";

                txtBushoS.Text = "";    // 2019/02/20
                cmbKeishoS.Text = "";   // 2019/02/20
                cmbKeishoS.SelectedIndex = -1;  // 2019/02/20

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

                //txtCode.Focus();
                txtName1.Focus();

                button2.Enabled = true; // 2019/04/02
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
                        case 0: // �V�K�o�^
                            if (MessageBox.Show("�V�K�o�^���܂��B��낵���ł����H", "�o�^�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                return;
                            }

                            dts.���Ӑ�.Add���Ӑ�Row(setClientRow(dts.���Ӑ�.New���Ӑ�Row()));
                            break;

                        case 1: // �X�V
                            if (MessageBox.Show("�X�V���܂��B��낵���ł����H", "�X�V�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {
                                return;
                            }

                            darwinDataSet.���Ӑ�Row r = dts.���Ӑ�.Single(a => a.ID == fMode.ID);

                            if (!r.HasErrors)
                            {
                                setClientRow(r);
                            }
                            else
                            {
                                MessageBox.Show(fMode.ID + "���L�[�s�݂ł��F�f�[�^�̍X�V�Ɏ��s���܂���", "�X�V�G���[", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }

                            break;
                    }

                    // �f�[�^�x�[�X�X�V
                    cAdp.Update(dts.���Ӑ�);

                    // ��ʏ�����
                    DispClear();

                    //�f�[�^�� 'darwinDataSet.���Ӑ�' �e�[�u���ɓǂݍ��݂܂��B
                    //dataGridView1.DataSource = null;
                    cAdp.Fill(dts.���Ӑ�);
                    gridSearch(dataGridView1);

                    //this.���Ӑ�TableAdapter.Fill(this.darwinDataSet.���Ӑ�);
                    //dataGridView1.DataSource = this.darwinDataSet.���Ӑ�;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"�X�V����",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     darwinDataSet.���Ӑ�Row�s�f�[�^�쐬 </summary>
        /// <param name="r">
        ///     darwinDataSet.���Ӑ�Row �I�u�W�F�N�g</param>
        /// <returns>
        ///     darwinDataSet.���Ӑ�Row �I�u�W�F�N�g</returns>
        /// ------------------------------------------------------------------
        private darwinDataSet.���Ӑ�Row setClientRow(darwinDataSet.���Ӑ�Row r)
        {
            r.���� = txtName1.Text;
            r.�t���K�i = txtFuri.Text;
            r.���� = txtName2.Text;
            r.�h�� = cmbKeisho.Text;
            r.�S���Җ� = txtTantou.Text;
            r.������ = txtBusho.Text;

            if (((mtxtZipCode.Text).Replace("-", "")).Trim() == "")
            {
                r.�X�֔ԍ� = "";
            }
            else
            {
                r.�X�֔ԍ� = mtxtZipCode.Text;
            }

            r.�s���{�� = cmbCity.Text;
            r.�Z��1 = txtAddress1.Text;
            r.�Z��2 = txtAddress2.Text;
            r.�d�b�ԍ� = txtTel.Text;
            r.FAX�ԍ� = txtFax.Text;
            r.���[���A�h���X = txtEmail.Text;

            Utility.ComboShain cmb1 = new Utility.ComboShain();

            if (cmbShain.SelectedIndex == -1)
            {
                r.�S���Ј��R�[�h = 0;
            }
            else
            {
                cmb1 = (Utility.ComboShain)cmbShain.SelectedItem;
                r.�S���Ј��R�[�h = cmb1.ID;
            }

            r.���� = Int32.Parse(txtShimebi.Text);
            r.�Œʒm = cmbTax.Text.ToString();

            if (((mtxtZipCodeS.Text).Replace("-", "")).Trim() == "")
            {
                r.������X�֔ԍ� = "";
            }
            else
            {
                r.������X�֔ԍ� = mtxtZipCodeS.Text;
            }

            r.������s���{�� = cmbCityS.Text;
            r.������Z��1 = txtAddress1S.Text;
            r.������Z��2 = txtAddress2S.Text;
            r.������S���Җ� = txtTantouS.Text;

            r.���l = txtMemo.Text;

            if (fMode.Mode == 0) r.�o�^�N���� = DateTime.Now;

            r.�ύX�N���� = DateTime.Now;
            r.�����於�� = txtNameSeikyu.Text;   // 2015/07/05

            r.�����敔���� = txtBushoS.Text;  // 2019/02/20
            r.������h�� = cmbKeishoS.Text;  // 2019/02/20

            return r;
        }

        // �o�^�f�[�^�`�F�b�N
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
                    //Control.���Ӑ� Client = new Control.���Ӑ�();
                    //OleDbDataReader dr;

                    //sqlStr = " where ID = " + txtCode.Text.ToString();
                    //dr = Client.FillBy(sqlStr);

                    //if (dr.HasRows == true)
                    //{
                    //    txtCode.Focus();
                    //    dr.Close();
                    //    Client.Close();
                    //    throw new Exception("���ɓo�^�ς݂̃R�[�h�ł�");
                    //}

                    //dr.Close();
                    //Client.Close();

                }

                //���̃`�F�b�N
                if (txtName1.Text.Trim().Length < 1)
                {
                    txtName1.Focus();
                    throw new Exception("���̂���͂��Ă�������");
                }


                //�����F�������H
                if (txtShimebi.Text == null)
                {
                    this.txtShimebi.Focus();
                    throw new Exception("�����͐����œ��͂��Ă�������");
                }

                str = this.txtShimebi.Text;

                if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d))
                {
                }
                else
                {
                    this.txtShimebi.Focus();
                    throw new Exception("�����͐����œ��͂��Ă�������");
                }

                ////�[���͕s��
                //if (Convert.ToInt32(this.txtCode.Text.ToString()) == 0)
                //{
                //    this.txtCode.Focus();
                //    throw new Exception("�[���͓o�^�ł��܂���");
                //}

                //�N���X�Ƀf�[�^�Z�b�g
                //cMaster.ID = Convert.ToInt32(txtCode.Text.ToString());

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

            if (sender == txtName1)
            {
                objtxt = txtName1;

                // 2019/03/16
                MyTextKana.TextKana.textEnter(txtName1);
            }

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

            if (sender == txtShimebi)
            {
                objtxt = txtShimebi;
            }

            if (sender == mtxtZipCodeS)
            {
                objMtxt = mtxtZipCodeS;
            }

            if (sender == txtAddress1S)
            {
                objtxt = txtAddress1S;
            }

            if (sender == txtAddress2S)
            {
                objtxt = txtAddress2S;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }

            // ������S���Җ� 2015/11/20
            if (sender == txtTantouS)
            {
                objtxt = txtTantouS;
            }

            // �����於�� 2015/07/05
            if (sender == txtNameSeikyu)
            {
                objtxt = txtNameSeikyu;
            }

            // �����d�b�ԍ� 2015/07/05
            if (sender == txtsTel)
            {
                objtxt = txtsTel;
            }

            // �������ԍ� 2015/07/05
            if (sender == txtsZip)
            {
                objtxt = txtsZip;
            }

            // ���������於�� 2015/07/05
            if (sender == textBox2)
            {
                objtxt = textBox2;
            }

            // ���������敔���� 2019/02/20
            if (sender == txtBushoS)
            {
                objtxt = txtBushoS;
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

            // ������S���Җ� 2015/11/20
            if (sender == txtTantouS)
            {
                objtxt = txtTantouS;
            }

            if (sender == txtName1)
            {
                objtxt = txtName1;

                // 2019/03/16
                MyTextKana.TextKana.textLeave(txtName1);
            }

            if (sender == txtFuri)
            {
                objtxt = txtFuri;
            }

            if (sender == txtName2)
            {
                objtxt = txtName2;

                //�����於�̂փR�s�[ 2015/07/05
                if (txtNameSeikyu.Text == "")
                {
                    txtNameSeikyu.Text = txtName2.Text;
                }
            }

            if (sender == txtTantou)
            {
                objtxt = txtTantou;
            }

            if (sender == txtBusho)
            {
                objtxt = txtBusho;

                // ������փR�s�[ : 2019/02/20
                if (txtBushoS.Text == "")
                {
                    txtBushoS.Text = txtBusho.Text;
                }
            }

            if (sender == mtxtZipCode)
            {
                objMtxt = mtxtZipCode;

                //������փR�s�[
                if (mtxtZipCodeS.Text.Trim() == "-")
                {
                    mtxtZipCodeS.Text = mtxtZipCode.Text;
                }
            }

            if (sender == cmbCity)
            {
                //������փR�s�[
                if (cmbCityS.Text == "")
                {
                    cmbCityS.Text = cmbCity.Text;
                }
            }

            if (sender == txtAddress1)
            {
                objtxt = txtAddress1;

                //������փR�s�[
                if (txtAddress1S.Text == "")
                {
                    txtAddress1S.Text = txtAddress1.Text;
                }
            }

            if (sender == txtAddress2)
            {
                objtxt = txtAddress2;

                //������փR�s�[
                if (txtAddress2S.Text == "")
                {
                    txtAddress2S.Text = txtAddress2.Text;
                }
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

            if (sender == txtShimebi)
            {
                objtxt = txtShimebi;
            }

            if (sender == mtxtZipCodeS)
            {
                objMtxt = mtxtZipCodeS;
            }

            if (sender == txtAddress1S)
            {
                objtxt = txtAddress1S;
            }

            if (sender == txtAddress2S)
            {
                objtxt = txtAddress2S;
            }

            if (sender == txtMemo)
            {
                objtxt = txtMemo;
            }

            if (sender == textBox1)
            {
                objtxt = textBox1;
            }
            
            // �����於�� 2015/07/05
            if (sender == txtNameSeikyu)
            {
                objtxt = txtNameSeikyu;
            }

            // �����d�b�ԍ� 2015/07/05
            if (sender == txtsTel)
            {
                objtxt = txtsTel;
            }

            // �������ԍ� 2015/07/05
            if (sender == txtsZip)
            {
                objtxt = txtsZip;
            }

            // ���������於�� 2015/07/05
            if (sender == textBox2)
            {
                objtxt = textBox2;
            }

            // ���������敔���� 2019/02/20
            if (sender == txtBushoS)
            {
                objtxt = txtBushoS;
            }

            objtxt.BackColor = Color.White;
            objMtxt.BackColor = Color.White;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            // �폜�m�F�F2018/01/03
            if (MessageBox.Show("�폜�������I������܂����B���s���Ă�낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // �f�[�^�ǂݍ��݁F2018/01/03
            if (!jStatus)
            {
                Cursor = Cursors.WaitCursor;
                jAdp.Fill(dts.��1);
                jStatus = true;
                Cursor = Cursors.Default;
            }

            // �󒍃f�[�^�ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���
            if (dts.��1.Any(a => a.���Ӑ�ID == fMode.ID))
            {
                MessageBox.Show(txtName1.Text.ToString() + "�̎󒍃f�[�^���o�^����Ă��܂�", txtName1.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // �f�[�^�ǂݍ��݁F2018/01/03
            if (!sStatus)
            {
                Cursor = Cursors.WaitCursor;
                rAdp.Fill(dts.������);
                sStatus = true;
                Cursor = Cursors.Default;
            }

            // �����f�[�^�ɓo�^����Ă���Ƃ��͍폜�s�Ƃ���
            if (dts.������.Any(a => a.���Ӑ�ID == fMode.ID))
            {
                MessageBox.Show(txtName1.Text.ToString() + "�̐����f�[�^���o�^����Ă��܂�", txtName1.Text.ToString() + "�͍폜�ł��܂���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //�폜�m�F
            if (MessageBox.Show("�폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            //�f�[�^�폜
            try
            {
                darwinDataSet.���Ӑ�Row r = dts.���Ӑ�.Single(a => a.ID == fMode.ID);
                if (!r.HasErrors)
                {
                    r.Delete();

                    // �f�[�^�x�[�X�X�V
                    cAdp.Update(dts.���Ӑ�);

                    MessageBox.Show("�폜����܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // �f�[�^�� 'darwinDataSet.���Ӑ�' �e�[�u���ɓǂݍ��݂܂��B
                    cAdp.Fill(dts.���Ӑ�);
                    gridSearch(dataGridView1);
                }
                else
                {
                    MessageBox.Show("�폜�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show("�폜�Ɏ��s���܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // ��ʏ�����
            DispClear();
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

        private void cmbKeishoSet()
        {
            cmbKeisho.Items.Clear();
            cmbKeisho.Items.Add("�l");
            cmbKeisho.Items.Add("�a");
            cmbKeisho.Items.Add("�䒆");
        }
        private void cmbKeishoSSet()
        {
            cmbKeishoS.Items.Clear();
            cmbKeishoS.Items.Add("�l");
            cmbKeishoS.Items.Add("�a");
            cmbKeishoS.Items.Add("�䒆");
        }

        private void cmbTaxSet()
        {
            cmbTax.Items.Clear();
            cmbTax.Items.Add("�`�[�v");
            cmbTax.Items.Add("������");
            cmbTax.Items.Add("��ې�");
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     �s���{���R���{�l�Z�b�g </summary>
        /// ------------------------------------------------------------------
        private void cmbCitySet(ComboBox tempCmb)
        {
            tempCmb.Items.Add("�k�C��");
            tempCmb.Items.Add("�X��");
            tempCmb.Items.Add("��茧");
            tempCmb.Items.Add("�{�錧");
            tempCmb.Items.Add("�H�c��");
            tempCmb.Items.Add("�R�`��");
            tempCmb.Items.Add("������");
            tempCmb.Items.Add("��錧");
            tempCmb.Items.Add("�Ȗ،�");
            tempCmb.Items.Add("�Q�n��");
            tempCmb.Items.Add("��ʌ�");
            tempCmb.Items.Add("��t��");
            tempCmb.Items.Add("�����s");
            tempCmb.Items.Add("�_�ސ쌧");
            tempCmb.Items.Add("�R����");
            tempCmb.Items.Add("���쌧");
            tempCmb.Items.Add("�V����");
            tempCmb.Items.Add("�x�R��");
            tempCmb.Items.Add("�ΐ쌧");
            tempCmb.Items.Add("���䌧");
            tempCmb.Items.Add("�򕌌�");
            tempCmb.Items.Add("�É���");
            tempCmb.Items.Add("���m��");
            tempCmb.Items.Add("�O�d��");
            tempCmb.Items.Add("���ꌧ");
            tempCmb.Items.Add("���s�{");
            tempCmb.Items.Add("���{");
            tempCmb.Items.Add("���Ɍ�");
            tempCmb.Items.Add("�ޗǌ�");
            tempCmb.Items.Add("�a�̎R��");
            tempCmb.Items.Add("���挧");
            tempCmb.Items.Add("������");
            tempCmb.Items.Add("���R��");
            tempCmb.Items.Add("�L����");
            tempCmb.Items.Add("�R����");
            tempCmb.Items.Add("������");
            tempCmb.Items.Add("���쌧");
            tempCmb.Items.Add("���Q��");
            tempCmb.Items.Add("���m��");
            tempCmb.Items.Add("������");
            tempCmb.Items.Add("���ꌧ");
            tempCmb.Items.Add("���茧");
            tempCmb.Items.Add("�F�{��");
            tempCmb.Items.Add("�啪��");
            tempCmb.Items.Add("�{�茧");
            tempCmb.Items.Add("��������");
            tempCmb.Items.Add("���ꌧ");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // �O���b�h���Ӑ�f�[�^��\������
            gridSearch(dataGridView1);
        }

        /// ------------------------------------------------------------------
        /// <summary>
        ///     �O���b�h���Ӑ�f�[�^��\������ </summary>
        /// <param name="g">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        /// ------------------------------------------------------------------
        private void gridSearch(DataGridView g)
        {
            g.Rows.Clear();
            int iX = 0;

            // ���̌���
            var s = dts.���Ӑ�.Where(a => a.����.Contains(textBox1.Text.Trim())).OrderBy(a => a.ID);

            // �d�b�ԍ�����
            s = s.Where(a => a.�d�b�ԍ�.Contains(txtsTel.Text.Trim())).OrderBy(a => a.ID);

            // �X�֔ԍ�����
            s = s.Where(a => a.�X�֔ԍ�.Contains(txtsZip.Text.Trim())).OrderBy(a => a.ID);

            // �����於�̌���
            s = s.Where(a => a.�����於��.Contains(textBox2.Text.Trim())).OrderBy(a => a.ID);

            // �O���b�h�ɕ\��
            foreach (var t in s)
            {
                g.Rows.Add();
                g[colID, iX].Value = t.ID.ToString();
                g[colRyaku, iX].Value = t.����;
                g[colFuri, iX].Value = t.�t���K�i;
                g[colName, iX].Value = t.����;
                g[colKeisho, iX].Value = t.�h��;
                g[colTantou, iX].Value = t.�S���Җ�;
                g[colBusho, iX].Value = t.������;
                g[colZip, iX].Value = t.�X�֔ԍ�;
                g[colCity, iX].Value = t.�s���{��;
                g[colAdd1, iX].Value = t.�Z��1;
                g[colAdd2, iX].Value = t.�Z��2;
                g[colTel, iX].Value = t.�d�b�ԍ�;
                g[colFax, iX].Value = t.FAX�ԍ�;
                g[colEmail, iX].Value = t.���[���A�h���X;

                if (t.�Ј�Row == null)
                {
                    g[colEigyo, iX].Value = string.Empty;
                }
                else
                {
                    g[colEigyo, iX].Value = t.�Ј�Row.����;
                }

                g[colShimebi, iX].Value = t.����.ToString();
                g[colZei, iX].Value = t.�Œʒm;

                //if (t.Is�����於��Null())
                //{
                //    g[colSeName, iX].Value = string.Empty;
                //}
                //else
                //{
                //    g[colSeName, iX].Value = t.�����於��;
                //}

                g[colSeName, iX].Value = t.�����於��;

                g[colSeCity, iX].Value = t.������s���{��;
                g[colSeZip, iX].Value = t.������X�֔ԍ�;
                g[colSeAdd1, iX].Value = t.������Z��1;
                g[colSeAdd2, iX].Value = t.������Z��2;

                if (t.������S���Җ� == null)
                {
                    g[colTantouS, iX].Value = string.Empty;
                }
                else
                {
                    g[colTantouS, iX].Value = t.������S���Җ�;
                }

                g[colBiko, iX].Value = t.���l;
                g[colAddDt, iX].Value = t.�o�^�N����;
                g[colUpDt, iX].Value = t.�ύX�N����;

                iX++;
            }

            g.CurrentCell = null;
        }
        
        private void label14_Click(object sender, EventArgs e)
        {
            Form frm = new frmShain();

            frm.ShowDialog();
            Utility.ComboShain.load(cmbShain);
        }

        private void txtName1_TextChanged(object sender, EventArgs e)
        {
            if (MyTextKana.TextKana.textChanged(txtName1.Text) == false)
            {
                txtFuri.Text = "";
            }
        }

        private void txtName1_KeyDown(object sender, KeyEventArgs e)
        {
            // 2019/03/14
           string furi = MyTextKana.TextKana.textBox_KeyDown(txtName1, sender, e);

            if (furi != "")
            {
                //txtFuri.Text = furi;  // 2019/03/16 �R�����g��
                txtFuri.Text += furi;   // 2019/03/16
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
            ComboBox cmbC = null;
            bool mc = false;

            if (sender == mtxtZipCode)
            {
                mTxt = mtxtZipCode;
                txtAdd = txtAddress1;
                cmbC = cmbCity;
            }
            else if (sender == mtxtZipCodeS)
            {
                mTxt = mtxtZipCodeS;
                txtAdd = txtAddress1S;
                cmbC = cmbCityS;
            }

            string zipText = mTxt.Text.Replace("-", "");

            if (zipText.Length == 7)
            {
                foreach (var t in zipArray)
                {
                    string[] r = t.Split(',');

                    if (zipText == r[2].Replace("\"", ""))
                    {
                        // �s���{������\��
                        for (int i = 0; i < cmbC.Items.Count; i++)
                        {
                            if (cmbC.Items[i].ToString() == r[6].Replace("\"", ""))
                            {
                                cmbC.SelectedIndex = i;
                                break;
                            }
                        }

                        // �Z��
                        string ad = r[7].Replace("\"", "") + r[8].Replace("\"", "");
                            
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

                // �Y������X�֔ԍ��Ȃ��̂Ƃ�
                if (!mc)
                {
                    cmbC.SelectedIndex = -1;
                    txtAdd.Text = string.Empty;
                }
            }
        }

        private void txtShimebi_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void cmbKeisho_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ������h�̂ɃR�s�[ 2019/02/20
            if (cmbKeishoS.SelectedIndex == -1)
            {
                cmbKeishoS.SelectedIndex = cmbKeisho.SelectedIndex;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmChargeName frm = new frmChargeName();
            frm.ShowDialog();

            this.Close();
        }
    }
}
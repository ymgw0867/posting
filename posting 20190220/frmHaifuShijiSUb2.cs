using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace posting
{
    public partial class frmHaifuShijiSUb2 : Form
    {
        Utility.formMode fMode = new Utility.formMode();
        Entity.�z�z�G���A cMaster = new Entity.�z�z�G���A();

        const string MESSAGE_CAPTION = "�z�z�w���f�[�^�ҏW";
        const int MIHAIFU_ADD = 0;      //���z�z���V�K�o�^
        const int MIHAIFU_UPDATE = 1;   //���z�z���X�V

        public frmHaifuShijiSUb2()
        {
            InitializeComponent();
        }

        private void frmHaifuShijiSUb_Load(object sender, EventArgs e)
        {

            ShowData();

            GridviewSet.Setting(dataGridView1);
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));
        }

        // �f�[�^�O���b�h�r���[�N���X
        private class GridviewSet
        {

            //�z�z�w�����׃f�[�^�O���b�h
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
                    tempDGV.Columns.Add("col1", "PID");
                    tempDGV.Columns.Add("col2", "FID");
                    tempDGV.Columns.Add("col3", "�Ԓn��");
                    tempDGV.Columns.Add("col4", "�}���V������");
                    tempDGV.Columns.Add("col5", "���R");
                    tempDGV.Columns.Add("col6", "���̑����e");

                    tempDGV.Columns[0].Width = 80;
                    tempDGV.Columns[1].Width = 80;
                    tempDGV.Columns[2].Width = 140;
                    tempDGV.Columns[3].Width = 175;
                    tempDGV.Columns[4].Width = 60;
                    tempDGV.Columns[5].Width = 265;

                    tempDGV.Columns[0].Visible = false;
                    tempDGV.Columns[1].Visible = false;

                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    
                    //tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //tempDGV.Columns[19].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

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

            /// <summary>
            /// �f�[�^�O���b�h�r���[�̎w��s�̃f�[�^���擾����
            /// </summary>
            /// <param name="dgv">�ΏۂƂ���f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
            public static Boolean GetData(DataGridView dgv, ref Entity.���z�z��� tempC)
            {
                int iX = 0;
                string sqlStr;
                Control.���z�z��� mihaifu = new Control.���z�z���();
                OleDbDataReader dr;

                sqlStr = " where ���z�z���.ID = " + (int)dgv[0, dgv.SelectedRows[iX].Index].Value;
                dr = mihaifu.FillBy(sqlStr);

                if (dr.HasRows == true)
                {
                    while (dr.Read() == true)
                    {
                        tempC.ID = Convert.ToInt32(dr["ID"].ToString());
                        tempC.�z�z�G���AID = int.Parse(dr["�z�z�G���AID"].ToString());
                        tempC.�Ԓn�� = dr["�Ԓn��"].ToString();
                        tempC.�}���V������ = dr["�}���V������"].ToString();
                        tempC.���R = Int32.Parse(dr["���R"].ToString());
                        tempC.���̑����e = dr["���̑����e"].ToString();
                    }
                }
                else
                {
                    dr.Close();
                    mihaifu.Close();
                    return false;
                }

                dr.Close();
                mihaifu.Close();
                return true;
            }

            public static void ShowData(DataGridView tempDGV,int tempID)
            {
                string sqlSTRING = "";
                int iX;

                try
                {
                    tempDGV.RowCount = 0;

                    //���z�z���f�[�^�̃f�[�^���[�_�[���擾����
                    OleDbDataReader dR;
                    Control.���z�z��� cMi = new Control.���z�z���();
                    sqlSTRING = "where �z�z�G���AID = " + tempID.ToString();
                    dR = cMi.FillBy(sqlSTRING);

                    //�O���b�h�r���[�ɕ\������
                    iX = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = dR["ID"].ToString();
                        tempDGV[1, iX].Value = dR["�z�z�G���AID"].ToString();
                        tempDGV[2, iX].Value = dR["�Ԓn��"].ToString();
                        tempDGV[3, iX].Value = dR["�}���V������"].ToString();
                        tempDGV[4, iX].Value = Int32.Parse(dR["���R"].ToString());
                        tempDGV[5, iX].Value = dR["���̑����e"].ToString();
                        iX++;
                    }

                    dR.Close();
                    cMi.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }

                tempDGV.CurrentCell = null;
            }


        }


        private void ShowData()
        {
            lblsID.Text = _SID.ToString();
            lblhDate.Text = _hDate;
            lblName.Text = _staffName;
            lblID.Text = _ID.ToString();
            lblcName.Text = _cName;
            lblfJyouken.Text = _fJyouken;
            lblfKeitai.Text = _fKeitai;
            lblAdd.Text = _Add;
            txtTanka.Text = _Tanka.ToString("#,##0.0");
            lblyMaisu.Text = _yMaisu.ToString("#,##0");
            txthMaisu.Text = _hMaisu.ToString("#,##0");
            
            //txtBanchi.Text = _Banchi;
            //txtManshon.Text = _Manshon;
            //txtRiyu.Text = _Riyu.ToString();
            //txtRiyuName.Text = "";

            ////���R�E�v�\��
            //OleDbDataReader dR;
            //Control.���z�z���R cRiyu = new Control.���z�z���R();

            //dR = cRiyu.FillBy("where ID = " + txtRiyu.Text);

            ////�E�v����\��
            //while (dR.Read())
            //{
            //    txtRiyuName.Text = dR["�E�v"].ToString().Trim();
            //}

            //dR.Close();
            //cRiyu.Close();

            //txtSonota.Text = _Sonota;

            if (_kanryo == 1)
            {
                checkBox1.Checked = true;
                label12.Visible = true;
            }
            else
            {
                checkBox1.Checked = false;
                label12.Visible = false;
            }

            txtEdaban.Text = _Edaban;

            //�}�ԗL��
            if (_Edaban_Status == 0)
            {
                txtEdaban.Enabled = false;
            }
            else
            {
                txtEdaban.Enabled = true;
            }

            //���z�z�L��
            if (_Mihaifu_Status == 0)
            {
                //txtBanchi.Enabled = false;
                //txtManshon.Enabled = false;
                //txtRiyu.Enabled = false;
                //txtSonota.Enabled = false;

                dataGridView1.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
            }
            else
            {
                //txtBanchi.Enabled = true;
                //txtManshon.Enabled = true;
                //txtRiyu.Enabled = true;
                //txtSonota.Enabled = true;

                dataGridView1.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
            }
        }

        //�\���f�[�^�̃v���p�e�B
        private int _SID;
        private string _hDate;
        private string _staffName;
        private int _ID;
        private string _cName;
        private string _fJyouken;
        private string _fKeitai;
        private string _Add;
        private double _Tanka;
        private int _yMaisu;
        private int _hMaisu;
        private int _kanryo;
        private string _Edaban;
        private int _Edaban_Status;

        public int SID
        {
            set { this._SID = value; }
        }

        public string hDate
        {
            set { this._hDate = value; }
        }

        public string staffName
        {
            set { this._staffName = value; }
        }

        public int ID
        {
            set
            { this._ID = value; }
        }

        public string cName
        {
            set { this._cName = value; }
        }

        public string fJyouken
        {
            set { this._fJyouken = value; }
        }

        public string fKeitai
        {
            set { _fKeitai = value; }
        }

        public string Add
        {
            set { _Add = value; }
        }

        public double Tanka
        {
            get { return _Tanka; }
            set { _Tanka = value; }
        }

        public int yMaisu
        {
            set { _yMaisu = value; }
        }

        public int hMaisu
        {
            get { return _hMaisu; }
            set { _hMaisu = value; }
        }

        public int kanryo
        {
            get { return _kanryo; }
            set { _kanryo = value; }
        }

        public string Edaban
        {
            get { return _Edaban; }
            set { _Edaban = value; }
        }

        public int Edaban_Status
        {
            get { return _Edaban_Status; }
            set { _Edaban_Status = value; }
        }

        private string _Banchi;

        public string Banchi
        {
            get { return _Banchi; }
            set { _Banchi = value; }
        }

        private string _Manshon;

        public string Manshon
        {
            get { return _Manshon; }
            set { _Manshon = value; }
        }

        private int _Riyu;

        public int Riyu
        {
            get { return _Riyu; }
            set { _Riyu = value; }
        }

        private string _Sonota;

        public string Sonota
        {
            get { return _Sonota; }
            set { _Sonota = value; }
        }

        private int _Mihaifu_Status;

        public int Mihaifu_Status
        {
            get { return _Mihaifu_Status; }
            set { _Mihaifu_Status = value; }
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtTanka_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtEdaban) { txtObj = txtEdaban; }

            //if (sender == txtBanchi) { txtObj = txtBanchi; }
            //if (sender == txtManshon) { txtObj = txtManshon; }
            
            //if (sender == txtRiyu) 
            //{
            //    txtObj = txtRiyu;

            //    if (txtRiyu.Text == "0")
            //    {
            //        txtRiyu.Text = "";
            //    }
            //}

            //if (sender == txtSonota) { txtObj = txtSonota; }

            if (sender == txtTanka) { txtObj = txtTanka; }
            if (sender == txthMaisu) { txtObj = txthMaisu; }

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;

        }

        private void txtTanka_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtEdaban) { txtObj = txtEdaban; }

            //if (sender == txtBanchi) { txtObj = txtBanchi; }
            //if (sender == txtManshon) { txtObj = txtManshon; }
            //if (sender == txtRiyu) { txtObj = txtRiyu; }
            //if (sender == txtSonota) { txtObj = txtSonota; }

            if (sender == txtTanka) { txtObj = txtTanka; }
            if (sender == txthMaisu) { txtObj = txthMaisu; }

            txtObj.BackColor = Color.White;

        }

        private void txtTanka_Validating(object sender, CancelEventArgs e)
        {
            string str;
            double d;

            if (txtTanka.Text == null)
            {
                MessageBox.Show("�����œ��͂��Ă�������",MESSAGE_CAPTION,MessageBoxButtons.OK,MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txtTanka.Text;

            if (double.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        private void txthMaisu_Validating(object sender, CancelEventArgs e)
        {
            string str;
            int d;

            if (txthMaisu.Text == null)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            str = this.txthMaisu.Text;

            if (Int32.TryParse(str, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
            {
                MessageBox.Show("�����œ��͂��Ă�������", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if ((Int32.Parse(txthMaisu.Text, System.Globalization.NumberStyles.Any) != 0) && 
                (Int32.Parse(lblyMaisu.Text, System.Globalization.NumberStyles.Any) != Int32.Parse(txthMaisu.Text, System.Globalization.NumberStyles.Any)))
            {
                MessageBox.Show("�\�薇���ƈ�v���Ă��܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //�}�ԗL���`�F�b�N�F2018/01/11
            if (_Edaban_Status == 1 && txtEdaban.Text.Trim() == string.Empty)
            {
                if (MessageBox.Show("�}�ԋL���������o�^�ł�����낵���ł����H", "�}�ԊY���Č�", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                {
                    txtEdaban.Focus();
                    return;
                }
            }

            //�P��
            _Tanka = Double.Parse(txtTanka.Text,System.Globalization.NumberStyles.Any);

            //�z�z����
            _hMaisu = Int32.Parse(txthMaisu.Text, System.Globalization.NumberStyles.Any);
        
            //�����敪
            if (checkBox1.Checked == true)
            {
                _kanryo = 1;
            }
            else
            {
                _kanryo = 0;
            }

            //�}�ԋL��
            _Edaban = txtEdaban.Text;

            ////�Ԓn�E��
            //_Banchi = txtBanchi.Text;

            ////�}���V������
            //_Manshon = txtManshon.Text;

            ////���R
            //if (txtRiyu.Text == "")
            //{
            //    _Riyu = 0;
            //}
            //else
            //{
            //    _Riyu = int.Parse(txtRiyu.Text);
            //}

            ////���̑�
            //_Sonota = txtSonota.Text;

            Close();
        }

        private void txthMaisu_KeyDown(object sender, KeyEventArgs e)
        {
            if (txthMaisu.SelectedText == "0")
            {
                if (e.KeyCode == Keys.Return)
                {
                    txthMaisu.Text = lblyMaisu.Text;
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmHaifuShijiSUb3 frm3 = new frmHaifuShijiSUb3(MIHAIFU_ADD);
            frm3.FID = int.Parse(lblID.Text);       //�z�z�G���AID
            frm3.Add = lblAdd.Text;     //�z�z��Z��

            frm3.ShowDialog();

            //�ĕ\��
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void GridEnter()
        {
            frmHaifuShijiSUb3 frm3 = new frmHaifuShijiSUb3(MIHAIFU_UPDATE);
            frm3.ID = int.Parse(dataGridView1[0, dataGridView1.SelectedRows[0].Index].Value.ToString());
            frm3.Add = lblAdd.Text;     //�z�z��Z��

            //�ҏW��ʕ\��
            frm3.ShowDialog();

            //�ĕ\��
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode.ToString() != "Return") return;
            if (dataGridView1.Rows.Count == 0) return;
            if (dataGridView1.SelectedRows.Count == 0) return;

            GridEnter();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("���z�z���f�[�^���I������Ă��܂���", MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("�I�����ꂽ " + dataGridView1.SelectedRows.Count.ToString() + "���̖��z�z�����폜���܂��B��낵���ł����H", "�폜�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                dataGridView1.CurrentCell = null;
                return;
            }

            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
            {
                int aID;
                aID = int.Parse(dataGridView1[0, r.Index].Value.ToString(), System.Globalization.NumberStyles.Any);

                //���R�[�h�폜
                Control.���z�z��� dArea = new Control.���z�z���();

                if (dArea.DataDelete(aID) == false)
                {
                    MessageBox.Show("�폜�Ɏ��s���܂����BID�F" + aID.ToString(), MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                dArea.Close();
            }

            //���z�z���ĕ\��
            GridviewSet.ShowData(dataGridView1, int.Parse(_ID.ToString()));
 
        }
    }
}
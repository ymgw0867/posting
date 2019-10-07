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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // ���O�C��
            frmLogin fl = new frmLogin();
            fl.ShowDialog();

            // ���O�C���������̏ꍇ�͏I������
            if (!global.loginStatus) Environment.Exit(0);
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new frmMenuMST();
            frm.ShowDialog();
            this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmKadou2025 frm = new frmKadou2025();
            frm.ShowDialog();
            this.Show();
        }
        
        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmTesuuryou frm = new frmTesuuryou();
            frm.ShowDialog();
            this.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmTesuuryouMeisai frm = new frmTesuuryouMeisai(0);
            frm.ShowDialog();
            this.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Hide();

            //�����ꗗ
            frmSeikyuRep2015 frm = new frmSeikyuRep2015();
            frm.ShowDialog();
            this.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //�������o�^
            this.Hide();
            frmSeikyuAdd frm = new frmSeikyuAdd(1);
            frm.ShowDialog();
            this.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //�����ꗗ
            this.Hide();
            frmNyukinRep frm = new frmNyukinRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            //�󒍊m�菑�o�^
            Form frm = new frmOrder(0);

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            //�|�X�e�B���O�G���A�o�^
            frmPosting frm = new frmPosting();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //�|�X�e�B���O�G���A�o�^
            Form frm = new frmTantouOrderRep();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new frmHaifuShiji();
            frm.ShowDialog();
            this.Show();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new frmHaifuKanryoRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new frmHaifuShinchoku();
            frm.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("�V�X�e�����I�����܂��B��낵���ł���", "�m�F", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }

            this.Dispose();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmTaihiRep frm = new frmTaihiRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmRuikeiRep frm = new frmRuikeiRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmTenkou frm = new frmTenkou();
            frm.ShowDialog();
            this.Show();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmHaifuMax frm = new frmHaifuMax();
            frm.ShowDialog();
            this.Show();
        }

        private void button23_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmMihaifuRep frm = new frmMihaifuRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            //���̃}�V���Ŕz�z�w�����s�������؂���
            int cFlg = 0;

            OleDbDataReader dR;
            Control.��Џ�� cSystem = new Control.��Џ��();
            dR = cSystem.Fill();

            while (dR.Read())
            {
                cFlg = int.Parse(dR["�z�z�t���O"].ToString());
            }

            dR.Close();
            cSystem.Close();

            if (cFlg == 1)
            {
                MessageBox.Show("���݁A���̃}�V���Ŕz�z�w���o�^���ł��B", "�N���`�F�b�N", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //�z�z�w���o�^����
            this.Hide();
            frmHaifuShijiADD frm = new frmHaifuShijiADD();
            frm.ShowDialog();
            this.Show();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmChirashiZaiko frm = new frmChirashiZaiko();
            frm.ShowDialog();
            this.Show();
        }

        private void button26_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmHaifuShijiRep frm = new frmHaifuShijiRep();
            frm.ShowDialog();
            this.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string sqlString;

            Control.FreeSql fSql = new Control.FreeSql();

            // �󒍃e�[�u���Ɂu�ҏW���b�N�v�t�B�[���h��ǉ����� : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE �� add �ҏW���b�N int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // �󒍃e�[�u���Ɂu��������̍ς݁v�t�B�[���h��ǉ����� : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE �� add ��������̍ς� int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // ���O�C���^�C�v�w�b�_�e�[�u���Ɂu�󒍌ʃ��b�N�����v�t�B�[���h��ǉ����� : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE ���O�C���^�C�v�w�b�_ add �󒍌ʃ��b�N���� int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // ���O�C���^�C�v�w�b�_�e�[�u���Ɂu�󒍌ʐ����v�t�B�[���h��ǉ����� : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE ���O�C���^�C�v�w�b�_ add �󒍌ʐ��� int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // ���O�C���^�C�v�w�b�_�e�[�u���Ɂu��������̍ς݌����v�t�B�[���h��ǉ����� : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE ���O�C���^�C�v�w�b�_ add ��������̍ς݌��� int default 0 NOT NULL";
            fSql.Execute(sqlString);

            fSql.Close();

            //// ��Џ��e�[�u���Ɂu�󒍊m�菑���̓V�[�g�p�X�v�t�B�[���h��ǉ����� : 2019/03/06
            //sqlString = "";
            //sqlString += "ALTER TABLE ��Џ�� add �󒍊m�菑���̓V�[�g�p�X nvarchar(255) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���Ɂu�c�Ɣ��l�v�t�B�[���h��ǉ����� : 2019/03/01
            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �c�Ɣ��l nvarchar(255) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// ���Ӑ�e�[�u���Ɂu������E�������v�u������E�h�́v�t�B�[���h��ǉ����� : 2019/02/19
            //sqlString = "";
            //sqlString += "ALTER TABLE ���Ӑ� add �����敔���� nvarchar(50) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE ���Ӑ� add ������h�� nvarchar(5) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //fSql.Close();
            

            // �ȉ��A�R�����g�� 2019/02/19
            // �O����}�X�^�[�Ɂu�x�����v�t�B�[���h��ǉ�����F2018/01/03
            //sqlString = "";
            //sqlString += "ALTER TABLE �O���� add �x���� int default 0 NOT NULL";

            //fSql.Execute(sqlString);

            //fSql.Close();

            // �ȉ���ALTER TABLE SQL �R�����g�� 2018/01/03
            //// �V���������e�[�u���Ɂu�����v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �V������ add ���� nvarchar(10) default '' NOT NULL";

            //fSql.Execute(sqlString);

            //fSql.Close();

            ///* �󒍃e�[�u���Ɂu�O����h�c�x���Q�v�u�O���x�����x���Q�v�u�O�������x���Q�v�u�O����h�c�x���R�v
            //�u�O���x�����x���R�v�u�O�������x���R�v�t�B�[���h��ǉ����� */

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O����ID�x��2 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���x�����x��2 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O�������x��2 money default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O����ID�x��3 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���x�����x��3 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O�������x��3 money default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���Ɂu�O���˗����x���Q�v�u�O���˗����x���R�v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���˗����x��2 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���˗����x��3 datetime";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���Ɂu�O���ϑ������Q�v�u�O���ϑ������R�v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���ϑ�����2 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���ϑ�����3 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���Ɂu�O���n�����Q�v�u�O���n�����R�v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���n����2 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���n����3 datetime";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���Ɂu�O���󂯓n���S���҂Q�v�u�O���󂯓n���S���҂R�v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���󂯓n���S����2 nvarchar(50) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���󂯓n���S����3 nvarchar(50) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// �O���x���e�[�u���Ɂu�����z�v�u�������t�v�u�������l�v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �O���x�� add �����z int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �O���x�� add �������t nvarchar(10) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �O���x�� add �������l nvarchar(100) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���Ɂu�O���x��ID2�v�u�O���x��ID3�v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���x��ID2 nvarchar(14) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �O���x��ID3 nvarchar(14) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���Ɂu�[�i�����s�v�t�B�[���h��ǉ�����
            //sqlString = "";
            //sqlString += "ALTER TABLE �� add �[�i�����s int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //// �󒍃e�[�u���ɃC���f�b�N�X��ǉ�����@2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�󒍎��ID ";
            //sqlString += "ON ��(�󒍎��ID) ";
            //sqlString += "INCLUDE(ID,���Ӑ�ID,�`���V��,����,�ō����z)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�󒍃`���V�� ";
            //sqlString += "ON ��(�`���V�� ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�󒍓� ";
            //sqlString += "ON ��(�󒍓� ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�o�^���[�U�[ID ";
            //sqlString += "ON ��(�o�^���[�U�[ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�z�z�`�� ";
            //sqlString += "ON ��(�z�z�`�� ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_���Ӑ�ID ";
            //sqlString += "ON ��(���Ӑ�ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //// �z�z�G���A�e�[�u���ɃC���f�b�N�X��ǉ�����@2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�z�z�w��ID ";
            //sqlString += "ON �z�z�G���A(�z�z�w��ID) ";
            //sqlString += "INCLUDE(�\�薇��,��ID) ";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�����敪�z�z�w��ID ";
            //sqlString += "ON �z�z�G���A(�����敪,�z�z�w��ID) ";
            //sqlString += "INCLUDE(�\�薇��, ��ID)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�z�z�G���A��ID ";
            //sqlString += "ON �z�z�G���A(��ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�X�e�[�^�X ";
            //sqlString += "ON �z�z�G���A(�X�e�[�^�X ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�����敪 ";
            //sqlString += "ON �z�z�G���A(�����敪 ASC) ";
            //sqlString += "INCLUDE(�\�薇��, ��ID) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //// �z�z�w���e�[�u���ɃC���f�b�N�X��ǉ�����@2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�z�z�w�� ";
            //sqlString += "ON �z�z�w��(�z�z��ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_���[�U�[ID ";
            //sqlString += "ON �z�z�w��(���[�U�[ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //// ���Ӑ�e�[�u���ɃC���f�b�N�X��ǉ�����@2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_�S���Ј��R�[�h ";
            //sqlString += "ON ���Ӑ�(�S���Ј��R�[�h ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);


            //////�󒍃f�[�^�̐�����ID������������ 2010/02/17
            //////�c�`�e�N�j�J���T�[�r�X�@��ID�F201001150004�@������ID�F822 �� 726
            ////sqlString = "";
            ////sqlString += "update �� set ������ID = 726 ";
            ////sqlString += "where ID = 201001150004";

            ////fSql.Execute(sqlString);

            //////�󒍃f�[�^�̐�����ID������������ 2010/03/03
            //////HAKU�@��ID�F200911040013�@������ID�F776 �� 613
            ////sqlString = "";
            ////sqlString += "update �� set ������ID = 613 ";
            ////sqlString += "where ID = 200911040013";

            ////fSql.Execute(sqlString);

            //////�󒍃f�[�^�̐�����ID������������ 2010/03/03
            //////�\�Q���ڍ��@�@��ID�F200912210002�@������ID�F743 �� 660
            ////sqlString = "";
            ////sqlString += "update �� set ������ID = 660 ";
            ////sqlString += "where ID = 200912210002";

            ////fSql.Execute(sqlString);

            //////�󒍃f�[�^�̐�����ID������������ 2010/03/04
            //////���R�v���@��ID�F201001150002�@������ID�F899 �� 744 
            ////sqlString = "";
            ////sqlString += "update �� set ������ID = 744 ";
            ////sqlString += "where ID = 201001150002";

            ////fSql.Execute(sqlString);

            //////�󒍃f�[�^�̐�����ID������������ 2010/04/14
            //////���R�v���@��ID�F201003300007�@������ID�F0 �� 1018 
            ////sqlString = "";
            ////sqlString += "update �� set ������ID = 1018 ";
            ////sqlString += "where (ID = 201003300007) and (������ID = 0)";

            ////fSql.Execute(sqlString);

            //////�󒍃f�[�^�̐�����ID������������ 2010/04/14
            //////���R�v���@��ID�F201002190001�@������ID�F1017 �� 900 
            ////sqlString = "";
            ////sqlString += "update �� set ������ID = 900 ";
            ////sqlString += "where (ID = 201002190001) and (������ID = 1017)";

            ////fSql.Execute(sqlString);

            //fSql.Close();

            // ���j���[�^�C�g���N���X 2015/07/07
            clsMenu cm = new clsMenu();

            // ���j���[�^�C�g��CSV�̓Ǎ� 2015/07/07
            cm.loadMenu();

            // ���j���[�^�C�g�����Z�b�g 2015/07/07
            Utility.getMenuTittle(button14, cm);
            Utility.getMenuTittle(button15, cm);
            Utility.getMenuTittle(button16, cm);
            Utility.getMenuTittle(button24, cm);
            Utility.getMenuTittle(button18, cm);
            Utility.getMenuTittle(button17, cm);
            Utility.getMenuTittle(button23, cm);
            Utility.getMenuTittle(button2, cm);
            Utility.getMenuTittle(button25, cm);
            Utility.getMenuTittle(button26, cm);
            Utility.getMenuTittle(button5, cm);
            Utility.getMenuTittle(button4, cm);
            Utility.getMenuTittle(button8, cm);
            Utility.getMenuTittle(button7, cm);
            Utility.getMenuTittle(button19, cm);
            Utility.getMenuTittle(button20, cm);
            Utility.getMenuTittle(button21, cm);
            Utility.getMenuTittle(button22, cm);
            Utility.getMenuTittle(button9, cm);
            Utility.getMenuTittle(button10, cm);
            Utility.getMenuTittle(button11, cm);
            Utility.getMenuTittle(button6, cm);
            Utility.getMenuTittle(button12, cm);
            Utility.getMenuTittle(button27, cm);
            Utility.getMenuTittle(button28, cm);
            Utility.getMenuTittle(button29, cm);
            Utility.getMenuTittle(button30, cm);
            Utility.getMenuTittle(button31, cm);
            Utility.getMenuTittle(button13, cm);
            Utility.getMenuTittle(button32, cm);
            Utility.getMenuTittle(button33, cm);
            Utility.getMenuTittle(button34, cm);

            // ���j���[�{�^���\����ԏ�����
            button14.Enabled = false;
            button15.Enabled = false;
            button16.Enabled = false;
            button24.Enabled = false;
            button18.Enabled = false;
            button17.Enabled = false;
            button23.Enabled = false;
            button2.Enabled = false;
            button25.Enabled = false;
            button26.Enabled = false;
            button5.Enabled = false;
            button4.Enabled = false;
            button8.Enabled = false;
            button7.Enabled = false;
            button19.Enabled = false;
            button20.Enabled = false;
            button21.Enabled = false;
            button22.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = false;
            button11.Enabled = false;
            button6.Enabled = false;
            button12.Enabled = false;
            button27.Enabled = false;
            button28.Enabled = false;
            button29.Enabled = false;
            button30.Enabled = false;
            button31.Enabled = false;
            button13.Enabled = false;
            button32.Enabled = false;
            button33.Enabled = false;
            button34.Enabled = false;

            // ���O�C�����[�U�[���Ƃ̃��j���[����
            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.���O�C���^�C�v�w�b�_TableAdapter hAdp = new darwinDataSetTableAdapters.���O�C���^�C�v�w�b�_TableAdapter();
            darwinDataSetTableAdapters.���O�C���^�C�v�^�OTableAdapter tAdp = new darwinDataSetTableAdapters.���O�C���^�C�v�^�OTableAdapter();
            hAdp.Fill(dts.���O�C���^�C�v�w�b�_);
            tAdp.Fill(dts.���O�C���^�C�v�^�O);

            foreach (var h in dts.���O�C���^�C�v�w�b�_.Where(a => a.Id == global.loginType))
            {
                foreach (var item in h.Get���O�C���^�C�v�^�ORows())
                {
                    if (menuButtonStatus(button14, item.tag)) continue;
                    if (menuButtonStatus(button15, item.tag)) continue;
                    if (menuButtonStatus(button16, item.tag)) continue;
                    if (menuButtonStatus(button24, item.tag)) continue;
                    if (menuButtonStatus(button18, item.tag)) continue;
                    if (menuButtonStatus(button17, item.tag)) continue;
                    if (menuButtonStatus(button23, item.tag)) continue;
                    if (menuButtonStatus(button2, item.tag)) continue;
                    if (menuButtonStatus(button25, item.tag)) continue;
                    if (menuButtonStatus(button26, item.tag)) continue;
                    if (menuButtonStatus(button5, item.tag)) continue;
                    if (menuButtonStatus(button4, item.tag)) continue;
                    if (menuButtonStatus(button8, item.tag)) continue;
                    if (menuButtonStatus(button7, item.tag)) continue;
                    if (menuButtonStatus(button19, item.tag)) continue;
                    if (menuButtonStatus(button20, item.tag)) continue;
                    if (menuButtonStatus(button21, item.tag)) continue;
                    if (menuButtonStatus(button22, item.tag)) continue;
                    if (menuButtonStatus(button9, item.tag)) continue;
                    if (menuButtonStatus(button10, item.tag)) continue;
                    if (menuButtonStatus(button11, item.tag)) continue;
                    if (menuButtonStatus(button6, item.tag)) continue;
                    if (menuButtonStatus(button12, item.tag)) continue;
                    if (menuButtonStatus(button27, item.tag)) continue;
                    if (menuButtonStatus(button28, item.tag)) continue;
                    if (menuButtonStatus(button29, item.tag)) continue;
                    if (menuButtonStatus(button30, item.tag)) continue;
                    if (menuButtonStatus(button31, item.tag)) continue;
                    if (menuButtonStatus(button13, item.tag)) continue;
                    if (menuButtonStatus(button32, item.tag)) continue;
                    if (menuButtonStatus(button33, item.tag)) continue;
                    if (menuButtonStatus(button34, item.tag)) continue;
                }
            }

            // ���O�C�������[�U�[
            //lblLogin.Text = "���O�C�������[�U�[�F" + global.loginUser;
            lblLogin.Text = global.loginUser + "���񂪃��O�C�����ł�";


            //// �������g�̃o�[�W���������擾����@2016/11/08
            //System.Diagnostics.FileVersionInfo ver =
            //    System.Diagnostics.FileVersionInfo.GetVersionInfo(
            //    System.Reflection.Assembly.GetExecutingAssembly().Location);

            // �L���v�V�����Ƀo�[�W������ǉ��@2016/11/08
            this.Text += " ver " + Application.ProductVersion;
        }

        private bool menuButtonStatus(Button btn, int tag)
        {
            if (Utility.strToInt(btn.Tag.ToString()) == tag)
            {
                btn.Enabled = true;
                return true;
            }

            return false;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmOrderNumber frm = new frmOrderNumber();
            frm.ShowDialog();
            this.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmFuriRep frm = new frmFuriRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmShiharai frm = new frmShiharai(0);
            frm.ShowDialog();
            this.Show();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmKaikakeMenu frm = new frmKaikakeMenu();
            frm.ShowDialog();
            this.Show();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmShiharaiYotei frm = new frmShiharaiYotei();
            frm.ShowDialog();
            this.Show();
        }

        private void button29_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmOrderExcel frm = new frmOrderExcel();
            frm.ShowDialog();
            this.Show();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmNyukinRep2015 frm = new frmNyukinRep2015();
            frm.ShowDialog();
            this.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            //frmHansokuTeateRep frm = new frmHansokuTeateRep();
            frmEiUriageRep frm = new frmEiUriageRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button31_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmUrikakeMenu frm = new frmUrikakeMenu();
            frm.ShowDialog();
            this.Show();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmSeikyuShime frm = new frmSeikyuShime();
            frm.ShowDialog();
            this.Show();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmOrderRecord frm = new frmOrderRecord();
            frm.ShowDialog();
            this.Show();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmNouhinRep frm = new frmNouhinRep();
            frm.ShowDialog();
            this.Show();
        }
    }
}
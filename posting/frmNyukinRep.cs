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
    public partial class frmNyukinRep : Form
    {
        const string MESSAGE_CAPTION = "�N���C�A���g�ʐ����ꗗ";

        public frmNyukinRep()
        {
            InitializeComponent();

            // �f�[�^�ǂݍ���
            rAdp.Fill(dts.�V������);
            nAdp.Fill(dts.�V����);
            cAdp.Fill(dts.���Ӑ�);
            sAdp.Fill(dts.�Ј�);
            jAdp.Fill(dts.��1);
        }

        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.�V������TableAdapter rAdp = new darwinDataSetTableAdapters.�V������TableAdapter();
        darwinDataSetTableAdapters.�V����TableAdapter nAdp = new darwinDataSetTableAdapters.�V����TableAdapter();
        darwinDataSetTableAdapters.���Ӑ�TableAdapter cAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
        darwinDataSetTableAdapters.�Ј�TableAdapter sAdp = new darwinDataSetTableAdapters.�Ј�TableAdapter();
        darwinDataSetTableAdapters.��1TableAdapter jAdp = new darwinDataSetTableAdapters.��1TableAdapter();

        bool mukouStatus = false;

        private void form_Load(object sender, EventArgs e)
        {
            //�E�B���h�E�Y�ŏ��T�C�Y
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            ////�s�v�������f�[�^�폜 2010/05/31
            //DataDelete();

            //�O���b�h��`
            GridviewSet.Setting(dataGridView1);

            //��ʃN���A
            DispClear();
        }

        private void DataDelete()
        {
            string sqlString;

            Control.FreeSql fSql = new Control.FreeSql();
            
            //sqlSTRING += " and ((������.ID != 709) and (������.ID != 771) and (������.ID != 723) and 
            //(������.ID != 822) and (������.ID != 765) and (������.ID != 776) and (������.ID != 743) and
            //(������.ID != 899) and (������.ID != 1017) and (������.ID != 847) and (������.ID != 849) and
            //(������.ID != 990) and (������.ID != 1190)) ";

            //�\�V�G���[���h
            sqlString = "delete from ������ where ID = 709";
            fSql.Execute(sqlString);

            //�����ς�
            sqlString = "delete from ������ where ID = 771";
            fSql.Execute(sqlString);

            //�L�����������N
            sqlString = "delete from ������ where ID = 723";
            fSql.Execute(sqlString);

            //�c�`�e�N�j�J���T�[�r�X
            sqlString = "delete from ������ where ID = 822";
            fSql.Execute(sqlString);

            //
            sqlString = "delete from ������ where ID = 765";
            fSql.Execute(sqlString);

            //�g�`�j�t
            sqlString = "delete from ������ where ID = 776";
            fSql.Execute(sqlString);

            //�\�Q���ڍ��@
            sqlString = "delete from ������ where ID = 743";
            fSql.Execute(sqlString);

            //���R�v��
            sqlString = "delete from ������ where ID = 899";
            fSql.Execute(sqlString);

            //���R�v��
            sqlString = "delete from ������ where ID = 1017";
            fSql.Execute(sqlString);

            //�v���Y�~�b�N
            sqlString = "delete from ������ where ID = 847";
            fSql.Execute(sqlString);

            //�c�`�e�N�j�J���T�[�r�X
            sqlString = "delete from ������ where ID = 849";
            fSql.Execute(sqlString);

            //��̉�
            sqlString = "delete from ������ where ID = 990";
            fSql.Execute(sqlString);

            //����c�[�~�i�[��
            sqlString = "delete from ������ where ID = 1190";
            fSql.Execute(sqlString);

            fSql.Close();
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
                    tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                    tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    // ��w�b�_�[�\���ʒu�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // ��w�b�_�[�t�H���g�w��
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                    // �f�[�^�t�H���g�w��
                    tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);
                    
                    // �s�̍���
                    tempDGV.ColumnHeadersHeight = 18;
                    tempDGV.RowTemplate.Height = 18;

                    // �S�̂̍���
                    tempDGV.Height = 470;

                    // ��s�̐F
                    tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                    // �e�񕝎w��
                    tempDGV.Columns.Add("col1", "�ԍ�");
                    tempDGV.Columns.Add("col2", "�N���C�A���g��");
                    tempDGV.Columns.Add("col13", "�t���K�i");
                    tempDGV.Columns.Add("col3", "�S����");
                    tempDGV.Columns.Add("col4", "���s��");
                    tempDGV.Columns.Add("col5", "�������z");
                    tempDGV.Columns.Add("col6", "�����\���");
                    tempDGV.Columns.Add("col7", "������");
                    tempDGV.Columns.Add("col8", "�����z");
                    tempDGV.Columns.Add("col9", "�����c��");

                    DataGridViewCheckBoxColumn cbc = new DataGridViewCheckBoxColumn();
                    tempDGV.Columns.Add(cbc);
                    tempDGV.Columns[10].HeaderText = "������";
                    tempDGV.Columns[10].Name = "col10";

                    tempDGV.Columns.Add("col11", "������ID");
                    tempDGV.Columns.Add("col12", "���l");   // 2015/07/22

                    //tempDGV.Columns[1].Frozen = true;   // 2015/07/22

                    tempDGV.Columns[0].Width = 60;
                    tempDGV.Columns[1].Width = 220;
                    tempDGV.Columns[2].Width = 160;
                    tempDGV.Columns[3].Width = 100;
                    tempDGV.Columns[4].Width = 100;
                    tempDGV.Columns[5].Width = 100;
                    tempDGV.Columns[6].Width = 100;
                    tempDGV.Columns[7].Width = 100;
                    tempDGV.Columns[8].Width = 100;
                    tempDGV.Columns[9].Width = 100;
                    tempDGV.Columns[10].Width = 60;
                    tempDGV.Columns[12].Width = 200;

                    tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    tempDGV.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    //tempDGV.Columns[3].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[4].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[5].DefaultCellStyle.Format = "yyyy/M/dd";
                    //tempDGV.Columns[7].DefaultCellStyle.Format = "#,##0";
                    //tempDGV.Columns[8].DefaultCellStyle.Format = "#,##0";

                    //������ID���\���Ƃ���
                    tempDGV.Columns[11].Visible = false;

                    // �s�w�b�_��\�����Ȃ�
                    tempDGV.RowHeadersVisible = false;

                    // �I�����[�h
                    tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                    tempDGV.MultiSelect = false;

                    // �ҏW�s�Ƃ���
                    tempDGV.ReadOnly = false;

                    // �ҏW�E�s�̎w��
                    foreach (DataGridViewColumn d in tempDGV.Columns)
                    {
                        if (d.Name == "col10" || d.Name == "col12")
                        {
                            d.ReadOnly = false;
                        }
                        else
                        {
                            d.ReadOnly = true;
                        }
                    }

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
                    tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                    tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[���b�Z�[�W", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public static void ShowData(DataGridView tempDGV, int tempKBN, TextBox tempTxt, TextBox tempTotal, DateTime tempSDate, DateTime tempEDate,string tempCstr)
            {
                string sqlSTRING = "";
                int sZan = 0;
                int sTotal = 0;

                try
                {
                    // �f�[�^���[�_�[���擾����
                    OleDbDataReader dR;

                    sqlSTRING += "SELECT ������.ID,������.���Ӑ�ID,������.�������z,������.�����,";
                    sqlSTRING += "������.�l���z,������.������z,������.�ŗ�,";
                    sqlSTRING += "������.�����\���,������.���s��,������.�����c,������.�����敪,";
                    sqlSTRING += "������.�U������ID1,������.�U������ID2,������.���l,";
                    sqlSTRING += "������.�o�^�N����,������.�ύX�N����,���Ӑ�.����,�Ј�.����,n_tbl.������,n_tbl.�����z ";
                    sqlSTRING += "from ������ LEFT OUTER JOIN ���Ӑ� ";
                    sqlSTRING += "ON ������.���Ӑ�ID = ���Ӑ�.ID LEFT OUTER JOIN �Ј� ";
                    sqlSTRING += "ON ���Ӑ�.�S���Ј��R�[�h = �Ј�.ID LEFT OUTER JOIN ";
                    sqlSTRING += "(SELECT ������ID, MAX(�����N����) AS ������,sum(���z) as �����z ";
                    sqlSTRING += "FROM ���� ";
                    sqlSTRING += "GROUP BY ������ID) AS n_tbl ON ������.ID = n_tbl.������ID ";
                    sqlSTRING += "where ";

                    if (tempKBN == 1)
                    {
                        sqlSTRING += "(������.�����敪 = 0) and ";
                    }

                    sqlSTRING += "(������.�����\��� >= '" + tempSDate + "') and (������.�����\��� <= '" + tempEDate + "') ";

                    if (tempCstr != "")
                    {
                        sqlSTRING += " and (���Ӑ�.���� like '%" + tempCstr + "%') " ;
                    }

                    // �s�v�f�[�^ 2010/02/16
                    sqlSTRING += " and ((������.ID != 709) and (������.ID != 771) and (������.ID != 723) and (������.ID != 822) and (������.ID != 765) and (������.ID != 776) and (������.ID != 743) and (������.ID != 899) and (������.ID != 1017) and (������.ID != 847) and (������.ID != 849) and (������.ID != 990) and (������.ID != 1190)) ";
                    
                    sqlSTRING += "ORDER BY ������.���Ӑ�ID, ������.���s�� DESC ";

                    Control.FreeSql fCon = new Control.FreeSql();
                    dR = fCon.free_dsReader(sqlSTRING);

                    // �O���b�h�r���[�ɕ\������
                    int iX = 0;

                    tempDGV.RowCount = 0;

                    while (dR.Read())
                    {
                        tempDGV.Rows.Add();

                        tempDGV[0, iX].Value = int.Parse(dR["���Ӑ�ID"].ToString());
                        tempDGV[1, iX].Value = dR["����"].ToString();
                        tempDGV[2, iX].Value = dR["����"].ToString();
                        tempDGV[3, iX].Value = DateTime.Parse(dR["���s��"].ToString()).ToShortDateString();
                        tempDGV[4, iX].Value = int.Parse(dR["�������z"].ToString(),System.Globalization.NumberStyles.Any);

                        if (dR["�����\���"] == DBNull.Value)
                        {
                            tempDGV[5, iX].Value = "";
                        }
                        else
                        {
                            tempDGV[5, iX].Value = DateTime.Parse(dR["�����\���"].ToString()).ToShortDateString();
                        }

                        if (dR["������"] == DBNull.Value )
                        {
                            tempDGV[6, iX].Value = "";
                        }
                        else
                        {
                            tempDGV[6, iX].Value = DateTime.Parse(dR["������"].ToString()).ToShortDateString();
                        }

                        if (dR["�����z"] == DBNull.Value)
                        {
                            tempDGV[7, iX].Value = 0;
                        }
                        else
                        {
                            tempDGV[7, iX].Value = int.Parse(dR["�����z"].ToString(), System.Globalization.NumberStyles.Any);
                        }

                        tempDGV[8, iX].Value = int.Parse(dR["�����c"].ToString(), System.Globalization.NumberStyles.Any);

                        // ������ID 2010/2/16
                        tempDGV[9, iX].Value = dR["ID"].ToString();

                        // ���l 2015/07/22
                        tempDGV[10, iX].Value = dR["���l"].ToString();

                        // �������v
                        if (dR["�����z"] != DBNull.Value)
                        sTotal += int.Parse(dR["�����z"].ToString(), System.Globalization.NumberStyles.Any);
                        
                        // �����c���v
                        sZan += int.Parse(dR["�����c"].ToString(), System.Globalization.NumberStyles.Any);
                        
                        iX++;
                    }

                    dR.Close();
                    fCon.Close();

                    //if (tempDGV.RowCount > 25)
                    //    //if (tempDGV.RowCount > 25)
                    //{
                    //    tempDGV.Columns[1].Width = 253;
                    //}
                    //else
                    //{
                    //    tempDGV.Columns[1].Width = 270;
                    //}

                    tempTotal.Text = sTotal.ToString("#,##0");
                    tempTxt.Text = sZan.ToString("#,##0");
                    tempDGV.CurrentCell = null;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
                }
            }
        }
        
        private void ShowDataLinq(DataGridView g)
        {
            int sZan = 0;
            int sTotal = 0;
            int sMinyu = 0;
            int sUrikake = 0;

            try
            {
                // �O���b�h�r���[�ɕ\������
                int iX = 0;

                g.RowCount = 0;

                // �����f�[�^�͑ΏۊO
                //var s = dts.�V������.Where(a => a.���� == global.FLGOFF).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);

                // 2015/12/09  
                var s = dts.�V������
                        .Where(a => a.���� == global.FLGOFF && a.Get��1Rows().Count() > 0)
                        .OrderBy(a => a.�x������).ThenBy(a => a.���������s��);

                // �x������
                if (tDate.Checked)
                {
                    s = s.Where(a => a.�x������ >= DateTime.Parse(tDate.Value.ToShortDateString())).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);
                }

                if (tDate2.Checked)
                {
                    s = s.Where(a => a.�x������ <= DateTime.Parse(tDate2.Value.ToShortDateString())).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);
                }

                // ���������s��
                if (hDate.Checked)
                {
                    s = s.Where(a => a.���������s�� >= DateTime.Parse(hDate.Value.ToShortDateString())).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);
                }

                if (hDate2.Checked)
                {
                    s = s.Where(a => a.���������s�� <= DateTime.Parse(hDate2.Value.ToShortDateString())).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);
                }

                // ���������̂�
                if (radioButton2.Checked)
                {
                    if (chkMishu.Checked)
                    {
                        s = s.Where(a => a.�������� == global.FLGOFF).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);
                    }
                    else
                    {
                        s = s.Where(a => a.�������� == global.FLGOFF || a.�c�� > 0).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);
                    }
                }
                else if (radioButton3.Checked)  // �����m��̂�
                {
                    s = s.Where(a => a.�������� == global.FLGON && a.�c�� > 0).OrderBy(a => a.�x������).ThenBy(a => a.���������s��);
                }

                foreach (var t in s)
                {
                    // �N���C�A���g���w��
                    if (txtsClientName.Text.Trim() != string.Empty)
                    {
                        // ���Ӑ於
                        if (t.���Ӑ�Row == null)
                        {
                            continue;
                        }
                        else if (!t.���Ӑ�Row.����.Contains(txtsClientName.Text.Trim()))
                        {
                            continue;
                        }
                    }

                    g.Rows.Add();

                    g[0, iX].Value = t.���Ӑ�ID.ToString();    // ���Ӑ�ID

                    // ���Ӑ於�E�S����
                    if (t.���Ӑ�Row == null)
                    {
                        g[1, iX].Value = string.Empty;
                        g[2, iX].Value = string.Empty;
                        g[3, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[1, iX].Value = t.���Ӑ�Row.����;
                        g[2, iX].Value = t.���Ӑ�Row.�t���K�i;

                        if (t.���Ӑ�Row.�Ј�Row == null)
                        {
                            g[3, iX].Value = string.Empty;
                        }
                        else
                        {
                            g[3, iX].Value = t.���Ӑ�Row.�Ј�Row.����;
                        }
                    }

                    g[4, iX].Value =  t.���������s��.ToShortDateString();     // ���������s��
                    g[5, iX].Value = t.�������z.ToString("#,##0");            // �������z

                    // �����\���
                    if (t.Is�x������Null())
                    {
                        g[6, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[6, iX].Value = t.�x������.ToShortDateString();
                    }

                    // �����z�E������
                    DateTime nDt = DateTime.Parse("1900/01/01");
                    int nyukin = 0;

                    foreach (var n in t.Get�V����Rows())
                    {
                        nyukin += n.���z;

                        if (nDt < n.�����N����)
                        {
                            nDt = n.�����N����;
                        }
                    }

                    if (t.Get�V����Rows().Any())
                    {
                        g[7, iX].Value = nDt.ToShortDateString();
                        g[8, iX].Value = nyukin.ToString("#,##0");
                    }
                    else
                    {
                        g[7, iX].Value = string.Empty;
                        g[8, iX].Value = "0";
                    }

                    g[9, iX].Value = t.�c��.ToString("#,##0");

                    if (t.�������� == global.FLGOFF)
                    {
                        g[10, iX].Value = false;
                    }
                    else
                    {
                        g[10, iX].Value = true;
                    }

                    g[11, iX].Value = t.ID.ToString();

                    if (t.Is���lNull())
                    {
                        g[12, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[12, iX].Value = t.���l;
                    }

                    sTotal += nyukin;   // �������v

                    //if (chkMishu.Checked)
                    //{
                    //    if ()
                    //}
                    //else
                    //{
                    //    sZan += t.�c��;     // �����c���v
                    //}

                    // �c������
                    if (t.�c�� != 0)
                    {
                        if (chkMishu.Checked)�@// ���������͎c���Ȃ��Ƃ���
                        {                            
                            if (t.�������� == global.FLGOFF)
                            {
                                // �c���v
                                sZan += t.�c��;

                                // ������ or ���|��
                                if (DateTime.Today > t.�x������)
                                {
                                    // �x���������߂��Ă����疢����
                                    sMinyu += t.�c��;
                                }
                                else
                                {
                                    // �x�������ȑO�ł���Δ��|��
                                    sUrikake += t.�c��;
                                }
                            }
                        }
                        else
                        {
                            // �c���v
                            sZan += t.�c��;

                            // ������ or ���|��
                            if (DateTime.Today > t.�x������)
                            {
                                // �x���������߂��Ă����疢����
                                sMinyu += t.�c��;
                            }
                            else
                            {
                                // �x�������ȑO�ł���Δ��|��
                                sUrikake += t.�c��;
                            }
                        }
                    }

                    iX++;
                }
                    
                txtTotal.Text = sTotal.ToString("#,##0");
                txtZan.Text = sZan.ToString("#,##0");
                txtMinyu.Text = sMinyu.ToString("#,##0");
                txtUrikake.Text = sUrikake.ToString("#,##0");

                g.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "�G���[", MessageBoxButtons.OK);
            }
        }
        
        
        /// <summary>
        /// ��ʂ��N���A����
        /// </summary>
        private void DispClear()
        {
            try
            {
                chkMishu.Checked = true;
                radioButton1.Checked = true;
                btnPrn.Enabled = false;
                button1.Enabled = false;    // 2017/08/14
                hDate.Checked = false;
                hDate2.Checked = false;
                tDate.Checked = false;
                tDate2.Checked = false;
                txtsClientName.Text = "";
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "��ʃN���A", MessageBoxButtons.OK);
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                //��ʕ\��
                mukouStatus = false;
                ShowDataLinq(dataGridView1);
                mukouStatus = true;

                if (dataGridView1.RowCount > 0)
                {
                    btnPrn.Enabled = true;
                    button1.Enabled = true;
                }
                else
                {
                    btnPrn.Enabled = false;
                    button1.Enabled = false;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"�I��",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }    
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Gengo_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            EigyoReport(dataGridView1);
        }
        
        private void EigyoReport(DataGridView tempDGV)
        {
            const int S_GYO = 4;        // �G�N�Z���t�@�C�����׈���J�n�s
            const int S_ROWSMAX = 11;   // �G�N�Z���t�@�C����ő�l

            try
            {
                //�}�E�X�|�C���^��ҋ@�ɂ���
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + Properties.Settings.Default.�G�N�Z���N���C�A���g�ʐ����ꗗ, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    // ���v������
                    StringBuilder sb = new StringBuilder();
                    sb.Append("�����v�F").Append(txtTotal.Text).Append("  ");
                    sb.Append("�������F").Append(txtMinyu.Text).Append("  ");
                    sb.Append("���|���F").Append(txtUrikake.Text).Append("  ");
                    sb.Append("�����c�v�F").Append(txtZan.Text);
                    oxlsSheet.Cells[1, 6] = sb.ToString();


                    // �O���b�h������
                    for (int iX = 0; iX <= tempDGV.RowCount - 1; iX++)
                    {
                        oxlsSheet.Cells[S_GYO - 3, S_ROWSMAX] = int.Parse(this.txtZan.Text, System.Globalization.NumberStyles.Any);
                        oxlsSheet.Cells[iX + S_GYO, 1] = tempDGV[0, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 2] = tempDGV[1, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 3] = tempDGV[3, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 4] = tempDGV[4, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 5] = tempDGV[5, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 6] = tempDGV[6, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 7] = tempDGV[7, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 8] = tempDGV[8, iX].Value.ToString();
                        oxlsSheet.Cells[iX + S_GYO, 9] = tempDGV[9, iX].Value.ToString();

                        if (tempDGV[10, iX].Value.ToString() == "True")
                        {
                            oxlsSheet.Cells[iX + S_GYO, 10] = "*";
                        }
                        else
                        {
                            oxlsSheet.Cells[iX + S_GYO, 10] = string.Empty;
                        }

                        oxlsSheet.Cells[iX + S_GYO, 11] = tempDGV[12, iX].Value.ToString();
                    }

                    ////////�Z���㕔�֎������R�r��������
                    //////rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    //////rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    //////oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //�Z�������֎������R�r��������
                    rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, S_ROWSMAX];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

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
                    oXls.Visible = true;

                    //���
                    oxlsSheet.PrintPreview(true);

                    // �E�B���h�E���\���ɂ���
                    oXls.Visible = false;

                    //�ۑ�����
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //�_�C�A���O�{�b�N�X�̏����ݒ�
                    saveFileDialog1.Title = "�N���C�A���g�ʐ����ꗗ";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = "�N���C�A���g�ʐ����ꗗ";
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
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {

                    //Book���N���[�Y
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excel���I��
                    oXls.Quit();

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
                MessageBox.Show(e.Message, MESSAGE_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //�}�E�X�|�C���^�����ɖ߂�
            this.Cursor = Cursors.Default;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrn.Enabled = false;
            dataGridView1.RowCount = 0;
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.SelectAll();
            txtObj.BackColor = Color.LightGray;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();

            if (sender == txtsClientName) txtObj = txtsClientName;

            txtObj.BackColor = Color.White;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (mukouStatus)
            {
                // �����ς݁A���l
                if (e.ColumnIndex == 10 || e.ColumnIndex == 12)
                {
                    // ID�擾
                    int sID = Utility.strToInt(dataGridView1["col11", e.RowIndex].Value.ToString());

                    // �f�[�^�擾
                    var s = dts.�V������.Single(a => a.ID == sID);

                    // �����σ`�F�b�N
                    if (e.ColumnIndex == 10)
                    {
                        if (dataGridView1["col10", e.RowIndex].Value.ToString() == "True")
                        {
                            s.�������� = global.FLGON;
                        }
                        else
                        {
                            s.�������� = global.FLGOFF;
                        }
                    }
                    
                    // ���l
                    if (e.ColumnIndex == 12)
                    {
                        if (dataGridView1["col12", e.RowIndex].Value == null)
                        {
                            s.���l = string.Empty;
                        }
                        else
                        {
                            s.���l = dataGridView1["col12", e.RowIndex].Value.ToString();
                        }
                    }
                    
                    s.�ύX�N���� = DateTime.Now;
                    s.���[�U�[ID = global.loginUserID;

                    // �V�������f�[�^�X�V
                    rAdp.Update(dts.�V������);

                    // �f�[�^�ǂݍ���
                    rAdp.Fill(dts.�V������);
                }
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCellAddress.X == 10)
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            putKaikeiCsv();
        }


        private void putKaikeiCsv()
        {
            //�_�C�A���O�{�b�N�X�̏����ݒ�
            SaveFileDialog sf = new SaveFileDialog();
            sf.Title = "��vCSV�o��";
            sf.OverwritePrompt = true;
            sf.RestoreDirectory = true;
            sf.FileName = "�����s�ėp�f�[�^_����";
            sf.Filter = "Microsoft Office Excel�t�@�C��(*.csv)|*.csv";

            //�_�C�A���O�{�b�N�X��\�����u�ۑ��v�{�^�����I�����ꂽ��t�@�C������\��
            string fileName = string.Empty;
            DialogResult ret = sf.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                SaveData(sf.FileName, dataGridView1);
            }

        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     ��v�f�[�^�o�͏��� </summary>
        /// <param name="fName">
        ///     �o�̓t�@�C����</param>
        /// <param name="g">
        ///     �f�[�^�O���b�h�r���[�I�u�W�F�N�g</param>
        ///---------------------------------------------------------------------
        private void SaveData(string fName, DataGridView g)
        {
            string wrkOutputData = string.Empty;
            Boolean iniFlg = true;
            Boolean pblFirstGyouFlg = true;

            //�o�̓t�@�C���C���X�^���X�쐬
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(fName, false, System.Text.Encoding.GetEncoding(932));

            try
            {
                // �\�����f�[�^��ǂݏo��
                for (int i = 0; i < g.RowCount; i++)
                {
                    //// �������f�[�^�̓l�O��
                    //if (Utility.strToInt(g[8, i].Value.ToString()) <= 0)
                    //{
                    //    continue;
                    //}

                    //�w�b�_�t�@�C���o��
                    if (pblFirstGyouFlg)
                    {
                        wrkOutputData = string.Empty;
                        wrkOutputData += Entity.OutPutHeader.dn01 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd01 + ",";
                        wrkOutputData += Entity.OutPutHeader.hd02 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr02 + ",";

                        wrkOutputData += Entity.OutPutHeader.kr06 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks02 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks06 + ",";
                        wrkOutputData += Entity.OutPutHeader.kr01 + ",";

                        wrkOutputData += Entity.OutPutHeader.kr55 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks52 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks53 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks54 + ",";

                        wrkOutputData += Entity.OutPutHeader.ks55 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks04 + ",";
                        wrkOutputData += Entity.OutPutHeader.tk01 + ",";
                        wrkOutputData += Entity.OutPutHeader.ks01;

                        outFile.WriteLine(wrkOutputData);
                    }

                    //�o�̓f�[�^�쐬
                    if (pblFirstGyouFlg)
                    {
                        wrkOutputData = "*,";
                    }
                    else
                    {
                        wrkOutputData = ",";
                    }

                    wrkOutputData += "0,";          // �`�[����R�[�h
                    wrkOutputData += g[4, i].Value.ToString() + ",";  // ���s��
                    wrkOutputData += "135,";        // �ؕ�����ȖڃR�[�h
                    wrkOutputData += Utility.strToInt(g[5, i].Value.ToString()) + ",";    // �ؕ��{�̋��z
                    wrkOutputData += "501,";        // �ݕ�����ȖڃR�[�h
                    wrkOutputData += Utility.strToInt(g[5, i].Value.ToString()) + ",";    // �ݕ��{�̋��z
                    wrkOutputData += "20,";         // �ؕ�����R�[�h
                    wrkOutputData += (Utility.strToInt(g[0, i].Value.ToString()) + 990000000) + ",";      // �����R�[�h
                    wrkOutputData += ",";       // �ݕ��⏕�ȖڃR�[�h
                    wrkOutputData += "3,";      // �ݕ��ŗ��敪�R�[�h
                    wrkOutputData += "8,";      // �ݕ��ŗ�
                    wrkOutputData += "2,";      // �ݕ�����Ōv�Z
                    wrkOutputData += "2,";      // �ݕ��[������
                    wrkOutputData += g[2, i].Value.ToString().Replace(",","") + ",";        // �E�v
                    wrkOutputData += "20";      // �ݕ�����R�[�h

                    // �t�@�C���֏o��            
                    outFile.WriteLine(wrkOutputData);
                    pblFirstGyouFlg = false;
                }

                //�t�@�C���N���[�Y
                outFile.Close();

                // �I�����b�Z�[�W
                MessageBox.Show("�쐬�I��", "�����s�ėp�f�[�^", MessageBoxButtons.OK);

                ////�o�̓t�@�C���폜
                //utility.FileDelete(global.WorkDir + global.DIR_OK, global.OUTFILE);

                ////�ꎞ�t�@�C�����o�̓t�@�C���ɃR�s�[
                //File.Copy(global.WorkDir + global.DIR_OK + global.tmpFile, global.WorkDir + global.DIR_OK + global.OUTFILE);

                ////�ꎞ�t�@�C���폜
                //utility.FileDelete(global.WorkDir + global.DIR_OK, global.tmpFile);
            }
            catch (Exception e)
            {
                MessageBox.Show("�f�[�^�ϊ���" + Environment.NewLine + e.Message, "�G���[", MessageBoxButtons.OK);
            }
        }
    }
}
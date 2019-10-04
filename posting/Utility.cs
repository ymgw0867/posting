using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;

namespace posting
{
    class Utility
    {
        public class DBConnect
        {
            OleDbConnection cn = new OleDbConnection() ;

            public OleDbConnection Cn
            {
                get
                {
                return cn;
                }
            }
            
            private string sServer;
            private string sDataBase;
            private string sUI;
            private string sPS;

            public DBConnect()
            {
                try
                {
                    // MySeting���ڂ̎擾
                    //�T�[�o��
                    sServer = Properties.Settings.Default.sServer;

                    // �f�[�^�x�[�X��
                    sDataBase = Properties.Settings.Default.sDataBase;

                    // ���[�U�[ID
                    sUI = Properties.Settings.Default.sUI;

                    // �p�X���[�h
                    sPS = Properties.Settings.Default.sPs;

                    // �f�[�^�x�[�X�ڑ�������
                    //////cn.ConnectionString += "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                    //////cn.ConnectionString += sDBPath;
                    //////cn.ConnectionString += @"\";
                    //////cn.ConnectionString += sDBName;
                    //////cn.ConnectionString += ";Jet OLEDB:Database Password=";

                    ////////cn.ConnectionString += sDBPword;

                    cn.ConnectionString = "Provider=SQLOLEDB;";
                    cn.ConnectionString += "Data Source=" + sServer + ";";
                    cn.ConnectionString += "Initial Catalog=" + sDataBase + ";";
                    cn.ConnectionString += "Persist Security Info=True;";
                    cn.ConnectionString += "User ID=" + sUI + ";";
                    cn.ConnectionString += "Password=" + sPS +";";

                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     ������̒l���������`�F�b�N���� </summary>
        /// <param name="tempStr">
        ///     ���؂��镶����</param>
        /// <returns>
        ///     ����:true,�����łȂ�:false</returns>
        ///----------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     �����^��Int�֕ϊ����ĕԂ��i���l�łȂ��Ƃ��͂O��Ԃ��j</summary>
        /// <param name="tempStr">
        ///     �����^�̒l</param>
        /// <returns>
        ///     Long�^�̒l</returns>
        /// --------------------------------------------------------------------------------
        public static long strToLong(string tempStr)
        {
            tempStr = tempStr.Replace(",", "");
            if (tempStr == "0.0000") tempStr = "0";

            if (NumericCheck(tempStr))
            {
                return long.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     �����^��decimal�֕ϊ����ĕԂ��i���l�łȂ��Ƃ��͂O��Ԃ��j</summary>
        /// <param name="tempStr">
        ///     �����^�̒l</param>
        /// <returns>
        ///     decimal�^�̒l</returns>
        /// --------------------------------------------------------------------------------
        public static decimal strToDecimal(string tempStr)
        {
            tempStr = tempStr.Replace(",", "");

            if (NumericCheck(tempStr))
            {
                return decimal.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     �����^��Int�֕ϊ����ĕԂ��i���l�łȂ��Ƃ��͂O��Ԃ��j</summary>
        /// <param name="tempStr">
        ///     �����^�̒l</param>
        /// <returns>
        ///     Int�^�̒l</returns>
        /// --------------------------------------------------------------------------------
        public static int strToInt(string tempStr)
        {
            tempStr = tempStr.Replace(",", "");

            if (NumericCheck(tempStr))
            {
                return int.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     �����^��Int�֕ϊ����ĕԂ��i���l�łȂ��Ƃ��͂O��Ԃ��j</summary>
        /// <param name="tempStr">
        ///     �����^�̒l</param>
        /// <returns>
        ///     Int�^�̒l</returns>
        /// --------------------------------------------------------------------------------
        public static Int64 strToInt64(string tempStr)
        {
            tempStr = tempStr.Replace(",", "");

            if (NumericCheck(tempStr))
            {
                return Int64.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }
        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     �����^��double�֕ϊ����ĕԂ��i���l�łȂ��Ƃ��͂O��Ԃ��j</summary>
        /// <param name="tempStr">
        ///     �����^�̒l</param>
        /// <returns>
        ///     double�^�̒l</returns>
        /// --------------------------------------------------------------------------------
        public static double strToDouble(string tempStr)
        {
            tempStr = tempStr.Replace(",", "");

            if (NumericCheck(tempStr))
            {
                return double.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     �I�u�W�F�N�g��NULL�܂��͐����ȊO�Ȃ�Ȃ�[����Ԃ��A�����Ȃ��int�ɕϊ����ĕԂ� </summary>
        /// <param name="obj">
        ///     �I�u�W�F�N�g</param>
        /// <returns>
        ///     �߂�l</returns>
        ///--------------------------------------------------------------------------------------------------
        public static int nullToInt(object obj)
        {
            int rtn = 0; 
            if (obj != null)
            {
                if (!int.TryParse(obj.ToString().Replace(",", ""), out rtn))
                {
                    rtn = 0;
                }
            }
            return rtn;
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     �I�u�W�F�N�g��NULL�Ȃ�l�Ȃ���Ԃ��BNUll�ł͂Ȃ��Ƃ���string�^�ɕϊ����ĕԂ� </summary>
        /// <param name="obj">
        ///     �I�u�W�F�N�g</param>
        /// <returns>
        ///     �߂�l</returns>
        ///--------------------------------------------------------------------------------------------------
        public static string nullToStr(object obj)
        {
            string rtn = string.Empty;

            if (obj == null)
            {
                return rtn;
            }
            else
            {
                return obj.ToString();
            }
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     �I�u�W�F�N�g��NULL�܂��͐����ȊO�Ȃ�Ȃ�[����Ԃ��A�����Ȃ��long�ɕϊ����ĕԂ� </summary>
        /// <param name="obj">
        ///     �I�u�W�F�N�g</param>
        /// <returns>
        ///     �߂�l</returns>
        ///--------------------------------------------------------------------------------------------------
        public static long nullTolng(object obj)
        {
            int rtn = 0;
            if (obj != null)
            {
                if (!int.TryParse(obj.ToString().Replace(",", ""), out rtn))
                {
                    rtn = 0;
                }
            }
            return rtn;
        }


        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     �I�u�W�F�N�g��NULL�܂��͐����ȊO�Ȃ�Ȃ�[����Ԃ��A�����Ȃ��Double�ɕϊ����ĕԂ� </summary>
        /// <param name="obj">
        ///     �I�u�W�F�N�g</param>
        /// <returns>
        ///     �߂�l</returns>
        ///--------------------------------------------------------------------------------------------------
        public static double nullToDouble(object obj)
        {
            double rtn = 0;

            if (obj != null)
            {
                if (!double.TryParse(obj.ToString().Replace(",", ""), out rtn))
                {
                    rtn = 0;
                }
            }
            return rtn;
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     �I�u�W�F�N�g��NULL�܂��͓��t�ȊO�Ȃ�Ȃ�l�Ȃ���Ԃ��A���t�Ȃ��ShortDateString�ɕϊ����ĕԂ� </summary>
        /// <param name="obj">
        ///     �I�u�W�F�N�g</param>
        /// <returns>
        ///     �߂�l</returns>
        ///--------------------------------------------------------------------------------------------------
        public static string nullToDate(object obj)
        {
            DateTime dt = DateTime.Parse("1900/01/01");

            if (obj != null)
            {
                if (!DateTime.TryParse(obj.ToString(), out dt))
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }

            return dt.ToShortDateString();
        }



        // �������[�h
        public class formMode
        {
            //private int F_Mode;
            public int Mode { get; set; }
            public int ID { get; set; }
            public long jID { get; set; }
            public string sID { get; set; }
            public int kingaku { get; set; }
        }

        // �������[�h
        public class areaMode
        {
            private int F_areaMode;
            private int F_rowIndex;

            public int Mode
            {
                get
                {
                    return F_areaMode;
                }
                set
                {
                    F_areaMode = value;
                }
            }

            public int RowIndex
            {
                get
                {
                    return F_rowIndex;
                }
                set
                {
                    F_rowIndex = value;
                }
            }
        }

        // ����ŗ�
        public class ����ŗ�
        {
            private int F_RT;

            public int Ritsu
            {
                get
                {
                    return F_RT;
                }
                set
                {
                    F_RT = value;
                }

            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     ���O�C�����[�U�[�R���{�{�b�N�X�N���X </summary>
        ///------------------------------------------------------------
        public class comboLoginUser
        {
            public int ID { get; set; }
            public string Name { get; set; }

            ///-------------------------------------------------------------
            /// <summary>
            ///     ���O�C�����[�U�[�R���{�{�b�N�Xitem���[�h </summary>
            /// <param name="cmbObj">
            ///     �R���{�{�b�N�X�I�u�W�F�N�g</param>
            ///-------------------------------------------------------------
            public static void itemLoad(ComboBox cmbObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.���O�C�����[�U�[TableAdapter adp = new darwinDataSetTableAdapters.���O�C�����[�U�[TableAdapter();
                    adp.Fill(dts.���O�C�����[�U�[);

                    comboLoginUser[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.���O�C�����[�U�[)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboLoginUser();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.���O�C�����[�U�[;
                        iX++;
                    }

                    cmbObj.DataSource = sList;
                    cmbObj.DisplayMember = "Name";
                    cmbObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���O�C�����[�U�[ �R���{�{�b�N�X���[�h");
                }
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     ���O�C���^�C�v�w�b�_�R���{�{�b�N�X�N���X </summary>
        ///------------------------------------------------------------
        public class comboLogintype
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public int Lock { get; set; }
            public int seigen { get; set; }
            public int Jyuryo { get; set; }

            ///-------------------------------------------------------------
            /// <summary>
            ///     ���O�C���^�C�v�w�b�_�R���{�{�b�N�Xitem���[�h </summary>
            /// <param name="cmbObj">
            ///     �R���{�{�b�N�X�I�u�W�F�N�g</param>
            ///-------------------------------------------------------------
            public static void itemLoad(ComboBox cmbObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.���O�C���^�C�v�w�b�_TableAdapter adp = new darwinDataSetTableAdapters.���O�C���^�C�v�w�b�_TableAdapter();
                    adp.Fill(dts.���O�C���^�C�v�w�b�_);

                    comboLogintype[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.���O�C���^�C�v�w�b�_)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboLogintype();
                        sList[iX].ID = t.Id;
                        sList[iX].Name = t.����;

                        iX++;
                    }

                    cmbObj.DataSource = sList;
                    cmbObj.DisplayMember = "Name";
                    cmbObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���O�C���^�C�v�w�b�_ �R���{�{�b�N�X���[�h");
                }
            }


            ///-------------------------------------------------------------
            /// <summary>
            ///     ���O�C���^�C�v�w�b�_�R���{�{�b�N�Xitem���[�h </summary>
            /// <param name="cmbObj">
            ///     �R���{�{�b�N�X�I�u�W�F�N�g</param>
            ///-------------------------------------------------------------
            public static void itemLoad(CheckedListBox lstObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.���O�C���^�C�v�w�b�_TableAdapter adp = new darwinDataSetTableAdapters.���O�C���^�C�v�w�b�_TableAdapter();
                    adp.Fill(dts.���O�C���^�C�v�w�b�_);

                    comboLogintype[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.���O�C���^�C�v�w�b�_)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboLogintype();
                        sList[iX].ID = t.Id;
                        sList[iX].Name = t.����;
                        sList[iX].Lock = t.�󒍌ʃ��b�N����;
                        sList[iX].seigen = t.�󒍌ʐ���;
                        sList[iX].Jyuryo = t.��������̍ς݌���;

                        iX++;
                    }

                    lstObj.DataSource = sList;
                    lstObj.DisplayMember = "Name";
                    lstObj.ValueMember = "ID";
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���O�C���^�C�v�w�b�_ �`�F�b�N���X�g�{�b�N�X���[�h");
                }
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     �O����}�X�^�[�R���{�{�b�N�X�N���X </summary>
        ///------------------------------------------------------------
        public class comboGaichu
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public int shiharaibi { get; set; }     // �x���� 2018/01/04

            public static comboGaichuUser[] getArrayGaichu()
            {
                comboGaichuUser[] sList = null;

                darwinDataSet dts = new darwinDataSet();
                darwinDataSetTableAdapters.�O����TableAdapter adp = new darwinDataSetTableAdapters.�O����TableAdapter();
                adp.Fill(dts.�O����);

                int iX = 0;

                foreach (var t in dts.�O����)
                {
                    Array.Resize(ref sList, iX + 1);
                    sList[iX] = new comboGaichuUser();
                    sList[iX].ID = t.ID;
                    sList[iX].Name = t.����;
                    sList[iX].shiharaibi = t.�x����;
                    iX++;
                }

                return sList;
            }


            ///-------------------------------------------------------------
            /// <summary>
            ///     �O����R���{�{�b�N�Xitem���[�h </summary>
            /// <param name="cmbObj">
            ///     �R���{�{�b�N�X�I�u�W�F�N�g</param>
            ///-------------------------------------------------------------
            public static void itemLoad(ComboBox cmbObj)
            {
                try 
	            {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.�O����TableAdapter adp = new darwinDataSetTableAdapters.�O����TableAdapter();
                    adp.Fill(dts.�O����);

                    comboGaichu[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.�O����)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboGaichu();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.����;
                        sList[iX].shiharaibi = t.�x����;
                        iX++;
                    }

                    cmbObj.DataSource = sList;
                    cmbObj.DisplayMember = "Name";
                    cmbObj.ValueMember = "ID";
	            }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�O����R���{�{�b�N�X���[�h");
                }
            }
        }

        public class comboGaichuUser
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public int shiharaibi { get; set; }     // �x���� 2018/01/04
        }


        // �����R���{�{�b�N�X�N���X
        public class ComboShozoku
        {
            private int F_ID;
            private string F_Name;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            // �����}�X�^�[���[�h
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboShozoku cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.���� Shozoku = new Control.����();
                    dR = Shozoku.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboShozoku();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["������1"].ToString() + " " + dR["������2"].ToString();
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    Shozoku.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�����R���{�{�b�N�X���[�h");
                }

            }

            // �������R���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboShozoku cmbS = new ComboShozoku();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboShozoku)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }


        /// �Ј��R���{�{�b�N�X�N���X
        ///     
        public class ComboShain
        {
            public int ID { get; set; }
            public string Name { get; set; }

            /// ------------------------------------------------------------------
            /// <summary>
            ///     �Ј��}�X�^�[���[�h </summary>
            /// <param name="tempObj">
            ///     �R���{�{�b�N�X�I�u�W�F�N�g</param>
            /// ------------------------------------------------------------------
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.�Ј�TableAdapter adp = new darwinDataSetTableAdapters.�Ј�TableAdapter();
                    adp.Fill(dts.�Ј�);

                    ComboShain[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.�Ј�)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new ComboShain();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.����;
                        iX++;
                    }

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�Ј��R���{�{�b�N�X���[�h");
                }
            }
             
            // �Ј��R���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboShain cmbS = new ComboShain();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboShain)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }
            }
        }

        // ������ʃR���{�{�b�N�X�N���X
        public class ComboKouza
        {
            private int F_ID;
            private string F_Name;

            private const int kCode_F = 1;
            private const string kName_F = "����";

            private const int kCode_T = 2;
            private const string kName_T = "����";

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            // ������ʃZ�b�g
            public static void load(ComboBox tempObj)
            {
                try
                {
                    ComboKouza cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    //����
                    cmb1 = new ComboKouza();
                    cmb1.ID = kCode_F;
                    cmb1.Name = kName_F;
                    tempObj.Items.Add(cmb1);

                    //����
                    cmb1 = new ComboKouza();
                    cmb1.ID = kCode_T;
                    cmb1.Name = kName_T;
                    tempObj.Items.Add(cmb1);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "������ʃZ�b�g");
                }

            }

            //������ʃR���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboKouza cmbS = new ComboKouza();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboKouza)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // ���Ə��R���{�{�b�N�X�N���X : 2015/06/24
        public class ComboOffice
        {
            public int ID { get; set; }
            public string Name { get; set; }

            //���Ə��}�X�^�[���[�h : 2015/06/24
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.���Ə�TableAdapter adp = new darwinDataSetTableAdapters.���Ə�TableAdapter();
                    adp.Fill(dts.���Ə�);

                    ComboOffice[] fList = null;
                    int iX = 0;

                    foreach (var t in dts.���Ə�.OrderBy(a => a.ID))
                    {
                        Array.Resize(ref fList, iX + 1);
                        fList[iX] = new ComboOffice();
                        fList[iX].ID = t.ID;
                        fList[iX].Name = t.����;
                        iX++;
                    }

                    tempObj.DataSource = fList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���Ə��R���{�{�b�N�X���[�h");
                }

            }

            //���Ə��R���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboOffice cmbS = new ComboOffice();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboOffice)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // ���Ӑ�R���{�{�b�N�X�N���X : 2015/06/24
        public class ComboClient
        {
            private int F_ID;
            private string F_Name;
            private string F_NameShow;

            public int ID { get; set; }
            public string Name { get; set; }
            public string NameShow { get; set; }

            // ���Ӑ�}�X�^�[���[�h
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboClient cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.���Ӑ� Client = new Control.���Ӑ�();
                    dR = Client.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboClient();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["ID"].ToString() + " �F " + dR["����"].ToString() + "";
                        cmb1.NameShow = dR["����"].ToString() + "";
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    Client.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���Ӑ�R���{�{�b�N�X���[�h");
                }

            }

            /// ------------------------------------------------------------------
            /// <summary>
            ///     �R���{�{�b�N�X�ɓ��Ӑ�}�X�^�[�����[�h���� </summary>
            /// <param name="cmbObj">
            ///     �R���{�{�b�N�X�I�u�W�F�N�g</param>
            ///      : 2015/06/24
            /// ------------------------------------------------------------------
            public static void itemsLoad(ComboBox cmbObj)
            {
                ComboClient[] cList = null;
                int iX = 0;

                darwinDataSet dts = new darwinDataSet();
                darwinDataSetTableAdapters.���Ӑ�TableAdapter tAdp = new darwinDataSetTableAdapters.���Ӑ�TableAdapter();
                tAdp.Fill(dts.���Ӑ�);

                foreach (var t in dts.���Ӑ�)
                {
                    Array.Resize(ref cList, iX + 1);
                    cList[iX] = new ComboClient();
                    cList[iX].ID = t.ID;
                    cList[iX].Name = t.ID.ToString() + " �F " + t.����.ToString();
                    cList[iX].NameShow = t.����.ToString();
                    iX++;
                }

                cmbObj.DataSource = cList;
                cmbObj.DisplayMember = "Name";
                cmbObj.ValueMember = "ID";
            }

        }

        // �z�z�`�ԃR���{�{�b�N�X�N���X : 2015/06/24
        public class ComboFkeitai
        {
            public int ID { get; set; }
            public string  Name { get; set; }

            // �z�z�`�ԃ}�X�^�[���[�h : 2015/06/24
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.�z�z�`��TableAdapter adp = new darwinDataSetTableAdapters.�z�z�`��TableAdapter();
                    adp.Fill(dts.�z�z�`��);

                    ComboFkeitai[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.�z�z�`��)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new ComboFkeitai();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.����;
                        iX++;
                    }

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�z�z�`�ԃR���{�{�b�N�X���[�h");
                }
            }

            // �z�z�`�ԃR���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboFkeitai cmbS = new ComboFkeitai();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboFkeitai)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }
            }
        }

        // �Č���ʃR���{�{�b�N�X�N���X : 2015/06/30
        public class ComboAnshu
        {
            public int ID { get; set; }
            public string Name { get; set; }

            // �Č���ʃA�C�e���Z�b�g : 2015/06/30
            public static void load(ComboBox tempObj)
            {
                try
                {
                    ComboAnshu[] sList = new ComboAnshu[3];
                    sList[0] = new ComboAnshu();
                    sList[0].ID = 1;
                    sList[0].Name = "���Ѓ|�X";

                    sList[1] = new ComboAnshu();
                    sList[1].ID = 2;
                    sList[1].Name = "���Ѓ|�X";

                    sList[2] = new ComboAnshu();
                    sList[2].ID = 3;
                    sList[2].Name = "�|�X�ȊO�Č�";

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�Č���ʃA�C�e���Z�b�g");
                }
            }
        }

        // ���^�R���{�{�b�N�X�N���X : 2015/06/24
        public class ComboSize
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public decimal Tanka { get; set; }

            // ���^�}�X�^�[���[�h : 2015/06/24
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.���^TableAdapter adp = new darwinDataSetTableAdapters.���^TableAdapter();
                    adp.Fill(dts.���^);

                    ComboSize[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.���^)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new ComboSize();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.����;
                        sList[iX].Tanka = t.���P��1;
                        iX++;
                    }

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "���^�R���{�{�b�N�X���[�h");
                }
            }

            // ���^�R���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboSize cmbS = new ComboSize();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboSize)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // �����p�^�[���R���{�{�b�N�X�N���X
        public class ComboShimebi
        {
            private int F_ID;
            private string F_Name;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            // �����p�^�[���}�X�^�[���[�h
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboShimebi cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.�����p�^�[�� cPat = new Control.�����p�^�[��();
                    dR = cPat.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboShimebi();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["�E�v"].ToString() + "";
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    cPat.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�����p�^�[���R���{�{�b�N�X���[�h");
                }

            }

            // �����p�^�[���R���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboShimebi cmbS = new ComboShimebi();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboShimebi)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // �U�������R���{�{�b�N�X�N���X
        public class ComboFuri
        {
            private int F_ID;
            private string F_Name;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            // �U�������}�X�^�[���[�h
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboFuri cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.�U������ cFuri = new Control.�U������();
                    dR = cFuri.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboFuri();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["���Z�@�֖�"].ToString() + "�@" + dR["�x�X��"].ToString() + "�@";

                        switch (Int32.Parse(dR["�������"].ToString()))
                        {
                            case 1:
                                cmb1.Name += "����" + "�@";
                                break;
                            case 2:
                                cmb1.Name += "����" + "�@";
                                break;
                        }

                        cmb1.Name += dR["�����ԍ�"].ToString();

                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    cFuri.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�U�������R���{�{�b�N�X���[�h");
                }

            }

            // �U�������R���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboFuri cmbS = new ComboFuri();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboFuri)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }
        
        public class comboOri
        {
            public int ID { get; set; }
            public string Name { get; set; }

            // �܃}�X�^�[���[�h
            public static void load(DataGridViewComboBoxColumn tempObj)
            {
                try
                {
                    comboOri[] jList = new comboOri[2];

                    jList[0] = new comboOri();
                    jList[0].ID = 1;
                    jList[0].Name = "�L";
                    jList[1] = new comboOri();
                    jList[1].ID = 2;
                    jList[1].Name = "��";

                    tempObj.DataSource = jList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�܃R���{�{�b�N�X���[�h");
                }

            }

        }

        // �󒍎�ʃR���{�{�b�N�X�N���X
        public class ComboJshubetsu
        {
            public int ID { get; set; }
            public string Name { get; set; }

            // �󒍎�ʃ}�X�^�[���[�h
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.�󒍎��TableAdapter jAdp = new darwinDataSetTableAdapters.�󒍎��TableAdapter();
                    jAdp.Fill(dts.�󒍎��);

                    ComboJshubetsu[] jList = null;
                    int iX = 0;

                    foreach (var t in dts.�󒍎��)
                    {
                        Array.Resize(ref jList, iX + 1);
                        jList[iX] = new ComboJshubetsu();
                        jList[iX].ID = t.ID;
                        jList[iX].Name = t.����;
                        iX++;
                    }

                    tempObj.DataSource = jList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�󒍎�ʃR���{�{�b�N�X���[�h");
                }

            }

            // �󒍎�ʃ}�X�^�[���[�h
            public static void load(DataGridViewComboBoxColumn tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.�󒍎��TableAdapter jAdp = new darwinDataSetTableAdapters.�󒍎��TableAdapter();
                    jAdp.Fill(dts.�󒍎��);

                    ComboJshubetsu[] jList = null;
                    int iX = 0;

                    foreach (var t in dts.�󒍎��)
                    {
                        Array.Resize(ref jList, iX + 1);
                        jList[iX] = new ComboJshubetsu();
                        jList[iX].ID = t.ID;
                        jList[iX].Name = t.����;
                        iX++;
                    }

                    tempObj.DataSource = jList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�󒍎�ʃR���{�{�b�N�X���[�h");
                }

            }

            // �󒍎�ʃR���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboJshubetsu cmbS = new ComboJshubetsu();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboJshubetsu)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // �z�z���R���{�{�b�N�X�N���X
        public class ComboStaff
        {
            private int F_ID;
            private string F_Name;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            // �z�z���}�X�^�[���[�h
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboStaff cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.�z�z�� Staff = new Control.�z�z��();
                    dR = Staff.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboStaff();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["ID"].ToString() + " �F " +dR["����"].ToString();
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    Staff.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�z�z���R���{�{�b�N�X���[�h");
                }

            }

            // �z�z���R���{�\��
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboStaff cmbS = new ComboStaff();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboStaff)tempObj.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        Sh = true;
                        break;
                    }

                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // �V�󃊃X�g�{�b�N�X�N���X
        public class ComboTenkou 
        {
            private DateTime F_DATE;
            private string F_Name;

            public DateTime sDATE
            {
                get
                {
                    return F_DATE;
                }
                set
                {
                    F_DATE = value;
                }
            }

            public string Name
            {
                get
                {
                    return F_Name;
                }
                set
                {
                    F_Name = value;
                }
            }

            // �V��}�X�^�[���[�h
            public static void load(ComboBox tempObj)
            {
                try
                {
                    string SqlSTR;
                    OleDbDataReader dR;
                    ComboTenkou cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "sDATE";

                    Control.FreeSql cTenkou = new Control.FreeSql();

                    SqlSTR = "";
                    SqlSTR += "select distinct �V�� from �V��";
                    dR = cTenkou.free_dsReader(SqlSTR);

                    while (dR.Read())
                    {
                        cmb1 = new ComboTenkou();
                        cmb1.Name = dR["�V��"].ToString() + "";
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    cTenkou.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "�V�󃊃X�g�{�b�N�X���[�h");
                }

            }


        }
        public static void FlgOnOff(int tempFLG)
        {
            string sqlSTR;
            sqlSTR = "";
            sqlSTR += "update ��Џ�� ";
            sqlSTR += "set �z�z�t���O = " + tempFLG.ToString() + ",";
            sqlSTR += "�ύX�N���� = '" + DateTime.Today + "'";

            Control.FreeSql fCon = new Control.FreeSql();
            fCon.Execute(sqlSTR);
        }

        /// <summary>
        /// �E�B���h�E�ŏ��T�C�Y�̐ݒ�
        /// </summary>
        /// <param name="tempFrm">�ΏۂƂ���E�B���h�E�I�u�W�F�N�g</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">Height</param>
        public static void WindowsMinSize(Form tempFrm,int wSize,int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize); 
        }

        /// <summary>
        /// �E�B���h�E�ŏ��T�C�Y�̐ݒ�
        /// </summary>
        /// <param name="tempFrm">�ΏۂƂ���E�B���h�E�I�u�W�F�N�g</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">height</param>
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }
        
        /// ----------------------------------------------------------------------------
        /// <summary>
        ///     ���j���[�^�C�g���z�񂩂烁�j���[�{�^���̃e�L�X�g���Z�b�g���� </summary>
        /// <param name="btn">
        ///     �{�^���I�u�W�F�N�g </param>
        /// <param name="cm">
        ///     ���j���[�^�C�g���N���X </param>
        /// ----------------------------------------------------------------------------
        public static void getMenuTittle(Button btn, clsMenu cm)
        {
            foreach (var t in cm.menuCsv)
            {
                // �J���}��؂��1�s�̃f�[�^�z����擾
                string[] k = t.Split(',');

                // �z��̗v�f����2�ȏ゠�邩
                if (k.Length > 1)
                {
                    // Tag���ƍ����ă��j���[�^�C�g�����Z�b�g����
                    if (k[0].Equals(btn.Tag))
                    {
                        btn.Text = k[1];
                        break;
                    }
                }
            }
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     �󒍔ԍ����̔Ԃ��� </summary>
        /// <param name="dt">
        ///     ���ɓ�</param>
        /// <returns>
        ///     �󒍔ԍ�</returns>
        /// --------------------------------------------------------------------
        public static long getOrderNum(DateTime dt)
        {
            //ID���̔�
            long gID = 0;
            DateTime sJDate = dt;

            // �`�[�ԍ��ŏ��l
            Int64 denNumMin = ((long)sJDate.Year * 100000000) + (sJDate.Month * 1000000) + (sJDate.Day * 10000);
            denNumMin++;

            // �`�[�ԍ��ő�l
            long denNumMax = denNumMin + 9998;

            // �󒍓��̓`�[�����邩�H
            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.��1TableAdapter rAdp = new darwinDataSetTableAdapters.��1TableAdapter();
            rAdp.Fill(dts.��1);

            long s = 0;

            if (dts.��1.Any(a => a.ID >= denNumMin && a.ID <= denNumMax))
            {
                s = dts.��1.Where(a => a.ID >= denNumMin && a.ID <= denNumMax).Max(a => a.ID);
            }

            // �̔ԍς݂��H
            darwinDataSetTableAdapters.�󒍔ԍ��̔�TableAdapter jAdp = new darwinDataSetTableAdapters.�󒍔ԍ��̔�TableAdapter();
            jAdp.Fill(dts.�󒍔ԍ��̔�);

            long j = 0;
            
            if (dts.�󒍔ԍ��̔�.Any(a => a.�󒍔ԍ� >= denNumMin && a.�󒍔ԍ� <= denNumMax))
            {
                j = dts.�󒍔ԍ��̔�.Where(a => a.�󒍔ԍ� >= denNumMin && a.�󒍔ԍ� <= denNumMax).Max(a => a.�󒍔ԍ�);
            }

            if (s == 0 && j == 0)   // �Y�����̎󒍔ԍ������̔Ԃ̂Ƃ�
            {
                gID = denNumMin;
            }
            else if (s > j)    // �󒍔ԍ���r�F�傫������1���Z�������̂�V�󒍔ԍ��Ƃ���
            {
                gID = s + 1;
            }
            else if (j > s)
            {
                gID = j + 1;
            }

            // �󒍔ԍ���Ԃ�
            return gID;
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     DataGridViewMaskedTextBoxCell�I�u�W�F�N�g�̗��\���܂��B </summary>
        /// ------------------------------------------------------------------------------
        public class DataGridViewMaskedTextBoxColumn :
            DataGridViewColumn
        {
            //CellTemplate�Ƃ���DataGridViewMaskedTextBoxCell�I�u�W�F�N�g���w�肵��
            //��{�N���X�̃R���X�g���N�^���Ăяo��
            public DataGridViewMaskedTextBoxColumn()
                : base(new DataGridViewMaskedTextBoxCell())
            {
            }

            private string maskValue = "";
            /// <summary>
            /// MaskedTextBox��Mask�v���p�e�B�ɓK�p����l
            /// </summary>
            public string Mask
            {
                get
                {
                    return this.maskValue;
                }
                set
                {
                    this.maskValue = value;
                }
            }

            //�V�����v���p�e�B��ǉ����Ă��邽�߁A
            // Clone���\�b�h���I�[�o�[���C�h����K�v������
            public override object Clone()
            {
                DataGridViewMaskedTextBoxColumn col =
                    (DataGridViewMaskedTextBoxColumn)base.Clone();
                col.Mask = this.Mask;
                return col;
            }

            //CellTemplate�̎擾�Ɛݒ�
            public override DataGridViewCell CellTemplate
            {
                get
                {
                    return base.CellTemplate;
                }
                set
                {
                    //DataGridViewMaskedTextBoxCell����
                    // CellTemplate�ɐݒ�ł��Ȃ��悤�ɂ���
                    if (!(value is DataGridViewMaskedTextBoxCell))
                    {
                        throw new InvalidCastException(
                            "DataGridViewMaskedTextBoxCell�I�u�W�F�N�g��" +
                            "�w�肵�Ă��������B");
                    }
                    base.CellTemplate = value;
                }
            }
        }
        
        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     MaskedTextBox�ŕҏW�ł���e�L�X�g����
        ///     DataGridView�R���g���[���ɕ\�����܂��B</summary>
        /// ------------------------------------------------------------------------------
        public class DataGridViewMaskedTextBoxCell :
            DataGridViewTextBoxCell
        {
            //�R���X�g���N�^
            public DataGridViewMaskedTextBoxCell()
            {
            }

            //�ҏW�R���g���[��������������
            //�ҏW�R���g���[���͕ʂ̃Z�����ł��g���܂킳��邽�߁A�������̕K�v������
            public override void InitializeEditingControl(
                int rowIndex, object initialFormattedValue,
                DataGridViewCellStyle dataGridViewCellStyle)
            {
                base.InitializeEditingControl(rowIndex,
                    initialFormattedValue, dataGridViewCellStyle);

                //�ҏW�R���g���[���̎擾
                DataGridViewMaskedTextBoxEditingControl maskedBox =
                    this.DataGridView.EditingControl as
                    DataGridViewMaskedTextBoxEditingControl;
                if (maskedBox != null)
                {
                    //Text��ݒ�
                    string maskedText = initialFormattedValue as string;
                    maskedBox.Text = maskedText != null ? maskedText : "";
                    //�J�X�^����̃v���p�e�B�𔽉f������
                    DataGridViewMaskedTextBoxColumn column =
                        this.OwningColumn as DataGridViewMaskedTextBoxColumn;
                    if (column != null)
                    {
                        maskedBox.Mask = column.Mask;
                    }
                }
            }

            //�ҏW�R���g���[���̌^���w�肷��
            public override Type EditType
            {
                get
                {
                    return typeof(DataGridViewMaskedTextBoxEditingControl);
                }
            }

            //�Z���̒l�̃f�[�^�^���w�肷��
            //�����ł́AObject�^�Ƃ���
            //��{�N���X�Ɠ����Ȃ̂ŁA�I�[�o�[���C�h�̕K�v�Ȃ�
            public override Type ValueType
            {
                get
                {
                    return typeof(object);
                }
            }

            //�V�������R�[�h�s�̃Z���̊���l���w�肷��
            public override object DefaultNewRowValue
            {
                get
                {
                    return base.DefaultNewRowValue;
                }
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     DataGridViewMaskedTextBoxCell�Ńz�X�g�����
        ///     MaskedTextBox�R���g���[����\���܂��B</summary>
        /// ------------------------------------------------------------------------------
        public class DataGridViewMaskedTextBoxEditingControl :
            MaskedTextBox, IDataGridViewEditingControl
        {
            //�ҏW�R���g���[�����\������Ă���DataGridView
            DataGridView dataGridView;
            //�ҏW�R���g���[�����\������Ă���s
            int rowIndex;
            //�ҏW�R���g���[���̒l�ƃZ���̒l���Ⴄ���ǂ���
            bool valueChanged;

            //�R���X�g���N�^
            public DataGridViewMaskedTextBoxEditingControl()
            {
                this.TabStop = false;
            }

            #region IDataGridViewEditingControl �����o

            //�ҏW�R���g���[���ŕύX���ꂽ�Z���̒l
            public object GetEditingControlFormattedValue(
                DataGridViewDataErrorContexts context)
            {
                return this.Text;
            }

            //�ҏW�R���g���[���ŕύX���ꂽ�Z���̒l
            public object EditingControlFormattedValue
            {
                get
                {
                    return this.GetEditingControlFormattedValue(
                        DataGridViewDataErrorContexts.Formatting);
                }
                set
                {
                    this.Text = (string)value;
                }
            }

            //�Z���X�^�C����ҏW�R���g���[���ɓK�p����
            //�ҏW�R���g���[���̑O�i�F�A�w�i�F�A�t�H���g�Ȃǂ��Z���X�^�C���ɍ��킹��
            public void ApplyCellStyleToEditingControl(
                DataGridViewCellStyle dataGridViewCellStyle)
            {
                this.Font = dataGridViewCellStyle.Font;
                this.ForeColor = dataGridViewCellStyle.ForeColor;
                this.BackColor = dataGridViewCellStyle.BackColor;
                switch (dataGridViewCellStyle.Alignment)
                {
                    case DataGridViewContentAlignment.BottomCenter:
                    case DataGridViewContentAlignment.MiddleCenter:
                    case DataGridViewContentAlignment.TopCenter:
                        this.TextAlign = HorizontalAlignment.Center;
                        break;
                    case DataGridViewContentAlignment.BottomRight:
                    case DataGridViewContentAlignment.MiddleRight:
                    case DataGridViewContentAlignment.TopRight:
                        this.TextAlign = HorizontalAlignment.Right;
                        break;
                    default:
                        this.TextAlign = HorizontalAlignment.Left;
                        break;
                }
            }

            //�ҏW����Z��������DataGridView
            public DataGridView EditingControlDataGridView
            {
                get
                {
                    return this.dataGridView;
                }
                set
                {
                    this.dataGridView = value;
                }
            }

            //�ҏW���Ă���s�̃C���f�b�N�X
            public int EditingControlRowIndex
            {
                get
                {
                    return this.rowIndex;
                }
                set
                {
                    this.rowIndex = value;
                }
            }

            //�l���ύX���ꂽ���ǂ���
            //�ҏW�R���g���[���̒l�ƃZ���̒l���Ⴄ���ǂ���
            public bool EditingControlValueChanged
            {
                get
                {
                    return this.valueChanged;
                }
                set
                {
                    this.valueChanged = value;
                }
            }


            /// ------------------------------------------------------------------------------
            /// <summary>
            ///     �w�肳�ꂽ�L�[��DataGridView���������邩�A�ҏW�R���g���[�����������邩
            ///     True��Ԃ��ƁA�ҏW�R���g���[������������
            ///     dataGridViewWantsInputKey��True�̎��́ADataGridView�������ł��� </summary>
            /// <param name="keyData"></param>
            /// <param name="dataGridViewWantsInputKey"></param>
            /// <returns></returns>
            /// ------------------------------------------------------------------------------
            public bool EditingControlWantsInputKey(
                Keys keyData, bool dataGridViewWantsInputKey)
            {
                //Keys.Left�ARight�AHome�AEnd�̎��́ATrue��Ԃ�
                //���̂悤�ɂ��Ȃ��ƁA�����̃L�[�ŕʂ̃Z���Ƀt�H�[�J�X���ڂ��Ă��܂�
                switch (keyData & Keys.KeyCode)
                {
                    case Keys.Right:
                    case Keys.End:
                    case Keys.Left:
                    case Keys.Home:
                        return true;
                    default:
                        return !dataGridViewWantsInputKey;
                }
            }

            //�}�E�X�J�[�\����EditingPanel��ɂ���Ƃ��̃J�[�\�����w�肷��
            //EditingPanel�͕ҏW�R���g���[�����z�X�g����p�l���ŁA
            //�ҏW�R���g���[�����Z����菬�����ƃR���g���[���ȊO�̕������p�l���ƂȂ�
            public Cursor EditingPanelCursor
            {
                get
                {
                    return base.Cursor;
                }
            }

            //�R���g���[���ŕҏW���鏀��������
            //�e�L�X�g��I����Ԃɂ�����A�}���|�C���^�𖖔��ɂ����肷��
            public void PrepareEditingControlForEdit(bool selectAll)
            {
                if (selectAll)
                {
                    //�I����Ԃɂ���
                    this.SelectAll();
                }
                else
                {
                    //�}���|�C���^�𖖔��ɂ���
                    this.SelectionStart = this.TextLength;
                }
            }

            //�l���ύX�������ɁA�Z���̈ʒu��ύX���邩�ǂ���
            //�l���ύX���ꂽ���ɕҏW�R���g���[���̑傫�����ύX����鎞��True
            public bool RepositionEditingControlOnValueChange
            {
                get
                {
                    return false;
                }
            }

            #endregion

            //�l���ύX���ꂽ��
            protected override void OnTextChanged(EventArgs e)
            {
                base.OnTextChanged(e);
                //�l���ύX���ꂽ���Ƃ�DataGridView�ɒʒm����
                this.valueChanged = true;
                this.dataGridView.NotifyCurrentCellDirty(true);
            }
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
        public static int GetTaxRT(DateTime tempDate)
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
        public static decimal GetTax(decimal tempKin, int tempTax)
        {
            decimal GakuD;
            //int GakuI;

            // �[���؎̂� 2015/07/01
            GakuD = Math.Floor(tempKin * tempTax / 100);

            //GakuD += (decimal)0.5;
            //GakuI = (int)GakuD;

            return GakuD;
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     �X�֔ԍ�CSV��z��ɓǂݍ��� </summary>
        /// <param name="z">
        ///     �z��</param>
        /// ----------------------------------------------------------------------
        public static void zipCsvLoad(ref string[] z)
        {
            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.��Џ��TableAdapter adp = new darwinDataSetTableAdapters.��Џ��TableAdapter();
            adp.Fill(dts.��Џ��);

            string denHeadCsvPath = string.Empty;

            // ��Џ�񂩂�X�֔ԍ�CSV�p�X���擾
            foreach (var t in dts.��Џ��)
            {
                denHeadCsvPath = t.�X�֔ԍ�CSV�p�X;
            }

            // �X�֔ԍ�CSV�t�@�C�������݂��Ă��Ȃ�������I��
            if (!System.IO.File.Exists(denHeadCsvPath))
            {
                return;
            }

            // �X�֔ԍ�CSV�ǂݍ���
            z = System.IO.File.ReadAllLines(denHeadCsvPath, Encoding.Default);
        }
    }
}

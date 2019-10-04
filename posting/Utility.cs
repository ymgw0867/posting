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
                    // MySeting項目の取得
                    //サーバ名
                    sServer = Properties.Settings.Default.sServer;

                    // データベース名
                    sDataBase = Properties.Settings.Default.sDataBase;

                    // ユーザーID
                    sUI = Properties.Settings.Default.sUI;

                    // パスワード
                    sPS = Properties.Settings.Default.sPs;

                    // データベース接続文字列
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
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
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
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Long型の値</returns>
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
        ///     文字型をdecimalへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     decimal型の値</returns>
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
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
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
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
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
        ///     文字型をdoubleへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     double型の値</returns>
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
        ///     オブジェクトがNULLまたは数字以外ならならゼロを返す、数字ならばintに変換して返す </summary>
        /// <param name="obj">
        ///     オブジェクト</param>
        /// <returns>
        ///     戻り値</returns>
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
        ///     オブジェクトがNULLなら値なしを返す。NUllではないときはstring型に変換して返す </summary>
        /// <param name="obj">
        ///     オブジェクト</param>
        /// <returns>
        ///     戻り値</returns>
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
        ///     オブジェクトがNULLまたは数字以外ならならゼロを返す、数字ならばlongに変換して返す </summary>
        /// <param name="obj">
        ///     オブジェクト</param>
        /// <returns>
        ///     戻り値</returns>
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
        ///     オブジェクトがNULLまたは数字以外ならならゼロを返す、数字ならばDoubleに変換して返す </summary>
        /// <param name="obj">
        ///     オブジェクト</param>
        /// <returns>
        ///     戻り値</returns>
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
        ///     オブジェクトがNULLまたは日付以外ならなら値なしを返す、日付ならばShortDateStringに変換して返す </summary>
        /// <param name="obj">
        ///     オブジェクト</param>
        /// <returns>
        ///     戻り値</returns>
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



        // 処理モード
        public class formMode
        {
            //private int F_Mode;
            public int Mode { get; set; }
            public int ID { get; set; }
            public long jID { get; set; }
            public string sID { get; set; }
            public int kingaku { get; set; }
        }

        // 処理モード
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

        // 消費税率
        public class 消費税率
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
        ///     ログインユーザーコンボボックスクラス </summary>
        ///------------------------------------------------------------
        public class comboLoginUser
        {
            public int ID { get; set; }
            public string Name { get; set; }

            ///-------------------------------------------------------------
            /// <summary>
            ///     ログインユーザーコンボボックスitemロード </summary>
            /// <param name="cmbObj">
            ///     コンボボックスオブジェクト</param>
            ///-------------------------------------------------------------
            public static void itemLoad(ComboBox cmbObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.ログインユーザーTableAdapter adp = new darwinDataSetTableAdapters.ログインユーザーTableAdapter();
                    adp.Fill(dts.ログインユーザー);

                    comboLoginUser[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.ログインユーザー)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboLoginUser();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.ログインユーザー;
                        iX++;
                    }

                    cmbObj.DataSource = sList;
                    cmbObj.DisplayMember = "Name";
                    cmbObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ログインユーザー コンボボックスロード");
                }
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダコンボボックスクラス </summary>
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
            ///     ログインタイプヘッダコンボボックスitemロード </summary>
            /// <param name="cmbObj">
            ///     コンボボックスオブジェクト</param>
            ///-------------------------------------------------------------
            public static void itemLoad(ComboBox cmbObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter adp = new darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter();
                    adp.Fill(dts.ログインタイプヘッダ);

                    comboLogintype[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.ログインタイプヘッダ)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboLogintype();
                        sList[iX].ID = t.Id;
                        sList[iX].Name = t.名称;

                        iX++;
                    }

                    cmbObj.DataSource = sList;
                    cmbObj.DisplayMember = "Name";
                    cmbObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ログインタイプヘッダ コンボボックスロード");
                }
            }


            ///-------------------------------------------------------------
            /// <summary>
            ///     ログインタイプヘッダコンボボックスitemロード </summary>
            /// <param name="cmbObj">
            ///     コンボボックスオブジェクト</param>
            ///-------------------------------------------------------------
            public static void itemLoad(CheckedListBox lstObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter adp = new darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter();
                    adp.Fill(dts.ログインタイプヘッダ);

                    comboLogintype[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.ログインタイプヘッダ)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboLogintype();
                        sList[iX].ID = t.Id;
                        sList[iX].Name = t.名称;
                        sList[iX].Lock = t.受注個別ロック権限;
                        sList[iX].seigen = t.受注個別制限;
                        sList[iX].Jyuryo = t.注文書受領済み権限;

                        iX++;
                    }

                    lstObj.DataSource = sList;
                    lstObj.DisplayMember = "Name";
                    lstObj.ValueMember = "ID";
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ログインタイプヘッダ チェックリストボックスロード");
                }
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     外注先マスターコンボボックスクラス </summary>
        ///------------------------------------------------------------
        public class comboGaichu
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public int shiharaibi { get; set; }     // 支払日 2018/01/04

            public static comboGaichuUser[] getArrayGaichu()
            {
                comboGaichuUser[] sList = null;

                darwinDataSet dts = new darwinDataSet();
                darwinDataSetTableAdapters.外注先TableAdapter adp = new darwinDataSetTableAdapters.外注先TableAdapter();
                adp.Fill(dts.外注先);

                int iX = 0;

                foreach (var t in dts.外注先)
                {
                    Array.Resize(ref sList, iX + 1);
                    sList[iX] = new comboGaichuUser();
                    sList[iX].ID = t.ID;
                    sList[iX].Name = t.名称;
                    sList[iX].shiharaibi = t.支払日;
                    iX++;
                }

                return sList;
            }


            ///-------------------------------------------------------------
            /// <summary>
            ///     外注先コンボボックスitemロード </summary>
            /// <param name="cmbObj">
            ///     コンボボックスオブジェクト</param>
            ///-------------------------------------------------------------
            public static void itemLoad(ComboBox cmbObj)
            {
                try 
	            {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.外注先TableAdapter adp = new darwinDataSetTableAdapters.外注先TableAdapter();
                    adp.Fill(dts.外注先);

                    comboGaichu[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.外注先)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new comboGaichu();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.名称;
                        sList[iX].shiharaibi = t.支払日;
                        iX++;
                    }

                    cmbObj.DataSource = sList;
                    cmbObj.DisplayMember = "Name";
                    cmbObj.ValueMember = "ID";
	            }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "外注先コンボボックスロード");
                }
            }
        }

        public class comboGaichuUser
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public int shiharaibi { get; set; }     // 支払日 2018/01/04
        }


        // 所属コンボボックスクラス
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

            // 所属マスターロード
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboShozoku cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.所属 Shozoku = new Control.所属();
                    dR = Shozoku.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboShozoku();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["所属名1"].ToString() + " " + dR["所属名2"].ToString();
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    Shozoku.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "所属コンボボックスロード");
                }

            }

            // 所属名コンボ表示
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


        /// 社員コンボボックスクラス
        ///     
        public class ComboShain
        {
            public int ID { get; set; }
            public string Name { get; set; }

            /// ------------------------------------------------------------------
            /// <summary>
            ///     社員マスターロード </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// ------------------------------------------------------------------
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.社員TableAdapter adp = new darwinDataSetTableAdapters.社員TableAdapter();
                    adp.Fill(dts.社員);

                    ComboShain[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.社員)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new ComboShain();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.氏名;
                        iX++;
                    }

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "社員コンボボックスロード");
                }
            }
             
            // 社員コンボ表示
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

        // 口座種別コンボボックスクラス
        public class ComboKouza
        {
            private int F_ID;
            private string F_Name;

            private const int kCode_F = 1;
            private const string kName_F = "普通";

            private const int kCode_T = 2;
            private const string kName_T = "当座";

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

            // 口座種別セット
            public static void load(ComboBox tempObj)
            {
                try
                {
                    ComboKouza cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    //普通
                    cmb1 = new ComboKouza();
                    cmb1.ID = kCode_F;
                    cmb1.Name = kName_F;
                    tempObj.Items.Add(cmb1);

                    //当座
                    cmb1 = new ComboKouza();
                    cmb1.ID = kCode_T;
                    cmb1.Name = kName_T;
                    tempObj.Items.Add(cmb1);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "口座種別セット");
                }

            }

            //口座種別コンボ表示
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

        // 事業所コンボボックスクラス : 2015/06/24
        public class ComboOffice
        {
            public int ID { get; set; }
            public string Name { get; set; }

            //事業所マスターロード : 2015/06/24
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.事業所TableAdapter adp = new darwinDataSetTableAdapters.事業所TableAdapter();
                    adp.Fill(dts.事業所);

                    ComboOffice[] fList = null;
                    int iX = 0;

                    foreach (var t in dts.事業所.OrderBy(a => a.ID))
                    {
                        Array.Resize(ref fList, iX + 1);
                        fList[iX] = new ComboOffice();
                        fList[iX].ID = t.ID;
                        fList[iX].Name = t.名称;
                        iX++;
                    }

                    tempObj.DataSource = fList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "事業所コンボボックスロード");
                }

            }

            //事業所コンボ表示
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

        // 得意先コンボボックスクラス : 2015/06/24
        public class ComboClient
        {
            private int F_ID;
            private string F_Name;
            private string F_NameShow;

            public int ID { get; set; }
            public string Name { get; set; }
            public string NameShow { get; set; }

            // 得意先マスターロード
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboClient cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.得意先 Client = new Control.得意先();
                    dR = Client.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboClient();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["ID"].ToString() + " ： " + dR["略称"].ToString() + "";
                        cmb1.NameShow = dR["名称"].ToString() + "";
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    Client.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "得意先コンボボックスロード");
                }

            }

            /// ------------------------------------------------------------------
            /// <summary>
            ///     コンボボックスに得意先マスターをロードする </summary>
            /// <param name="cmbObj">
            ///     コンボボックスオブジェクト</param>
            ///      : 2015/06/24
            /// ------------------------------------------------------------------
            public static void itemsLoad(ComboBox cmbObj)
            {
                ComboClient[] cList = null;
                int iX = 0;

                darwinDataSet dts = new darwinDataSet();
                darwinDataSetTableAdapters.得意先TableAdapter tAdp = new darwinDataSetTableAdapters.得意先TableAdapter();
                tAdp.Fill(dts.得意先);

                foreach (var t in dts.得意先)
                {
                    Array.Resize(ref cList, iX + 1);
                    cList[iX] = new ComboClient();
                    cList[iX].ID = t.ID;
                    cList[iX].Name = t.ID.ToString() + " ： " + t.略称.ToString();
                    cList[iX].NameShow = t.略称.ToString();
                    iX++;
                }

                cmbObj.DataSource = cList;
                cmbObj.DisplayMember = "Name";
                cmbObj.ValueMember = "ID";
            }

        }

        // 配布形態コンボボックスクラス : 2015/06/24
        public class ComboFkeitai
        {
            public int ID { get; set; }
            public string  Name { get; set; }

            // 配布形態マスターロード : 2015/06/24
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.配布形態TableAdapter adp = new darwinDataSetTableAdapters.配布形態TableAdapter();
                    adp.Fill(dts.配布形態);

                    ComboFkeitai[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.配布形態)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new ComboFkeitai();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.名称;
                        iX++;
                    }

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "配布形態コンボボックスロード");
                }
            }

            // 配布形態コンボ表示
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

        // 案件種別コンボボックスクラス : 2015/06/30
        public class ComboAnshu
        {
            public int ID { get; set; }
            public string Name { get; set; }

            // 案件種別アイテムセット : 2015/06/30
            public static void load(ComboBox tempObj)
            {
                try
                {
                    ComboAnshu[] sList = new ComboAnshu[3];
                    sList[0] = new ComboAnshu();
                    sList[0].ID = 1;
                    sList[0].Name = "自社ポス";

                    sList[1] = new ComboAnshu();
                    sList[1].ID = 2;
                    sList[1].Name = "他社ポス";

                    sList[2] = new ComboAnshu();
                    sList[2].ID = 3;
                    sList[2].Name = "ポス以外案件";

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "案件種別アイテムセット");
                }
            }
        }

        // 判型コンボボックスクラス : 2015/06/24
        public class ComboSize
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public decimal Tanka { get; set; }

            // 判型マスターロード : 2015/06/24
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.判型TableAdapter adp = new darwinDataSetTableAdapters.判型TableAdapter();
                    adp.Fill(dts.判型);

                    ComboSize[] sList = null;
                    int iX = 0;

                    foreach (var t in dts.判型)
                    {
                        Array.Resize(ref sList, iX + 1);
                        sList[iX] = new ComboSize();
                        sList[iX].ID = t.ID;
                        sList[iX].Name = t.名称;
                        sList[iX].Tanka = t.卸単価1;
                        iX++;
                    }

                    tempObj.DataSource = sList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "判型コンボボックスロード");
                }
            }

            // 判型コンボ表示
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

        // 締日パターンコンボボックスクラス
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

            // 締日パターンマスターロード
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboShimebi cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.締日パターン cPat = new Control.締日パターン();
                    dR = cPat.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboShimebi();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["摘要"].ToString() + "";
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    cPat.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "締日パターンコンボボックスロード");
                }

            }

            // 締日パターンコンボ表示
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

        // 振込口座コンボボックスクラス
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

            // 振込口座マスターロード
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboFuri cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.振込口座 cFuri = new Control.振込口座();
                    dR = cFuri.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboFuri();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["金融機関名"].ToString() + "　" + dR["支店名"].ToString() + "　";

                        switch (Int32.Parse(dR["口座種別"].ToString()))
                        {
                            case 1:
                                cmb1.Name += "普通" + "　";
                                break;
                            case 2:
                                cmb1.Name += "当座" + "　";
                                break;
                        }

                        cmb1.Name += dR["口座番号"].ToString();

                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    cFuri.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "振込口座コンボボックスロード");
                }

            }

            // 振込口座コンボ表示
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

            // 折マスターロード
            public static void load(DataGridViewComboBoxColumn tempObj)
            {
                try
                {
                    comboOri[] jList = new comboOri[2];

                    jList[0] = new comboOri();
                    jList[0].ID = 1;
                    jList[0].Name = "有";
                    jList[1] = new comboOri();
                    jList[1].ID = 2;
                    jList[1].Name = "無";

                    tempObj.DataSource = jList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "折コンボボックスロード");
                }

            }

        }

        // 受注種別コンボボックスクラス
        public class ComboJshubetsu
        {
            public int ID { get; set; }
            public string Name { get; set; }

            // 受注種別マスターロード
            public static void load(ComboBox tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.受注種別TableAdapter jAdp = new darwinDataSetTableAdapters.受注種別TableAdapter();
                    jAdp.Fill(dts.受注種別);

                    ComboJshubetsu[] jList = null;
                    int iX = 0;

                    foreach (var t in dts.受注種別)
                    {
                        Array.Resize(ref jList, iX + 1);
                        jList[iX] = new ComboJshubetsu();
                        jList[iX].ID = t.ID;
                        jList[iX].Name = t.名称;
                        iX++;
                    }

                    tempObj.DataSource = jList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "受注種別コンボボックスロード");
                }

            }

            // 受注種別マスターロード
            public static void load(DataGridViewComboBoxColumn tempObj)
            {
                try
                {
                    darwinDataSet dts = new darwinDataSet();
                    darwinDataSetTableAdapters.受注種別TableAdapter jAdp = new darwinDataSetTableAdapters.受注種別TableAdapter();
                    jAdp.Fill(dts.受注種別);

                    ComboJshubetsu[] jList = null;
                    int iX = 0;

                    foreach (var t in dts.受注種別)
                    {
                        Array.Resize(ref jList, iX + 1);
                        jList[iX] = new ComboJshubetsu();
                        jList[iX].ID = t.ID;
                        jList[iX].Name = t.名称;
                        iX++;
                    }

                    tempObj.DataSource = jList;
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "受注種別コンボボックスロード");
                }

            }

            // 受注種別コンボ表示
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

        // 配布員コンボボックスクラス
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

            // 配布員マスターロード
            public static void load(ComboBox tempObj)
            {
                try
                {
                    OleDbDataReader dR;
                    ComboStaff cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "Name";
                    tempObj.ValueMember = "ID";

                    Control.配布員 Staff = new Control.配布員();
                    dR = Staff.Fill();

                    while (dR.Read())
                    {
                        cmb1 = new ComboStaff();
                        cmb1.ID = Int32.Parse(dR["ID"].ToString());
                        cmb1.Name = dR["ID"].ToString() + " ： " +dR["氏名"].ToString();
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    Staff.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "配布員コンボボックスロード");
                }

            }

            // 配布員コンボ表示
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

        // 天候リストボックスクラス
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

            // 天候マスターロード
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
                    SqlSTR += "select distinct 天候 from 天候";
                    dR = cTenkou.free_dsReader(SqlSTR);

                    while (dR.Read())
                    {
                        cmb1 = new ComboTenkou();
                        cmb1.Name = dR["天候"].ToString() + "";
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();

                    cTenkou.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "天候リストボックスロード");
                }

            }


        }
        public static void FlgOnOff(int tempFLG)
        {
            string sqlSTR;
            sqlSTR = "";
            sqlSTR += "update 会社情報 ";
            sqlSTR += "set 配布フラグ = " + tempFLG.ToString() + ",";
            sqlSTR += "変更年月日 = '" + DateTime.Today + "'";

            Control.FreeSql fCon = new Control.FreeSql();
            fCon.Execute(sqlSTR);
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">Height</param>
        public static void WindowsMinSize(Form tempFrm,int wSize,int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize); 
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">height</param>
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }
        
        /// ----------------------------------------------------------------------------
        /// <summary>
        ///     メニュータイトル配列からメニューボタンのテキストをセットする </summary>
        /// <param name="btn">
        ///     ボタンオブジェクト </param>
        /// <param name="cm">
        ///     メニュータイトルクラス </param>
        /// ----------------------------------------------------------------------------
        public static void getMenuTittle(Button btn, clsMenu cm)
        {
            foreach (var t in cm.menuCsv)
            {
                // カンマ区切りで1行のデータ配列を取得
                string[] k = t.Split(',');

                // 配列の要素数が2つ以上あるか
                if (k.Length > 1)
                {
                    // Tagを照合してメニュータイトルをセットする
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
        ///     受注番号を採番する </summary>
        /// <param name="dt">
        ///     入庫日</param>
        /// <returns>
        ///     受注番号</returns>
        /// --------------------------------------------------------------------
        public static long getOrderNum(DateTime dt)
        {
            //IDを採番
            long gID = 0;
            DateTime sJDate = dt;

            // 伝票番号最小値
            Int64 denNumMin = ((long)sJDate.Year * 100000000) + (sJDate.Month * 1000000) + (sJDate.Day * 10000);
            denNumMin++;

            // 伝票番号最大値
            long denNumMax = denNumMin + 9998;

            // 受注日の伝票があるか？
            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.受注1TableAdapter rAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
            rAdp.Fill(dts.受注1);

            long s = 0;

            if (dts.受注1.Any(a => a.ID >= denNumMin && a.ID <= denNumMax))
            {
                s = dts.受注1.Where(a => a.ID >= denNumMin && a.ID <= denNumMax).Max(a => a.ID);
            }

            // 採番済みか？
            darwinDataSetTableAdapters.受注番号採番TableAdapter jAdp = new darwinDataSetTableAdapters.受注番号採番TableAdapter();
            jAdp.Fill(dts.受注番号採番);

            long j = 0;
            
            if (dts.受注番号採番.Any(a => a.受注番号 >= denNumMin && a.受注番号 <= denNumMax))
            {
                j = dts.受注番号採番.Where(a => a.受注番号 >= denNumMin && a.受注番号 <= denNumMax).Max(a => a.受注番号);
            }

            if (s == 0 && j == 0)   // 該当日の受注番号が未採番のとき
            {
                gID = denNumMin;
            }
            else if (s > j)    // 受注番号比較：大きい方に1加算したものを新受注番号とする
            {
                gID = s + 1;
            }
            else if (j > s)
            {
                gID = j + 1;
            }

            // 受注番号を返す
            return gID;
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     DataGridViewMaskedTextBoxCellオブジェクトの列を表します。 </summary>
        /// ------------------------------------------------------------------------------
        public class DataGridViewMaskedTextBoxColumn :
            DataGridViewColumn
        {
            //CellTemplateとするDataGridViewMaskedTextBoxCellオブジェクトを指定して
            //基本クラスのコンストラクタを呼び出す
            public DataGridViewMaskedTextBoxColumn()
                : base(new DataGridViewMaskedTextBoxCell())
            {
            }

            private string maskValue = "";
            /// <summary>
            /// MaskedTextBoxのMaskプロパティに適用する値
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

            //新しいプロパティを追加しているため、
            // Cloneメソッドをオーバーライドする必要がある
            public override object Clone()
            {
                DataGridViewMaskedTextBoxColumn col =
                    (DataGridViewMaskedTextBoxColumn)base.Clone();
                col.Mask = this.Mask;
                return col;
            }

            //CellTemplateの取得と設定
            public override DataGridViewCell CellTemplate
            {
                get
                {
                    return base.CellTemplate;
                }
                set
                {
                    //DataGridViewMaskedTextBoxCellしか
                    // CellTemplateに設定できないようにする
                    if (!(value is DataGridViewMaskedTextBoxCell))
                    {
                        throw new InvalidCastException(
                            "DataGridViewMaskedTextBoxCellオブジェクトを" +
                            "指定してください。");
                    }
                    base.CellTemplate = value;
                }
            }
        }
        
        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     MaskedTextBoxで編集できるテキスト情報を
        ///     DataGridViewコントロールに表示します。</summary>
        /// ------------------------------------------------------------------------------
        public class DataGridViewMaskedTextBoxCell :
            DataGridViewTextBoxCell
        {
            //コンストラクタ
            public DataGridViewMaskedTextBoxCell()
            {
            }

            //編集コントロールを初期化する
            //編集コントロールは別のセルや列でも使いまわされるため、初期化の必要がある
            public override void InitializeEditingControl(
                int rowIndex, object initialFormattedValue,
                DataGridViewCellStyle dataGridViewCellStyle)
            {
                base.InitializeEditingControl(rowIndex,
                    initialFormattedValue, dataGridViewCellStyle);

                //編集コントロールの取得
                DataGridViewMaskedTextBoxEditingControl maskedBox =
                    this.DataGridView.EditingControl as
                    DataGridViewMaskedTextBoxEditingControl;
                if (maskedBox != null)
                {
                    //Textを設定
                    string maskedText = initialFormattedValue as string;
                    maskedBox.Text = maskedText != null ? maskedText : "";
                    //カスタム列のプロパティを反映させる
                    DataGridViewMaskedTextBoxColumn column =
                        this.OwningColumn as DataGridViewMaskedTextBoxColumn;
                    if (column != null)
                    {
                        maskedBox.Mask = column.Mask;
                    }
                }
            }

            //編集コントロールの型を指定する
            public override Type EditType
            {
                get
                {
                    return typeof(DataGridViewMaskedTextBoxEditingControl);
                }
            }

            //セルの値のデータ型を指定する
            //ここでは、Object型とする
            //基本クラスと同じなので、オーバーライドの必要なし
            public override Type ValueType
            {
                get
                {
                    return typeof(object);
                }
            }

            //新しいレコード行のセルの既定値を指定する
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
        ///     DataGridViewMaskedTextBoxCellでホストされる
        ///     MaskedTextBoxコントロールを表します。</summary>
        /// ------------------------------------------------------------------------------
        public class DataGridViewMaskedTextBoxEditingControl :
            MaskedTextBox, IDataGridViewEditingControl
        {
            //編集コントロールが表示されているDataGridView
            DataGridView dataGridView;
            //編集コントロールが表示されている行
            int rowIndex;
            //編集コントロールの値とセルの値が違うかどうか
            bool valueChanged;

            //コンストラクタ
            public DataGridViewMaskedTextBoxEditingControl()
            {
                this.TabStop = false;
            }

            #region IDataGridViewEditingControl メンバ

            //編集コントロールで変更されたセルの値
            public object GetEditingControlFormattedValue(
                DataGridViewDataErrorContexts context)
            {
                return this.Text;
            }

            //編集コントロールで変更されたセルの値
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

            //セルスタイルを編集コントロールに適用する
            //編集コントロールの前景色、背景色、フォントなどをセルスタイルに合わせる
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

            //編集するセルがあるDataGridView
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

            //編集している行のインデックス
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

            //値が変更されたかどうか
            //編集コントロールの値とセルの値が違うかどうか
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
            ///     指定されたキーをDataGridViewが処理するか、編集コントロールが処理するか
            ///     Trueを返すと、編集コントロールが処理する
            ///     dataGridViewWantsInputKeyがTrueの時は、DataGridViewが処理できる </summary>
            /// <param name="keyData"></param>
            /// <param name="dataGridViewWantsInputKey"></param>
            /// <returns></returns>
            /// ------------------------------------------------------------------------------
            public bool EditingControlWantsInputKey(
                Keys keyData, bool dataGridViewWantsInputKey)
            {
                //Keys.Left、Right、Home、Endの時は、Trueを返す
                //このようにしないと、これらのキーで別のセルにフォーカスが移ってしまう
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

            //マウスカーソルがEditingPanel上にあるときのカーソルを指定する
            //EditingPanelは編集コントロールをホストするパネルで、
            //編集コントロールがセルより小さいとコントロール以外の部分がパネルとなる
            public Cursor EditingPanelCursor
            {
                get
                {
                    return base.Cursor;
                }
            }

            //コントロールで編集する準備をする
            //テキストを選択状態にしたり、挿入ポインタを末尾にしたりする
            public void PrepareEditingControlForEdit(bool selectAll)
            {
                if (selectAll)
                {
                    //選択状態にする
                    this.SelectAll();
                }
                else
                {
                    //挿入ポインタを末尾にする
                    this.SelectionStart = this.TextLength;
                }
            }

            //値が変更した時に、セルの位置を変更するかどうか
            //値が変更された時に編集コントロールの大きさが変更される時はTrue
            public bool RepositionEditingControlOnValueChange
            {
                get
                {
                    return false;
                }
            }

            #endregion

            //値が変更された時
            protected override void OnTextChanged(EventArgs e)
            {
                base.OnTextChanged(e);
                //値が変更されたことをDataGridViewに通知する
                this.valueChanged = true;
                this.dataGridView.NotifyCurrentCellDirty(true);
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     消費税率取得 </summary>
        /// <param name="tempDate">
        ///     基準日付</param>
        /// <returns>
        ///     税率</returns>
        ///  
        /// 2015/06/24
        ///------------------------------------------------------------------
        public static int GetTaxRT(DateTime tempDate)
        {
            //税率取得
            int Ritsu = 0;

            darwinDataSet dts = new posting.darwinDataSet();
            darwinDataSetTableAdapters.税率TableAdapter adp = new darwinDataSetTableAdapters.税率TableAdapter();
            adp.Fill(dts.税率);

            foreach (var t in dts.税率.Where(a => a.開始年月日 <= tempDate).OrderByDescending(a => a.開始年月日))
            {
                Ritsu = t.税率;
                break;
            }

            return Ritsu;
        }

        /// -------------------------------------------------------------
        /// <summary>
        ///     消費税計算 </summary>
        /// <param name="tempKin">
        ///     対象金額</param>
        /// <param name="tempTax">
        ///     税率</param>
        /// <returns>
        ///     消費税額</returns>
        /// -------------------------------------------------------------
        public static decimal GetTax(decimal tempKin, int tempTax)
        {
            decimal GakuD;
            //int GakuI;

            // 端数切捨て 2015/07/01
            GakuD = Math.Floor(tempKin * tempTax / 100);

            //GakuD += (decimal)0.5;
            //GakuI = (int)GakuD;

            return GakuD;
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     郵便番号CSVを配列に読み込む </summary>
        /// <param name="z">
        ///     配列</param>
        /// ----------------------------------------------------------------------
        public static void zipCsvLoad(ref string[] z)
        {
            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.会社情報TableAdapter adp = new darwinDataSetTableAdapters.会社情報TableAdapter();
            adp.Fill(dts.会社情報);

            string denHeadCsvPath = string.Empty;

            // 会社情報から郵便番号CSVパスを取得
            foreach (var t in dts.会社情報)
            {
                denHeadCsvPath = t.郵便番号CSVパス;
            }

            // 郵便番号CSVファイルが存在していなかったら終了
            if (!System.IO.File.Exists(denHeadCsvPath))
            {
                return;
            }

            // 郵便番号CSV読み込み
            z = System.IO.File.ReadAllLines(denHeadCsvPath, Encoding.Default);
        }
    }
}

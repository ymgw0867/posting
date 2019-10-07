using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows;

namespace posting
{
    class Control
    {
        /// <summary>
        /// DataControlクラスの基本クラス
        /// </summary>
        public class BaseControl
        {
            private Utility.DBConnect dbConnect;

            /// <summary>
            /// データベース接続メソッド
            /// </summary>
            /// <returns>データベース接続情報を取得します</returns>
            public OleDbConnection GetConnection()
            {
                //return dbConnect;
                return dbConnect.Cn;
            }

            /// <summary>
            /// BaseControlのコンストラクタ。DBConnectクラスのインスタンスを作成します。
            /// </summary>
            public BaseControl()
            {
                dbConnect = new Utility.DBConnect();
            }
        }

        /// <summary>
        /// データコントロールクラス BaseControlを継承します
        /// </summary>
        public class DataControl : BaseControl 
        {
            private Access.DataAccess dAccess;
            public OleDbConnection cn = new OleDbConnection();

            /// <summary>
            /// DataControlクラスのコンストラクタ。データアクセスクラスのインスタンスを作成します。
            /// </summary>
            public DataControl()
            {
                // データアクセスクラスのインスタンスを作成する
                dAccess = new Access.DataAccess();
            }

            /// <summary>
            /// データベースの接続を解除します
            /// </summary>
            public void Close()
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                }
            }

            /// <summary>
            /// データリーダー取得インターフェイスを引数としたメソッド
            /// </summary>
            /// <param name="IDR">データリーダーを取得するインターフェイス</param>
            /// <returns>引数で指定したマスターのデータリーダー</returns>
            public OleDbDataReader FillAccess(Access.DataAccess.IFill IDR)
            {
                // データベース接続情報を取得する
                cn = this.GetConnection();

                return IDR.GetdReader(cn);
                
            }

            /// <summary>
            /// 条件付きデータリーダを取得します
            /// </summary>
            /// <param name="tempString">SQL文を記述した文字列</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader free_dsReader(string tempString)
            {

                try
                {
                    return FillByAccess(new Access.DataAccess.free_dsReader(), tempString);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            /// <summary>
            /// 条件付きデータリーダー取得インターフェイスを引数としたメソッド
            /// </summary>
            /// <param name="IDSR">データリーダーを取得するインターフェイス</param>
            /// <param name="tempString">SQL文のwhere以下の条件を記述した文字列</param>
            /// <returns>条件式に一致する引数で指定されたマスターのデータリーダー</returns>
            public OleDbDataReader FillByAccess(Access.DataAccess.IFillBy IDSR,string tempString)
            {
                // データベース接続情報を取得する
                cn = this.GetConnection();

                return IDSR.GetdsReader(cn,tempString);

            }
        }

        /// <summary>
        /// フリーＳＱＬ実行
        /// </summary>
        public class FreeSql : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();

            public Boolean Execute(string tempSql)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    SCom.CommandText = tempSql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }
        }

        /// <summary>
        /// 会社情報クラス
        /// </summary>
        public class 会社情報 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 会社情報マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cKaisha">Entityクラスの会社情報</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.会社情報 cKaisha)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 会社情報 ";
                    mySql += "(会社名,代表者氏名,役職名,電話番号,FAX番号,住所1,住所2,";
                    mySql += "郵便番号,メールアドレス,部署名,担当者名,特記事項1,特記事項2,";
                    mySql += "依頼人コード,依頼人名,金融機関コード,金融機関名,支店コード,支店名,";
                    mySql += "口座種別,口座番号,配布フラグ,登録年月日,変更年月日,郵便番号CSVパス, 受注確定書入力シートパス) ";
                    mySql += "values ('" + cKaisha.会社名 + "',";
                    mySql += "'" + cKaisha.代表者氏名 + "',";
                    mySql += "'" + cKaisha.役職名 + "',";
                    mySql += "'" + cKaisha.電話番号 + "',";
                    mySql += "'" + cKaisha.FAX番号 + "',";
                    mySql += "'" + cKaisha.住所1 + "',";
                    mySql += "'" + cKaisha.住所2 + "',";
                    mySql += "'" + cKaisha.郵便番号 + "',";
                    mySql += "'" + cKaisha.メールアドレス + "',";
                    mySql += "'" + cKaisha.部署名 + "',";
                    mySql += "'" + cKaisha.担当者名 + "',";
                    mySql += "'" + cKaisha.特記事項1 + "',";
                    mySql += "'" + cKaisha.特記事項2 + "',";
                    mySql += "'" + cKaisha.依頼人コード + "',";
                    mySql += "'" + cKaisha.依頼人名 + "',";
                    mySql += "'" + cKaisha.金融機関コード + "',";
                    mySql += "'" + cKaisha.金融機関名 + "',";
                    mySql += "'" + cKaisha.支店コード + "',";
                    mySql += "'" + cKaisha.支店名 + "',";
                    mySql += cKaisha.口座種別 + ",";
                    mySql += "'" + cKaisha.口座番号 + "',";
                    mySql += cKaisha.配布フラグ + ",";
                    mySql += "'" + cKaisha.登録年月日 + "',";
                    mySql += "'" + cKaisha.変更年月日 + "',";
                    mySql += "'" + cKaisha.郵便番号CSVパス + "',";
                    mySql += "'" + cKaisha.受注確定書入力シートパス + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 会社情報更新
            /// </summary>
            /// <param name="cKaisha">会社情報Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.会社情報 cKaisha)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 会社情報 set ";
                    mySql += "会社名 = '" + cKaisha.会社名 + "',";
                    mySql += "代表者氏名 = '" + cKaisha.代表者氏名 + "',";
                    mySql += "役職名 = '" + cKaisha.役職名 + "',";
                    mySql += "電話番号 = '" + cKaisha.電話番号 + "',";
                    mySql += "FAX番号 = '" + cKaisha.FAX番号 + "',";
                    mySql += "住所1 = '" + cKaisha.住所1 + "',";
                    mySql += "住所2 = '" + cKaisha.住所2 + "',";
                    mySql += "郵便番号 = '" + cKaisha.郵便番号 + "',";
                    mySql += "メールアドレス = '" + cKaisha.メールアドレス + "',";
                    mySql += "部署名 = '" + cKaisha.部署名 + "',";
                    mySql += "担当者名 = '" + cKaisha.担当者名 + "',";
                    mySql += "特記事項1 = '" + cKaisha.特記事項1 + "',";
                    mySql += "特記事項2 = '" + cKaisha.特記事項2 + "',";
                    mySql += "依頼人コード = '" + cKaisha.依頼人コード + "',";
                    mySql += "依頼人名 = '" + cKaisha.依頼人名 + "',";
                    mySql += "金融機関コード = '" + cKaisha.金融機関コード + "',";
                    mySql += "金融機関名 = '" + cKaisha.金融機関名 + "',";
                    mySql += "支店コード = '" + cKaisha.支店コード + "',";
                    mySql += "支店名 = '" + cKaisha.支店名 + "',";
                    mySql += "口座種別 = " + cKaisha.口座種別 + ",";
                    mySql += "口座番号 = '" + cKaisha.口座番号 + "',";
                    mySql += "配布フラグ = " + cKaisha.配布フラグ + ",";
                    mySql += "変更年月日 = '" + DateTime.Today + "',";
                    mySql += "郵便番号CSVパス = '" + cKaisha.郵便番号CSVパス + "',";
                    mySql += "受注確定書入力シートパス = '" + cKaisha.受注確定書入力シートパス + "' ";
                    mySql += "where ID = " + cKaisha.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 会社情報マスターレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 会社情報 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 会社情報マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>会社情報マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 会社情報マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill会社情報());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 会社情報マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 会員情報マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy会社情報(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// 事業所クラス
        /// </summary>
        public class 事業所 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 事業所マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cJigyosho">Entityクラスの言語</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.事業所 cJigyosho)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 事業所 ";
                    mySql += "(ID,名称,郵便番号,住所1,住所2,電話番号,FAX番号,";
                    mySql += "備考,登録年月日,変更年月日) ";
                    mySql += "values (" + cJigyosho.ID + ",";
                    mySql += "'" + cJigyosho.名称 + "',";
                    mySql += "'" + cJigyosho.郵便番号 + "',";
                    mySql += "'" + cJigyosho.住所1 + "',";
                    mySql += "'" + cJigyosho.住所2 + "',";
                    mySql += "'" + cJigyosho.電話番号 + "',";
                    mySql += "'" + cJigyosho.FAX番号 + "',";
                    mySql += "'" + cJigyosho.備考 + "',";
                    mySql += "'" + cJigyosho.登録年月日 + "',";
                    mySql += "'" + cJigyosho.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 事業所マスター更新
            /// </summary>
            /// <param name="cJigyosho">事業所マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.事業所 cJigyosho)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 事業所 set ";
                    mySql += "名称 = '" + cJigyosho.名称 + "',";

                    mySql += "郵便番号 = '" + cJigyosho.郵便番号 + "',";
                    mySql += "住所1 = '" + cJigyosho.住所1 + "',";
                    mySql += "住所2 = '" + cJigyosho.住所2 + "',";
                    mySql += "電話番号 = '" + cJigyosho.電話番号 + "',";
                    mySql += "FAX番号 = '" + cJigyosho.FAX番号 + "',";
                    mySql += "備考 = '" + cJigyosho.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cJigyosho.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 事業所マスターレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 事業所 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 事業所マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>事業所マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 事業所マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill事業所());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 事業所マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 事業所マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy事業所(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 受注クラス
        /// </summary>
        public class 受注 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            ///--------------------------------------------------------------------
            /// <summary>
            ///     受注データに新規にデータを登録する </summary>
            /// <param name="cJyuchu">
            ///     Entityクラスの受注</param>
            /// <returns>
            ///     </returns>
            ///--------------------------------------------------------------------
            public Boolean DataInsert(Entity.受注 cJyuchu)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 受注 ";
                    mySql += "(ID,事業所ID,受注日,受注区分,得意先ID,社員ID,チラシ名,受注種別ID,";
                    mySql += "単価,枚数,金額,消費税,税込金額,値引額,売上金額,税率,判型,配布単価,依頼先,原価,配布形態,";
                    mySql += "配布条件,配布開始日,配布終了日,納品予定日,請求書,";
                    mySql += "請求書ID,請求書発行日,入金方法,入金予定日,";
                    mySql += "振込口座ID,未配布情報有無,枝番有無,特記事項,エリア備考,";
                    mySql += "完了区分,併配除外,登録年月日,変更年月日,外注先ID営業,外注支払日営業,外注原価営業,外注依頼日営業,";
                    mySql += "外注先ID支払,外注支払日支払,外注原価支払,外注依頼日支払,ユーザーID,案件種別,";
                    mySql += "配布猶予,納品形態,報告時期,報告精度,報告方法,メールアドレス,登録ユーザーID, 外注渡し日,";
                    mySql += "外注受け渡し担当者,外注委託枚数,業種,";
                    mySql += "外注先ID支払2,外注支払日支払2,外注原価支払2,外注先ID支払3,外注支払日支払3,外注原価支払3,";
                    mySql += "外注依頼日支払2,外注依頼日支払3,外注委託枚数2,外注委託枚数3,外注渡し日2,外注渡し日3,外注受け渡し担当者2,外注受け渡し担当者3,";
                    mySql += "営業備考, 編集ロック, 注文書受領済み) ";
                    mySql += "values (";
                    mySql += cJyuchu.ID + ",";
                    mySql += cJyuchu.事業所ID + ",";
                    mySql += "'" + cJyuchu.受注日 + "',";
                    mySql += "'" + cJyuchu.受注区分 + "',";
                    mySql += cJyuchu.得意先ID + ",";
                    mySql += cJyuchu.社員ID + ",";
                    mySql += "'" + cJyuchu.チラシ名 + "',";
                    mySql += cJyuchu.受注種別ID + ",";
                    mySql += cJyuchu.単価 + ",";
                    mySql += cJyuchu.枚数 + ",";
                    mySql += cJyuchu.金額 + ",";
                    mySql += cJyuchu.消費税 + ",";
                    mySql += cJyuchu.税込金額 + ",";
                    mySql += cJyuchu.値引額 + ",";
                    mySql += cJyuchu.売上金額 + ",";
                    mySql += cJyuchu.税率 + ",";
                    mySql += cJyuchu.判型 + ",";
                    mySql += cJyuchu.配布単価 + ",";
                    mySql += "'" + cJyuchu.依頼先 + "',";
                    mySql += cJyuchu.原価 + ",";
                    mySql += cJyuchu.配布形態 + ",";
                    mySql += "'" + cJyuchu.配布条件 + "',";

                    if (cJyuchu.配布開始日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.配布開始日 + "',";
                    }

                    if (cJyuchu.配布終了日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.配布終了日 + "',";
                    }
                    
                    if (cJyuchu.納品予定日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.納品予定日 + "',";
                    }
                    
                    mySql += cJyuchu.請求書 + ",";
                    mySql += cJyuchu.請求書ID + ",";

                    //mySql += "Null,";   //請求書発行日

                    // 2015/07/01
                    if (cJyuchu.請求書発行日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.請求書発行日 + "',";
                    }
                    
                    mySql += "'" + cJyuchu.入金方法 + "',";

                    if (cJyuchu.入金予定日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.入金予定日 + "',";
                    }
                    
                    mySql += cJyuchu.振込口座ID + ",";
                    mySql += cJyuchu.未配布情報有無 + ",";
                    mySql += cJyuchu.枝番有無 + ",";
                    mySql += "'" + cJyuchu.特記事項 + "',";
                    mySql += "'" + cJyuchu.エリア備考 + "',";
                    mySql += cJyuchu.完了区分 + ",";
                    mySql += cJyuchu.併配除外 + ",";    // 2014/11/26
                    mySql += "'" + cJyuchu.登録年月日 + "',";
                    mySql += "'" + cJyuchu.変更年月日 + "',";
                    mySql += cJyuchu.外注先ID営業.ToString() + ",";  // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.外注先支払日営業 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先支払日営業 + "',";
                    }

                    mySql += cJyuchu.外注先原価営業 + ","; // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.外注先依頼日営業 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先依頼日営業 + "',";
                    }

                    mySql += cJyuchu.外注先ID支払.ToString() + ",";  // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.外注先支払日支払 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先支払日支払 + "',";
                    }

                    mySql += cJyuchu.外注先原価支払 + ",";     // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.外注先依頼日支払 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先依頼日支払 + "',";
                    }

                    mySql += "'" + cJyuchu.ユーザーID + "',";   // 2015/06/30  2015/07/10
                    mySql += cJyuchu.案件種別 + ",";  // 2015/06/30
                    mySql += "'" + cJyuchu.配布猶予 + "',";   // 2015/07/01
                    mySql += "'" + cJyuchu.納品形態 + "',";   // 2015/07/01
                    mySql += "'" + cJyuchu.報告時期 + "',";     // 2015/07/01
                    mySql += "'" + cJyuchu.報告精度 + "',";     // 2015/07/01
                    mySql += "'" + cJyuchu.報告方法 + "',";     // 2015/07/01
                    mySql += "'" + cJyuchu.メールアドレス + "',";  // 2015/07/01
                    mySql += "'" + cJyuchu.ユーザーID + "',";       // 2015/08/10

                    // 2015/09/02
                    if (cJyuchu.外注渡し日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注渡し日 + "',";
                    }

                    mySql += "'" + cJyuchu.外注受け渡し担当者 + "',";  // 2015/08/11
                    mySql += cJyuchu.外注委託枚数 + ",";              // 2015/09/20
                    mySql += "'" + cJyuchu.業種 + "',";               // 2015/09/20
                    
                    mySql += cJyuchu.外注先ID支払2.ToString() + ",";  // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.外注先支払日支払2 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先支払日支払2 + "',";
                    }

                    mySql += cJyuchu.外注先原価支払2 + ",";     // 2016/10/14

                    mySql += cJyuchu.外注先ID支払3.ToString() + ",";  // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.外注先支払日支払3 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先支払日支払3 + "',";
                    }

                    mySql += cJyuchu.外注先原価支払3 + ",";     // 2016/10/14

                    // 2016/10/15
                    if (cJyuchu.外注先依頼日支払2 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先依頼日支払2 + "',";
                    }

                    if (cJyuchu.外注先依頼日支払3 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注先依頼日支払3 + "',";
                    }

                    mySql += cJyuchu.外注委託枚数2 + ",";
                    mySql += cJyuchu.外注委託枚数3 + ",";

                    if (cJyuchu.外注渡し日2 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注渡し日2 + "',";
                    }

                    if (cJyuchu.外注渡し日3 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.外注渡し日3 + "',";
                    }

                    mySql += "'" + cJyuchu.外注受け渡し担当者2 + "',";
                    mySql += "'" + cJyuchu.外注受け渡し担当者3 + "',";
                    mySql += "'" + cJyuchu.営業備考 + "',";       // 2019/03/01
                    mySql += cJyuchu.編集ロック + ",";            // 2019/10/05
                    mySql += cJyuchu.注文書受領済み + ")";         // 2019/10/05

                    //System.Windows.Forms.MessageBox.Show(mySql);

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }
                catch(Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message, "登録エラー", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            ///-------------------------------------------------------------
            /// <summary>
            ///     受注データ更新 </summary>
            /// <param name="cJyuchu">
            ///     受注データEntityクラス</param>
            /// <returns>
            ///     成功；true、失敗：false</returns>
            ///--------------------------------------------------------------
            public Boolean DataUpdate(Entity.受注 cJyuchu)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 受注 set ";
                    mySql += "事業所ID = " + cJyuchu.事業所ID + ",";
                    mySql += "受注日 = '" + cJyuchu.受注日 + "',";
                    mySql += "受注区分 = '" + cJyuchu.受注区分 + "',";
                    mySql += "得意先ID = " + cJyuchu.得意先ID + ",";
                    mySql += "社員ID = " + cJyuchu.社員ID + ",";
                    mySql += "チラシ名 = '" + cJyuchu.チラシ名 + "',";
                    mySql += "受注種別ID = " + cJyuchu.受注種別ID + ",";
                    mySql += "単価 = " + cJyuchu.単価 + ",";
                    mySql += "枚数 = " + cJyuchu.枚数 + ",";
                    mySql += "金額 = " + cJyuchu.金額 + ",";
                    mySql += "消費税 = " + cJyuchu.消費税 + ",";
                    mySql += "税込金額 = " + cJyuchu.税込金額 + ",";
                    mySql += "値引額 = " + cJyuchu.値引額 + ",";
                    mySql += "売上金額 = " + cJyuchu.売上金額 + ",";
                    mySql += "税率 = " + cJyuchu.税率 + ",";
                    mySql += "判型 = " + cJyuchu.判型 + ",";
                    mySql += "配布単価 = " + cJyuchu.配布単価 + ",";
                    mySql += "依頼先 = '" + cJyuchu.依頼先 + "',";
                    mySql += "原価 = " + cJyuchu.原価 + ",";
                    mySql += "配布形態 = " + cJyuchu.配布形態 + ",";
                    mySql += "配布条件 = '" + cJyuchu.配布条件 + "',";

                    if (cJyuchu.配布開始日 == "")
                    {
                        mySql += "配布開始日 = Null,";
                    }
                    else
                    {
                        mySql += "配布開始日 = '" + cJyuchu.配布開始日 + "',";
                    }

                    if (cJyuchu.配布終了日 == "")
                    {
                        mySql += "配布終了日 = Null,";
                    }
                    else
                    {
                        mySql += "配布終了日 = '" + cJyuchu.配布終了日 + "',";
                    }

                    //mySql += "配布猶予 = '" + cJyuchu.配布猶予 + "',";    // 2015/06/23

                    if (cJyuchu.納品予定日 == "")
                    {
                        mySql += "納品予定日 = Null,";
                    }
                    else
                    {
                        mySql += "納品予定日 = '" + cJyuchu.納品予定日 + "',";
                    }

                    //mySql += "納品形態 = '" + cJyuchu.納品形態 + "',";    // 2015/06/23

                    mySql += "請求書 = " + cJyuchu.請求書 + ",";
                    mySql += "請求書ID = " + cJyuchu.請求書ID + ",";

                    if (cJyuchu.請求書発行日 == "")
                    {
                        mySql += "請求書発行日 = Null,";
                    }
                    else
                    {
                        mySql += "請求書発行日 = '" + cJyuchu.請求書発行日 + "',";
                    }

                    mySql += "入金方法 = '" + cJyuchu.入金方法 + "',";

                    if (cJyuchu.入金予定日 == "")
                    {
                        mySql += "入金予定日 = Null,";
                    }
                    else
                    {
                        mySql += "入金予定日 = '" + cJyuchu.入金予定日 + "',";
                    }

                    //mySql += "報告時期 = '" + cJyuchu.報告時期 + "',";                // 2015/06/23
                    //mySql += "報告精度 = '" + cJyuchu.報告精度 + "',";                // 2015/06/23
                    //mySql += "報告方法 = '" + cJyuchu.報告方法 + "',";                // 2015/06/23
                    //mySql += "メールアドレス = '" + cJyuchu.メールアドレス + "',";     // 2015/06/23

                    mySql += "振込口座ID = " + cJyuchu.振込口座ID + ",";
                    mySql += "未配布情報有無 = " + cJyuchu.未配布情報有無 + ",";
                    mySql += "枝番有無 = " + cJyuchu.枝番有無 + ",";
                    mySql += "特記事項 = '" + cJyuchu.特記事項 + "',";
                    mySql += "エリア備考 = '" + cJyuchu.エリア備考 + "',";
                    mySql += "完了区分 = " + cJyuchu.完了区分 + ",";
                    mySql += "併配除外 = " + cJyuchu.併配除外 + ",";
                    mySql += "変更年月日 = '" + DateTime.Now + "',";

                    mySql += "外注先ID営業 = " + cJyuchu.外注先ID営業 + ",";          // 2015/06/30

                    // 2015/07/17
                    if (cJyuchu.外注先支払日営業 == "")
                    {
                        mySql += "外注支払日営業 = Null,";
                    }
                    else
                    {
                        mySql += "外注支払日営業 = '" + cJyuchu.外注先支払日営業 + "',";
                    }

                    mySql += "外注原価営業 = " + cJyuchu.外注先原価営業 + ",";         // 2015/06/30

                    // 以下、外注先依頼日営業使用しないのでコメント化 2015/08/11
                    //// 2015/07/17
                    //if (cJyuchu.外注先依頼日営業 == "")
                    //{
                    //    mySql += "外注依頼日営業 = Null,";
                    //}
                    //else
                    //{
                    //    mySql += "外注依頼日営業 = '" + cJyuchu.外注先依頼日営業 + "',";
                    //}
                    
                    mySql += "外注先ID支払 = " + cJyuchu.外注先ID支払 + ",";          // 2015/06/30

                    // 2015/07/17
                    if (cJyuchu.外注先支払日支払 == "")
                    {
                        mySql += "外注支払日支払 = Null,";
                    }
                    else
                    {
                        mySql += "外注支払日支払 = '" + cJyuchu.外注先支払日支払 + "',";
                    }

                    mySql += "外注原価支払 = " + cJyuchu.外注先原価支払 + ",";         // 2015/06/30

                    // 2015/07/17
                    if (cJyuchu.外注先依頼日支払 == "")
                    {
                        mySql += "外注依頼日支払 = Null,";
                    }
                    else
                    {
                        mySql += "外注依頼日支払 = '" + cJyuchu.外注先依頼日支払 + "',";
                    }

                    mySql += "ユーザーID = '" + cJyuchu.ユーザーID + "',";                  // 2015/06/30 2015/07/10
                    mySql += "案件種別 = " + cJyuchu.案件種別 + ",";                        // 2015/06/30

                    // 2015/08/11
                    if (cJyuchu.外注渡し日 == "")
                    {
                        mySql += "外注渡し日 = Null,";
                    }
                    else
                    {
                        mySql += "外注渡し日 = '" + cJyuchu.外注渡し日 + "',";
                    }

                    mySql += "外注受け渡し担当者 = '" + cJyuchu.外注受け渡し担当者 + "',";    // 2015/08/11
                    mySql += "外注委託枚数 = " + cJyuchu.外注委託枚数 + ",";    // 2015/09/20
                    mySql += "業種 = '" + cJyuchu.業種 + "',";                 // 2015/09/20


                    mySql += "外注先ID支払2 = " + cJyuchu.外注先ID支払2 + ",";          // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.外注先支払日支払2 == "")
                    {
                        mySql += "外注支払日支払2 = Null,";
                    }
                    else
                    {
                        mySql += "外注支払日支払2 = '" + cJyuchu.外注先支払日支払2 + "',";
                    }

                    mySql += "外注原価支払2 = " + cJyuchu.外注先原価支払2 + ",";         // 2016/10/14


                    mySql += "外注先ID支払3 = " + cJyuchu.外注先ID支払3 + ",";          // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.外注先支払日支払3 == "")
                    {
                        mySql += "外注支払日支払3 = Null,";
                    }
                    else
                    {
                        mySql += "外注支払日支払3 = '" + cJyuchu.外注先支払日支払3 + "',";
                    }

                    mySql += "外注原価支払3 = " + cJyuchu.外注先原価支払3 + ",";         // 2016/10/14

                    // 2016/10/15
                    if (cJyuchu.外注先依頼日支払2 == "")
                    {
                        mySql += "外注依頼日支払2 = Null,";
                    }
                    else
                    {
                        mySql += "外注依頼日支払2 = '" + cJyuchu.外注先依頼日支払2 + "',";
                    }

                    if (cJyuchu.外注先依頼日支払3 == "")
                    {
                        mySql += "外注依頼日支払3 = Null,";
                    }
                    else
                    {
                        mySql += "外注依頼日支払3 = '" + cJyuchu.外注先依頼日支払3 + "',";
                    }

                    mySql += "外注委託枚数2 = " + cJyuchu.外注委託枚数2 + ",";
                    mySql += "外注委託枚数3 = " + cJyuchu.外注委託枚数3 + ",";

                    if (cJyuchu.外注渡し日2 == "")
                    {
                        mySql += "外注渡し日2 = Null,";
                    }
                    else
                    {
                        mySql += "外注渡し日2 = '" + cJyuchu.外注渡し日2 + "',";
                    }

                    if (cJyuchu.外注渡し日3 == "")
                    {
                        mySql += "外注渡し日3 = Null,";
                    }
                    else
                    {
                        mySql += "外注渡し日3 = '" + cJyuchu.外注渡し日3 + "',";
                    }

                    mySql += "外注受け渡し担当者2 = '" + cJyuchu.外注受け渡し担当者2 + "',";
                    mySql += "外注受け渡し担当者3 = '" + cJyuchu.外注受け渡し担当者3 + "',";
                    mySql += "営業備考 = '" + cJyuchu.営業備考 + "',";
                    mySql += "編集ロック = " + cJyuchu.編集ロック + ",";
                    mySql += "注文書受領済み = " + cJyuchu.注文書受領済み + " ";

                    mySql += "where ID = " + cJyuchu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch(Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message, "更新エラー", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            ///-----------------------------------------------------------
            /// <summary>
            ///     受注データレコード削除 </summary>
            /// <param name="tempCode">
            ///     ID</param>
            /// <returns>
            ///     成功：true、失敗：false</returns>
            ///-----------------------------------------------------------
            public Boolean DataDelete(long tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 受注 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 受注データ全件のデータリーダを取得します
            /// </summary>
            /// <returns>受注データ全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 受注データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill受注());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 受注データ条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 受注データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy受注(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 社員クラス
        /// </summary>
        public class 社員 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 社員マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cShain">Entityクラスの社員</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.社員 cShain)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 社員 ";
                    mySql += "(ID,氏名,フリガナ,所属コード,役職,入社年月日,";
                    mySql += "備考,登録年月日,変更年月日) ";
                    mySql += "values (" + cShain.ID + ",";
                    mySql += "'" + cShain.氏名 + "',";
                    mySql += "'" + cShain.フリガナ + "',";
                    mySql += cShain.所属コード + ",";
                    mySql += "'" + cShain.役職 + "',";

                    if (cShain.入社年月日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cShain.入社年月日 + "',";
                    }

                    mySql += "'" + cShain.備考 + "',";
                    mySql += "'" + cShain.登録年月日 + "',";
                    mySql += "'" + cShain.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 社員マスター更新
            /// </summary>
            /// <param name="cShain">社員マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.社員 cShain)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 社員 set ";
                    mySql += "氏名 = '" + cShain.氏名 + "',";
                    mySql += "フリガナ = '" + cShain.フリガナ + "',";
                    mySql += "所属コード = " + cShain.所属コード + ",";
                    mySql += "役職 = '" + cShain.役職 + "',";

                    if (cShain.入社年月日 == "")
                    {
                        mySql += "入社年月日 = Null,";
                    }
                    else
                    {
                        mySql += "入社年月日 = '" + cShain.入社年月日 + "',";
                    }

                    mySql += "備考 = '" + cShain.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShain.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 社員マスターレコード削除
            /// </summary>
            /// <param name="tempCode">社員ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 社員 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 社員マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>社員マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 社員マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill社員());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 社員マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 社員マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy社員(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 受注種別クラス
        /// </summary>
        public class 受注種別 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 受注種別マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cJtype">Entityクラスの受注種別</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.受注種別 cJtype)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 受注種別 ";
                    mySql += "(名称,備考,登録年月日,変更年月日) ";
                    mySql += "values ('" + cJtype.名称 + "',";
                    mySql += "'" + cJtype.備考 + "',";
                    mySql += "'" + cJtype.登録年月日 + "',";
                    mySql += "'" + cJtype.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 受注種別マスター更新
            /// </summary>
            /// <param name="cJtype">受注種別マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.受注種別 cJtype)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 受注種別 set ";
                    mySql += "名称 = '" + cJtype.名称 + "',";
                    mySql += "備考 = '" + cJtype.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cJtype.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 受注種別マスターレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 受注種別 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 受注種別マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>受注種別マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 受注種別マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill受注種別());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 受注種別マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 受注種別マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy受注種別(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 所属クラス
        /// </summary>
        public class 所属 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 所属マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cShozoku">Entityクラスの所属</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.所属 cShozoku)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 所属 ";
                    mySql += "(ID,所属名1,所属名2,備考,登録年月日,変更年月日) ";
                    mySql += "values (" + cShozoku.ID + ",";
                    mySql += "'" + cShozoku.所属名1 + "',";
                    mySql += "'" + cShozoku.所属名2 + "',";
                    mySql += "'" + cShozoku.備考 + "',";
                    mySql += "'" + cShozoku.登録年月日 + "',";
                    mySql += "'" + cShozoku.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 所属マスター更新
            /// </summary>
            /// <param name="cShozoku">所属マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.所属 cShozoku)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 所属 set ";
                    mySql += "所属名1 = '" + cShozoku.所属名1 + "',";
                    mySql += "所属名2 = '" + cShozoku.所属名2 + "',";
                    mySql += "備考 = '" + cShozoku.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShozoku.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 所属マスターレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 所属 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 所属マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>所属マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 所属マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill所属());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 所属マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 所属マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy所属(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 振込口座クラス
        /// </summary>
        public class 振込口座 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 振込口座マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cFurikomi">Entityクラスの振込口座</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.振込口座 cFurikomi)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 振込口座 ";
                    mySql += "(金融機関名,支店名,口座種別,口座番号,口座名義,登録年月日,変更年月日) ";
                    mySql += "values ('" + cFurikomi.金融機関名 + "',";
                    mySql += "'" + cFurikomi.支店名 + "',";
                    mySql += cFurikomi.口座種別 + ",";
                    mySql += "'" + cFurikomi.口座番号 + "',";
                    mySql += "'" + cFurikomi.口座名義 + "',";
                    mySql += "'" + cFurikomi.登録年月日 + "',";
                    mySql += "'" + cFurikomi.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 振込口座マスター更新
            /// </summary>
            /// <param name="cFurikomi">振込口座マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.振込口座 cFurikomi)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 振込口座 set ";
                    mySql += "金融機関名 = '" + cFurikomi.金融機関名 + "',";
                    mySql += "支店名 = '" + cFurikomi.支店名 + "',";
                    mySql += "口座種別 = " + cFurikomi.口座種別 + ",";
                    mySql += "口座番号 = '" + cFurikomi.口座番号 + "',";
                    mySql += "口座名義 = '" + cFurikomi.口座名義 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cFurikomi.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 振込口座マスターレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 振込口座 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 振込口座マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>振込口座マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 振込口座マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill振込口座());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 振込口座マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 振込口座マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy振込口座(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 請求書クラス
        /// </summary>
        public class 請求書 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 請求書データに新規にデータを登録する
            /// </summary>
            /// <param name="cSeikyu">Entityクラスの請求書</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.請求書 cSeikyu)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 請求書 ";
                    mySql += "(ID,得意先ID,請求金額,消費税,値引額,売上金額,税率,入金予定日,発行日,";
                    mySql += "入金残,完了区分,振込口座ID1,振込口座ID2,登録年月日,変更年月日) ";
                    mySql += "values (";
                    mySql += cSeikyu.ID + ",";
                    mySql += cSeikyu.得意先ID + ",";
                    mySql += cSeikyu.請求金額 + ",";
                    mySql += cSeikyu.消費税 + ",";
                    mySql += cSeikyu.値引額 + ",";
                    mySql += cSeikyu.売上金額 + ",";
                    mySql += cSeikyu.税率 + ",";
                    mySql += "'" + cSeikyu.入金予定日 + "',";
                    mySql += "'" + cSeikyu.発行日 + "',";
                    mySql += cSeikyu.入金残 + ",";
                    mySql += cSeikyu.完了区分 + ",";
                    mySql += cSeikyu.振込口座ID1 + ",";
                    mySql += cSeikyu.振込口座ID2 + ",";
                    mySql += "'" + cSeikyu.登録年月日 + "',";
                    mySql += "'" + cSeikyu.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 請求書データ更新
            /// </summary>
            /// <param name="cSeikyu">請求書Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.請求書 cSeikyu)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 請求書 set ";
                    mySql += "得意先ID = " + cSeikyu.得意先ID + ",";
                    mySql += "請求金額 = " + cSeikyu.請求金額 + ",";
                    mySql += "消費税 = " + cSeikyu.消費税 + ",";
                    mySql += "値引額 = " + cSeikyu.値引額 + ",";
                    mySql += "売上金額 = " + cSeikyu.売上金額 + ",";
                    mySql += "税率 = " + cSeikyu.税率 + ",";
                    mySql += "入金予定日 = '" + cSeikyu.入金予定日 + "',";
                    mySql += "発行日 = '" + cSeikyu.発行日 + "',";
                    mySql += "入金残 = " + cSeikyu.入金残 + ",";
                    mySql += "完了区分 = " + cSeikyu.完了区分 + ",";
                    mySql += "振込口座ID1 = " + cSeikyu.振込口座ID1 + ",";
                    mySql += "振込口座ID2 = " + cSeikyu.振込口座ID2 + ",";
                    mySql += "備考 = '" + cSeikyu.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cSeikyu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 請求書レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 請求書 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 請求書データ全件のデータリーダを取得します
            /// </summary>
            /// <returns>請求書データ全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 請求書データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill請求書());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 請求書データ条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 請求書データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy請求書(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 税率クラス
        /// </summary>
        public class 税率 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 税率データに新規にデータを登録する
            /// </summary>
            /// <param name="cTax">Entityクラスの税率</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.税率 cTax)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 税率 ";
                    mySql += "(税率,開始年月日,備考,登録年月日,変更年月日) ";
                    mySql += "values (" + cTax.設定税率 + ",";
                    mySql += "'" + cTax.開始年月日 + "',";
                    mySql += "'" + cTax.備考 + "',";
                    mySql += "'" + cTax.登録年月日 + "',";
                    mySql += "'" + cTax.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 税率マスター更新
            /// </summary>
            /// <param name="cTax">税率Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.税率 cTax)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 税率 set ";
                    mySql += "税率 = " + cTax.設定税率 + ",";
                    mySql += "開始年月日 = '" + cTax.開始年月日 + "',";
                    mySql += "備考 = '" + cTax.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cTax.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 税率レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 税率 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 税率マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>税率マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 税率マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill税率());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 税率マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 税率マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy税率(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 町名クラス
        /// </summary>
        public class 町名 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 町名マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cTown">Entityクラスの町名</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.町名 cTown)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 町名 ";
                    mySql += "(ID,名称,市区町村コード,備考,登録年月日,変更年月日) ";
                    mySql += "values (" + cTown.ID + ",";
                    mySql += "'" + cTown.名称 + "',";
                    mySql += cTown.市区町村コード + ",";
                    mySql += "'" + cTown.備考 + "',";
                    mySql += "'" + cTown.登録年月日 + "',";
                    mySql += "'" + cTown.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 町名マスター更新
            /// </summary>
            /// <param name="cTown">町名Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.町名 cTown)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 町名 set ";
                    mySql += "名称 ='" + cTown.名称 + "',";
                    mySql += "市区町村コード = " + cTown.市区町村コード + ",";
                    mySql += "備考 ='" + cTown.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cTown.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 町名レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 町名 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 町名マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>町名マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 町名マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill町名());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 町名マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 町名マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy町名(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 締日パターンクラス
        /// </summary>
        public class 締日パターン : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 締日パターンマスターに新規にデータを登録する
            /// </summary>
            /// <param name="cShime">Entityクラスの締日パターン</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.締日パターン cShime)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 締日パターン ";
                    mySql += "(摘要,備考,登録年月日,変更年月日) ";
                    mySql += "values ('" + cShime.摘要 + "',";
                    mySql += "'" + cShime.備考 + "',";
                    mySql += "'" + cShime.登録年月日 + "',";
                    mySql += "'" + cShime.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 締日パターンマスター更新
            /// </summary>
            /// <param name="cShime">締日パターンEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.締日パターン cShime)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 締日パターン set ";
                    mySql += "摘要 ='" + cShime.摘要 + "',";
                    mySql += "備考 ='" + cShime.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShime.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 締日パターンレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 締日パターン ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 締日パターンマスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>締日パターンマスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 締日パターンマスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill締日パターン());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 締日パターンマスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 締日パターンマスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy締日パターン(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 得意先クラス
        /// </summary>
        public class 得意先 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 得意先マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cClient">Entityクラスの得意先</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.得意先 cClient)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 得意先 ";
                    mySql += "(略称,フリガナ,名称,敬称,担当者名,部署名,郵便番号,都道府県,";
                    mySql += "住所1,住所2,電話番号,FAX番号,メールアドレス,担当社員コード,締日,";
                    mySql += "税通知,請求先郵便番号,請求先都道府県,請求先住所1,請求先住所2,備考,";
                    mySql += "登録年月日,変更年月日) ";
                    mySql += "values ('" + cClient.略称 + "',";
                    mySql += "'" + cClient.フリガナ + "',";
                    mySql += "'" + cClient.名称 + "',";
                    mySql += "'" + cClient.敬称 + "',";
                    mySql += "'" + cClient.担当者名 + "',";
                    mySql += "'" + cClient.部署名 + "',";
                    mySql += "'" + cClient.郵便番号 + "',";
                    mySql += "'" + cClient.都道府県 + "',";
                    mySql += "'" + cClient.住所1 + "',";
                    mySql += "'" + cClient.住所2 + "',";
                    mySql += "'" + cClient.電話番号 + "',";
                    mySql += "'" + cClient.FAX番号 + "',";
                    mySql += "'" + cClient.メールアドレス + "',";
                    mySql += cClient.担当社員コード + ",";
                    mySql += cClient.締日 + ",";
                    mySql += "'" + cClient.税通知 + "',";
                    mySql += "'" + cClient.請求先郵便番号 + "',";
                    mySql += "'" + cClient.請求先都道府県 + "',";
                    mySql += "'" + cClient.請求先住所1 + "',";
                    mySql += "'" + cClient.請求先住所2 + "',";
                    mySql += "'" + cClient.備考 + "',";
                    mySql += "'" + cClient.登録年月日 + "',";
                    mySql += "'" + cClient.変更年月日 + "')";
                   
                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch(Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.ToString());
                    return false;
                }

            }

            /// <summary>
            /// 得意先マスター更新
            /// </summary>
            /// <param name="cClient">得意先Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.得意先 cClient)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 得意先 set ";
                    mySql += "略称 = '" + cClient.略称 + "',";
                    mySql += "フリガナ = '" + cClient.フリガナ + "',";
                    mySql += "名称 = '" + cClient.名称 + "',";
                    mySql += "敬称 = '" + cClient.敬称 + "',";
                    mySql += "担当者名 = '" + cClient.担当者名 + "',";
                    mySql += "部署名 = '" + cClient.部署名 + "',";
                    mySql += "郵便番号 = '" + cClient.郵便番号 + "',";
                    mySql += "都道府県 = '" + cClient.都道府県 + "',";
                    mySql += "住所1 = '" + cClient.住所1 + "',";
                    mySql += "住所2 = '" + cClient.住所2 + "',";
                    mySql += "電話番号 = '" + cClient.電話番号 + "',";
                    mySql += "FAX番号 = '" + cClient.FAX番号 + "',";
                    mySql += "メールアドレス = '" + cClient.メールアドレス + "',";
                    mySql += "担当社員コード = " + cClient.担当社員コード + ",";
                    mySql += "締日 = " + cClient.締日 + ",";
                    mySql += "税通知 = '" + cClient.税通知 + "',";
                    mySql += "請求先郵便番号 = '" + cClient.請求先郵便番号 + "',";
                    mySql += "請求先都道府県 = '" + cClient.請求先都道府県 + "',";
                    mySql += "請求先住所1 = '" + cClient.請求先住所1 + "',";
                    mySql += "請求先住所2 = '" + cClient.請求先住所2 + "',";
                    mySql += "備考 ='" + cClient.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cClient.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 得意先レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 得意先 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 得意先マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>得意先マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 得意先マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill得意先());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 得意先マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 得意先マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy得意先(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 入金クラス
        /// </summary>
        public class 入金 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 入金データに新規にデータを登録する
            /// </summary>
            /// <param name="cNyukin">Entityクラスの入金</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.入金 cNyukin)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 入金 ";
                    mySql += "(請求書ID,入金年月日,金額,備考,登録年月日,変更年月日) ";
                    mySql += "values (" + cNyukin.請求書ID + ",";
                    mySql += "'" + cNyukin.入金年月日 + "',";
                    mySql += cNyukin.金額 + ",";
                    mySql += "'" + cNyukin.備考 + "',";
                    mySql += "'" + cNyukin.登録年月日 + "',";
                    mySql += "'" + cNyukin.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 入金マスター更新
            /// </summary>
            /// <param name="cNyukin">入金Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.入金 cNyukin)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 入金 set ";
                    mySql += "請求書ID = " + cNyukin.請求書ID + ",";
                    mySql += "入金年月日 = '" + cNyukin.入金年月日 + "',";
                    mySql += "金額 = " + cNyukin.金額 + ",";
                    mySql += "備考 ='" + cNyukin.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cNyukin.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 入金レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 入金 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 入金データ全件のデータリーダを取得します
            /// </summary>
            /// <returns>入金データ全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 入金データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill入金());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 入金データ条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 入金データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy入金(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 配布エリアクラス
        /// </summary>
        public class 配布エリア : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 配布エリアに新規にデータを登録する
            /// </summary>
            /// <param name="cArea">Entityクラスの配布エリア</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.配布エリア cArea)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 配布エリア ";
                    mySql += "(町名ID,予定枚数,受注ID,配布指示ID,配布単価,配布日,実配布枚数,実残数,";
                    mySql += "報告枚数,報告残数,併配区分,枝番記入,完了区分,ステータス,登録年月日,変更年月日) ";
                    mySql += "values (" + cArea.町名ID + ",";
                    mySql += cArea.予定枚数 + ",";
                    mySql += cArea.受注ID+ ",";
                    mySql += cArea.配布指示ID + ",";
                    mySql += cArea.配布単価 + ",";

                    if (cArea.配布日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cArea.配布日 + "',";
                    }

                    mySql += cArea.実配布枚数 + ",";
                    mySql += cArea.実残数 + ",";
                    mySql += cArea.報告枚数 + ",";
                    mySql += cArea.報告残数 + ",";
                    mySql += cArea.併配区分 + ",";
                    mySql += "'" + cArea.枝番記入 + "',";
                    mySql += cArea.完了区分 + ",";
                    mySql += cArea.ステータス + ",";
                    mySql += "'" + cArea.登録年月日 + "',";
                    mySql += "'" + cArea.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布エリア更新
            /// </summary>
            /// <param name="cArea">配布エリアEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.配布エリア cArea)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 配布エリア set ";
                    mySql += "町名ID = " + cArea.町名ID + ",";
                    mySql += "予定枚数 = " + cArea.予定枚数 + ",";
                    mySql += "受注ID = " + cArea.受注ID+ ",";
                    mySql += "配布指示ID = " + cArea.配布指示ID + ",";
                    mySql += "配布単価 = " + cArea.配布単価 + ",";


                    if (cArea.配布日 == "")
                    {
                        mySql += "配布日 = Null,";
                    }
                    else
                    {
                        mySql += "配布日 = '" + cArea.配布日 + "',";
                    }

                    mySql += "実配布枚数 = " + cArea.実配布枚数 + ",";
                    mySql += "実残数 = " + cArea.実残数 + ",";
                    mySql += "報告枚数 = " + cArea.報告枚数 + ",";
                    mySql += "報告残数 = " + cArea.報告残数 + ",";
                    mySql += "併配区分 = " + cArea.併配区分 + ",";
                    mySql += "枝番記入 = '" + cArea.枝番記入 + "',";
                    mySql += "完了区分 = " + cArea.完了区分 + ",";
                    mySql += "ステータス = " + cArea.ステータス + ",";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cArea.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布エリアレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 配布エリア ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布エリアデータ全件のデータリーダを取得します
            /// </summary>
            /// <returns>入金データ全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 配布エリアデータのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill配布エリア());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 配布エリアデータ条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 配布エリアデータのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy配布エリア(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 未配布情報クラス
        /// </summary>
        public class 未配布情報 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 未配布情報に新規にデータを登録する
            /// </summary>
            /// <param name="cMihaifu">Entityクラスの未配布情報</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.未配布情報 cMihaifu)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 未配布情報 ";
                    mySql += "(配布エリアID,番地号,マンション名,理由,その他内容,登録年月日,変更年月日) ";
                    mySql += "values (" + cMihaifu.配布エリアID + ",";
                    mySql += "'" + cMihaifu.番地号  + "',";
                    mySql += "'" + cMihaifu.マンション名 + "',";
                    mySql += cMihaifu.理由 + ",";
                    mySql += "'" + cMihaifu.その他内容 + "',";
                    mySql += "'" + cMihaifu.登録年月日 + "',";
                    mySql += "'" + cMihaifu.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 未配布情報更新
            /// </summary>
            /// <param name="cMihaifu">未配布情報Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.未配布情報 cMihaifu)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 未配布情報 set ";
                    mySql += "配布エリアID = " + cMihaifu.配布エリアID + ",";
                    mySql += "番地号 = '" + cMihaifu.番地号 + "',";
                    mySql += "マンション名 = '" + cMihaifu.マンション名 + "',";
                    mySql += "理由 = " + cMihaifu.理由 + ",";
                    mySql += "その他内容 = '" + cMihaifu.その他内容 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cMihaifu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 未配布情報レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 未配布情報 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 未配布情報データ全件のデータリーダを取得します
            /// </summary>
            /// <returns>未配布情報全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 未配布情報データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill未配布情報());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 未配布情報データ条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 未配布情報データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy未配布情報(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }


        /// <summary>
        /// 配布員クラス
        /// </summary>
        public class 配布員 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// -------------------------------------------------------------------------
            /// <summary>
            ///     配布員マスターに新規にデータを登録する </summary>
            /// <param name="cStaff">
            ///     Entityクラスの配布員</param>
            /// <returns>
            ///     登録成功:true, 登録失敗:false</returns>
            ///     
            ///     2015/07/16 マイナンバー,ユーザーID追加
            /// -------------------------------------------------------------------------
            public Boolean DataInsert(Entity.配布員 cStaff)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 配布員 ";
                    mySql += "(ID,氏名,フリガナ,郵便番号,住所,携帯電話番号,自宅電話番号,PCアドレス,";
                    mySql += "携帯アドレス,登録日,勤務区分,街頭配布区分,街頭配布備考,支払区分,";
                    mySql += "事業所コード,金融機関コード,金融機関名,金融機関名カナ,支店コード,支店名,支店名カナ,";
                    mySql += "口座種別,口座番号,口座名義カナ,備考,";
                    mySql += "登録年月日,変更年月日,マイナンバー,ユーザーID) ";
                    mySql += "values (" + cStaff.ID + ",";
                    mySql += "'" + cStaff.氏名 + "',";
                    mySql += "'" + cStaff.フリガナ + "',";
                    mySql += "'" + cStaff.郵便番号 + "',";
                    mySql += "'" + cStaff.住所 + "',";
                    mySql += "'" + cStaff.携帯電話番号 + "',";
                    mySql += "'" + cStaff.自宅電話番号 + "',";
                    mySql += "'" + cStaff.PCアドレス + "',";
                    mySql += "'" + cStaff.携帯アドレス + "',";

                    if (cStaff.登録日 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cStaff.登録日 + "',";
                    }

                    mySql += cStaff.勤務区分 + ",";
                    mySql += cStaff.街頭配布区分 + ",";
                    mySql += "'" + cStaff.街頭配布備考 + "',";
                    mySql += "'" + cStaff.支払区分 + "',";
                    mySql += cStaff.事業所コード + ",";
                    mySql += "'" + cStaff.金融機関コード + "',";
                    mySql += "'" + cStaff.金融機関名 + "',";
                    mySql += "'" + cStaff.金融機関名カナ + "',";
                    mySql += "'" + cStaff.支店コード + "',";
                    mySql += "'" + cStaff.支店名 + "',";
                    mySql += "'" + cStaff.支店名カナ + "',";
                    mySql += cStaff.口座種別 + ",";
                    mySql += "'" + cStaff.口座番号 + "',";
                    mySql += "'" + cStaff.口座名義カナ + "',";
                    mySql += "'" + cStaff.備考 + "',";
                    mySql += "'" + cStaff.登録年月日 + "',";
                    mySql += "'" + cStaff.変更年月日 + "',";
                    mySql += "'" + cStaff.マイナンバー + "',";    // 2015/07/16
                    mySql += cStaff.ユーザーID + ")";             // 2015/07/16

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            /// -------------------------------------------------------------------------
            /// <summary>
            ///     配布員更新 </summary>
            /// <param name="cStaff">
            ///     配布エリアEntityクラス</param>
            /// <returns>
            ///     成功；true、失敗：false</returns>
            ///     
            ///     2015/07/16 マイナンバー,ユーザーID追加
            /// -------------------------------------------------------------------------
            public Boolean DataUpdate(Entity.配布員 cStaff)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 配布員 set ";
                    mySql += "氏名 = '" + cStaff.氏名 + "',";
                    mySql += "フリガナ = '" + cStaff.フリガナ + "',";
                    mySql += "郵便番号 = '" + cStaff.郵便番号 + "',";
                    mySql += "住所 = '" + cStaff.住所 + "',";
                    mySql += "携帯電話番号 = '" + cStaff.携帯電話番号 + "',";
                    mySql += "自宅電話番号 = '" + cStaff.自宅電話番号 + "',";
                    mySql += "PCアドレス = '" + cStaff.PCアドレス + "',";
                    mySql += "携帯アドレス = '" + cStaff.携帯アドレス + "',";

                    if (cStaff.登録日 == "")
                    {
                        mySql += "登録日 = Null,";
                    }
                    else
                    {
                        mySql += "登録日 = '" + cStaff.登録日 + "',";
                    }

                    mySql += "勤務区分 = " + cStaff.勤務区分 + ",";
                    mySql += "街頭配布区分 = " + cStaff.街頭配布区分 + ",";
                    mySql += "街頭配布備考 = '" + cStaff.街頭配布備考 + "',";
                    mySql += "支払区分 = '" +cStaff.支払区分 + "',";
                    mySql += "事業所コード = " + cStaff.事業所コード + ",";
                    mySql += "金融機関コード = '" + cStaff.金融機関コード + "',";
                    mySql += "金融機関名 = '" + cStaff.金融機関名 + "',";
                    mySql += "金融機関名カナ = '" + cStaff.金融機関名カナ + "',";
                    mySql += "支店コード = '" + cStaff.支店コード + "',";
                    mySql += "支店名 = '" + cStaff.支店名 + "',";
                    mySql += "支店名カナ = '" + cStaff.支店名カナ + "',";
                    mySql += "口座種別 = " + cStaff.口座種別 + ",";
                    mySql += "口座番号 = '" + cStaff.口座番号 + "',";
                    mySql += "口座名義カナ = '" + cStaff.口座名義カナ + "',";
                    mySql += "備考 = '" + cStaff.備考 + "',";
                    mySql += "変更年月日 = '" + cStaff.変更年月日 + "',";
                    mySql += "マイナンバー = '" + cStaff.マイナンバー + "',";   // 2015/07/16
                    mySql += "ユーザーID = " + cStaff.ユーザーID + " ";         // 2015/07/16
                    mySql += "where ID = " + cStaff.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            ///-------------------------------------------------------------------
            /// <summary>
            ///     配布員レコード削除 </summary>
            /// <param name="tempCode">
            ///     ID</param>
            /// <returns>
            ///     成功：true、失敗：false</returns>
            ///-------------------------------------------------------------------
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 配布員 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            /// <summary>
            /// 配布員マスターデータ全件のデータリーダを取得します
            /// </summary>
            /// <returns>配布員マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 配布員マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill配布員());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 配布員マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 配布員マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy配布員(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 配布形態クラス
        /// </summary>
        public class 配布形態 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 配布形態マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cHaifu">Entityクラスの配布形態</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.配布形態 cHaifu)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 配布形態 ";
                    mySql += "(名称,一人当たり枚数,備考,登録年月日,変更年月日) ";
                    mySql += "values ('" + cHaifu.名称 + "',";
                    mySql += cHaifu.一人当たり枚数 + ",";
                    mySql += "'" + cHaifu.備考 + "',";
                    mySql += "'" + cHaifu.登録年月日 + "',";
                    mySql += "'" + cHaifu.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布形態マスター更新
            /// </summary>
            /// <param name="cHaifu">入金Entityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.配布形態 cHaifu)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 配布形態 set ";
                    mySql += "名称 = '" + cHaifu.名称 + "',";
                    mySql += "一人当たり枚数 = " + cHaifu.一人当たり枚数 + ",";
                    mySql += "備考 ='" + cHaifu.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cHaifu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布形態レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 配布形態 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布形態マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>配布形態マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 配布形態マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill配布形態());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 配布形態マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 配布形態マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy配布形態(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 配布指示クラス
        /// </summary>
        public class 配布指示 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            ///------------------------------------------------------------------------------
            /// <summary>
            ///     配布指示データに新規にデータを登録する </summary>
            /// <param name="cShiji">
            ///     Entityクラスの配布指示</param>
            /// <returns>
            ///     </returns>
            ///------------------------------------------------------------------------------
            public Boolean DataInsert(Entity.配布指示 cShiji)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 配布指示 ";
                    mySql += "(ID,配布日,入力日,配布員ID,交通費,交通区間開始,交通区間終了,";
                    mySql += "配布開始時刻,配布終了時刻,終了レポート,未配布区分,未配布理由,";
                    mySql += "注意事項,";
                    mySql += "登録年月日,変更年月日) ";
                    mySql += "values (";
                    mySql += cShiji.ID + ",";
                    
                    //if (cShiji.配布日 == "")
                    //{
                    //    mySql += "Null,";
                    //}
                    //else
                    //{
                    //    mySql += "'" + cShiji.配布日 + "',";
                    //}

                    mySql += "'" + cShiji.配布日 + "',";
                    mySql += "'" + cShiji.入力日 + "',";
                    mySql += cShiji.配布員ID + ",";
                    mySql += cShiji.交通費 + ",";
                    mySql += "'" + cShiji.交通区間開始 + "',";
                    mySql += "'" + cShiji.交通区間終了 + "',";
                    mySql += "'" + cShiji.配布開始時刻 + "',";
                    mySql += "'" + cShiji.配布終了時刻 + "',";
                    mySql += "'" + cShiji.終了レポート + "',";
                    mySql += "'" + cShiji.未配布区分 + "',";
                    mySql += "'" + cShiji.未配布理由 + "',";
                    mySql += "'" + cShiji.注意事項 + "',";
                    mySql += "'" + cShiji.登録年月日 + "',";
                    mySql += "'" + cShiji.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布指示データ更新
            /// </summary>
            /// <param name="cShiji">配布指示データEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.配布指示 cShiji)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 配布指示 set ";

                    //if (cShiji.配布日 == "")
                    //{
                    //    mySql += "配布日 = Null,";
                    //}
                    //else
                    //{
                    //    mySql += "配布日 = '" + cShiji.配布日 + "',";
                    //}

                    mySql += "配布日 = '" + cShiji.配布日 + "',";
                    mySql += "入力日 = '" + cShiji.入力日 + "',";
                    mySql += "配布員ID = " + cShiji.配布員ID + ",";
                    mySql += "交通費 = " + cShiji.交通費 + ",";
                    mySql += "交通区間開始 = '" + cShiji.交通区間開始 + "',";
                    mySql += "交通区間終了 = '" + cShiji.交通区間終了 + "',";
                    mySql += "配布開始時刻 = '" + cShiji.配布開始時刻 + "',";
                    mySql += "配布終了時刻 = '" + cShiji.配布終了時刻 + "',";
                    mySql += "終了レポート = '" + cShiji.終了レポート + "',";
                    mySql += "未配布区分 = '" + cShiji.未配布区分 + "',";
                    mySql += "未配布理由 = '" + cShiji.未配布理由 + "',";
                    mySql += "注意事項 = '" + cShiji.注意事項 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShiji.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布指示レコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 配布指示 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 配布指示データ全件のデータリーダを取得します
            /// </summary>
            /// <returns>配布指示データ全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 配布指示データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill配布指示());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 配布指示データ条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 配布指示データのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy配布指示(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// 判型クラス
        /// </summary>
        public class 判型 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 判型マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cSize">Entityクラスの判型</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.判型 cSize)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 判型 ";
                    mySql += "(名称,卸単価1,卸単価2,卸単価3,備考,登録年月日,変更年月日) ";
                    mySql += "values ('" + cSize.名称 + "',";
                    mySql += cSize.卸単価1 + ",";
                    mySql += cSize.卸単価2 + ",";
                    mySql += cSize.卸単価3 + ",";
                    mySql += "'" + cSize.備考 + "',";
                    mySql += "'" + cSize.登録年月日 + "',";
                    mySql += "'" + cSize.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 判型マスター更新
            /// </summary>
            /// <param name="cSize">判型マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.判型 cSize)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 判型 set ";
                    mySql += "名称 = '" + cSize.名称 + "',";
                    mySql += "卸単価1 = " + cSize.卸単価1 + ",";
                    mySql += "卸単価2 = " + cSize.卸単価2 + ",";
                    mySql += "卸単価3 = " + cSize.卸単価3 + ",";
                    mySql += "備考 = '" + cSize.備考 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cSize.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 判型マスターレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 判型 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 判型マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>判型マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 判型マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill判型());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 判型マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 判型マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy判型(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// 支給控除クラス
        /// </summary>
        public class 支給控除 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 支給控除マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cShikyu">Entityクラスの判型</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.支給控除 cShikyu)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 支給控除 ";
                    mySql += "(日付,配布員ID,摘要,単価,数量,金額,支給控除区分,登録年月日,変更年月日) ";
                    mySql += "values ('" + cShikyu.日付 + "',";
                    mySql += cShikyu.配布員ID + ",";
                    mySql += "'" + cShikyu.摘要 + "',";
                    mySql += cShikyu.単価 + ",";
                    mySql += cShikyu.数量 + ",";
                    mySql += cShikyu.金額 + ",";
                    mySql += cShikyu.支給控除区分 + ",";
                    mySql += "'" + cShikyu.登録年月日 + "',";
                    mySql += "'" + cShikyu.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 支給控除マスター更新
            /// </summary>
            /// <param name="cShikyu">支給控除マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.支給控除 cShikyu)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 支給控除 set ";
                    mySql += "日付 = '" + cShikyu.日付 + "',";
                    mySql += "配布員ID = " + cShikyu.配布員ID + ",";
                    mySql += "摘要 = '" + cShikyu.摘要 + "',";
                    mySql += "単価 = " + cShikyu.単価 + ",";
                    mySql += "数量 = " + cShikyu.数量 + ",";
                    mySql += "金額 = " + cShikyu.金額 + ",";
                    mySql += "支給控除区分 = " + cShikyu.支給控除区分 + ",";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShikyu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 支給控除マスターレコード削除
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 支給控除 ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 支給控除マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>支給控除マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // 判型マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill支給控除());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 支給控除マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 支給控除マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy支給控除(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        public class 天候 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 支給控除マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cShikyu">Entityクラスの判型</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.天候 cTenkou)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 天候 ";
                    mySql += "(日付,天候,登録年月日,変更年月日) ";
                    mySql += "values ('" + cTenkou.日付 + "',";
                    mySql += "'" + cTenkou.天候名 + "',";
                    mySql += "'" + cTenkou.登録年月日 + "',";
                    mySql += "'" + cTenkou.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 天候マスター更新
            /// </summary>
            /// <param name="cShikyu">天候マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.天候 cTenkou)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 天候 set ";
                    mySql += "天候 = '" + cTenkou.天候名 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where 日付 = '" + cTenkou.日付 + "'";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 天候マスターレコード削除
            /// </summary>
            /// <param name="tempDate">日付</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(DateTime tempDate)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 天候 ";
                    mySql += "where 日付 = " + tempDate;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 天候マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>天候マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    //天候マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill天候());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 天候マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 天候マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy天候(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        public class 未配布理由 : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// 未配布理由マスターに新規にデータを登録する
            /// </summary>
            /// <param name="cRiyu">Entityクラスの未配布理由</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.未配布理由 cRiyu)
            {

                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    //登録処理
                    mySql = "";
                    mySql += "insert into 未配布理由 ";
                    mySql += "(ID,摘要,登録年月日,変更年月日) ";
                    mySql += "values ('" + cRiyu.ID + "',";
                    mySql += "'" + cRiyu.摘要 + "',";
                    mySql += "'" + cRiyu.登録年月日 + "',";
                    mySql += "'" + cRiyu.変更年月日 + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 未配布理由マスター更新
            /// </summary>
            /// <param name="cRiyu">未配布理由マスターEntityクラス</param>
            /// <returns>成功；true、失敗：false</returns>
            public Boolean DataUpdate(Entity.未配布理由 cRiyu)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update 未配布理由 set ";
                    mySql += "摘要 = '" + cRiyu.摘要 + "',";
                    mySql += "変更年月日 = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cRiyu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 未配布理由マスターレコード削除
            /// </summary>
            /// <param name="tempDate">日付</param>
            /// <returns>成功：true、失敗：false</returns>
            public Boolean DataDelete(int tempID)
            {
                try
                {
                    //データベース接続情報の取得
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from 未配布理由 ";
                    mySql += "where ID = " + tempID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// 未配布理由マスター全件のデータリーダを取得します
            /// </summary>
            /// <returns>未配布理由マスター全件データリーダー</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    //未配布理由マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillAccess(new Access.DataAccess.Fill未配布理由());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// 未配布理由マスター条件付データリーダー取得
            /// </summary>
            /// <param name="tempString">条件式：SQL文のwhere句以下を記述します</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // 天候マスターのデータリーダー取得クラスのインスタンスを引数で指定
                    return FillByAccess(new Access.DataAccess.FillBy未配布理由(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

    }
}

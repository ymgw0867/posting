using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using MyLibrary;

namespace posting
{
    class Access
    {
        public class DataAccess
        {
            // コンストラクタ
            public DataAccess()
            {
            }

            // データリーダー取得インターフェイス
            public interface IFill
            {
                // 抽象メソッド
                OleDbDataReader GetdReader(OleDbConnection tempConnection);
            }

            public class Fill会社情報 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 会社情報マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    mySql = "";
                    mySql += "select * from 会社情報 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill事業所 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 事業所マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 事業所 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill社員 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 社員マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 社員 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill受注 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 受注データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 受注 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill受注種別 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 受注種別マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 受注種別 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill所属 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 所属マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 所属 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill振込口座 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 振込口座マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 振込口座 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill請求書 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 請求書データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 請求書 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill税率 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 税率マスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 税率 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill町名 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 町名マスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 町名 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill締日パターン : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 締日パターンマスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 締日パターン order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill得意先 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 得意先マスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 得意先 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill入金 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 入金データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 入金 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill配布エリア : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 配布エリアデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 配布エリア order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill配布員 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 配布員データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 配布員 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill配布形態 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 配布形態マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 配布形態 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill配布指示 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 配布指示データデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 配布指示 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill判型 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 判型マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 判型 order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill支給控除 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 支給控除マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select 支給控除.*,配布員.氏名 from ";
                    mySql += "支給控除 left join 配布員 ";
                    mySql += "on 支給控除.配布員ID = 配布員.ID ";
                    mySql += "order by 支給控除.ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill天候 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 天候マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {

                    mySql = "";
                    mySql += "select 天候.* from 天候 ";
                    mySql += "order by 天候.日付";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill未配布理由 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 未配布理由マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {

                    mySql = "";
                    mySql += "select 未配布理由.* from 未配布理由 ";
                    mySql += "order by 未配布理由.ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill未配布情報 : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 未配布情報マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {

                    mySql = "";
                    mySql += "select 未配布情報.* from 未配布情報 ";
                    mySql += "order by 未配布情報.ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }


            // 条件付きデータリーダー取得インターフェイス
            public interface IFillBy
            {
                // 抽象メソッド
                OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString);
            }

            // データリーダー取得クラス
            public class free_dsReader : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文</param>
                /// <returns>データリーダー</returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += tempString;
                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;
                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy会社情報 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き会社情報マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 会社情報 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy事業所 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き事業所マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from 事業所 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy社員 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き社員マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {

                    mySql = "";
                    mySql += "select * from 社員 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy受注 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き受注データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {

                    mySql = "";
                    mySql += "select * from 受注 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy受注種別 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き受注種別データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 受注種別 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy所属 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き所属データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 所属 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy振込口座 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き振込口座マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 振込口座 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy請求書 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き請求書データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 請求書 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy税率 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き税率マスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 税率 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy町名 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き町名マスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 町名 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy締日パターン : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き締日パターンマスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 締日パターン ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy得意先 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き得意先マスターリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 得意先 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy入金 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き入金データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 入金 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy配布エリア : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き配布エリアデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 配布エリア ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy配布員 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き配布員マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 配布員 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy配布形態 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き配布形態マスターデータリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 配布形態 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy配布指示 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き配布指示データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 配布指示 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy判型: IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き判型データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 判型 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy支給控除 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き支給控除データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select 支給控除.*,配布員.氏名 from ";
                    mySql += "支給控除 left join 配布員 ";
                    mySql += "on 支給控除.配布員ID = 配布員.ID ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy天候 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き天候データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select 天候.* from 天候 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy未配布理由 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き未配布理由データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 未配布理由 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy未配布情報 : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// 条件付き未配布情報データリーダー取得
                /// </summary>
                /// <param name="tempConnection">データベース接続情報</param>
                /// <param name="tempString">SQL文の where以下の条件式を記述します</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from 未配布情報 ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

        }
    }
} 


   

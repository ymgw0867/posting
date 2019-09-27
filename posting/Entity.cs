using System;
using System.Collections.Generic;
using System.Text;

namespace posting
{
    class Entity
    {
        public class 会社情報
        {
            public int ID { get; set; }
            public string 会社名 { get; set; }
            public string 代表者氏名 { get; set; }
            public string 役職名 { get; set; }
            public string 電話番号 { get; set; }
            public string FAX番号 { get; set; }
            public string 住所1 { get; set; }
            public string 住所2 { get; set; }
            public string 郵便番号 { get; set; }
            public string メールアドレス { get; set; }
            public string 部署名 { get; set; }
            public string 担当者名 { get; set; }
            public string 特記事項1 { get; set; }
            public string 特記事項2 { get; set; }
            public string 依頼人コード { get; set; }
            public string 依頼人名 { get; set; }
            public string 金融機関コード { get; set; }
            public string 金融機関名 { get; set; }
            public string 支店コード { get; set; }
            public string 支店名 { get; set; }
            public int 口座種別 { get; set; }
            public string 口座番号 { get; set; }
            public int 配布フラグ { get; set; }
            public DateTime 登録年月日 { get; set; }
            public DateTime 変更年月日 { get; set; }
            public string 郵便番号CSVパス { get; set; } 
            public string 受注確定書入力シートパス { get; set; }    // 2019/03/06
        }

        public class 事業所
        {
            private int F_ID;
            private string F_名称;
            private string F_郵便番号;
            private string F_住所1;
            private string F_住所2;
            private string F_電話番号;
            private string F_FAX番号;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string 名称
            {
                set
                {
                    F_名称 = value;
                }
                get
                {
                    return F_名称;
                }
            }

            public string 郵便番号
            {
                set
                {
                    F_郵便番号 = value;
                }
                get
                {
                    return F_郵便番号;
                }
            }

            public string 住所1
            {
                set
                {
                    F_住所1 = value;
                }
                get
                {
                    return F_住所1;
                }
            }

            public string 住所2
            {
                set
                {
                    F_住所2 = value;
                }
                get
                {
                    return F_住所2;
                }
            }

            public string 電話番号
            {
                set
                {
                    F_電話番号 = value;
                }
                get
                {
                    return F_電話番号;
                }
            }

            public string FAX番号
            {
                set
                {
                    F_FAX番号 = value;
                }
                get
                {
                    return F_FAX番号;
                }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }
            
            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 社員
        {
            private int F_ID;
            private string F_氏名;
            private string F_フリガナ;
            private int F_所属コード;
            private string F_役職;
            private string F_入社年月日;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;
 
            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string 氏名
            {
                set
                {
                    F_氏名 = value;
                }
                get
                {
                    return F_氏名;
                }
            }

            public string フリガナ
            {
                set
                {
                    F_フリガナ = value;
                }
                get
                {
                    return F_フリガナ;
                }
            }

            public int 所属コード
            {
                set
                {
                    F_所属コード = value;
                }
                get
                {
                    return F_所属コード;
                }
            }

            public string 役職
            {
                set
                {
                    F_役職 = value;
                }
                get
                {
                    return F_役職;
                }
            }

            public string 入社年月日
            {
                set
                {
                    F_入社年月日 = value;
                }
                get
                {
                    return F_入社年月日;
                }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 受注
        {
            public long ID { get; set; }            
            public int 事業所ID { get; set; }            
            public DateTime 受注日 { get; set; }            
            public string 受注区分 { get; set; }            
            public int 得意先ID { get; set; }            
            public int 社員ID { get; set; }            
            public string チラシ名 { get; set; }            
            public int 受注種別ID { get; set; }            
            public double 単価 { get; set; }            
            public int 枚数 { get; set; }            
            public int 金額 { get; set; }            
            public int 消費税 { get; set; }            
            public int 税込金額 { get; set; }            
            public int 値引額 { get; set; }            
            public int 売上金額 { get; set; }            
            public int 税率 { get; set; }            
            public int 判型 { get; set; }            
            public double 配布単価 { get; set; }            
            public string 依頼先 { get; set; }            
            public double 原価 { get; set; }            
            public int 配布形態 { get; set; }            
            public string 配布条件 { get; set; }            
            public string 配布開始日 { get; set; }            
            public string 配布終了日 { get; set; }            
            public string 配布猶予 { get; set; }            
            public string 納品予定日 { get; set; }            
            public string 納品形態 { get; set; }            
            public int 請求書 { get; set; }            
            public int 請求書ID { get; set; }            
            public string 請求書発行日 { get; set; }            
            public string 入金方法 { get; set; }            
            public string 入金予定日 { get; set; }            
            public string 報告時期 { get; set; }            
            public string 報告精度 { get; set; }            
            public string 報告方法 { get; set; }            
            public string メールアドレス { get; set; }            
            public int 振込口座ID { get; set; }            
            public int 未配布情報有無 { get; set; }            
            public int 枝番有無 { get; set; }            
            public string 特記事項 { get; set; }          
            public string エリア備考 { get; set; }
            public int 完了区分 { get; set; }
            public DateTime 登録年月日 { get; set; }
            public DateTime 変更年月日 { get; set; }
            public long IDTemplateS { get; set; }
            public long IDTemplateE { get; set; }
            public int 併配除外 { get; set; }

            // 2015/06/30 追加フィールド
            public int 外注先ID営業 { get; set; }
            public string 外注先支払日営業 { get; set; }
            public double 外注先原価営業 { get; set; }
            public string 外注先依頼日営業 { get; set; }    // 2015/07/17

            public int 外注先ID支払 { get; set; }
            public string 外注先支払日支払 { get; set; }
            public double 外注先原価支払 { get; set; }
            public int 外注先ID支払2 { get; set; }           // 2016/10/14
            public string 外注先支払日支払2 { get; set; }   // 2016/10/14
            public double 外注先原価支払2 { get; set; }    // 2016/10/14
            public int 外注先ID支払3 { get; set; }       // 2016/10/14
            public string 外注先支払日支払3 { get; set; }   // 2016/10/14
            public double 外注先原価支払3 { get; set; }    // 2016/10/14
            public string 外注先依頼日支払 { get; set; }    // 2015/07/17
            public string 外注先依頼日支払2 { get; set; }    // 2016/10/15
            public string 外注先依頼日支払3 { get; set; }    // 2016/10/15

            public int ユーザーID { get; set; }
            public int 案件種別 { get; set; }
            public string 外注渡し日 { get; set; }           // 2015/08/11
            public string 外注渡し日2 { get; set; }           // 2016/10/15
            public string 外注渡し日3 { get; set; }           // 2016/10/15
            public string 外注受け渡し担当者 { get; set; }   // 2015/08/11
            public string 外注受け渡し担当者2 { get; set; }   // 2016/10/15
            public string 外注受け渡し担当者3 { get; set; }   // 2016/10/15

            public int 外注委託枚数 { get; set; }     // 2015/09/20
            public int 外注委託枚数2 { get; set; }     // 2016/10/15
            public int 外注委託枚数3 { get; set; }     // 2016/10/15
            public string 業種 { get; set; }         // 2015/09/20      
            public string 営業備考 { get; set; }   // 2019/03/01
        }

        public class 受注種別
        {
            private int F_ID;
            private string F_名称;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string 名称
            {
                set
                {
                    F_名称 = value;
                }
                get
                {
                    return F_名称;
                }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 所属
        {
            private int F_ID;
            private string F_所属名1;
            private string F_所属名2;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string 所属名1
            {
                set
                {
                    F_所属名1 = value;
                }
                get
                {
                    return F_所属名1;
                }
            }

            public string 所属名2
            {
                set
                {
                    F_所属名2 = value;
                }
                get
                {
                    return F_所属名2;
                }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 振込口座
        {
            private int F_ID;
            private string F_金融機関名;
            private string F_支店名;
            private int F_口座種別;
            private string F_口座番号;
            private string F_口座名義;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

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

            public string 金融機関名
            {
                get
                {
                    return F_金融機関名;
                }
                set
                {
                    F_金融機関名 = value;
                }
            }

            public string 支店名
            {
                get
                {
                    return F_支店名;
                }
                set
                {
                    F_支店名 = value;
                }
            }

            public int 口座種別
            {
                get
                {
                    return F_口座種別;
                }
                set
                {
                    F_口座種別 = value;
                }
            }

            public string 口座番号
            {
                get
                {
                    return F_口座番号;
                }
                set
                {
                    F_口座番号 = value;
                }
            }

            public string 口座名義
            {
                get
                {
                    return F_口座名義;
                }
                set
                {
                    F_口座名義 = value;
                }
            }

            public DateTime 登録年月日
            {
                get
                {
                    return F_登録年月日;
                }
                set
                {
                    F_登録年月日 = value;
                }
            }

            public DateTime 変更年月日
            {
                get
                {
                    return F_変更年月日;
                }
                set
                {
                    F_変更年月日 = value;
                }
            }
       }
       
        public class 請求書
        {
            private int F_ID;
            private int F_得意先ID;
            private int F_請求金額;
            private int F_消費税;
            private int F_値引額;
            private int F_売上金額;
            private int F_税率;
            private DateTime F_入金予定日;
            private DateTime F_発行日;
            private int F_入金残;
            private int F_完了区分;
            private int F_振込口座ID1;
            private int F_振込口座ID2;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

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
            public int 得意先ID
            {
                get
                {
                    return F_得意先ID;
                }
                set
                {
                    F_得意先ID = value;
                }
            }
            public int 請求金額
            {
                get
                {
                    return F_請求金額;
                }
                set
                {
                    F_請求金額 = value;
                }
            }
            public int 消費税
            {
                get
                {
                    return F_消費税;
                }
                set
                {
                    F_消費税 = value;
                }
            }
            public int 値引額
            {
                get
                {
                    return F_値引額;
                }
                set
                {
                    F_値引額 = value;
                }
            }

            public int 売上金額
            {
                get
                {
                    return F_売上金額;
                }
                set
                {
                    F_売上金額 = value;
                }
            }
            public int 税率
            {
                get
                {
                    return F_税率;
                }
                set
                {
                    F_税率 = value;
                }
            }
            public DateTime 入金予定日
            {
                get
                {
                    return F_入金予定日;
                }
                set
                {
                    F_入金予定日 = value;
                }
            }
            public DateTime 発行日
            {
                get
                {
                    return F_発行日;
                }
                set
                {
                    F_発行日 = value;
                }
            }
            public int 入金残
            {
                get
                {
                    return F_入金残;
                }
                set
                {
                    F_入金残 = value;
                }
            }
            public int 完了区分
            {
                get
                {
                    return F_完了区分;
                }
                set
                {
                    F_完了区分 = value;
                }
            }
            public int 振込口座ID1
            {
                get
                {
                    return F_振込口座ID1;
                }
                set
                {
                    F_振込口座ID1 = value;
                }
            }
            public int 振込口座ID2
            {
                get
                {
                    return F_振込口座ID2;
                }
                set
                {
                    F_振込口座ID2 = value;
                }
            }
            public string 備考
            {
                get
                {
                    return F_備考;
                }
                set
                {
                    F_備考 = value;
                }
            }
            public DateTime 登録年月日
            {
                get
                {
                    return F_登録年月日;
                }
                set
                {
                    F_登録年月日 = value;
                }
            }
            public DateTime 変更年月日
            {
                get
                {
                    return F_変更年月日;
                }
                set
                {
                    F_変更年月日 = value;
                }
            }
        }

        public class 税率
        {
            private int F_ID;
            private int F_税率;
            private DateTime F_開始年月日;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public int 設定税率
            {
                set
                {
                    F_税率 = value;
                }
                get
                {
                    return F_税率;
                }
            }

            public DateTime 開始年月日
            {
                set
                {
                    F_開始年月日 = value;
                }
                get
                {
                    return F_開始年月日;
                }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 町名
        {
            private int F_ID;
            private string F_名称;
            private int F_市区町村コード;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string 名称
            {
                set
                {
                    F_名称 = value;
                }
                get
                {
                    return F_名称;
                }
            }

            public int 市区町村コード
            {
                get { return F_市区町村コード; }
                set { F_市区町村コード = value; }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 締日パターン
        {
            private int F_ID;
            private string F_摘要;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string 摘要
            {
                set
                {
                    F_摘要 = value;
                }
                get
                {
                    return F_摘要;
                }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 得意先
        {
            public int ID { set; get; }
            public string 略称 { set; get; }
            public string フリガナ { set; get; }
            public string 名称 { set; get; }   
            public string 敬称 { set; get; }
            public string 担当者名 { set; get; }
            public string 部署名 { set; get; }
            public string 郵便番号 { set; get; }
            public string 都道府県 { set; get; }
            public string 住所1 { set; get; }
            public string 住所2 { set; get; }
            public string 電話番号 { set; get; }
            public string FAX番号 { set; get; }
            public string メールアドレス { set; get; }
            public int 担当社員コード { set; get; }
            public int 締日 { set; get; }
            public string 税通知 { set; get; }
            public string 請求先郵便番号 { set; get; }
            public string 請求先都道府県 { set; get; }
            public string 請求先住所1 { set; get; }
            public string 請求先住所2 { set; get; }
            public string 備考 { set; get; }
            public DateTime 登録年月日 { set; get; }
            public DateTime 変更年月日 { set; get; }
            public string 請求先名称 { set; get; }
            public string 請求先担当者名 { set; get; }
        }

        public class 入金
        {
            private int F_ID;
            private int F_請求書ID;
            private DateTime F_入金年月日;
            private int F_金額;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public int 請求書ID
            {
                set
                {
                    F_請求書ID = value;
                }
                get
                {
                    return F_請求書ID;
                }
            }

            public DateTime 入金年月日
            {
                set
                {
                    F_入金年月日 = value;
                }
                get
                {
                    return F_入金年月日;
                }
            }

            public int 金額
            {
                set
                {
                    F_金額 = value;
                }
                get
                {
                    return F_金額;
                }
            }

            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 配布エリア
        {
            private int F_ID;
            private int F_町名ID;
            private int F_予定枚数;
            private long F_受注ID;
            private int F_配布指示ID;
            private double F_配布単価;
            private string F_配布日;
            private int F_実配布枚数;
            private int F_実残数;
            private int F_報告枚数;
            private int F_報告残数;
            private int F_併配区分;
            private string F_枝番記入;
            private int F_完了区分;
            private int F_ステータス;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }
            public int 町名ID
            {
                set
                {
                    F_町名ID = value;
                }
                get
                {
                    return F_町名ID;
                }
            }
            public int 予定枚数
            {
                set
                {
                    F_予定枚数 = value;
                }
                get
                {
                    return F_予定枚数;
                }
            }
            public long 受注ID
            {
                set
                {
                    F_受注ID = value;
                }
                get
                {
                    return F_受注ID;
                }
            }
            public int 配布指示ID
            {
                set
                {
                    F_配布指示ID = value;
                }
                get
                {
                    return F_配布指示ID;
                }
            }
            public double 配布単価
            {
                set
                {
                    F_配布単価 = value;
                }
                get
                {
                    return F_配布単価;
                }
            }
            public string 配布日
            {
                set
                {
                    F_配布日 = value;
                }
                get
                {
                    return F_配布日;
                }
            }
            public int 実配布枚数
            {
                set
                {
                    F_実配布枚数 = value;
                }
                get
                {
                    return F_実配布枚数;
                }
            }
            public int 実残数
            {
                set
                {
                    F_実残数 = value;
                }
                get
                {
                    return F_実残数;
                }
            }
            public int 報告枚数
            {
                set
                {
                    F_報告枚数 = value;
                }
                get
                {
                    return F_報告枚数;
                }
            }
            public int 報告残数
            {
                set
                {
                    F_報告残数 = value;
                }
                get
                {
                    return F_報告残数;
                }
            }
            public int 併配区分
            {
                set
                {
                    F_併配区分 = value;
                }
                get
                {
                    return F_併配区分;
                }
            }
            public string 枝番記入
            {
                set
                {
                    F_枝番記入 = value;
                }
                get
                {
                    return F_枝番記入;
                }
            }
            public int 完了区分
            {
                set
                {
                    F_完了区分 = value;
                }
                get
                {
                    return F_完了区分;
                }
            }
            public int ステータス
            {
                set
                {
                    F_ステータス = value;
                }
                get
                {
                    return F_ステータス;
                }
            }
            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }
            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 未配布情報
        {
            private int F_ID;

            public int ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }

            private int F_配布エリアID;

            public int 配布エリアID
            {
                get { return F_配布エリアID; }
                set { F_配布エリアID = value; }
            }

            private string F_番地号;

            public string 番地号
            {
                get { return F_番地号; }
                set { F_番地号 = value; }
            }

            private string F_マンション名;

            public string マンション名
            {
                get { return F_マンション名; }
                set { F_マンション名 = value; }
            }

            private int F_理由;

            public int 理由
            {
                get { return F_理由; }
                set { F_理由 = value; }
            }

            private string F_その他内容;

            public string その他内容
            {
                get { return F_その他内容; }
                set { F_その他内容 = value; }
            }

            private DateTime F_登録年月日;

            public DateTime 登録年月日
            {
                get { return F_登録年月日; }
                set { F_登録年月日 = value; }
            }

            private DateTime F_変更年月日;

            public DateTime 変更年月日
            {
                get { return F_変更年月日; }
                set { F_変更年月日 = value; }
            }

        }

        public class 配布員
        {
            public int ID { set; get; }
            public string 氏名 { set; get; }
            public string フリガナ { set; get; }
            public string 郵便番号 { set; get; }
            public string 住所 { set; get; }
            public string 携帯電話番号 { set; get; }
            public string 自宅電話番号 { set; get; }
            public string PCアドレス { set; get; }
            public string 携帯アドレス { set; get; }
            public string 登録日 { set; get; }
            public int 勤務区分 { set; get; }
            public int 街頭配布区分 { set; get; }
            public string 街頭配布備考 { set; get; }
            public string 支払区分 { set; get; }
            public int 事業所コード { set; get; }
            public string 金融機関コード { set; get; }
            public string 金融機関名 { set; get; }
            public string 金融機関名カナ { set; get; }
            public string 支店コード { set; get; }
            public string 支店名 { set; get; }
            public string 支店名カナ { set; get; }
            public int 口座種別 { set; get; }
            public string 口座番号 { set; get; }
            public string 口座名義カナ { set; get; }
            public string 備考 { set; get; }
            public DateTime 登録年月日 { set; get; }
            public DateTime 変更年月日 { set; get; }            
            public int ユーザーID { get; set; }
            public string マイナンバー { get; set; }            
        }

        public class 配布形態
        {
            public int ID { set; get; }
            public string 名称 { set; get; }
            public int 一人当たり枚数 { set; get; }
            public string 備考 { set; get; }
            public DateTime 登録年月日 { set; get; }
            public DateTime 変更年月日 { set; get; }            
        }

        public class 配布指示
        {
            private int F_ID;
            private DateTime F_配布日;
            private DateTime F_入力日;
            private int F_配布員ID;
            private int F_交通費;
            private string F_交通区間開始;
            private string F_交通区間終了;
            private string F_配布開始時刻;
            private string F_配布終了時刻;
            private string F_終了レポート;
            private string F_未配布区分;
            private string F_未配布理由;
            private string F_注意事項;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;
            private int F_ユーザーID;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public DateTime 配布日
            {
                set
                {
                    F_配布日 = value;
                }
                get
                {
                    return F_配布日;
                }
            }

            public DateTime 入力日
            {
                set
                {
                    F_入力日 = value;
                }
                get
                {
                    return F_入力日;
                }
            }

            public int 配布員ID
            {
                set
                {
                    F_配布員ID = value;
                }
                get
                {
                    return F_配布員ID;
                }
            }

            public int 交通費
            {
                set
                {
                    F_交通費 = value;
                }
                get
                {
                    return F_交通費;
                }
            }

            public string 交通区間開始
            {
                set
                {
                    F_交通区間開始 = value;
                }
                get
                {
                    return F_交通区間開始;
                }
            }

            public string 交通区間終了
            {
                set
                {
                    F_交通区間終了 = value;
                }
                get
                {
                    return F_交通区間終了;
                }
            }

            public string 配布開始時刻
            {
                set
                {
                    F_配布開始時刻 = value;
                }
                get
                {
                    return F_配布開始時刻;
                }
            }

            public string 配布終了時刻
            {
                set
                {
                    F_配布終了時刻 = value;
                }
                get
                {
                    return F_配布終了時刻;
                }
            }

            public string 終了レポート
            {
                set
                {
                    F_終了レポート = value;
                }
                get
                {
                    return F_終了レポート;
                }
            }

            public string 未配布区分
            {
                set
                {
                    F_未配布区分 = value;
                }
                get
                {
                    return F_未配布区分;
                }
            }

            public string 未配布理由
            {
                set
                {
                    F_未配布理由 = value;
                }
                get
                {
                    return F_未配布理由;
                }
            }

            public string 注意事項
            {
                set
                {
                    F_注意事項 = value;
                }
                get
                {
                    return F_注意事項;
                }
            }


            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
            public int ユーザーID
            {
                set
                {
                    F_ユーザーID = value;
                }
                get
                {
                    return F_ユーザーID;
                }
            }
        }

        public class 判型
        {
            private int F_ID;
            private string F_名称;
            private double F_卸単価1;
            private double F_卸単価2;
            private double F_卸単価3;
            private string F_備考;
            private DateTime F_登録年月日;
            private DateTime F_変更年月日;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string 名称
            {
                set
                {
                    F_名称 = value;
                }
                get
                {
                    return F_名称;
                }
            }

            public double 卸単価1
            {
                set
                {
                    F_卸単価1 = value;
                }
                get
                {
                    return F_卸単価1;
                }
            }

            public double 卸単価2
            {
                set
                {
                    F_卸単価2 = value;
                }
                get
                {
                    return F_卸単価2;
                }
            }

            public double 卸単価3
            {
                set
                {
                    F_卸単価3 = value;
                }
                get
                {
                    return F_卸単価3;
                }
            }
            
            public string 備考
            {
                set
                {
                    F_備考 = value;
                }
                get
                {
                    return F_備考;
                }
            }

            public DateTime 登録年月日
            {
                set
                {
                    F_登録年月日 = value;
                }
                get
                {
                    return F_登録年月日;
                }
            }

            public DateTime 変更年月日
            {
                set
                {
                    F_変更年月日 = value;
                }
                get
                {
                    return F_変更年月日;
                }
            }
        }

        public class 支給控除
        {
            //ID
            private int F_ID;

            public int ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }
	
            //日付
            private DateTime F_日付;

            public DateTime 日付
            {
                get { return F_日付; }
                set { F_日付 = value; }
            }

            private int F_配布員ID;

            public int 配布員ID
            {
                get { return F_配布員ID; }
                set { F_配布員ID = value; }
            }

            private string F_配布員名;

            public string 配布員名
            {
                get { return F_配布員名; }
                set { F_配布員名 = value; }
            }
	
            private string F_摘要;

            public string 摘要
            {
                get { return F_摘要; }
                set { F_摘要 = value; }
            }

            private double F_単価;

            public double 単価
            {
                get { return F_単価; }
                set { F_単価 = value; }
            }

            private int F_数量;

            public int 数量
            {
                get { return F_数量; }
                set { F_数量 = value; }
            }

            private double F_金額;

            public double 金額
            {
                get { return F_金額; }
                set { F_金額 = value; }
            }

            private int F_支給控除区分;

            public int 支給控除区分
            {
                get { return F_支給控除区分; }
                set { F_支給控除区分 = value; }
            }

            private DateTime F_登録年月日;

            public DateTime 登録年月日
            {
                get { return F_登録年月日; }
                set { F_登録年月日 = value; }
            }

            private DateTime F_変更年月日;

            public DateTime 変更年月日
            {
                get { return F_変更年月日; }
                set { F_変更年月日 = value; }
            }
	
        }

        public class 天候
        {
            private DateTime F_日付;

            public DateTime 日付
            {
                get { return F_日付; }
                set { F_日付 = value; }
            }

            private string F_天候名;

            public string 天候名
            {
                get { return F_天候名; }
                set { F_天候名 = value; }
            }

            private DateTime F_登録年月日;

            public DateTime 登録年月日
            {
                get { return F_登録年月日; }
                set { F_登録年月日 = value; }
            }

            private DateTime F_変更年月日;

            public DateTime 変更年月日
            {
                get { return F_変更年月日; }
                set { F_変更年月日 = value; }
            }
        }

        public class 未配布理由
        {
            private int F_ID;

            public int ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }

            private string F_摘要;

            public string 摘要
            {
                get { return F_摘要; }
                set { F_摘要 = value; }
            }

            private DateTime F_登録年月日;

            public DateTime 登録年月日
            {
                get { return F_登録年月日; }
                set { F_登録年月日 = value; }
            }

            private DateTime F_変更年月日;

            public DateTime 変更年月日
            {
                get { return F_変更年月日; }
                set { F_変更年月日 = value; }
            }
        }

        public class 全銀ヘッダ
        {
            private string F_データ区分;

            public string データ区分
            {
                get { return F_データ区分; }
                set { F_データ区分 = value; }
            }

            private string F_種別コード;

            public string 種別コード
            {
                get { return F_種別コード; }
                set { F_種別コード = value; }
            }

            private string F_コード区分;

            public string コード区分
            {
                get { return F_コード区分; }
                set { F_コード区分 = value; }
            }

            private string F_振込依頼人コード;

            public string 振込依頼人コード
            {
                get { return F_振込依頼人コード; }
                set { F_振込依頼人コード = value; }
            }

            private string F_振込依頼人名;

            public string 振込依頼人名
            {
                get { return F_振込依頼人名; }
                set { F_振込依頼人名 = value; }
            }

            private string F_取組日;

            public string 取組日
            {
                get { return F_取組日; }
                set { F_取組日 = value; }
            }

            private string F_仕向銀行番号;

            public string 仕向銀行番号
            {
                get { return F_仕向銀行番号; }
                set { F_仕向銀行番号 = value; }
            }

            private string F_仕向銀行名;

            public string 仕向銀行名
            {
                get { return F_仕向銀行名; }
                set { F_仕向銀行名 = value; }
            }

            private string F_仕向支店番号;

            public string 仕向支店番号
            {
                get { return F_仕向支店番号; }
                set { F_仕向支店番号 = value; }
            }

            private string F_仕向支店名;

            public string 仕向支店名
            {
                get { return F_仕向支店名; }
                set { F_仕向支店名 = value; }
            }

            private string F_預金種目;

            public string 預金種目
            {
                get { return F_預金種目; }
                set { F_預金種目 = value; }
            }

            private string F_口座番号;

            public string 口座番号
            {
                get { return F_口座番号; }
                set { F_口座番号 = value; }
            }

            private string F_ダミー;

            public string ダミー
            {
                get { return F_ダミー; }
                set { F_ダミー = value; }
            }
	
        }

        public class 全銀データレコード
        {
            private string F_データ区分;

            public string データ区分
            {
                get { return F_データ区分; }
                set { F_データ区分 = value; }
            }

            private string F_被仕向銀行番号;

            public string 被仕向銀行番号
            {
                get { return F_被仕向銀行番号; }
                set { F_被仕向銀行番号 = value; }
            }

            private string F_被仕向銀行名;

            public string 被仕向銀行名
            {
                get { return F_被仕向銀行名; }
                set { F_被仕向銀行名 = value; }
            }

            private string F_被仕向支店番号;

            public string 被仕向支店番号
            {
                get { return F_被仕向支店番号; }
                set { F_被仕向支店番号 = value; }
            }

            private string F_被仕向支店名;

            public string 被仕向支店名
            {
                get { return F_被仕向支店名; }
                set { F_被仕向支店名 = value; }
            }

            private string F_手形交換所番号;

            public string 手形交換所番号
            {
                get { return F_手形交換所番号; }
                set { F_手形交換所番号 = value; }
            }

            private string F_預金種目;

            public string 預金種目
            {
                get { return F_預金種目; }
                set { F_預金種目 = value; }
            }

            private string F_口座番号;

            public string 口座番号
            {
                get { return F_口座番号; }
                set { F_口座番号 = value; }
            }

            private string F_受取人名;

            public string 受取人名
            {
                get { return F_受取人名; }
                set { F_受取人名 = value; }
            }

            private string F_振込金額;

            public string 振込金額
            {
                get { return F_振込金額; }
                set { F_振込金額 = value; }
            }

            private string F_新規コード;

            public string 新規コード
            {
                get { return F_新規コード; }
                set { F_新規コード = value; }
            }

            private string F_顧客コード1;

            public string 顧客コード1
            {
                get { return F_顧客コード1; }
                set { F_顧客コード1 = value; }
            }

            private string F_顧客コード2;

            public string 顧客コード2
            {
                get { return F_顧客コード2; }
                set { F_顧客コード2 = value; }
            }

            private string F_EDI情報;

            public string EDI情報
            {
                get { return F_EDI情報; }
                set { F_EDI情報 = value; }
            }

            private string F_振込指定区分;

            public string 振込指定区分
            {
                get { return F_振込指定区分; }
                set { F_振込指定区分 = value; }
            }

            private string F_識別表示;

            public string 識別表示
            {
                get { return F_識別表示; }
                set { F_識別表示 = value; }
            }

            private string F_ダミー;

            public string ダミー
            {
                get { return F_ダミー; }
                set { F_ダミー = value; }
            }
	
        }

        public class 全銀トレーラーレコード
        {
            private string F_データ区分;

            public string データ区分
            {
                get { return F_データ区分; }
                set { F_データ区分 = value; }
            }

            private string F_合計件数;

            public string 合計件数
            {
                get { return F_合計件数; }
                set { F_合計件数 = value; }
            }

            private string F_合計金額;

            public string 合計金額
            {
                get { return F_合計金額; }
                set { F_合計金額 = value; }
            }

            private string F_ダミー;

            public string ダミー
            {
                get { return F_ダミー; }
                set { F_ダミー = value; }
            }
        }

        public class 全銀エンドレコード
        {
            private string F_データ区分;

            public string データ区分
            {
                get { return F_データ区分; }
                set { F_データ区分 = value; }
            }

            private string F_ダミー;

            public string ダミー
            {
                get { return F_ダミー; }
                set { F_ダミー = value; }
            }
        }

        //汎用データヘッダ項目
        public class OutPutHeader
        {
            public const string dn01 = @"""OBCD001""";  // 伝票区切

            public const string hd00 = @"""CSJS003""";  // 部門指定方法 
            public const string hd01 = @"""CSJS004""";  // 伝票部門コード 
            public const string hd02 = @"""CSJS005""";  // 日付 
            public const string hd03 = @"""CSJS007""";  // 伝票№

            public const string kr01 = @"""CSJS200""";  // 借方部門コード
            public const string kr02 = @"""CSJS201""";  // 借方勘定科目コード
            public const string kr03 = @"""CSJS202""";  // 借方補助科目コード
            public const string kr04 = @"""CSJS205""";  // 借方事業区分コード
            public const string kr05 = @"""CSJS207""";  // 借方端数処理
            public const string kr55 = @"""CSJS208""";  // 借方取引先コード
            public const string kr06 = @"""CSJS213""";  // 借方本体金額

            public const string ks01 = @"""CSJS300""";  // 貸方部門コード
            public const string ks02 = @"""CSJS301""";  // 貸方勘定科目コード
            public const string ks52 = @"""CSJS302""";  // 貸方補助科目コード・・・固定値（0） 2019/09/27
            public const string ks53 = @"""CSJS304""";  // 貸方税率区分コード
            public const string ks03 = @"""CSJS305""";  // 貸方事業区分コード
            public const string ks55 = @"""CSJS306""";  // 貸方消費税計算
            public const string ks04 = @"""CSJS307""";  // 貸方端数処理
            public const string ks05 = @"""CSJS308""";  // 貸方取引先コード
            public const string ks06 = @"""CSJS313""";  // 貸方本体金額
            public const string ks54 = @"""CSJS320""";  // 貸方税率
            public const string ks56 = @"""CSJS322""";  // 税率種別 ・・・固定値（0：標準） 2019/09/27

            public const string tk01 = @"""CSJS100""";  // 摘要
            
        }


    }
}

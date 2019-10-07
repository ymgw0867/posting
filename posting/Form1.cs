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

            // ログイン
            frmLogin fl = new frmLogin();
            fl.ShowDialog();

            // ログイン未完了の場合は終了する
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

            //請求一覧
            frmSeikyuRep2015 frm = new frmSeikyuRep2015();
            frm.ShowDialog();
            this.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //請求書登録
            this.Hide();
            frmSeikyuAdd frm = new frmSeikyuAdd(1);
            frm.ShowDialog();
            this.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //請求一覧
            this.Hide();
            frmNyukinRep frm = new frmNyukinRep();
            frm.ShowDialog();
            this.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            //受注確定書登録
            Form frm = new frmOrder(0);

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            //ポスティングエリア登録
            frmPosting frm = new frmPosting();

            this.Hide();
            frm.ShowDialog();
            this.Show();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //ポスティングエリア登録
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
            if (MessageBox.Show("システムを終了します。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
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
            //他のマシンで配布指示実行中か検証する
            int cFlg = 0;

            OleDbDataReader dR;
            Control.会社情報 cSystem = new Control.会社情報();
            dR = cSystem.Fill();

            while (dR.Read())
            {
                cFlg = int.Parse(dR["配布フラグ"].ToString());
            }

            dR.Close();
            cSystem.Close();

            if (cFlg == 1)
            {
                MessageBox.Show("現在、他のマシンで配布指示登録中です。", "起動チェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //配布指示登録処理
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

            // 受注テーブルに「編集ロック」フィールドを追加する : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE 受注 add 編集ロック int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // 受注テーブルに「注文書受領済み」フィールドを追加する : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE 受注 add 注文書受領済み int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // ログインタイプヘッダテーブルに「受注個別ロック権限」フィールドを追加する : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE ログインタイプヘッダ add 受注個別ロック権限 int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // ログインタイプヘッダテーブルに「受注個別制限」フィールドを追加する : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE ログインタイプヘッダ add 受注個別制限 int default 0 NOT NULL";
            fSql.Execute(sqlString);

            // ログインタイプヘッダテーブルに「注文書受領済み権限」フィールドを追加する : 2019/10/03
            sqlString = "";
            sqlString += "ALTER TABLE ログインタイプヘッダ add 注文書受領済み権限 int default 0 NOT NULL";
            fSql.Execute(sqlString);

            fSql.Close();

            //// 会社情報テーブルに「受注確定書入力シートパス」フィールドを追加する : 2019/03/06
            //sqlString = "";
            //sqlString += "ALTER TABLE 会社情報 add 受注確定書入力シートパス nvarchar(255) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// 受注テーブルに「営業備考」フィールドを追加する : 2019/03/01
            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 営業備考 nvarchar(255) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// 得意先テーブルに「請求先・部署名」「請求先・敬称」フィールドを追加する : 2019/02/19
            //sqlString = "";
            //sqlString += "ALTER TABLE 得意先 add 請求先部署名 nvarchar(50) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 得意先 add 請求先敬称 nvarchar(5) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //fSql.Close();
            

            // 以下、コメント化 2019/02/19
            // 外注先マスターに「支払日」フィールドを追加する：2018/01/03
            //sqlString = "";
            //sqlString += "ALTER TABLE 外注先 add 支払日 int default 0 NOT NULL";

            //fSql.Execute(sqlString);

            //fSql.Close();

            // 以下のALTER TABLE SQL コメント化 2018/01/03
            //// 新請求書示テーブルに「口座」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 新請求書 add 口座 nvarchar(10) default '' NOT NULL";

            //fSql.Execute(sqlString);

            //fSql.Close();

            ///* 受注テーブルに「外注先ＩＤ支払２」「外注支払日支払２」「外注原価支払２」「外注先ＩＤ支払３」
            //「外注支払日支払３」「外注原価支払３」フィールドを追加する */

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注先ID支払2 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注支払日支払2 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注原価支払2 money default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注先ID支払3 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注支払日支払3 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注原価支払3 money default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //// 受注テーブルに「外注依頼日支払２」「外注依頼日支払３」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注依頼日支払2 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注依頼日支払3 datetime";
            //fSql.Execute(sqlString);

            //// 受注テーブルに「外注委託枚数２」「外注委託枚数３」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注委託枚数2 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注委託枚数3 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //// 受注テーブルに「外注渡し日２」「外注渡し日３」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注渡し日2 datetime";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注渡し日3 datetime";
            //fSql.Execute(sqlString);

            //// 受注テーブルに「外注受け渡し担当者２」「外注受け渡し担当者３」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注受け渡し担当者2 nvarchar(50) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注受け渡し担当者3 nvarchar(50) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// 外注支払テーブルに「調整額」「調整日付」「調整備考」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 外注支払 add 調整額 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 外注支払 add 調整日付 nvarchar(10) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 外注支払 add 調整備考 nvarchar(100) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// 受注テーブルに「外注支払ID2」「外注支払ID3」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注支払ID2 nvarchar(14) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 外注支払ID3 nvarchar(14) default '' NOT NULL";
            //fSql.Execute(sqlString);

            //// 受注テーブルに「納品書発行」フィールドを追加する
            //sqlString = "";
            //sqlString += "ALTER TABLE 受注 add 納品書発行 int default 0 NOT NULL";
            //fSql.Execute(sqlString);

            //// 受注テーブルにインデックスを追加する　2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_受注種別ID ";
            //sqlString += "ON 受注(受注種別ID) ";
            //sqlString += "INCLUDE(ID,得意先ID,チラシ名,枚数,税込金額)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_受注チラシ名 ";
            //sqlString += "ON 受注(チラシ名 ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_受注日 ";
            //sqlString += "ON 受注(受注日 ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_登録ユーザーID ";
            //sqlString += "ON 受注(登録ユーザーID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_配布形態 ";
            //sqlString += "ON 受注(配布形態 ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_得意先ID ";
            //sqlString += "ON 受注(得意先ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //// 配布エリアテーブルにインデックスを追加する　2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_配布指示ID ";
            //sqlString += "ON 配布エリア(配布指示ID) ";
            //sqlString += "INCLUDE(予定枚数,受注ID) ";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_完了区分配布指示ID ";
            //sqlString += "ON 配布エリア(完了区分,配布指示ID) ";
            //sqlString += "INCLUDE(予定枚数, 受注ID)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_配布エリア受注ID ";
            //sqlString += "ON 配布エリア(受注ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_ステータス ";
            //sqlString += "ON 配布エリア(ステータス ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_完了区分 ";
            //sqlString += "ON 配布エリア(完了区分 ASC) ";
            //sqlString += "INCLUDE(予定枚数, 受注ID) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //// 配布指示テーブルにインデックスを追加する　2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_配布指示 ";
            //sqlString += "ON 配布指示(配布員ID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_ユーザーID ";
            //sqlString += "ON 配布指示(ユーザーID ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);

            //// 得意先テーブルにインデックスを追加する　2016/11/02
            //sqlString = "";
            //sqlString += "CREATE NONCLUSTERED INDEX IX_担当社員コード ";
            //sqlString += "ON 得意先(担当社員コード ASC) ";
            //sqlString += "WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)";
            //fSql.Execute(sqlString);


            //////受注データの請求書IDを書き換える 2010/02/17
            //////ＤＡテクニカルサービス　受注ID：201001150004　請求書ID：822 → 726
            ////sqlString = "";
            ////sqlString += "update 受注 set 請求書ID = 726 ";
            ////sqlString += "where ID = 201001150004";

            ////fSql.Execute(sqlString);

            //////受注データの請求書IDを書き換える 2010/03/03
            //////HAKU　受注ID：200911040013　請求書ID：776 → 613
            ////sqlString = "";
            ////sqlString += "update 受注 set 請求書ID = 613 ";
            ////sqlString += "where ID = 200911040013";

            ////fSql.Execute(sqlString);

            //////受注データの請求書IDを書き換える 2010/03/03
            //////表参道接骨院　受注ID：200912210002　請求書ID：743 → 660
            ////sqlString = "";
            ////sqlString += "update 受注 set 請求書ID = 660 ";
            ////sqlString += "where ID = 200912210002";

            ////fSql.Execute(sqlString);

            //////受注データの請求書IDを書き換える 2010/03/04
            //////レコプロ　受注ID：201001150002　請求書ID：899 → 744 
            ////sqlString = "";
            ////sqlString += "update 受注 set 請求書ID = 744 ";
            ////sqlString += "where ID = 201001150002";

            ////fSql.Execute(sqlString);

            //////受注データの請求書IDを書き換える 2010/04/14
            //////レコプロ　受注ID：201003300007　請求書ID：0 → 1018 
            ////sqlString = "";
            ////sqlString += "update 受注 set 請求書ID = 1018 ";
            ////sqlString += "where (ID = 201003300007) and (請求書ID = 0)";

            ////fSql.Execute(sqlString);

            //////受注データの請求書IDを書き換える 2010/04/14
            //////レコプロ　受注ID：201002190001　請求書ID：1017 → 900 
            ////sqlString = "";
            ////sqlString += "update 受注 set 請求書ID = 900 ";
            ////sqlString += "where (ID = 201002190001) and (請求書ID = 1017)";

            ////fSql.Execute(sqlString);

            //fSql.Close();

            // メニュータイトルクラス 2015/07/07
            clsMenu cm = new clsMenu();

            // メニュータイトルCSVの読込 2015/07/07
            cm.loadMenu();

            // メニュータイトルをセット 2015/07/07
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

            // メニューボタン表示状態初期化
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

            // ログインユーザーごとのメニュー制御
            darwinDataSet dts = new darwinDataSet();
            darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter hAdp = new darwinDataSetTableAdapters.ログインタイプヘッダTableAdapter();
            darwinDataSetTableAdapters.ログインタイプタグTableAdapter tAdp = new darwinDataSetTableAdapters.ログインタイプタグTableAdapter();
            hAdp.Fill(dts.ログインタイプヘッダ);
            tAdp.Fill(dts.ログインタイプタグ);

            foreach (var h in dts.ログインタイプヘッダ.Where(a => a.Id == global.loginType))
            {
                foreach (var item in h.GetログインタイプタグRows())
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

            // ログイン中ユーザー
            //lblLogin.Text = "ログイン中ユーザー：" + global.loginUser;
            lblLogin.Text = global.loginUser + "さんがログイン中です";


            //// 自分自身のバージョン情報を取得する　2016/11/08
            //System.Diagnostics.FileVersionInfo ver =
            //    System.Diagnostics.FileVersionInfo.GetVersionInfo(
            //    System.Reflection.Assembly.GetExecutingAssembly().Location);

            // キャプションにバージョンを追加　2016/11/08
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
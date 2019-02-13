using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows;

namespace posting
{
    class clsSeikyuData
    {
        public clsSeikyuData(Form _frm)
        {
            frm = _frm;
        }

        Form frm;

        // データセット・テーブルアダプタ
        darwinDataSet dts = new darwinDataSet();
        darwinDataSetTableAdapters.受注1TableAdapter jAdp = new darwinDataSetTableAdapters.受注1TableAdapter();
        darwinDataSetTableAdapters.新請求書TableAdapter nAdp = new darwinDataSetTableAdapters.新請求書TableAdapter();

        ///-------------------------------------------------------
        /// <summary>
        ///     請求書データ集計 </summary>
        ///-------------------------------------------------------
        public void Summary(int sYear, int sMonth)
        {
            Cursor.Current = Cursors.WaitCursor;

            // オーナーフォームを無効にする
            frm.Enabled = false;

            // プログレスバーを表示する
            int rCnt = 0;
            frmPrg frmP = new frmPrg();
            frmP.Owner = frm;
            frmP.Show();

            // データ読み込み
            jAdp.Fill(dts.受注1);
            nAdp.Fill(dts.新請求書);

            // クライアント別,請求書発行日別,支払期日別で請求金額を集計
            var s = dts.受注1
                .Where(a => a.完了区分 == 0 && !a.Is請求書発行日Null() )
                    .GroupBy(a => a.得意先ID)
                    .Select(cg => new
                    {
                        cCode = cg.Key,
                        seDt = cg.GroupBy(b => b.請求書発行日)
                        .Select(seg => new
                        {
                            sss = seg.Key,
                            siDt = seg.GroupBy(b => b.入金予定日)
                            .Select(h => new
                            {
                                nnn = h.Key,
                                kingaku = h.Sum(a => a.金額),     // 単価×枚数（税抜）
                                nebiki = h.Sum(a => a.値引額),
                                tax = h.Sum(a => a.消費税),
                                cnt = h.Count()
                            })
                        })
                    });

            // 件数取得
            int cTotal = s.Count();
            //frmP.progressMax = cTotal;
            //frmP.setProgressMax();

            // 新請求書テーブル登録
            foreach (var item in s)
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "請求データ作成中..." + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                //frmP.progressValue = rCnt;
                frmP.ProgressStep();
                
                // クライアント
                int clientCode = item.cCode;

                foreach (var k in item.seDt)
                {
                    // 請求書発行日が指定年月外のとき、ネグる
                    if (k.sss.Year != sYear || k.sss.Month != sMonth)
                    {
                        continue;
                    }

                    // 請求書発行日
                    DateTime seikyuDt = k.sss;

                    foreach (var j in k.siDt)
                    {
                        // 既に新請求書データが登録されているか得意先ID,請求書発行日,支払期日（入金予定日）で検索する
                        if (dts.新請求書.Any(a => a.得意先ID == clientCode && a.請求書発行日 == seikyuDt && a.支払期日 == j.nnn))
                        {
                            // 登録済みのため更新処理
                            darwinDataSet.新請求書Row r = dts.新請求書.Single(a => a.得意先ID == clientCode && a.請求書発行日 == seikyuDt && a.支払期日 == j.nnn);

                            // 請求金額
                            decimal kin = j.kingaku - j.nebiki + j.tax;     

                            // 金額もしくは値引額が変更になっているとき新請求書データを更新します
                            if (j.kingaku != r.売上金額 || j.nebiki != r.値引額 || kin != r.請求金額)
                            {
                                // 入金額計算
                                int nkin = r.請求金額 - r.残金;

                                //decimal kin = j.kingaku - j.nebiki + j.tax;     // 請求金額
                                r.請求金額 = (int)kin;
                                r.消費税 = (int)j.tax;
                                r.値引額 = (int)j.nebiki;
                                r.売上金額 = (int)j.kingaku;
                                r.明細数 = j.cnt;
                                r.残金 = (int)(kin - nkin);   // 再計算
                                r.変更年月日 = DateTime.Now; 
                                r.ユーザーID = global.loginUserID;
                            }
                        }
                        else
                        {
                            // 新請求書データ新規登録
                            decimal kin = j.kingaku - j.nebiki + j.tax; // 請求金額
                            darwinDataSet.新請求書Row r = dts.新請求書.New新請求書Row();
                            r.得意先ID = clientCode;
                            r.請求金額 = (int)kin;
                            r.消費税 = (int)j.tax;
                            r.値引額 = (int)j.nebiki;
                            r.売上金額 = (int)j.kingaku;
                            r.税率 = 0;
                            r.請求書発行日 = seikyuDt;
                            r.支払期日 = j.nnn;
                            r.残金 = (int)kin;
                            r.入金完了 = 0;
                            r.請求書発行済 = 0;
                            r.明細数 = j.cnt;
                            r.備考 = string.Empty;
                            r.登録年月日 = DateTime.Now;
                            r.変更年月日 = DateTime.Now;
                            r.ユーザーID = global.loginUserID;
                            r.無効 = global.FLGOFF;
                            r.精算備考 = string.Empty;
                            r.精算額 = 0;
                            r.精算日付 = string.Empty;
                            dts.新請求書.Add新請求書Row(r);
                        }
                    }
                }
            }

            //MessageBox.Show("!!");
            Application.DoEvents();

            // 100ミリ秒遅らせる
            System.Threading.Thread.Sleep(100);

            // データベース更新
            nAdp.Update(dts.新請求書);
            nAdp.Fill(dts.新請求書);

            rCnt = 0;

            // 新請求データ取得
            var ss = dts.新請求書.Where(a => a.入金完了 == global.FLGOFF);
            cTotal = ss.Count();
            //frmP.progressMax = cTotal;
            //frmP.setProgressMax();

            // 受注データと新請求データの紐付
            foreach (var t in ss)
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "受注データと請求データを最適化中 Step.1 ..." + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                //frmP.progressValue = rCnt;
                frmP.ProgressStep();

                //// 1ミリ秒停止する
                //System.Threading.Thread.Sleep(1);

                foreach (var d in dts.受注1.Where(a => a.完了区分 == global.FLGOFF && !a.Is請求書発行日Null() && 
                                     !a.Is入金予定日Null() && a.得意先ID == t.得意先ID &&
                                     a.請求書発行日 == t.請求書発行日 && a.入金予定日 == t.支払期日))
                {
                    d.請求書ID = t.ID;
                }
            }

            //MessageBox.Show("!!");
            Application.DoEvents();

            // 100ミリ秒遅らせる
            System.Threading.Thread.Sleep(100);

            rCnt = 0;

            // 締め処理後新請求書データに紐付されない受注データの請求書IDを初期化(0)にする
            foreach (var t in ss)
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "受注データと請求データを最適化中 Step.2 ..." + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                foreach (var m in t.Get受注1Rows())
                {
                    if (m.Is請求書発行日Null())
                    {
                        m.請求書ID = 0;
                    }
                    else if (m.Is入金予定日Null())
                    {
                        m.請求書ID = 0;
                    }
                    else if (t.得意先ID != m.得意先ID || t.請求書発行日 != m.請求書発行日 || 
                             t.支払期日 != m.入金予定日)
                    {
                        m.請求書ID = 0;
                    }
                }
            }

            Application.DoEvents();

            // 100ミリ秒遅らせる
            System.Threading.Thread.Sleep(100);


            rCnt = 0;

            // 締め処理後受注データに紐付されない新請求書データを削除にする
            foreach (var t in ss)
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "受注データと請求データを最適化中 Step.3 ..." + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();
                
                // 締め処理後受注データに紐付されない新請求書データを削除にする 2015/12/19
                if (t.Get受注1Rows().Count() == 0)
                {
                    t.Delete();
                }
            }

            Application.DoEvents();

            // データベース更新
            jAdp.Update(dts.受注1);
            nAdp.Update(dts.新請求書);
            nAdp.Fill(dts.新請求書);

            Application.DoEvents();

            // いったんオーナーをアクティブにする
            frm.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            frm.Enabled = true;

            Cursor.Current = Cursors.Default; 
        }
    }
}

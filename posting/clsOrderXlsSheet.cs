using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Windows.Forms;

namespace posting
{
    ///----------------------------------------------------------------------
    /// <summary>
    ///     受注確定書入力シートへデータを書き込む：2019/03/06　</summary>
    ///----------------------------------------------------------------------
    public partial class frmOrderExcel
    {
        private void setData2Sheet()
        {
            XLWorkbook bk = null;

            try
            {
                using (bk = new XLWorkbook(sheetPath, XLEventTracking.Disabled))
                {
                    var sheet1 = bk.Worksheet(1);
                    //var tbl = sheet1.RangeUsed().AsTable();
                    //var rCnt = sheet1.RowsUsed().Count();

                    IXLRow row = null;

                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        if (sheet1.RowsUsed().Count() > 2)
                        {
                            // 前回最終行下罫線を引き直す
                            row = sheet1.Row(sheet1.RowsUsed().Count());

                            sheet1.Range(row.Cell(1), row.Cell(sheet1.ColumnsUsed().Count())).Style
                                .Border.SetBottomBorder(XLBorderStyleValues.Hair);
                        }

                        row = sheet1.Row(sheet1.RowsUsed().Count() + 1);

                        // 追加行罫線を引く
                        sheet1.Range(row.Cell(1), row.Cell(sheet1.ColumnsUsed().Count())).Style
                            //.Border.SetTopBorder(XLBorderStyleValues.Thin)
                            //.Border.SetBottomBorder(XLBorderStyleValues.Thin)
                            .Border.SetBottomBorder(XLBorderStyleValues.Hair)
                            .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                            .Border.SetRightBorder(XLBorderStyleValues.Thin);

                        // 追加行書式設定
                        for (int ic = 1; ic <= sheet1.ColumnsUsed().Count(); ic++)
                        {
                            row.Cell(ic).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)
                                .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                                .Font.SetFontName("游ゴシック")
                                .Font.SetFontSize(11);

                            if (ic == 1 || ic == 8 || ic == 34)
                            {
                                row.Cell(ic).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            }

                            if (ic == 2 || ic == 3 || ic == 4 || ic == 32 || ic == 33)
                            {
                                row.Cell(ic).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                            }

                            if (ic == 5 || ic == 6 || (ic >= 9 && ic <= 30) || ic == 35)
                            {
                                row.Cell(ic).Style.NumberFormat.SetFormat("#,##0");
                            }

                            if (ic == 7)
                            {
                                // 粗利率計算：数式
                                string formula = "=F" + row.RowNumber() + "/E" + row.RowNumber();
                                row.Cell(ic).FormulaA1 = formula;

                                row.Cell(ic).Style.NumberFormat.SetFormat("#,##.0%");
                            }

                            // 罫線
                            row.Cell(1).Style.Border.LeftBorder = XLBorderStyleValues.Medium;

                            if (ic == 8 || ic == 30)
                            {
                                row.Cell(ic).Style.Border.RightBorder = XLBorderStyleValues.MediumDashDotDot;
                            }

                            if (ic == 9 || ic == 11 || ic == 13 || ic == 15 || ic == 17 || ic == 19 || ic == 21 || ic == 23 ||
                                ic == 25 || ic == 27 || ic == 29)
                            {
                                row.Cell(ic).Style.Border.RightBorder = XLBorderStyleValues.Hair;
                                row.Cell(ic + 1).Style.Border.LeftBorder = XLBorderStyleValues.None;
                            }

                            if (ic == 32 || ic == 35)
                            {
                                row.Cell(ic).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                            }
                        }

                        // 売上
                        decimal uri = Utility.strToInt(Utility.nullToStr(dataGridView2[colKingaku, i].Value)) - Utility.strToInt(Utility.nullToStr(dataGridView2[colNebiki, i].Value));

                        // 原価
                        decimal genka = Utility.strToInt(Utility.nullToStr(dataGridView2[colGenka2, i].Value)) +
                                    Utility.strToInt(Utility.nullToStr(dataGridView2[colGenka3, i].Value)) +
                                    Utility.strToInt(Utility.nullToStr(dataGridView2[colGenka4, i].Value));

                        // 粗利
                        decimal arari = uri - genka;

                        row.Cell(1).Value = DateTime.Today.ToShortDateString();
                        row.Cell(2).Value = txtEiTantou.Text;
                        row.Cell(3).Value = txtClient.Text;
                        row.Cell(4).Value = Utility.nullToStr(dataGridView2[colNaiyou, i].Value);
                        row.Cell(5).Value = uri;   // 税抜合計
                        row.Cell(6).Value = arari;     // 粗利

                        //decimal arariRt = arari / uri;     // 粗利率
                        //row.Cell(7).Value = arariRt.ToString("0.0%");
                        //row.Cell(7).Value = arariRt;

                        // 注文書回収フラグ
                        if (chkKaishuFlg.Checked)
                        {
                            row.Cell(8).Value = "○";
                        }
                        else
                        {
                            row.Cell(8).Value = "";
                        }

                        // 受注種別毎
                        switch (dataGridView2[colCmb, i].FormattedValue.ToString())
                        {
                            case "ポスティング":

                                if (Utility.nullToStr(dataGridView2[colGaichu2, i].Value).Contains("自社"))
                                {
                                    // 自社POS
                                    row.Cell(9).Value = uri;
                                    row.Cell(10).Value = uri - genka;
                                }
                                else
                                {
                                    // 他社POS
                                    row.Cell(11).Value = uri;
                                    row.Cell(12).Value = uri - genka;
                                }
                                break;

                            case "印刷":
                                row.Cell(13).Value = uri;
                                row.Cell(14).Value = uri - genka;
                                break;

                            case "折代":
                                row.Cell(15).Value = uri;
                                row.Cell(16).Value = uri - genka;
                                break;

                            case "デザイン":
                                row.Cell(17).Value = uri;
                                row.Cell(18).Value = uri - genka;
                                break;

                            case "新聞折込":
                                row.Cell(19).Value = uri;
                                row.Cell(20).Value = uri - genka;
                                break;

                            case "リスティング広告":
                                row.Cell(21).Value = uri;
                                row.Cell(22).Value = uri - genka;
                                break;

                            case "内職":
                                row.Cell(23).Value = uri;
                                row.Cell(24).Value = uri - genka;
                                break;

                            case "その他":
                                row.Cell(25).Value = uri;
                                row.Cell(26).Value = uri - genka;
                                break;

                            case "web":
                                row.Cell(27).Value = uri;
                                row.Cell(28).Value = uri - genka;
                                break;

                            default:
                                row.Cell(29).Value = uri;
                                row.Cell(30).Value = uri - genka;
                                break;
                        }

                        // 支払区分


                        // 新・既
                        row.Cell(32).Value = cmbNewRep.Text;

                        // 実施日
                        row.Cell(33).Value = dataGridView2[colDtHaifuS, i].Value.ToString();

                        // サイズ
                        row.Cell(34).Value = dataGridView2[colSize, i].Value.ToString();

                        // 枚数
                        row.Cell(35).Value = Utility.strToInt(dataGridView2[colMaisu, i].Value.ToString().Replace(",", ""));
                    }

                    if (row != null)
                    {
                        // 最終行下罫線を引く
                        sheet1.Range(row.Cell(1), row.Cell(sheet1.ColumnsUsed().Count())).Style
                            .Border.SetBottomBorder(XLBorderStyleValues.Medium);
                    }

                    // 更新
                    bk.SaveAs(sheetPath);
                }
            }
            catch (Exception ex)
            {
                // エラーメッセージ 2019/06/19
                MessageBox.Show(ex.Message + Environment.NewLine + "受注確定書入力シートへの書き込みに失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                // 2019/06/19
                try
                {
                    if (bk != null)
                    {
                        bk.SaveAs(sheetPath);

                        // 後片付け
                        bk.Dispose();
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }
}

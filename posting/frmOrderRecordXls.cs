using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace posting
{
    public partial class frmOrderRecordXls : Form
    {
        public frmOrderRecordXls(string fileName)
        {
            InitializeComponent();

            _fileName = fileName;

            oXls = new Excel.Application();
            oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(_fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            // シート数を取得
            sCnt = oXlsBook.Sheets.Count;
            oxlsSheet = new Excel.Worksheet[sCnt];
            
            for (int i = 0; i < sCnt; i++)
			{
                oxlsSheet[i] = new Excel.Worksheet();
                oxlsSheet[i] = (Excel.Worksheet)oXlsBook.Sheets[i + 1];
			}
        }

        string _fileName = string.Empty;

        // Excelブック、シート
        Excel.Application oXls = null;
        Excel.Workbook oXlsBook = null;
        Excel.Worksheet [] oxlsSheet = null;
        Excel.Range dRng;
 
        int sCnt = 0;   // シート数
        int nPage = 1;  // 現在のページ

        private void frmOrderRecordXls_Load(object sender, EventArgs e)
        {
            // 表示
            xlsSheetShow(nPage - 1);

            this.Text = System.IO.Path.GetFileName(_fileName);
        }
        
        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     受注確定書シートフォームに表示 </summary>
        /// <param name="sNum">
        ///     インデックス</param>
        /// -------------------------------------------------------------------------------
        private void xlsSheetShow(int sNum)
        {
            Excel.Range stRange = (Excel.Range)oxlsSheet[sNum].Cells[1, 1];   // 開始セル
            Excel.Range edRange = (Excel.Range)oxlsSheet[sNum].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);   // 最終セル

            Excel.Range dRng;

            // シートのレンジを取得
            dRng = oxlsSheet[sNum].Cells.get_Range(stRange, edRange);

            //コピー
            dRng.Copy(Type.Missing);

            //クリップボードからデータ取得し画面表示
            Image iData = Clipboard.GetImage();
            pictureBox1.Image = iData;

            //保存処理
            oXls.DisplayAlerts = false;

            // 移動ボタン
            if (sCnt == 1)
            {
                button1.Enabled = false;
                button2.Enabled = false;
            }
            else
            {
                if (nPage == sCnt)
                {
                    button2.Enabled = false;
                }
                else
                {
                    button2.Enabled = true;
                }

                if (nPage == 1)
                {
                    button1.Enabled = false;
                }
                else
                {
                    button1.Enabled = true;
                }
            }

            // ページ表示
            label1.Text = nPage.ToString() + "／" + sCnt.ToString() + "シート";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (nPage > 1)
            {
                nPage--;
                xlsSheetShow(nPage - 1);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (nPage < sCnt)
            {
                nPage++;
                xlsSheetShow(nPage - 1);
            }
        }

        private void frmOrderRecordXls_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Bookをクローズ
            oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

            //Excelを終了
            oXls.Quit();

            // COM オブジェクトの参照カウントを解放する 
            for (int i = 0; i < oxlsSheet.Length; i++)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet[i]);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}

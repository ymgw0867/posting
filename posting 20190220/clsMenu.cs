using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace posting
{
    class clsMenu
    {
        public string[] menuCsv;    // メニュータイトル配列

        public void loadMenu()
        {
            // メニュータイトル読み込み
            string sPath = System.IO.Directory.GetCurrentDirectory() + @"\" + Properties.Settings.Default.menuCsv;

            if (System.IO.File.Exists(sPath))
            {
                menuCsv = System.IO.File.ReadAllLines(sPath, Encoding.Default);
            }
        }

    }
}

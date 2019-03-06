using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace posting
{
    class global
    {
        // ログインステータス
        public static bool loginStatus;

        // ログイン情報
        public static int loginUserID;
        public static string loginUser;
        public static int loginType;
        public static int loginOrderMntType;

        // フラグ
        public const int FLGON = 1;
        public const int FLGOFF = 0;

        // 受注編集ロックテーブルキー
        public const int lockKey = 1;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace MNBS.Common
{
    class global
    {
        // 表示画像名
        public static string pblImageFile;
        // MDBファイル名
        public const string MDBFILENAME = "MBS.mdb";
        // MDBバックアップファイル名
        public const string MDBBACKUP = "MBS_Back.mdb";
        // MDB一時ファイル名
        public const string MDBTEMPFILE = "MBS_temp.mdb";

        // ＯＣＲ処理スキャナ、ＦＡＸ指定
        public const int SCAN_SELECT = 0;
        public const int FAX_SELECT = 1;
        public const int END_SELECT = 9;

        // ＰＣ指定
        public const int PC1SELECT = 1;
        public const int PC2SELECT = 2;

        public const int STAFF_SELECT = 1;
        public const int PART_SELECT = 2;


        // 環境設定データ
        public const int sIDKEY = 1;
        public static int sYear = 0;
        public static int rekiYear = 0;
        public static int sMonth = 0;
        public static string sMSTJ = string.Empty;      // パートマスターパス
        public static string sMSTS = string.Empty;      // スタッフマスターパス
        public static string sTIF = string.Empty;       // スタッフ画像バックアップ作成先パス
        public static string sTIF2 = string.Empty;      // パート画像バックアップ作成先パス
        public static string sDAT = string.Empty;       // スタッフ受け渡しデータ作成先パス
        public static string sDAT2 = string.Empty;      // パート受け渡しデータ作成先パス
        public static int sBKDELS = 0;                  // スタッフバックアップデータ保存月数
        public static int sBKDELP = 0;                  // パートバックアップデータ保存月数

        // フラグオン・オフ
        public const string FLGON = "1";
        public const string FLGOFF = "0";

        // ○印
        public const string MARU = "○";

        //エラーチェック関連
        public static int errID;            //エラーデータID
        public static int errNumber;        //エラー項目番号
        public static int errRow;           //エラー行
        public static string errMsg;        //エラーメッセージ

        //エラー項目番号
        public const int eNothing = 0;      // エラーなし
        public const int eYear = 1;         // 対象年
        public const int eMonth = 2;        // 対象月
        public const int eShainNo = 3;      // 個人番号
        public const int eOrderCode = 4;    // オーダーコード
        //public const int eRiseki = 5;       // 離席時間
        public const int eKyukei = 6;       // 休憩
        public const int eKyuka = 7;        // 休暇
        //public const int eHiru1 = 8;        // 昼１
        //public const int eHiru2 = 9;        // 昼２
        public const int eSH = 10;          // 開始時
        public const int eSM = 11;          // 開始分
        public const int eEH = 12;          // 終了時
        public const int eEM = 13;          // 終了分
        //public const int eMark = 14;        // マーク
        public const int eDays = 15;        // 出勤日数
        //public const int eLunch = 16;       // 昼食回数
        public const int eTeisei = 17;      // 訂正

        //表示関係
        public static float miMdlZoomRate = 0;      //現在の表示倍率

        //表示倍率（%）
        public static float ZOOM_RATE_FAX = 0.25f;  // 標準倍率：ＦＡＸ画像用
        public static float ZOOM_RATE = 0.165f;     // 標準倍率：スキャン画像用
        public static float ZOOM_MAX = 2.00f;       // 最大倍率
        public static float ZOOM_MIN = 0.05f;       // 最小倍率
        public static float ZOOM_STEP = 0.02f;      // ステップ倍率
        public static float ZOOM_Total = 0f;        // 現在の倍率
        public static float ScaleFactor_Fax = 0f;   // ＦＡＸ画像用スケールファクタ
        public static float ScaleFactor_Scan = 0f;  // スキャン画像用スケールファクタ

        // 表示色関連
        public static Color lBackColorE;
        public static Color lBackColorN;

        // DataGridChangeValueイベント発生制御
        public static bool dg1ChabgeValueStatus;

        // 所定時間情報
        public static string ShoS = string.Empty;
        public static string ShoE = string.Empty;
        public static string AmS = string.Empty;
        public static string AmE = string.Empty;
        public static string PmS = string.Empty;
        public static string PmE = string.Empty;

        // 休暇区分
        public static string eYUKYU = "1";
        public static string eKEKKIN = "2";
        public static string eFURIDE = "3";
        public static string eFURIKYU = "4";

        // 休日区分
        public static int hWEEKDAY = 0;
        public static int hSATURDAY = 1;
        public static int hHOLIDAY = 2;

        // 勤務データ登録区分
        public const int sADDMODE = 1;
        public const int sEDITMODE = 0;

        //時間単位有休休暇関連
        public const double YUKYU0125 = 125;
        public const double YUKYU10 = 10;

        public const string SheetID = "1";

        //datagridview表示行数
        public static int _MULTIGYO = 31;

    }
}

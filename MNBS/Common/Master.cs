using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace MNBS.Common
{
    public class Master
    {
        public OleDbCommand sCom = new OleDbCommand();
        protected StringBuilder sb = new StringBuilder();

        public Master()
        {
            // データベース接続文字列
            OleDbConnection Cn = new OleDbConnection();
            sb.Clear();
            sb.Append("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
            sb.Append(Properties.Settings.Default.PathInst);
            sb.Append(Properties.Settings.Default.PathMDB);
            sb.Append(global.MDBFILENAME);
            Cn.ConnectionString = sb.ToString();
            Cn.Open();

            sCom.Connection = Cn;
        }
    }

    public class msConfig : Master
    {
        /// <summary>
        /// 環境設定マスター新規登録
        /// </summary>
        /// <param name="sID">キー</param>
        /// <param name="sMSTS">マスターファイルパス</param>
        /// <param name="sTIF">受け渡しデータ作成先パス</param>
        /// <param name="sDAT">ＯＣＲデータバックアップパス</param>
        /// <param name="sBKDELS">バックアップデータ保存月数</param>
        /// <param name="sSCAN">使用スキャナー</param>
        /// <param name="sSYEAR">年</param>
        /// <param name="sSMONTH">月</param>
        /// <param name="sDate">更新年月日</param>
        public void Insert(int sID, string sMSTS, string sTIF, string sDAT, int sBKDELS, 
                                    string sSCAN, string sSYEAR, string sSMONTH, DateTime sDate)
        {
            try
            {
                sb.Clear();
                sb.Append("insert into 環境設定 (");
                sb.Append("ID,MSTS,TIF,DAT,BKDELS,SCAN,SYEAR,SMONTH,更新年月日) values (");
                sb.Append("?,?,?,?,?,?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID);
                sCom.Parameters.AddWithValue("@msts", sMSTS);
                sCom.Parameters.AddWithValue("@tif", sTIF);
                sCom.Parameters.AddWithValue("@dat", sDAT);
                sCom.Parameters.AddWithValue("@bkdels", sBKDELS);
                sCom.Parameters.AddWithValue("@scan", sSCAN);
                sCom.Parameters.AddWithValue("@year", sSYEAR);
                sCom.Parameters.AddWithValue("@Month", sSMONTH);
                sCom.Parameters.AddWithValue("@update", sDate);
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }            
        }

        /// <summary>
        /// 環境設定マスター更新
        /// </summary>
        /// <param name="sMSTS">マスターファイルパス</param>
        /// <param name="sTIF">受け渡しデータ作成先パス</param>
        /// <param name="sDAT">ＯＣＲデータバックアップパス</param>
        /// <param name="sBKDELS">バックアップデータ保存月数</param>
        /// <param name="sSCAN">使用スキャナー</param>
        /// <param name="sSYEAR">年</param>
        /// <param name="sSMONTH">月</param>
        /// <param name="sDate">更新年月日</param>
        public void UpDate(string sMSTS, string sTIF, string sDAT, int sBKDELS, string sSCAN, 
                           string sSYEAR, string sSMONTH, DateTime sDate)
        {
            try
            {
                sb.Clear();
                sb.Append("update 環境設定 set ");
                sb.Append("MSTS=?,TIF=?,DAT=?,BKDELS=?,SCAN=?,SYEAR=?,SMONTH=?,更新年月日=?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@msts", sMSTS);
                sCom.Parameters.AddWithValue("@tif", sTIF);
                sCom.Parameters.AddWithValue("@dat", sDAT);
                sCom.Parameters.AddWithValue("@bkdels", sBKDELS);
                sCom.Parameters.AddWithValue("@scan", sSCAN);
                sCom.Parameters.AddWithValue("@year", sSYEAR);
                sCom.Parameters.AddWithValue("@Month", sSMONTH);
                sCom.Parameters.AddWithValue("@update", sDate);
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }


        /// <summary>
        /// 環境設定マスター取得
        /// </summary>
        /// <param name="sMSTS">マスターファイルパス</param>
        /// <param name="sTIF">受け渡しデータ作成先パス</param>
        /// <param name="sDAT">ＯＣＲデータバックアップパス</param>
        /// <param name="sBKDELS">バックアップデータ保存月数</param>
        /// <param name="sSCAN">使用スキャナー</param>
        /// <param name="sSYEAR">年</param>
        /// <param name="sSMONTH">月</param>
        /// <param name="sDate">更新年月日</param>
        public OleDbDataReader Select(int sID)
        {
            OleDbDataReader dr = null;

            try
            {
                sb.Clear();
                sb.Append("select * from 環境設定 ");
                sb.Append("where ID = ?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID);
                dr = sCom.ExecuteReader();
                return dr;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return dr;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 環境設定データを取得する
        /// </summary>
        public void GetCommonYearMonth()
        {
            OleDbDataReader dr = Select(global.sIDKEY);

            try
            {
                while (dr.Read())
                {
                    global.sYear = int.Parse(dr["SYEAR"].ToString());
                    global.sMonth = int.Parse(dr["SMONTH"].ToString());
                    global.sMSTS = dr["MSTS"].ToString();
                    global.sTIF = dr["TIF"].ToString();
                    global.sDAT = dr["DAT"].ToString();
                    global.sBKDELS = int.Parse(dr["BKDELS"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定年月取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (dr.IsClosed == false) dr.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }
    }

    public class msShain
    {
        public string _OrderCode { get; set; }  // オーダーコード
        public string _HaCode { get; set; }     // 派遣先コード
        public string _HaName { get; set; }     // 派遣先名
        public string _BuName { get; set; }     // 派遣先部署名
        public string _StaffCode { get; set; }  // スタッフコード
        public string _StaffName { get; set; }  // スタッフ名
        public string _STime { get; set; }      // 開始時間
        public string _ETime { get; set; }      // 終了時間
        public string _Kyukei { get; set; }     // 休憩時間
    }
}

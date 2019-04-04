using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace MNBS.Common
{
    class OCRData : Master
    {
        // ＯＣＲデータヘッダ
        public string OCRDATA_ID { get; set; }
        public string _SheetID { get; set; }
        public string _StaffCode { get; set; }
        public string _Year { get; set; }
        public string _Month { get; set; }
        public string _OrderCode { get; set; }
        public string _ShuDays { get; set; }
        public string _ImageName { get; set; }

        public class OCR_Item
        {
            public string _Day { get; set; }
            public string _Kyuka { get; set; }
            public string _Sh { get; set; }
            public string _Sm { get; set; }
            public string _eh { get; set; }
            public string _em { get; set; }
            public string _Kyukei { get; set; }
            public string _teisei { get; set; }
        }

        public OCR_Item[] itm = new OCR_Item[global._MULTIGYO];

        public OCRData()
        {
            // ヘッダ情報初期化
            OCRDATA_ID = string.Empty;
            _SheetID = string.Empty;
            _StaffCode = string.Empty;
            _Year = string.Empty;
            _Month = string.Empty;
            _OrderCode = string.Empty;
            _ShuDays = string.Empty;
            _ImageName = string.Empty;

            // 明細情報初期化
            for (int i = 0; i < global._MULTIGYO; i++)
            {
                itm[i] = new OCR_Item(); 
            }
        }

        /// <summary>
        /// 勤務票ヘッダデータ新規登録
        /// </summary>
        /// <param name="sID">ID</param>
        /// <param name="sSheID">シートＩＤ</param>
        /// <param name="sSNo">スタッフコード</param>
        /// <param name="sYear">年</param>
        /// <param name="sMonth">月</param>
        /// <param name="sOrderCode">オーダーコード</param>
        /// <param name="sShuDays">出勤日数</param>
        /// <param name="sImgName">画像名</param>
        public void HeadInsert(string sID, string sSheID, string sSNo, string sYear, string sMonth,
                           string sOrderCode, string sShuDays, string sImgName)
        {
            try
            {
                // 勤務票ヘッダ
                sb.Clear();
                sb.Append("insert into 勤務票ヘッダ ");
                sb.Append("(ID,シートID,個人番号,年,月,オーダーコード,");
                sb.Append("出勤日数合計,画像名,更新年月日) ");
                sb.Append("values (?,?,?,?,?,?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();

                sCom.Parameters.AddWithValue("@ID", sID);   // ID                
                sCom.Parameters.AddWithValue("@ID", sSheID);    // シートＩＤ                
                sCom.Parameters.AddWithValue("@kjn", sSNo); // 個人番号                
                sCom.Parameters.AddWithValue("@year", sYear);   // 年                
                sCom.Parameters.AddWithValue("@month", sMonth); // 月                
                sCom.Parameters.AddWithValue("@order", sOrderCode); // オーダーコード                
                sCom.Parameters.AddWithValue("@ShukinTL", sShuDays);    // 出勤日数合計                
                sCom.Parameters.AddWithValue("@IMG", sImgName); // 画像名                
                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());  // 更新年月日

                // テーブル書き込み
                sCom.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票ヘッダ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 勤務票ヘッダデータ更新
        /// </summary>
        /// <param name="sID">ID</param>
        /// <param name="sSheID">シートＩＤ</param>
        /// <param name="sSNo">スタッフコード</param>
        /// <param name="sYear">年</param>
        /// <param name="sMonth">月</param>
        /// <param name="sOrderCode">オーダーコード</param>
        /// <param name="sShuDays">出勤日数</param>       
        public void HeadUpdate(string sID, string sSheID, string sSNo, string sYear, string sMonth,
                           string sOrderCode, string sShuDays)
        {
            try
            {
                // 勤務票ヘッダ
                sb.Clear();
                sb.Append("update 勤務票ヘッダ set ");
                sb.Append("シートID=?,個人番号=?,年=?,月=?,オーダーコード=?,");
                sb.Append("出勤日数合計=?,更新年月日=? ");
                sb.Append("where ID=?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
           
                sCom.Parameters.AddWithValue("@sID", sSheID);   // シートＩＤ                
                sCom.Parameters.AddWithValue("@kjn", sSNo);     // 個人番号                
                sCom.Parameters.AddWithValue("@year", sYear);   // 年                
                sCom.Parameters.AddWithValue("@month", sMonth); // 月                
                sCom.Parameters.AddWithValue("@order", sOrderCode);     // オーダーコード                
                sCom.Parameters.AddWithValue("@ShukinTL", sShuDays);    // 出勤日数合計          
                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());  // 更新年月日
                sCom.Parameters.AddWithValue("@ID", sID);   // ID

                // テーブル書き込み
                sCom.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票ヘッダ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 勤務票データを削除する（ヘッダ、明細両方）
        /// </summary>
        /// <param name="_year">年</param>
        /// <param name="_month">月</param>
        /// <param name="_StaffCode">個人番号</param>
        public void DataDelete(string _year, string _month, string _StaffCode)
        {
            string sID = string.Empty;

            // 画像ファイル名を取得
            OleDbDataReader dR = HeaderSelect(_year, _month, _StaffCode);
            string sImgNm = string.Empty;

            while (dR.Read())
            {
                sID = dR["ID"].ToString();
            }
            dR.Close();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                // 勤務票ヘッダ
                sCom.CommandText = "delete from 勤務票ヘッダ where ID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID);
                sCom.ExecuteNonQuery();

                //勤務票明細データを削除します
                sCom.CommandText = "delete from 勤務票明細 where ヘッダID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID);
                sCom.ExecuteNonQuery();

                ////////画像ファイルを削除する
                //////if (System.IO.File.Exists(_InPath + sImgNm))
                //////{
                //////    System.IO.File.Delete(_InPath + sImgNm);
                //////}

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票データ削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
                // トランザクションロールバック
                sTran.Rollback(); 
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 基準年月以前の勤務票データを削除する
        /// </summary>
        /// <param name="sID">ヘッダID</param>
        /// <param name="_InPath">画像ファイルフォルダパス</param>
        public void DataDelete(int dYYMM)
        {
            string sID = string.Empty;

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                // 画像ファイル名を取得
                OleDbDataReader dR = HeaderSelect(dYYMM);
                string sImgNm = string.Empty;

                string[] sKey = new string[1];

                int iX = 0;

                while (dR.Read())
                {
                    // 配列作成
                    if (iX > 0) 
                    {
                        Array.Resize(ref sKey, iX + 1);
                    }

                    sKey[iX] = dR["ID"].ToString();

                    iX++;
                }
                dR.Close();

                for (int i = 0; i < sKey.Length; i++)
                {
                    if (sKey[i] != null)
                    {
                        //勤務票明細データを削除します
                        sCom.CommandText = "delete from 勤務票明細 where ヘッダID = ?";
                        sCom.Parameters.Clear();
                        sCom.Parameters.AddWithValue("@ID", sKey[i]);
                        sCom.ExecuteNonQuery();
                    }
                }

                // 勤務票ヘッダ
                sCom.CommandText = "delete from 勤務票ヘッダ where 年*100+月 <= ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@YearMonth", dYYMM);
                sCom.ExecuteNonQuery();

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票ＭＤＢデータ削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                // トランザクションロールバック
                sTran.Rollback();
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 勤務表明細新規登録
        /// </summary>
        /// <param name="shID">ヘッダＩＤ</param>
        /// <param name="sDay">日付</param>
        /// <param name="sSh">開始時</param>
        /// <param name="sSm">開始分</param>
        /// <param name="sEh">終了時</param>
        /// <param name="sEm">終了分</param>
        /// <param name="sHoliday">休暇</param>
        /// <param name="sKyukei">休憩時間</param>
        /// <param name="steisei">訂正</param>
        public void ItemInsert(string shID, string sDay, string sSh, string sSm, string sEh, string sEm,
                           string sHoliday, string sKyukei, string steisei)
        {
            try
            {
                // 勤務票明細
                sb.Clear();
                sb.Append("insert into 勤務票明細 ");
                sb.Append("(ヘッダID,日付,開始時,開始分,終了時,終了分,休暇,休憩時間,訂正,更新年月日) ");
                sb.Append("values (?,?,?,?,?,?,?,?,?,?)");
                sCom.CommandText = sb.ToString();

                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", shID);      // ヘッダID
                sCom.Parameters.AddWithValue("@Days", sDay);    // 日付 
                sCom.Parameters.AddWithValue("@sh", sSh);       // 開始・時間
                sCom.Parameters.AddWithValue("@sm", sSm);       // 開始・分
                sCom.Parameters.AddWithValue("@eh", sEh);       // 終了・時間
                sCom.Parameters.AddWithValue("@em", sEm);       // 終了・分
                sCom.Parameters.AddWithValue("@kyuka", sHoliday);   // 休暇
                sCom.Parameters.AddWithValue("@kyukei", sKyukei);   // 休憩時間
                sCom.Parameters.AddWithValue("@teisei", steisei);  // 訂正
                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());  // 更新年月日

                // テーブル書き込み
                sCom.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票明細", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 勤務表明細更新
        /// </summary>
        /// <param name="sSh">開始時</param>
        /// <param name="sSm">開始分</param>
        /// <param name="sEh">終了時</param>
        /// <param name="sEm">終了分</param>
        /// <param name="sHoliday">休暇</param>
        /// <param name="sKyukei">休憩時間</param>
        /// <param name="steisei">訂正</param>
        /// <param name="sID">ID</param>
        public void ItemUpdate(string sSh, string sSm, string sEh, string sEm,
                           string sKyuka, string sKyukei, string steisei, int sID)
        {
            try
            {
                // 勤務票明細
                sb.Clear();
                sb.Append("update 勤務票明細 set ");
                sb.Append("開始時=?,開始分=?,終了時=?,終了分=?,休暇=?,休憩時間=?,訂正=?,");
                sb.Append("更新年月日=? ");
                sb.Append("where ID=? ");
                sCom.CommandText = sb.ToString();

                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@sh", sSh);           // 開始・時間
                sCom.Parameters.AddWithValue("@sm", sSm);           // 開始・分
                sCom.Parameters.AddWithValue("@eh", sEh);           // 終了・時間
                sCom.Parameters.AddWithValue("@em", sEm);           // 終了・分
                sCom.Parameters.AddWithValue("@kyuka", sKyuka);     // 休暇
                sCom.Parameters.AddWithValue("@kyukei", sKyukei);   // 休憩時間
                sCom.Parameters.AddWithValue("@teisei", steisei);   // 訂正
                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());  // 更新年月日
                sCom.Parameters.AddWithValue("@ID", sID);           // ID

                // テーブル書き込み
                sCom.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票明細", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// ＣＳＶデータをＭＤＢに登録する
        /// </summary>
        /// <param name="_InPath">CSVデータパス</param>
        /// <param name="frmP">プログレスバーフォームオブジェクト</param>
        /// <param name="cLen">CSVデータ件数</param>
        public void CsvToMdb(string _InPath, frmPrg frmP, int cLen)
        {
            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt / cLen * 100;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // MDBへ登録する
                        // 勤務記録ヘッダテーブル
                        HeadInsert(Utility.GetStringSubMax(stCSV[0], 17), 
                                   Utility.GetStringSubMax(stCSV[2], 1),
                                   Utility.GetStringSubMax(stCSV[3], 6),
                                   Utility.GetStringSubMax(stCSV[4], 2),
                                   Utility.GetStringSubMax(stCSV[5], 2),
                                   Utility.GetStringSubMax(stCSV[6], 8),
                                   Utility.GetStringSubMax(stCSV[7], 2),
                                   Utility.GetStringSubMax(stCSV[1], 21));
                        
                        // 勤務票明細テーブル
                        int sDays = 0;
                        DateTime dt;

                        for (int i = 8; i <= 218; i += 7)
                        {
                            sCom.Parameters.Clear();
                            sDays++;

                            // 存在する日付のときにMDBへ登録する
                            string tempDt = (global.sYear + Properties.Settings.Default.RekiHosei).ToString() + "/" + global.sMonth.ToString() + "/" + sDays.ToString();
                            if (DateTime.TryParse(tempDt, out dt))
                            {
                                ItemInsert(Utility.GetStringSubMax(stCSV[0], 17),
                                           sDays.ToString(),
                                           Utility.GetStringSubMax(stCSV[i + 1], 2),
                                           Utility.GetStringSubMax(stCSV[i + 2], 2),
                                           Utility.GetStringSubMax(stCSV[i + 3], 2),
                                           Utility.GetStringSubMax(stCSV[i + 4], 2),
                                           Utility.GetStringSubMax(stCSV[i], 1),
                                           Utility.GetStringSubMax(stCSV[i + 5], 3),
                                           stCSV[i + 6]);
                            }
                        }
                    }
                }

                // トランザクションコミット
                sTran.Commit();

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// OCRDATAクラスから勤務票データをＭＤＢに登録する
        /// </summary>
        /// <param name="_InPath">CSVデータパス</param>
        /// <param name="frmP">プログレスバーフォームオブジェクト</param>
        /// <param name="cLen">CSVデータ件数</param>
        public void ClsToMdb(OCRData [] ocr, int idx)
        {
            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                // MDBへ登録する
                // 勤務記録ヘッダテーブル
                int yy = int.Parse(ocr[idx]._Year) + Properties.Settings.Default.RekiHosei;
                HeadInsert(ocr[idx].OCRDATA_ID, ocr[idx]._SheetID, ocr[idx]._StaffCode, yy.ToString(),
                            ocr[idx]._Month, ocr[idx]._OrderCode, ocr[idx]._ShuDays, ocr[idx]._ImageName);

                // 勤務票明細テーブル
                DateTime dt;

                for (int i = 0; i < global._MULTIGYO; i++)
                {
                    // 存在する日付のときにMDBへ登録する
                    string tempDt = (global.sYear + Properties.Settings.Default.RekiHosei).ToString() + "/" + global.sMonth.ToString() + "/" + (i + 1).ToString();
                    if (DateTime.TryParse(tempDt, out dt))
                    {
                        ItemInsert(ocr[idx].OCRDATA_ID,
                            (i + 1).ToString(),
                            ocr[idx].itm[i]._Sh,
                            ocr[idx].itm[i]._Sm,
                            ocr[idx].itm[i]._eh,
                            ocr[idx].itm[i]._em,
                            ocr[idx].itm[i]._Kyuka,
                            ocr[idx].itm[i]._Kyukei,
                            ocr[idx].itm[i]._teisei);
                    }                            
                }

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票インポート処理", MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// MDBデータの件数をカウントする
        /// </summary>
        /// <returns></returns>
        public int CountMDB()
        {
            int rCnt = 0;
            OleDbDataReader dR;
            sCom.CommandText = "select count(ID) as cnt from 勤務票ヘッダ";
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数取得
                rCnt = int.Parse(dR["cnt"].ToString());
                //rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            return rCnt;
        }
        
        /// <summary>
        /// 勤務票ヘッダのデータリーダーを取得する
        /// </summary>
        /// <returns>データリーダー</returns>
        public OleDbDataReader HeaderSelect()
        {
            OleDbDataReader dR;
            sCom.CommandText = "select * from 勤務票ヘッダ order by ID";
            dR = sCom.ExecuteReader();

            return dR;
        }

        /// <summary>
        /// 指定したIDの勤務票ヘッダのデータリーダーを取得する
        /// </summary>
        /// <param name="sID">ID</param>
        /// <returns>データリーダー</returns>
        public OleDbDataReader HeaderSelect(string sID)
        {
            OleDbDataReader dR;

            sb.Clear();
            sb.Append("select 勤務票ヘッダ.* from 勤務票ヘッダ left join スタッフ ");
            sb.Append("on 勤務票ヘッダ.個人番号 = スタッフ.スタッフコード ");
            sb.Append("where 勤務票ヘッダ.ID=?");

            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@1", sID);
            dR = sCom.ExecuteReader();

            return dR;
        }

        /// <summary>
        /// 指定した年・月・スタッフコードの勤務票ヘッダのデータリーダーを取得する
        /// </summary>
        /// <param name="sID">ID</param>
        /// <returns>データリーダー</returns>
        public OleDbDataReader HeaderSelect(string _sYear, string _sMonth, string _sStaffCode)
        {
            OleDbDataReader dR;

            sb.Clear();
            sb.Append("select * from 勤務票ヘッダ ");
            sb.Append("where 年=? and 月=? and 個人番号=?");

            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@1", _sYear);
            sCom.Parameters.AddWithValue("@2", _sMonth);
            sCom.Parameters.AddWithValue("@3", _sStaffCode);
            dR = sCom.ExecuteReader();

            return dR;
        }

        /// <summary>
        /// 指定した年月以前の勤務票ヘッダのデータリーダーを取得する
        /// </summary>
        /// <param name="dYYMM">年月</param>
        /// <returns>データリーダー</returns>
        public OleDbDataReader HeaderSelect(int dYYMM)
        {
            OleDbDataReader dR;

            sb.Clear();
            sb.Append("select * from 勤務票ヘッダ ");
            sb.Append("where 年*100+月 <= ?");

            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@1", dYYMM);
            dR = sCom.ExecuteReader();

            return dR;
        }

        /// <summary>
        /// 指定したIDの勤務表明細のデータリーダーを取得する
        /// </summary>
        /// <param name="sID">ID</param>
        /// <returns>データリーダー</returns>
        public OleDbDataReader itemSelect(string sID)
        {
            OleDbDataReader dR;

            sb.Clear();
            sb.Append("select 勤務票明細.* from 勤務票明細 ");
            sb.Append("where ヘッダID.ID=? order by ID");

            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@1", sID);
            dR = sCom.ExecuteReader();

            return dR;
        }
    }
}

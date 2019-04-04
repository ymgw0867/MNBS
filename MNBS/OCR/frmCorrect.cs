using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using MNBS.Common;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
using Leadtools.WinForms;

namespace MNBS.OCR
{
    public partial class frmCorrect : Form
    {
        public frmCorrect(int sMode)
        {
            InitializeComponent();

            // 開始時処理モード
            _sMode = sMode;

            //tagを初期化
            this.Tag = END_MAKEDATA;

            // フォームキャプション
            SetformText();

            // 環境設定情報取得
            msConfig ms = new msConfig();
            ms.GetCommonYearMonth();

            // 入力ファイルフォルダ取得
            _InPath = Utility.exisInFileDir();

            // 登録モードのとき新規データを追加します
            if (sMode == global.sADDMODE)
            {
                // 新規データを用意します
                AddNewData();
            }
        }

        //終了ステータス
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";

        // 開始時処理モード
        private int _sMode;

        // 入力パス
        private string _InPath;

        // データグリッドビューカラム定義
        private string cDay = "col1";
        private string cWeek = "col2";
        //private string cMark = "col3";
        private string cSH = "col4";
        private string cS = "col5";
        private string cSM = "col6";
        private string cEH = "col7";
        private string cE = "col8";
        private string cEM = "col9";
        private string cKyuka = "col10";
        private string cKyukei = "col11";
        //private string cHiru1 = "col12";
        //private string cHiru2 = "col13";
        private string cTeisei = "col14";
        private string cID = "col15";

        // OCRDATAクラス
        OCRData[] clsOCR;

        //カレントデータインデックス
        private int _cI;

        // 社員マスタークラスインスタンス
        msShain[] ms;

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 19;
                tempDGV.RowTemplate.Height = 19;

                // 全体の高さ
                tempDGV.Height = 610;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cDay, "日");
                tempDGV.Columns.Add(cWeek, "曜");
                tempDGV.Columns.Add(cKyuka, "休暇");
                tempDGV.Columns.Add(cSH, "開");
                tempDGV.Columns.Add(cS, string.Empty);
                tempDGV.Columns.Add(cSM, "始");
                tempDGV.Columns.Add(cEH, "終");
                tempDGV.Columns.Add(cE, string.Empty);
                tempDGV.Columns.Add(cEM, "了");
                tempDGV.Columns.Add(cKyukei, "休憩時間");

                DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(column);
                tempDGV.Columns[10].Name = cTeisei;
                tempDGV.Columns[10].HeaderText = "訂正";
                tempDGV.Columns.Add(cID, "ID");
                tempDGV.Columns[cID].Visible = false; // IDカラムは非表示とする

                // 各列の定義を行う
                foreach (DataGridViewColumn c in tempDGV.Columns)
                {
                    // 幅
                    c.Width = 40;

                    // 表示位置、編集可否
                    if (c.Name == cDay)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    
                    if (c.Name == cWeek)
                    {
                        c.Width = 28;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }

                    if (c.Name == cKyuka)
                    {
                        c.Width = 40;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    
                    if (c.Name == cSH)
                    {
                        c.Width = 24;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    
                    if (c.Name == cS)
                    {
                        c.Width = 12;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    
                    if (c.Name == cSM)
                    {
                        c.Width = 24;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    
                    if (c.Name == cEH)
                    {
                        c.Width = 24;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    
                    if (c.Name == cE)
                    {
                        c.Width = 12;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    
                    if (c.Name == cEM)
                    {
                        c.Width = 24;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    
                    if (c.Name == cKyukei)
                    {
                        c.Width = 40;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }

                    if (c.Name == cTeisei)
                    {
                        //c.Width = 40;
                        c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    
                    if (c.Name == cID)
                    {
                        c.ReadOnly = true;
                    }

                    // 入力可能桁数
                    if (c.Name != cTeisei)
                    {
                        DataGridViewTextBoxColumn col = (DataGridViewTextBoxColumn)c;

                        if (c.Name == cSH) col.MaxInputLength = 2;
                        if (c.Name == cSM) col.MaxInputLength = 2;
                        if (c.Name == cEH) col.MaxInputLength = 2;
                        if (c.Name == cEM) col.MaxInputLength = 2;
                        if (c.Name == cKyuka) col.MaxInputLength = 1;
                        if (c.Name == cKyukei) col.MaxInputLength = 3;
                    }
                }

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                tempDGV.MultiSelect = false;

                // 編集可とする
                //tempDGV.ReadOnly = false;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                tempDGV.StandardTab = false;

                // ソート禁止
                foreach (DataGridViewColumn c in tempDGV.Columns)
                {
                    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                //tempDGV.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

                // 編集モード
                tempDGV.EditMode = DataGridViewEditMode.EditOnEnter;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        private void SetformText()
        {
            //if (_usrSel == global.STAFF_SELECT)
            //    this.Text = "勤怠データ スタッフ ― 登録  処理年月：";
            //else if (_usrSel == global.PART_SELECT)
            //    this.Text = "勤怠データ パートタイマー ― 登録  処理年月：";

            // 2019/04/04 コメント化
            //this.Text = string.Format("{0}年{1}月", global.sYear.ToString(), global.sMonth.ToString());

            // 西暦表示 2019/04/04
            this.Text = string.Format("{0}年{1}月", global.sYear + Properties.Settings.Default.RekiHosei, global.sMonth.ToString());
        }

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // グリッド定義
            GridViewSetting(dg1);

            ////// 勤怠CSVデータを取得してMDBへ登録
            ////GetCsvDataToMDB();

            // OCRDATAクラスインスタンスをデータ件数分生成
            if (!GetCsvDataToOCR())
            {
                MessageBox.Show("対象となる勤務票データがありません", "勤務票データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                this.Close();
                return;
            }
            
            // スタッフCSVデータを社員クラスへ取り込む
            MstCsvDataToCls(global.sMSTS);

            //エラー情報初期化
            ErrInitial();

            // 開始時処理モードによって表示するデータをコントロールします
            switch (_sMode)
            {
                // 登録モードのとき最後のレコードを表示
                case global.sADDMODE:
                    _cI = clsOCR.Length - 1;
                    break;

                // 編集モードのとき先頭レコードを表示
                case global.sEDITMODE:
                    _cI = 0;
                    break;

                default:
                    break;
            }

            // データ表示
            DataShow(_cI, clsOCR, this.dg1);
        }

        /// <summary>
        /// OCRDATAクラスをデータ件数分生成する
        /// </summary>
        /// <param name="clsOCR">OCRDATAクラス配列</param>
        private void CreateOCRDATA(OCRData [] c)
        {
            OCRData ms = new OCRData();
            OleDbDataReader dR = ms.HeaderSelect();

            int iX = 0;
            while (dR.Read())
            {
                c[iX] = new OCRData();
                c[iX].OCRDATA_ID = dR["ID"].ToString();    // 勤務票ヘッダＩＤを取得
                iX++;
            }

            if (!dR.IsClosed) dR.Close();
            if (ms.sCom.Connection.State == ConnectionState.Open) ms.sCom.Connection.Close();
        }

        private void dg1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 3 || e.ColumnIndex == 4 || e.ColumnIndex == 6 || e.ColumnIndex == 7)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;

                if (e.ColumnIndex == 4 || e.ColumnIndex == 5 || e.ColumnIndex == 7 || e.ColumnIndex == 8)
                {
                    e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
                }
                //else
                //    e.AdvancedBorderStyle.Left = dg1.AdvancedCellBorderStyle.Left;
            }
        }

        /// <summary>
        /// スタッフCSVデータを社員クラスへ取り込む
        /// </summary>
        private void MstCsvDataToCls(string InPath)
        {
            //CSVファイルがなければ終了
            if (!System.IO.File.Exists(InPath)) return;

            ////CSVファイル数を取得
            //string ss = System.IO.Path.GetDirectoryName(InPath);
            //var inCsv = System.IO.Directory.GetFileSystemEntries(ss, "*.csv");

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーフォームを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            int cCnt = 0;
            int cTotal = 0;

            try
            {
                // ＣＳＶデータの行数を取得します
                foreach (var stBuffer in System.IO.File.ReadAllLines(InPath, Encoding.Default))
                {
                    cTotal++;
                }

                // CSVデータを社員クラスへ取込む
                bool hd = true; 

                // CSVファイルを１行ずつインポート
                //var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                foreach (var stBuffer in System.IO.File.ReadAllLines(InPath, Encoding.Default))
                {
                    // 最初の行はヘッダのため読み飛ばし
                    if (hd)
                    {
                        hd = false;
                        continue;
                    }
                    
                    // カンマ区切りで分割して配列に格納する
                    string[] stCSV = stBuffer.Split(',');

                    // 正規のフォーマットのときインポート
                    if (stCSV.Length > 83)
                    {
                        // 件数カウント
                        cCnt++;

                        // 社員クラスインスタンス
                        if (cCnt > 1) Array.Resize(ref ms, cCnt);
                        else if (cCnt == 1) ms = new msShain[1];

                        ms[cCnt - 1] = new msShain();

                        // プログレスバー表示
                        frmP.Text = "スタッフＣＳＶデータロード中　" + cCnt.ToString();
                        frmP.progressValue = cCnt * 100 / cTotal;
                        frmP.ProgressStep();

                        ms[cCnt - 1]._OrderCode = stCSV[0].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._HaCode = stCSV[7].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._HaName = stCSV[8].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._BuName = stCSV[9].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._StaffCode = stCSV[33].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._StaffName = stCSV[34].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._STime = stCSV[79].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._ETime = stCSV[80].Replace(@"""", string.Empty);
                        ms[cCnt - 1]._Kyukei = stCSV[83].Replace(@"""", string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + cCnt.ToString(), "スタッフＣＳＶインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
                //////1秒間待機する（時間のかかる処理があるものとする）
                ////System.Threading.Thread.Sleep(5000);
            }

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        /// <summary>
        /// CSVデータをMDBへインサートする
        /// </summary>
        private void GetCsvDataToMDB()
        {
            //CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(_InPath, "*.csv");

            //CSVファイルがなければ終了
            if (inCsv.Length == 0) return;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // CSVデータをMDBへ取り込む
            OCRData oc = new OCRData();
            oc.CsvToMdb(_InPath, frmP, inCsv.Length);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        /// <summary>
        /// CSVデータをOCRクラスへインサートする
        /// </summary>
        private bool GetCsvDataToOCR()
        {
            //CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(_InPath, "*.csv");

            //CSVファイルがなければ終了
            if (inCsv.Length == 0) return false;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // OCRDATAクラスインスタンスをデータ件数分生成
            clsOCR = new OCRData[inCsv.Length];

            // CSVデータからＯＣＲクラスへデータを格納する
            try
            {
                //CSVデータをOCRデータクラスへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + inCsv.Length.ToString();
                    frmP.progressValue = cCnt / inCsv.Length * 100;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // ヘッダ情報
                        clsOCR[cCnt - 1] = new OCRData();
                        clsOCR[cCnt - 1].OCRDATA_ID = Utility.GetStringSubMax(stCSV[0], 17);
                        clsOCR[cCnt - 1]._SheetID = Utility.GetStringSubMax(stCSV[2], 1);
                        clsOCR[cCnt - 1]._StaffCode = Utility.GetStringSubMax(stCSV[3], 6);
                        clsOCR[cCnt - 1]._Year = Utility.GetStringSubMax(stCSV[4], 2);
                        clsOCR[cCnt - 1]._Month = Utility.GetStringSubMax(stCSV[5], 2);
                        clsOCR[cCnt - 1]._OrderCode = Utility.GetStringSubMax(stCSV[6], 8);
                        clsOCR[cCnt - 1]._ShuDays = Utility.GetStringSubMax(stCSV[7], 2);
                        clsOCR[cCnt - 1]._ImageName = Utility.GetStringSubMax(stCSV[1], 21);                        

                        // 明細情報
                        int sDays = 0;

                        for (int i = 8; i <= 218; i += 7)
                        {
                            clsOCR[cCnt - 1].itm[sDays]._Day = sDays.ToString();
                            clsOCR[cCnt - 1].itm[sDays]._Kyuka = Utility.GetStringSubMax(stCSV[i], 1);
                            clsOCR[cCnt - 1].itm[sDays]._Sh = Utility.GetStringSubMax(stCSV[i + 1], 2);
                            clsOCR[cCnt - 1].itm[sDays]._Sm = Utility.GetStringSubMax(stCSV[i + 2], 2);
                            clsOCR[cCnt - 1].itm[sDays]._eh = Utility.GetStringSubMax(stCSV[i + 3], 2);
                            clsOCR[cCnt - 1].itm[sDays]._em = Utility.GetStringSubMax(stCSV[i + 4], 2);
                            clsOCR[cCnt - 1].itm[sDays]._Kyukei = Utility.GetStringSubMax(stCSV[i + 5], 3);
                            clsOCR[cCnt - 1].itm[sDays]._teisei = stCSV[i + 6];

                            sDays++;
                        }
                    }

                    // MDBへの接続を解除する　2013/11/05
                    if (clsOCR[cCnt - 1].sCom.Connection.State == ConnectionState.Open)
                        clsOCR[cCnt - 1].sCom.Connection.Close();
                }

                ////CSVファイルを削除する
                //foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                //{
                //    System.IO.File.Delete(files);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);
                return false;
            }
            finally
            {
                for (int i = 0; i < clsOCR.Length; i++)
                {
                    if (clsOCR[i].sCom.Connection.State == ConnectionState.Open)
                        clsOCR[i].sCom.Connection.Close();
                }
            }

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            return true;
        }

        private void ErrInitial()
        {
            //エラー情報初期化
            lblErrMsg.Visible = false;
            global.errNumber = global.eNothing;     //エラー番号
            global.errMsg = string.Empty;           //エラーメッセージ
            lblErrMsg.Text = string.Empty;
        }

        /// <summary>
        /// データ表示
        /// </summary>
        /// <param name="sIx">表示インデックス</param>
        /// <param name="rRec">OCRDATAクラス</param>
        /// <param name="dgv">DataGridViewオブジェクト</param>
        private void DataShow(int sIx, OCRData[] rRec, DataGridView dgv)
        {
            // 出勤日数合計
            int rDays = 0;

            // 画像ファイル名
            global.pblImageFile = string.Empty;

            // データグリッドビュー初期化
            dataGridInitial(this.dg1);

            //データ表示背景色初期化
            dsColorInitial(this.dg1);

            try
            {
                // ヘッダ情報
                txtYear.Text = Utility.EmptytoZero(rRec[sIx]._Year);
                txtMonth.Text = Utility.EmptytoZero(rRec[sIx]._Month);
                txtNo.Text = Utility.EmptytoZero(rRec[sIx]._StaffCode);
                txtOrderCode.Text = Utility.EmptytoZero(rRec[sIx]._OrderCode);
                txtShuTl.Text = Utility.EmptytoZero(rRec[sIx]._ShuDays);
                global.pblImageFile = rRec[sIx]._ImageName;
                                
                lblName.Text = string.Empty;
                txtShozokuCode.Text = string.Empty;
                lblShozoku.Text = string.Empty;

                // スタッフ情報取得
                GetStaffData(rRec[sIx]._StaffCode);

                //データ数表示
                lblPage.Text = (sIx + 1).ToString() + "/" + rRec.Length.ToString() + " 件目";

                // 明細情報
                int r = 0;
                for (int i = 0; i < global._MULTIGYO; i++)
                {
                    // データグリッド最下行まで達した（月末日以降のデータは強制消去する）
                    if (dgv.Rows.Count <= r)
                    {
                        rRec[sIx].itm[i]._Sh = string.Empty;
                        rRec[sIx].itm[i]._Sm = string.Empty;
                        rRec[sIx].itm[i]._eh = string.Empty;
                        rRec[sIx].itm[i]._em = string.Empty;
                        rRec[sIx].itm[i]._Kyuka = string.Empty;
                        rRec[sIx].itm[i]._Kyukei = string.Empty;
                        rRec[sIx].itm[i]._teisei = string.Empty;
                    }
                    else
                    {
                        //// ChangeValueイベント処理回避
                        //global.dg1ChabgeValueStatus = false;

                        // ChangeValueイベント処理実施
                        global.dg1ChabgeValueStatus = true;

                        if (rRec[sIx].itm[i]._teisei == global.FLGON) dgv[cTeisei, r].Value = true;
                        else dgv[cTeisei, r].Value = false;

                        dgv[cKyuka, r].Value = rRec[sIx].itm[i]._Kyuka;
                        dgv[cSH, r].Value = rRec[sIx].itm[i]._Sh;
                        dgv[cSM, r].Value = rRec[sIx].itm[i]._Sm;
                        dgv[cEH, r].Value = rRec[sIx].itm[i]._eh;
                        dgv[cEM, r].Value = rRec[sIx].itm[i]._em;
                        dgv[cKyukei, r].Value = rRec[sIx].itm[i]._Kyukei;

                        // ChangeValueイベント処理回避
                        global.dg1ChabgeValueStatus = false;

                        // 出勤日数加算
                        if (rRec[sIx].itm[i]._Sh != string.Empty) rDays++;

                        //dgv[cID, r].Value = dR["ID"].ToString();

                        // ChangeValueイベント処理実施
                        global.dg1ChabgeValueStatus = true;
                    }

                    r++;

                    //// データグリッド最下行まで達したら終了する（月末日以降のデータは無視する）
                    //if (dgv.Rows.Count == r) break;
                }

                //画像イメージ表示
                ShowImage(_InPath + global.pblImageFile);

                // ヘッダ情報
                txtYear.ReadOnly = false;
                txtMonth.ReadOnly = false;
                txtShozokuCode.ReadOnly = false;
                txtNo.ReadOnly = false;

                //最初のレコード
                if (sIx == 0)
                {
                    btnBefore.Enabled = false;
                    btnFirst.Enabled = false;
                }

                //最終レコード
                if ((sIx + 1) == rRec.Length)
                {
                    btnNext.Enabled = false;
                    btnEnd.Enabled = false;
                }

                //カレントセル選択状態としない
                dgv.CurrentCell = null;

                // その他のボタンを有効とする
                btnErrCheck.Enabled = true;
                btnDataMake.Enabled = true;
                btnDel.Enabled = true;

                // データグリッドビュー編集
                dg1.ReadOnly = false;

                //エラー情報表示
                ErrShow();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                global.dg1ChabgeValueStatus = true;
            }
        }

        /// <summary>
        /// スタッフ情報取得
        /// </summary>
        /// <param name="staffCode">スタッフコード</param>
        private void GetStaffData(string staffCode)
        {
            for (int i = 0; i < ms.Length; i++)
            {
                if (ms[i]._StaffCode == staffCode)
                {
                    lblName.Text = ms[i]._StaffName;
                    txtShozokuCode.Text = ms[i]._HaCode;
                    lblShozoku.Text = ms[i]._HaName;
                    break;
                }
            }
        }

        //表示初期化
        private void dataGridInitial(DataGridView dgv)
        {
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtNo.Text = string.Empty;
            txtShozokuCode.Text = string.Empty;
            txtOrderCode.Text = string.Empty;
            txtShuTl.Text = string.Empty;

            lblName.Text = string.Empty;
            lblShozoku.Text = string.Empty;

            txtYear.BackColor = Color.Empty;
            txtMonth.BackColor = Color.Empty;
            txtNo.BackColor = Color.Empty;
            txtShozokuCode.BackColor = Color.Empty;
            txtOrderCode.BackColor = Color.Empty;
            txtShuTl.BackColor = Color.Empty;

            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtNo.ForeColor = Color.Navy;
            txtShozokuCode.ForeColor = Color.Navy;
            txtOrderCode.ForeColor = Color.Navy;
            txtShuTl.ForeColor = Color.Navy;

            dgv.RowsDefaultCellStyle.ForeColor = Color.Navy;       //テキストカラーの設定
            dgv.DefaultCellStyle.SelectionBackColor = Color.Empty;
            dgv.DefaultCellStyle.SelectionForeColor = Color.Navy;

            //dgv.EditMode = EditMode.EditOnKeystrokeOrShortcutKey;

            //pictureBox1.Image = null;
            lblNoImage.Visible = false;
        }

        /// <summary>
        /// データ表示エリア背景色初期化
        /// </summary>
        /// <param name="dgv">データグリッドビューオブジェクト</param>
        private void dsColorInitial(DataGridView dgv)
        {
            // CellChangeValueイベント発生回避する
            global.dg1ChabgeValueStatus = false;

            txtYear.BackColor = Color.White;
            txtMonth.BackColor = Color.White;
            txtNo.BackColor = Color.White;
            txtShozokuCode.BackColor = Color.White;
            txtOrderCode.BackColor = Color.White;
            txtShuTl.BackColor = Color.White;

            // 行数
            dgv.RowCount = 0;

            // 行追加、日付セット
            //SysControl.SetDBConnect db = new SysControl.SetDBConnect();
            //OleDbCommand sCom = new OleDbCommand();
            //sCom.Connection = db.cnOpen();
            //OleDbDataReader dr = null;

            for (int i = 0; i < global._MULTIGYO; i++)
            {
                DateTime dt;
                if (DateTime.TryParse((global.sYear + Properties.Settings.Default.RekiHosei).ToString() + "/" + global.sMonth.ToString() + "/" + (i + 1).ToString(), out dt))
                {
                    // 行を追加
                    dgv.Rows.Add();
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
                                
                    // 日
                    dgv[cDay, i].Value = i + 1;
                    // 曜日
                    string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                    dgv[cWeek, i].Value = Youbi;

                    //// 土日の場合
                    //if (Youbi == "日" || Youbi == "土")
                    //{
                    //    dgv.Rows[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    //}
                    //else
                    //{
                    //// 休日テーブルを参照し休日に該当するか調べます
                    //sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                    //sCom.Parameters.Clear();
                    //sCom.Parameters.AddWithValue("@year", global.sYear);
                    //sCom.Parameters.AddWithValue("@Month", global.sMonth);
                    //sCom.Parameters.AddWithValue("@day", i + 1);
                    //dr = sCom.ExecuteReader();
                    //if (dr.HasRows)
                    //{
                    //    dgv.Rows[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    //}
                    //dr.Close();
                    //}

                    // 時分区切り記号
                    dgv[cS, i].Value = ":";
                    dgv[cE, i].Value = ":";

                }
            }
            //sCom.Connection.Close();

            // CellChangeValueイベントステータスもどす
            global.dg1ChabgeValueStatus = true;
        }

        /// <summary>
        /// 伝票画像表示
        /// </summary>
        /// <param name="iX">現在の伝票</param>
        /// <param name="tempImgName">画像名</param>
        public void ShowImage(string tempImgName)
        {
            string wrkFileName;

            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
                return;
            }

            //画像ファイルがあるときのみ表示
            wrkFileName = tempImgName;
            if (System.IO.File.Exists(wrkFileName))
            {
                leadImg.Visible = true;

                //画像ロード
                RasterCodecs.Startup();
                RasterCodecs cs = new RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                RasterPaintProperties prop = new RasterPaintProperties();
                prop = RasterPaintProperties.Default;
                prop.PaintDisplayMode = RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(wrkFileName, 0, CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    if (leadImg.ImageDpiX == 200) 
                        leadImg.ScaleFactor *= global.ZOOM_RATE_FAX;    // 200*200 画像のとき
                    else leadImg.ScaleFactor *= global.ZOOM_RATE;       // 300*300 画像のとき
                }
                else
                {
                    //leadImg.ScaleFactor *= global.miMdlZoomRate;

                    if (leadImg.ImageDpiX == 200)
                        leadImg.ScaleFactor *= global.ZOOM_RATE_FAX + global.ZOOM_Total;    // 200*200 画像のとき
                    else leadImg.ScaleFactor *= global.ZOOM_RATE + global.ZOOM_Total;       // 300*300 画像のとき
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = RasterViewerInteractiveMode.Pan;

                ////右へ90°回転させる
                //RotateCommand rc = new RotateCommand();
                //rc.Angle = 90 * 100;
                //rc.FillColor = new RasterColor(255, 255, 255);
                ////rc.Flags = RotateCommandFlags.Bicubic;
                //rc.Flags = RotateCommandFlags.Resize;
                //rc.Run(leadImg.Image);

                // グレースケールに変換
                GrayscaleCommand grayScaleCommand = new GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                RasterCodecs.Shutdown();
                global.pblImageFile = wrkFileName;

                lblNoImage.Visible = false;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;
            }
            else
            {
                //画像ファイルがないとき
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;
            }
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = 0;
            DataShow(_cI, clsOCR, dg1);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (_cI > 0)
            {
                _cI--;
                DataShow(_cI, clsOCR, dg1);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (_cI + 1 < clsOCR.Length)
            {
                _cI++;
                DataShow(_cI, clsOCR, dg1);
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = clsOCR.Length - 1;
            DataShow(_cI, clsOCR, dg1);
        }

        private void txtRiseki_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        /// <summary>
        ///  カレントデータの更新
        /// </summary>
        /// <param name="iX">カレントレコードのインデックス</param>
        private void CurDataUpDate(int iX)
        {
            // エラーメッセージ
            string errMsg = string.Empty;

            try
            {
                // ヘッダ情報
                clsOCR[iX]._StaffCode = Utility.NulltoStr(txtNo.Text).PadLeft(6, '0');
                clsOCR[iX]._Year = Utility.NulltoStr(txtYear.Text);
                clsOCR[iX]._Month = Utility.NulltoStr(txtMonth.Text);
                clsOCR[iX]._OrderCode = Utility.NulltoStr(txtOrderCode.Text);
                clsOCR[iX]._ShuDays = Utility.NulltoStr(txtShuTl.Text);

                // 明細情報
                for (int i = 0; i < dg1.Rows.Count; i++)
                {
                    // 訂正チェック
                    string sTeisei = string.Empty;
                    if (dg1[cTeisei, i].Value.ToString() == "True") sTeisei = global.FLGON;
                    else sTeisei = global.FLGOFF;

                    clsOCR[iX].itm[i]._Day = (i + 1).ToString();
                    clsOCR[iX].itm[i]._Kyuka = Utility.NulltoStr(dg1[cKyuka, i].Value).ToString().Trim();
                    clsOCR[iX].itm[i]._Sh = Utility.NulltoStr(dg1[cSH, i].Value).ToString().Trim();
                    clsOCR[iX].itm[i]._Sm = Utility.NulltoStr(dg1[cSM, i].Value).ToString().Trim();
                    clsOCR[iX].itm[i]._eh = Utility.NulltoStr(dg1[cEH, i].Value).ToString().Trim();
                    clsOCR[iX].itm[i]._em = Utility.NulltoStr(dg1[cEM, i].Value).ToString().Trim();
                    clsOCR[iX].itm[i]._Kyukei = Utility.NulltoStr(dg1[cKyukei, i].Value).ToString().Trim();
                    clsOCR[iX].itm[i]._teisei = sTeisei;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, errMsg, MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                global.ZOOM_Total += global.ZOOM_STEP;
                leadImg.ScaleFactor += global.ZOOM_STEP;

                //if (leadImg.ImageDpiX == 200)
                //    leadImg.ScaleFactor = global.ZOOM_RATE_FAX + global.ZOOM_Total;
                //else leadImg.ScaleFactor = global.ZOOM_RATE + global.ZOOM_Total;
            }

            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                global.ZOOM_Total -= global.ZOOM_STEP;
                leadImg.ScaleFactor -= global.ZOOM_STEP;

                //if (leadImg.ImageDpiX == 200)
                //    leadImg.ScaleFactor = global.ZOOM_RATE_FAX + global.ZOOM_Total;
                //else leadImg.ScaleFactor = global.ZOOM_RATE + global.ZOOM_Total;
            }

            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void dg1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                string ColName = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;

                if (ColName == cSH || ColName == cSM || ColName == cEH || ColName == cEM || ColName == cKyuka || ColName == cKyukei)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                    e.Control.KeyDown -= new KeyEventHandler(Control_KeyDown2);
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }

                //if (ColName == cMark || ColName == cKyukei || ColName == cHiru1 || ColName == cHiru2)
                //{
                //    //イベントハンドラが複数回追加されてしまうので最初に削除する
                //    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                //    e.Control.KeyDown -= new KeyEventHandler(Control_KeyDown2);
                //    //イベントハンドラを追加する
                //    e.Control.KeyDown += new KeyEventHandler(Control_KeyDown2);
                //    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress2);
                //}
            }
        }

        void Control_KeyDown2(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Space && e.KeyCode != Keys.Delete && e.KeyCode != Keys.Tab && e.KeyCode != Keys.Enter)
            {
                e.Handled = true;
                return;
            }
        }

        void Control_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)Keys.Space && e.KeyChar != '\b' && e.KeyChar != (char)Keys.Delete && e.KeyChar != (char)Keys.Tab && e.KeyChar != (char)Keys.Enter)
            {
                e.Handled = true;
                return;
            }
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void dg1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (global.dg1ChabgeValueStatus == false) return;
            global.dg1ChabgeValueStatus = false; // 自らChangeValueイベントを発生させない

            if (_cI == null) return;
            if (dg1.CurrentRow == null) return;

            // 該当カラム
            string Col = dg1.Columns[e.ColumnIndex].Name;

            // 該当行
            int Rn = e.RowIndex;

            if (Col == cDay || Col == cWeek || Col == cS || Col == cE)
            {
                global.dg1ChabgeValueStatus = true;
                return;
            }

            global.lBackColorE = Color.FromArgb(251, 228, 183);
            global.lBackColorN = Color.FromArgb(255, 255, 255);            

            switch (dg1[cWeek, Rn].Value.ToString())
            {
                case "土":
                    global.lBackColorN = Color.FromArgb(225, 244, 255);
                    break;
                case "日":
                    global.lBackColorN = Color.FromArgb(255, 223, 227);
                    break;
                default:
                    global.lBackColorN = Color.FromArgb(255, 255, 255);
                    break;
            }
            
            // 訂正チェック
            if (dg1[cTeisei, Rn].Value != null)
            {
                if (dg1[cTeisei, Rn].Value.ToString() == "True")
                {
                    global.lBackColorE = Color.FromArgb(251, 228, 183);
                    global.lBackColorN = Color.FromArgb(251, 228, 183);
                }

                // 行表示色
                TeiseiColor(Rn);
            }
            else
            {
                dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorN;
            }

            // 開始時間
            // 表示色をリセットします
            if (Col == cSH || Col == cSM)
            {
                dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
            }

            if ((Col == cSH || Col == cSM) && (dg1[cSH, Rn].Value != null && dg1[cSM, Rn].Value != null))
            {
                // 開始時
                if (Col == cSH)
                {
                    dg1[cSH, Rn].Style.BackColor = global.lBackColorN;

                    if (Utility.NumericCheck(dg1[cSH, Rn].Value.ToString()))
                        dg1[cSH, Rn].Value = dg1[cSH, Rn].Value.ToString().PadLeft(2, '0');

                    if (Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) < 8 && Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) != 0)
                        dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                    else dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                }

                // 開始分
                else if (Col == cSM)
                {
                    dg1[cSM, Rn].Style.BackColor = global.lBackColorN;

                    if (Utility.NumericCheck(dg1[cSM, Rn].Value.ToString()))
                    {
                        dg1[cSM, Rn].Value = dg1[cSM, Rn].Value.ToString().PadLeft(2, '0');

                        if (dg1[cSM, Rn].Value.ToString().Substring(1, 1) == "5" ||
                            dg1[cSM, Rn].Value.ToString().Substring(1, 1) == "0")
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                        else dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                    }
                }
            }

            // 終了時間
            if (Col == cEH || Col == cEM)
            {
                dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
            }

            if ((Col == cEH || Col == cEM) && (dg1[cEH, Rn].Value != null && dg1[cEM, Rn].Value != null))
            {
                if (Col == cEH)
                {
                    if (Utility.NumericCheck(dg1[cEH, Rn].Value.ToString()))
                        dg1[cEH, Rn].Value = dg1[cEH, Rn].Value.ToString().PadLeft(2, '0');

                    if (Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) < 8 && Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) >= 22)
                        dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                    else dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                }

                // 終了分
                else if (Col == cEM)
                {
                    if (Utility.NumericCheck(dg1[cEM, Rn].Value.ToString()))
                    {
                        dg1[cEM, Rn].Value = dg1[cEM, Rn].Value.ToString().PadLeft(2, '0');

                        if (dg1[cEM, Rn].Value.ToString().Substring(1, 1) == "5" ||
                            dg1[cEM, Rn].Value.ToString().Substring(1, 1) == "0")
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                        else dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                    }
                }
            }

            if (Col == cTeisei)
            {
                global.lBackColorN = Color.FromArgb(255, 255, 255);
                
                // 行表示色
                TeiseiColor(Rn);
            }

            // ChangeValueイベントステータスを戻す
            global.dg1ChabgeValueStatus = true;
        }

        private void TeiseiColor(int Rn)
        {

            if (dg1[cTeisei, Rn].Value.ToString() == "True")
            {
                dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorE;
                dg1[cDay, Rn].Style.BackColor = global.lBackColorE;
                dg1[cWeek, Rn].Style.BackColor = global.lBackColorE;
                dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cKyuka, Rn].Style.BackColor = global.lBackColorE;
                dg1[cKyukei, Rn].Style.BackColor = global.lBackColorE;
                dg1[cTeisei, Rn].Style.BackColor = global.lBackColorE;
            }
            else
            {
                dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorN;
                dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
                dg1[cWeek, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cKyuka, Rn].Style.BackColor = global.lBackColorN;
                dg1[cKyukei, Rn].Style.BackColor = global.lBackColorN;
                dg1[cTeisei, Rn].Style.BackColor = global.lBackColorN;
            }
        }

        private void dg1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;
            if (colName == cTeisei)
            {
                if (dg1.IsCurrentCellDirty)
                {
                    dg1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    dg1.RefreshEdit();
                }
            }
        }

        private void dg1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string colName = dg1.Columns[e.ColumnIndex].Name;
            //if (colName == cMark || colName == cHiru1 || colName == cHiru2)
            //{
            //    if (dg1[colName, dg1.CurrentRow.Index].Value == null)
            //    {
            //        dg1[colName, dg1.CurrentRow.Index].Value = global.MARU;
            //    }
            //    else if (dg1[colName, dg1.CurrentRow.Index].Value.ToString() == global.MARU)
            //    {
            //        dg1[colName, dg1.CurrentRow.Index].Value = string.Empty;
            //    }
            //    else
            //    {
            //        dg1[colName, dg1.CurrentRow.Index].Value = global.MARU;
            //    }

            //    dg1.RefreshEdit();

            //    // ChangeValueイベントステータスを戻す
            //    global.dg1ChabgeValueStatus = true; 
            //}
        }

        private void dg1_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            // 入力値がSpaceのときに有効とします
            if (e.Value.ToString().Trim() != string.Empty) return;

            DataGridViewEx dgv = (DataGridViewEx)sender;
            string colName = dg1.Columns[e.ColumnIndex].Name;

            //セルの列を調べる
            
            //if (colName == cMark || colName == cHiru1 || colName == cHiru2)
            //{
            //    if (dgv[colName, dgv.CurrentRow.Index].Value == null ||
            //        dgv[colName, dgv.CurrentRow.Index].Value.ToString() != global.MARU)
            //    {
            //        //○をセルの値とする
            //        e.Value = global.MARU;
            //    }
            //    else if (dgv[colName, dgv.CurrentRow.Index].Value.ToString() == global.MARU)
            //    {
            //        //セルをEmptyとする
            //        e.Value = string.Empty;
            //    }

            //    //解析が不要であることを知らせる
            //    e.ParsingApplied = true;

            //    // ChangeValueイベントステータスを戻す
            //    global.dg1ChabgeValueStatus = true; 
            //}
        }

        private void dg1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            string ColH = string.Empty;
            string ColM = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;

            // 開始時間または終了時間を判断
            if (ColM == cSM)
            {
                ColH = cSH;
            }
            else if (ColM == cEM)
            {
                ColH = cEH;
            }
            else
            {
                return;
            }

            // 開始時、終了時が入力済みで開始分、終了分が未入力のとき"00"を表示します
            if (dg1[ColH, dg1.CurrentRow.Index].Value != null)
            {
                if (dg1[ColH, dg1.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
                {
                    if (dg1[ColM, dg1.CurrentRow.Index].Value == null)
                    {
                        dg1[ColM, dg1.CurrentRow.Index].Value = "00";
                    }
                    else if (dg1[ColM, dg1.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
                    {
                        dg1[ColM, dg1.CurrentRow.Index].Value = "00";
                    }
                }
            }
        }

        private void btnErrCheck_Click(object sender, EventArgs e)
        {
            //カレントレコード更新
            CurDataUpDate(_cI);

            //エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(_cI, clsOCR.Length - 1) == false) return;

            //エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (_cI > 0)
            {
                if (ErrCheckMain(0, _cI - 1) == false) return;
            }

            // エラーなし
            DataShow(_cI, clsOCR, dg1);
            MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dg1.CurrentCell = null;
        }

        /// <summary>
        /// エラーチェックメイン処理
        /// </summary>
        /// <param name="sID">開始ID</param>
        /// <param name="eID">終了ID</param>
        /// <returns>True:エラーなし、false:エラーあり</returns>
        private Boolean ErrCheckMain(int sIx, int eIx)
        {
            int rCnt = 0;

            // オーナーフォームを無効にする
            this.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // レコード件数取得
            int cTotal = eIx - sIx + 1;

            // エラー情報初期化
            ErrInitial();

            // OCRDATAクラス読み出し
            Boolean eCheck = true;

            for (int i = sIx; i <= eIx; i++)
            {
                // データ件数加算
                rCnt++;

                // プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                // エラーチェックを実施する
                eCheck = ErrCheckData(i);

                //エラーがあったとき
                if (eCheck == false) break;
            }

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            //エラー有りの処理
            if (eCheck == false)
            {
                _cI = global.errID;
                DataShow(_cI, clsOCR, dg1);
            }

            return eCheck;
        }

        /// <summary>
        /// 項目別エラーチェック
        /// </summary>
        /// <param name="i">OCRDATAクラスインデックス</param>
        /// <returns>エラーなし：true, エラー有り：false</returns>
        private Boolean ErrCheckData(int i)
        {
            // 記入日数
            int dDays = 0;

            // 出勤日数
            int rDays = 0;

            //対象年
            if (clsOCR[i]._Year == string.Empty)
            {
                global.errID = i;
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "年を入力してください";

                return false;
            }

            if (Utility.NumericCheck(clsOCR[i]._Year) == false)
            {
                global.errID = i;
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "数字を入力してください";

                return false;
            }

            if (global.sYear != int.Parse(clsOCR[i]._Year))
            {
                global.errID = i;
                global.errNumber = global.eYear;
                global.errRow = 0;

                // 2019/04/04 コメント化
                //global.errMsg = "処理年と異なっています。 処理年月：" + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月";

                // 西暦表示 2019/04/04
                global.errMsg = "処理年と異なっています。 処理年月：" + (global.sYear + Properties.Settings.Default.RekiHosei) + "年 " + global.sMonth.ToString() + "月";

                return false;
            }

            //対象月
            if (clsOCR[i]._Month == string.Empty)
            {
                global.errID = i;
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "月を入力してください";

                return false;
            }

            if (Utility.NumericCheck(clsOCR[i]._Month) == false)
            {
                global.errID = i;
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "数字を入力してください";

                return false;
            }

            if (int.Parse(clsOCR[i]._Month) < 1 || int.Parse(clsOCR[i]._Month) > 12)
            {
                global.errID = i;
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "正しい月を入力してください";

                return false;
            }

            if (global.sMonth != int.Parse(clsOCR[i]._Month))
            {
                global.errID = i;
                global.errNumber = global.eMonth;
                global.errRow = 0;

                // コメント化 2019/04/04
                //global.errMsg = "処理月と異なっています。 処理年月：" + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月";

                // 西暦表示　2019/04/04
                global.errMsg = "処理月と異なっています。 処理年月：" + (global.sYear + Properties.Settings.Default.RekiHosei) + "年 " + global.sMonth.ToString() + "月";

                return false;
            }

            // 個人番号
            // 未入力のとき
            if (clsOCR[i]._StaffCode == null)
            {
                global.errID = i;
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号を入力してください";

                return false;
            }

            if (clsOCR[i]._StaffCode == string.Empty)
            {
                global.errID = i;
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号を入力してください";

                return false;
            }

            ////数字以外のとき
            //if (Utility.NumericCheck(cdR["個人番号"].ToString()) == false)
            //{
            //    global.errID = cdR["ID"].ToString();
            //    global.errNumber = global.eShainNo;
            //    global.errRow = 0;
            //    global.errMsg = "個人番号が不正です。";

            //    return false;
            //}

            //個人番号マスター登録検査
            bool rHas = false;

            for (int iX = 0; iX < ms.Length; iX++)
            {
                if (ms[iX]._StaffCode == clsOCR[i]._StaffCode)
                {
                    rHas = true;
                    break;
                }
            }

            if (!rHas)
            {
                global.errID = i;
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号がマスターに存在しません";
                return false;
            }

            // オーダーコード
            // 数字以外のとき
            if (Utility.NumericCheck(clsOCR[i]._OrderCode) == false)
            {
                global.errID = i;
                global.errNumber = global.eOrderCode;
                global.errRow = 0;
                global.errMsg = "オーダーコードが不正です。";

                return false;
            }

            rHas = false;

            // スタッフマスターを検索
            for (int iX = 0; iX < ms.Length; iX++)
            {
                if (ms[iX]._StaffCode == clsOCR[i]._StaffCode)
                {
                    if (ms[iX]._OrderCode == clsOCR[i]._OrderCode)
                    {
                        rHas = true;
                        break;
                    }
                }
            }

            if (!rHas)
            {
                global.errID = i;
                global.errNumber = global.eOrderCode;
                global.errRow = 0;
                global.errMsg = "オーダーコードとスタッフコードの組み合わせが一致しません。";
                return false;
            }

            //日付別データ
            for (int x = 0; x < global._MULTIGYO; x++)
            {
                // エラーチェックの実行フラグ
                bool _errChek = true;

                // エラーチェック対象外の行を判定します
                // 空白行のときエラーチェック対象外
                if (clsOCR[i].itm[x]._Sh == string.Empty && clsOCR[i].itm[x]._Sm == string.Empty &&
                    clsOCR[i].itm[x]._eh == string.Empty && clsOCR[i].itm[x]._em == string.Empty &&
                    clsOCR[i].itm[x]._Kyuka == string.Empty && clsOCR[i].itm[x]._Kyukei == string.Empty)
                {
                    _errChek = false;
                }

                // 明細行チェックの実施
                if (_errChek)
                {
                    if (!CheckMeisai(i, x))
                    {
                        global.errID = i;
                        global.errRow = x;
                        return false;
                    }
                }

                // 記入日数加算
                if (clsOCR[i].itm[x]._Sh != string.Empty ||
                    clsOCR[i].itm[x]._Kyuka == global.eYUKYU || clsOCR[i].itm[x]._Kyuka == global.eKEKKIN ||
                    clsOCR[i].itm[x]._Kyuka == global.eFURIDE || clsOCR[i].itm[x]._Kyuka == global.eFURIKYU)
                    dDays++;

                // 出勤日数加算
                if (clsOCR[i].itm[x]._Sh != string.Empty) rDays++;
            }

            // 出勤日数が記入回数と一致しているか
            if (dDays == 0)
            {
                global.errID = i;
                global.errNumber = global.eDays;
                global.errRow = 0;
                global.errMsg = "有効な勤怠データがありません";
                return false;
            }

            // 出勤日数が記入回数と一致しているか
            if (int.Parse(clsOCR[i]._ShuDays) != rDays)
            {
                global.errID = i;
                global.errNumber = global.eDays;
                global.errRow = 0;
                global.errMsg = "出勤日数合計が一致していません。プログラム計算値：" + rDays.ToString() + "日";
                return false;
            }

            return true;
        }

        /// <summary>
        /// 明細行毎のエラーチェック
        /// </summary>
        /// <param name="iH">OCRDATA配列インデックス</param>
        /// <param name="iX">OCRDATA行明細配列インデックス</param>
        /// <returns></returns>
        private bool CheckMeisai(int iH, int iX)
        {
            // 時間記入時チェック
            if (clsOCR[iH].itm[iX]._Kyuka != global.eYUKYU && 
                clsOCR[iH].itm[iX]._Kyuka != global.eKEKKIN &&
                clsOCR[iH].itm[iX]._Kyuka != global.eFURIKYU)
            {
                // 開始時記入
                if (clsOCR[iH].itm[iX]._Sh == string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "開始時刻を入力してください";
                    return false;
                }

                // 開始分記入
                if (clsOCR[iH].itm[iX]._Sm == string.Empty)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "開始時刻を入力してください";
                    return false;
                }

                // 開始時間 数字か？
                if (!Utility.NumericCheck(clsOCR[iH].itm[iX]._Sh))
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }

                // 開始時間 範囲外
                if (int.Parse(clsOCR[iH].itm[iX]._Sh) < 0 || int.Parse(clsOCR[iH].itm[iX]._Sh) > 23)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }

                // 開始分 数字か？
                if (!Utility.NumericCheck(clsOCR[iH].itm[iX]._Sm))
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }

                // 開始分 範囲
                if (int.Parse(clsOCR[iH].itm[iX]._Sm) < 0 || int.Parse(clsOCR[iH].itm[iX]._Sm) > 59)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }

                // 開始分は５分単位
                if ((int.Parse(clsOCR[iH].itm[iX]._Sm) % 5) != 0)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "5分単位で入力してください";
                    return false;
                }

                // 終了時記入
                if (clsOCR[iH].itm[iX]._eh == string.Empty)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "終了時刻を入力してください";
                    return false;
                }

                // 終了分記入
                if (clsOCR[iH].itm[iX]._eh == string.Empty)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "終了時刻を入力してください";
                    return false;
                }

                // 終了時　数字
                if (!Utility.NumericCheck(clsOCR[iH].itm[iX]._eh))
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }

                // 終了時間 範囲外
                if (int.Parse(clsOCR[iH].itm[iX]._eh) < 0 || int.Parse(clsOCR[iH].itm[iX]._eh) > 23)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }

                // 終了分 数字
                if (!Utility.NumericCheck(clsOCR[iH].itm[iX]._em))
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }

                // 終了分 範囲外
                if (int.Parse(clsOCR[iH].itm[iX]._em) < 0 || int.Parse(clsOCR[iH].itm[iX]._em) > 59)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }

                // 終了分は５分単位
                if ((int.Parse(clsOCR[iH].itm[iX]._em) % 5) != 0)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "5分単位で入力してください";
                    return false;
                }

                // 休憩時間
                if (clsOCR[iH].itm[iX]._Kyukei != string.Empty)
                {
                    // 数字
                    if (!Utility.NumericCheck(clsOCR[iH].itm[iX]._Kyukei))
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "休憩時間が不正です";
                        return false;
                    }

                    // ５分単位
                    if ((int.Parse(clsOCR[iH].itm[iX]._Kyukei) % 5) != 0)
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "5分単位で入力してください";
                        return false;
                    }
                }

                ////// 開始時間と終了時間
                ////int stm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());
                ////int etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                ////if (stm >= etm)
                ////{
                ////    global.errNumber = global.eEH;
                ////    global.errMsg = "終了時間が開始時間以前となっています";
                ////    return false;
                ////}

                ////// 勤務時間数
                ////DateTime sHM = DateTime.Parse(dR["開始時"].ToString() + ":" + dR["開始分"].ToString());
                ////DateTime eHM = DateTime.Parse(dR["終了時"].ToString() + ":" + dR["終了分"].ToString());

                ////ts = eHM - sHM;
                ////if ((ts.TotalMinutes <= 60) && (dR["休憩なし"].ToString() == global.FLGOFF))
                ////{
                ////    global.errNumber = global.eSH;
                ////    global.errMsg = "勤務時間数がマイナスになっています";
                ////    return false;
                ////}
            }

            // 休暇
            if (clsOCR[iH].itm[iX]._Kyuka != string.Empty)
            {
                if (clsOCR[iH].itm[iX]._Kyuka != global.eYUKYU &&
                    clsOCR[iH].itm[iX]._Kyuka != global.eKEKKIN &&
                    clsOCR[iH].itm[iX]._Kyuka != global.eFURIDE &&
                    clsOCR[iH].itm[iX]._Kyuka != global.eFURIKYU)
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休暇欄の数値が不正です";
                    return false;
                }
            }

            // 有休、欠勤、振替休日の場合
            string eMsg = string.Empty;
            if (clsOCR[iH].itm[iX]._Kyuka == global.eYUKYU) eMsg = "この日は有休なので入力できません";
            if (clsOCR[iH].itm[iX]._Kyuka == global.eKEKKIN) eMsg = "この日は欠勤なので入力できません";
            if (clsOCR[iH].itm[iX]._Kyuka == global.eFURIKYU) eMsg = "この日は振替休日なので入力できません";

            if (clsOCR[iH].itm[iX]._Kyuka == global.eYUKYU || 
                clsOCR[iH].itm[iX]._Kyuka == global.eKEKKIN ||
                clsOCR[iH].itm[iX]._Kyuka == global.eFURIKYU)
            {
                if (clsOCR[iH].itm[iX]._Sh != string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = eMsg;
                    return false;
                }

                if (clsOCR[iH].itm[iX]._Sm != string.Empty)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = eMsg;
                    return false;
                }

                if (clsOCR[iH].itm[iX]._eh != string.Empty)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = eMsg;
                    return false;
                }

                if (clsOCR[iH].itm[iX]._em != string.Empty)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = eMsg;
                    return false;
                }

                if (clsOCR[iH].itm[iX]._Kyukei != string.Empty)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = eMsg;
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// エラー情報表示
        /// </summary>
        private void ErrShow()
        {
            if (global.errNumber != global.eNothing)
            {
                lblErrMsg.Visible = true;
                lblErrMsg.Text = global.errMsg;

                //対象年
                if (global.errNumber == global.eYear)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                //対象月
                if (global.errNumber == global.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                //個人番号
                if (global.errNumber == global.eShainNo)
                {
                    txtNo.BackColor = Color.Yellow;
                    txtNo.Focus();
                }

                // オーダーコード
                if (global.errNumber == global.eOrderCode)
                {
                    txtOrderCode.BackColor = Color.Yellow;
                    txtOrderCode.Focus();
                }

                // 開始時
                if (global.errNumber == global.eSH)
                {
                    dg1[cSH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cSH, global.errRow];
                }

                // 開始分
                if (global.errNumber == global.eSM)
                {
                    dg1[cSM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cSM, global.errRow];
                }

                // 終了時
                if (global.errNumber == global.eEH)
                {
                    dg1[cEH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cEH, global.errRow];
                }

                // 終了分
                if (global.errNumber == global.eEM)
                {
                    dg1[cEM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cEM, global.errRow];
                }

                // 休暇
                if (global.errNumber == global.eKyuka)
                {
                    dg1[cKyuka, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cKyuka, global.errRow];
                }

                // 休憩なし
                if (global.errNumber == global.eKyukei)
                {
                    dg1[cKyukei, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cKyukei, global.errRow];
                }
                                
                // 訂正欄 2012/04/17
                if (global.errNumber == global.eTeisei)
                {
                    dg1[cTeisei, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cTeisei, global.errRow];
                }

                // 出勤日数
                if (global.errNumber == global.eDays)
                {
                    txtShuTl.BackColor = Color.Yellow;
                    txtShuTl.Focus();
                }
            }
        }

        private void dg1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
            //カレントレコード更新
            CurDataUpDate(_cI);

            //エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(_cI, clsOCR.Length - 1) == false)
            {
                return;
            }

            //エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (_cI > 0)
            {
                if (ErrCheckMain(0, _cI - 1) == false) return;
            }

            // 勤怠ＣＳＶデータ更新
            OCRDataUpdate(clsOCR);

            // 汎用データ作成
            if (MessageBox.Show("受け渡しデータを作成します。よろしいですか？", "勤怠データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            SaveData();
        }

        /// ---------------------------------------------
        /// <summary>
        ///     受け渡しデータ作成  </summary>
        /// ---------------------------------------------
        private void SaveData()
        {
            string OutFileName = string.Empty;
            try
            {
                //オーナーフォームを無効にする
                this.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();
                
                StringBuilder sb = new StringBuilder();

                // 出力ファイル名
                if (Properties.Settings.Default.PC == global.PC1SELECT.ToString())
                {
                    OutFileName = "PC1_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";
                }
                else
                {
                    OutFileName = "PC2_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";
                }

                ////出力先フォルダがあるか？なければ作成する
                string outPath = global.sDAT;
                if (!System.IO.Directory.Exists(outPath))
                {
                    System.IO.Directory.CreateDirectory(outPath);
                }

                // 出力ファイルインスタンス作成
                StreamWriter outFile = new StreamWriter(outPath + OutFileName, false, System.Text.Encoding.GetEncoding(932));
             
                //レコード件数取得
                int cTotal = clsOCR.Length;
                
                // 受け渡しデータ書き出し
                for (int i = 0; i < clsOCR.Length; i++)
                {
                    //プログレスバー表示
                    frmP.Text = "汎用データ作成中です・・・" + (i + 1).ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = (i + 1) / cTotal * 100;
                    frmP.ProgressStep();

                    // ヘッダ行を書き出す
                    sb.Clear();
                    sb.Append("*").Append(",");
                    sb.Append(clsOCR[i]._StaffCode).Append(",");
                    sb.Append(clsOCR[i]._OrderCode).Append(",");
                    //sb.Append(clsOCR[i]._Year).Append(",");   // 2019/04/04 コメント化
                    sb.Append(Utility.StrToInt(clsOCR[i]._Year) + Properties.Settings.Default.RekiHosei).Append(",");   // 西暦表示 2019/04/04
                    sb.Append(clsOCR[i]._Month.PadLeft(2, '0'));
                    outFile.WriteLine(sb.ToString());

                    // 明細データ
                    string sH = string.Empty;       // 開始時間
                    string eH = string.Empty;       // 終了時間
                    string sRest = string.Empty;    // 休憩時間
                    int rDays = 0;                  // 日付

                    for (int im = 0; im < global._MULTIGYO; im++)
			        {
                        // 日付カウント
                        rDays++;

                        // 開始時間
                        if (clsOCR[i].itm[im]._Sh == string.Empty)
                            sH = string.Empty;
                        else sH = clsOCR[i].itm[im]._Sh.PadLeft(2, '0') + ":" + clsOCR[i].itm[im]._Sm.PadLeft(2, '0');

                        // 終了時間
                        if (clsOCR[i].itm[im]._eh == string.Empty)
                            eH = string.Empty;
                        else eH = clsOCR[i].itm[im]._eh.PadLeft(2, '0') + ":" + clsOCR[i].itm[im]._em.PadLeft(2, '0');

                        // 休憩時間
                        if (clsOCR[i].itm[im]._Kyukei != string.Empty)
                            sRest = clsOCR[i].itm[im]._Kyukei;
                        else if (sH != string.Empty)
                            sRest = "0";
                        else sRest = string.Empty;

                        // 明細行を書き出す
                        sb.Clear();
                        sb.Append(rDays.ToString().PadLeft(2, '0')).Append(",");
                        sb.Append(sH).Append(",");
                        sb.Append(eH).Append(",");
                        sb.Append(sRest).Append(",");
                        sb.Append(clsOCR[i].itm[im]._Kyuka);

                        outFile.WriteLine(sb.ToString());
			        }
                }
                
                // 出力ファイルをクローズします
                outFile.Close();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;
                                
                // 画像ファイルを退避します
                FileMove(_InPath, "*.tif", global.sTIF);

                // CSVファイルを退避します
                FileMove(_InPath, "*.csv", global.sTIF);

                // 履歴ＭＤＢデータ作成
                SaveLastData();

                // 設定月数分経過した過去画像を削除する
                delBackUpFiles(global.sBKDELS, global.sTIF);
                
                // 基準年月以前の勤務票ＭＤＢデータを削除します
                mdbDataDelete(global.sBKDELS);

                //終了
                MessageBox.Show("受け渡しデータが作成されました。", "勤怠データ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //MDBファイル最適化
                mdbCompact();
            }
        }

        /// <summary>
        /// 画像ファイル退避処理
        /// </summary>
        private void tifFileMove(string tifPath)
        {
            //移動先フォルダがあるか？なければ作成する（TIFフォルダ）
            if (!System.IO.Directory.Exists(tifPath)) System.IO.Directory.CreateDirectory(tifPath);

            //画像を退避先フォルダへ移動する            
            foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.tif"))
            {
                File.Move(files, tifPath + @"\" + System.IO.Path.GetFileName(files));
            }
        }

        /// <summary>
        /// ファイル退避処理
        /// </summary>
        /// <param name="fromPath">移動元パス</param>
        /// <param name="fileType">ファイルタイプ(*.拡張子)</param>
        /// <param name="toPath">移動先パス</param>
        private void FileMove(string fromPath, string fileType, string toPath)
        {
            // 移動先フォルダがあるか？なければ作成する
            if (!System.IO.Directory.Exists(toPath)) System.IO.Directory.CreateDirectory(toPath);

            // ファイルを退避先フォルダへ移動する            
            foreach (string files in System.IO.Directory.GetFiles(fromPath, fileType))
            {
                File.Move(files, toPath + @"\" + System.IO.Path.GetFileName(files));
            }
        }

        /// <summary>
        /// 設定月数分経過した過去画像を削除する    
        /// </summary>
        private void imageDelete(int sdel)
        {
            //削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (sdel == 0) return;

            try
            {
                //削除年月の取得
                DateTime delDate = DateTime.Today.AddMonths(sdel * (-1));
                int _dYY = delDate.Year;            //基準年
                int _dMM = delDate.Month;           //基準月
                int _dYYMM = _dYY * 100 + _dMM;     //基準年月
                int _DataYYMM;
                string fileYYMM;

                // 設定月数分経過した過去画像を削除する
                // ①スタッフ
                foreach (string files in System.IO.Directory.GetFiles(global.sTIF, "*.tif"))
                {
                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }

                // ②パートタイマー
                foreach (string files in System.IO.Directory.GetFiles(global.sTIF2, "*.tif"))
                {
                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("過去画像削除中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
                return;
            }
        }

        /// <summary>
        /// 設定月数を経過したバックアップファイルを削除します
        /// </summary>
        /// <param name="sSpan">経過月数</param>
        /// <param name="sDir">ディレクトリ</param>
        private void delBackUpFiles(int sSpan, string sDir)
        {
            //削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (sSpan == int.Parse(global.FLGOFF)) return;

            if (System.IO.Directory.Exists(sDir))
            {
                if (sSpan > 0)
                {
                    foreach (string fName in System.IO.Directory.GetFiles(sDir, "*"))
                    {
                        string f = System.IO.Path.GetFileName(fName);

                        // ファイル名の長さを検証（日付情報あり？）
                        if (f.Length > 12)
                        {
                            // ファイル名から日付部分を取得します
                            DateTime dt;
                            string stDt = f.Substring(0, 4) + "/" + f.Substring(4, 2) + "/" + f.Substring(6, 2);
                            if (DateTime.TryParse(stDt, out dt))
                            {
                                // 設定月数を加算した日付を取得します
                                DateTime Fdt = dt.AddMonths(sSpan);

                                // 今日の日付と比較して設定月数を加算したファイル日付が既に経過している場合、ファイルを削除します
                                if (DateTime.Today.CompareTo(Fdt) == 1)
                                {
                                    System.IO.File.Delete(fName);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 基準年月以前の勤務票ＭＤＢデータを削除します
        /// </summary>
        /// <param name="sSpan">過去データの保存月数</param>
        private void mdbDataDelete(int sSpan)
        {
            //削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (sSpan == int.Parse(global.FLGOFF)) return;

            //削除基準年月の取得
            DateTime dt = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");
            DateTime delDate = dt.AddMonths(sSpan * (-1));
            int _dYY = delDate.Year;            //基準年
            int _dMM = delDate.Month;           //基準月
            int _dYYMM = _dYY * 100 + _dMM;     //基準年月

            // 基準年月以前の勤務票ＭＤＢデータを削除します
            OCRData ocr = new OCRData();
            ocr.DataDelete(_dYYMM);

            if (ocr.sCom.Connection.State == ConnectionState.Open)
                ocr.sCom.Connection.Close();
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Tag.ToString() != END_MAKEDATA)
            {
                if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                // カレントデータ更新
                CurDataUpDate(_cI);

                // 勤怠ＣＳＶデータ更新
                OCRDataUpdate(clsOCR);
            }

            // 解放する
            this.Dispose();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            //フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        /// <summary>
        /// MDBファイルを最適化する
        /// </summary>
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();

                string OldDb = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB + global.MDBFILENAME;

                string NewDb = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB + global.MDBTEMPFILE;

                // 最適化した一時MDBを作成する
                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB + global.MDBBACKUP);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB + global.MDBFILENAME,
                                    Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB + global.MDBBACKUP);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB + global.MDBTEMPFILE,
                                    Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB + global.MDBFILENAME);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void txtNo_Leave(object sender, EventArgs e)
        {
            txtNo.Text = txtNo.Text.PadLeft(6, '0');

            // スタッフ情報取得
            GetStaffData(txtNo.Text);
        }

        private void txtNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                frmStaffSelect frmS = new frmStaffSelect(ms);
                frmS.ShowDialog();
                if (frmS._msCode != string.Empty)
                {
                    txtNo.Text = frmS._msCode;
                    lblName.Text = frmS._msName;
                    txtShozokuCode.Text = frmS._msTenpoCode;
                    lblShozoku.Text = frmS._msTenpoName;
                    txtOrderCode.Text = frmS._msOrderCode;
                }
                frmS.Dispose();
            }
        }

        /// <summary>
        /// 勤務票CSVデータを追加する
        /// </summary>
        private void AddNewData()
        {
            // IDを取得します
            string _ID = string.Format("{0:0000}", DateTime.Today.Year) +
                         string.Format("{0:00}", DateTime.Today.Month) +
                         string.Format("{0:00}", DateTime.Today.Day) +
                         string.Format("{0:00}", DateTime.Now.Hour) +
                         string.Format("{0:00}", DateTime.Now.Minute) +
                         string.Format("{0:00}", DateTime.Now.Second) + "001";

            // 出力ファイルインスタンス作成
            StreamWriter outFile = new StreamWriter(_InPath + _ID + ".csv", false, System.Text.Encoding.GetEncoding(932));

            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();

                // ヘッダ情報
                sb.Append(_ID).Append(",");             // ファイル番号
                sb.Append(string.Empty).Append(",");    // 画像ファイル名
                sb.Append(string.Empty).Append(",");    // シートＩＤ
                sb.Append(string.Empty).Append(",");    // スタッフコード
                sb.Append(global.sYear).Append(",");    // 年
                sb.Append(global.sMonth).Append(",");   // 月
                sb.Append(string.Empty).Append(",");   // オーダーコード
                sb.Append(string.Empty).Append(",");    // 出勤日数

                // 明細情報
                for (int i = 0; i < global._MULTIGYO; i++)
                {
                    sb.Append(string.Empty).Append(",");    // 休暇
                    sb.Append(string.Empty).Append(",");    // 開始時
                    sb.Append(string.Empty).Append(",");    // 開始分
                    sb.Append(string.Empty).Append(",");    // 終了時
                    sb.Append(string.Empty).Append(",");    // 終了分
                    sb.Append(string.Empty).Append(",");    // 休憩時間
                    sb.Append(string.Empty);                // 訂正

                    if (i != 30) sb.Append(",");
                }

                outFile.WriteLine(sb.ToString());                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票データ新規登録処理", MessageBoxButtons.OK);
            }
            finally
            {
                outFile.Close();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の勤務票データを削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // 画像ファイルを削除する
            imgDelete(_InPath + clsOCR[_cI]._ImageName);

            // ＣＳＶファイルを削除する
            imgDelete(_InPath + clsOCR[_cI].OCRDATA_ID + ".csv");
            //imgDelete(_InPath + System.IO.Path.GetFileNameWithoutExtension(_InPath + clsOCR[_cI]._ImageName) + ".csv");

            //テーブル件数カウント：ゼロならばプログラム終了
            if (clsOCR.Length == 1)
            {
                MessageBox.Show("データが全て削除されました。プログラムを終了します", "勤務票削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                Environment.Exit(0);
            }

            // ＯＣＲデータクラス
            OCRDataDelete(clsOCR, _cI);

            //エラー情報初期化
            ErrInitial();

            //レコードを表示
            if (clsOCR.Length - 1 < _cI) _cI = clsOCR.Length - 1;
            DataShow(_cI, clsOCR, this.dg1);
        }

        /// <summary>
        /// ファイル削除
        /// </summary>
        /// <param name="imgPath">画像ファイルパス</param>
        private void imgDelete(string imgPath)
        {
            // ファイルを削除する
            if (System.IO.File.Exists(imgPath))
            {
                System.IO.File.Delete(imgPath);
            }
        }

        /// <summary>
        /// OCRDATAクラス配列リサイズ
        /// </summary>
        /// <param name="r">OCRDATAクラス</param>
        /// <param name="i">インデックス</param>
        private void OCRDataDelete(OCRData [] r, int i)
        {
            // 最終データでないとき
            if (i < (r.Length - 1))
            {
                for (int j = i; j < r.Length - 1; j++)
                {
                    r[j].OCRDATA_ID = r[j + 1].OCRDATA_ID;
                    r[j]._SheetID = r[j + 1]._SheetID;
                    r[j]._StaffCode = r[j + 1]._StaffCode;
                    r[j]._Year = r[j + 1]._Year;
                    r[j]._Month = r[j + 1]._Month;
                    r[j]._OrderCode = r[j + 1]._OrderCode;
                    r[j]._ShuDays = r[j + 1]._ShuDays;
                    r[j]._ImageName = r[j + 1]._ImageName;

                    for (int m = 0; m < global._MULTIGYO; m++)
                    {
                        r[j].itm[m]._Day = r[j + 1].itm[m]._Day;
                        r[j].itm[m]._Sh = r[j + 1].itm[m]._Sh;
                        r[j].itm[m]._Sm = r[j + 1].itm[m]._Sm;
                        r[j].itm[m]._eh = r[j + 1].itm[m]._eh;
                        r[j].itm[m]._em = r[j + 1].itm[m]._em;
                        r[j].itm[m]._Kyuka = r[j + 1].itm[m]._Kyuka;
                        r[j].itm[m]._Kyukei = r[j + 1].itm[m]._Kyukei;
                        r[j].itm[m]._teisei = r[j + 1].itm[m]._teisei;
                    }
                }
            }

            // 新しいクラス配列の要素数
            int newSize = clsOCR.Length - 1;
            Array.Resize(ref clsOCR, newSize);
        }

        /// <summary>
        /// 勤務票CSVデータを更新保存する
        /// </summary>
        private void OCRDataUpdate(OCRData[] r)
        {
            for (int i = 0; i < r.Length; i++)
            {
                // ファイル名を取得します
                string _ID = r[i].OCRDATA_ID;

                // 出力ファイルインスタンス作成
                StreamWriter outFile = new StreamWriter(_InPath + _ID + ".csv", false, System.Text.Encoding.GetEncoding(932));

                StringBuilder sb = new StringBuilder();

                try
                {
                    sb.Clear();

                    // ヘッダ情報
                    sb.Append(r[i].OCRDATA_ID).Append(",");     // ファイル番号
                    sb.Append(r[i]._ImageName).Append(",");     // 画像ファイル名
                    sb.Append(r[i]._SheetID).Append(",");       // シートＩＤ
                    sb.Append(r[i]._StaffCode).Append(",");     // スタッフコード
                    sb.Append(r[i]._Year).Append(",");          // 年
                    sb.Append(r[i]._Month).Append(",");         // 月
                    sb.Append(r[i]._OrderCode).Append(",");     // オーダーコード
                    sb.Append(r[i]._ShuDays).Append(",");       // 出勤日数

                    // 明細情報
                    for (int j = 0; j < global._MULTIGYO; j++)
                    {
                        sb.Append(r[i].itm[j]._Kyuka).Append(",");  // 休暇
                        sb.Append(r[i].itm[j]._Sh).Append(",");     // 開始時
                        sb.Append(r[i].itm[j]._Sm).Append(",");     // 開始分
                        sb.Append(r[i].itm[j]._eh).Append(",");     // 終了時
                        sb.Append(r[i].itm[j]._em).Append(",");     // 終了分
                        sb.Append(r[i].itm[j]._Kyukei).Append(","); // 休憩時間
                        sb.Append(r[i].itm[j]._teisei);             // 訂正

                        if (j != (global._MULTIGYO - 1)) sb.Append(",");
                    }

                    outFile.WriteLine(sb.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "勤務票CSVデータを更新保存", MessageBoxButtons.OK);
                }
                finally
                {
                    outFile.Close();
                }
            }
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     勤務票データをＭＤＢへ登録します </summary>
        ///--------------------------------------------------------
        private void SaveLastData()
        {
            for (int i = 0; i < clsOCR.Length; i++)
            {
                // 登録済みの同年、同月、同スタッフの勤務票データを削除する
                OCRData cr = new OCRData();
                cr.DataDelete((int.Parse(clsOCR[i]._Year) + Properties.Settings.Default.RekiHosei).ToString(),    // 年（西暦）
                               clsOCR[i]._Month,        // 月
                               clsOCR[i]._StaffCode);   // 個人番号

                // 勤務票データ登録
                OCRData cr2 = new OCRData();
                cr2.ClsToMdb(clsOCR, i);
            }
        }
    }
}

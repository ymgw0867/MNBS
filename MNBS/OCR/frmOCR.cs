﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Leadtools;
using Leadtools.Twain;
using Leadtools.Codecs;
using Leadtools.WinForms;
using Leadtools.WinForms.CommonDialogs.File;
using Leadtools.ImageProcessing;
using IdrFormEngine;
using MNBS.Common;

namespace MNBS.OCR
{
    public partial class frmOCR : Form
    {
        public frmOCR(int pcSel, int OcrSel)
        {
            InitializeComponent();
            InitClass();
            
            //ＰＣ指定
            _pcSel = pcSel;

            // ＯＣＲ指定
            _OcrSel = OcrSel;
        }

        // ＰＣ指定
        private int _pcSel;

        // ＯＣＲ指定
        private int _OcrSel;

        // TWAINでの取得用にTwainSessionを宣言します。
        private TwainSession _twainSession;

        // ＯＣＲ変換後のＣＳＶファイルと画像を登録するフォルダパス名変数を宣言します。
        private string _ocrPath;

        // 出力ファイル名保存用変数を宣言します。
        private string _fileName;

        // RasterImageViewerコントロールを宣言します。
        private RasterImageViewer _viewer;

        // イメージロード用RasterCodecsを宣言します。
        private RasterCodecs _codecs;

        // 出力ファイルのフォーマット保存用変数を宣言します。
        private RasterImageFormat _fileFormat;
        int _bitsPerPixel = 1;

        // 取得ページ数の保存用変数を宣言します。
        private int _pageNo;

        //スキャナから出力された画像枚数
        private int _sNumber = 0;

        //スキャナから出力された画像ファイル数
        private int _sFileNumber = 0;

        /// <summary>
        /// アプリケーションの初期化処理を行います。
        /// </summary>
        private void InitClass()
        {
            // フォームのタイトルを設定します。
            this.Text = "TWAIN 取得 【ＯＣＲ勤務票読み取り】";

            //自分自身のバージョン情報を取得する　2011/03/25
            //System.Diagnostics.FileVersionInfo ver =
            //    System.Diagnostics.FileVersionInfo.GetVersionInfo(
            //    System.Reflection.Assembly.GetExecutingAssembly().Location);

            //キャプションにバージョンを追加　2011/03/25
            //Messager.Caption += " ver " + ver.FileMajorPart.ToString() + "." + ver.FileMinorPart.ToString();

            //Text = Messager.Caption;

            // ロック解除状態を確認します。
            //Support.Unlock(false);

            // RasterImageViewerコントロールを初期化します。
            _viewer = new RasterImageViewer();
            //_viewer.Dock = DockStyle.Fill;
            _viewer.BackColor = Color.DarkGray;
            Controls.Add(_viewer);
            _viewer.BringToFront();
            _viewer.Visible = false;

            // コーデックパスを設定します。
            RasterCodecs.Startup();

            // RasterCodecsオブジェクトを初期化します。
            _codecs = new RasterCodecs();

            if (TwainSession.IsAvailable(this))
            {
                // TwainSessionオブジェクトを初期化します。
                _twainSession = new TwainSession();

                // TWAIN セッションを初期化します。 
                _twainSession.Startup(this, "FKDL", "LEADTOOLS", "Version16.5J", "LeadTools Twain", TwainStartupFlags.None);
            }
            else
            {
                //_miFileAcquire.Enabled = false;
                //_miFileSelectSource.Enabled = false;
            }

            // 各値を初期化します。
            _fileName = string.Empty;
            _fileFormat = RasterImageFormat.Tif;
            _pageNo = 1;
            _sFileNumber = 0;

            //UpdateMyControls();
            UpdateStatusBarText();
        }

        private void frmOCR_Load(object sender, EventArgs e)
        {
        }

        private void frmOCR_Shown(object sender, EventArgs e)
        {
        }

        private void frmOCR_FormClosing(object sender, FormClosingEventArgs e)
        {
            CleanUp();
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();

            // 勤務票スキャン処理
            if (_OcrSel == global.SCAN_SELECT)
            {
                while (true)
                {
                    ScanOcr();

                    if (_viewer.Image != null)
                    {
                        string msg = "続けて読込を行いますか？" + Environment.NewLine + "読込枚数 ： " + _viewer.Image.PageCount.ToString() + "枚";
                        if (MessageBox.Show(msg, "TWAIN取得", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                            break;
                    }
                    else
                    {
                        MessageBox.Show("処理を中断しました", "ＯＣＲ変換処理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                }
            }

            // ＯＣＲ処理
            this.Show();

            string inPath = string.Empty;
            string outPath = string.Empty;
            string wrkPath = string.Empty;
            string formatFileName = string.Empty;

            // ＯＣＲ出力先パス
            if (_pcSel == global.PC1SELECT)
                outPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.DATAS1;
            else outPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.DATAS2;

            // ＯＣＲ処理方法による入力フォルダの指定
            switch (_OcrSel)
	        {
                // スキャナ
                case global.SCAN_SELECT:

                    // ＯＣＲ書式ファイル
                    formatFileName = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathFMT + Properties.Settings.Default.fmt300;
                    inPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathSCAN;  // 画像フォルダ
                    wrkPath = string.Empty;
                    break;
                    
                // ＦＡＸ画像
                case global.FAX_SELECT:

                    // ＯＣＲ書式ファイル
                    formatFileName = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathFMT + Properties.Settings.Default.fmt200;
                    inPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathTRAY;
                    wrkPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathWORK;
                    break;
	        }

            // 勤務票画像を確認
            string[] intif = System.IO.Directory.GetFiles(inPath, "*.tif");
            if (intif.Length == 0)
            {
                MessageBox.Show("勤務票画像データがありません。処理を終了します。", "画像確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Close();
            }

            // マルチTifをページ毎に分割・解像度を調整します
            MultiTif(_OcrSel, inPath, wrkPath);

            // 画像数を取得します 
            var sTif = System.IO.Directory.GetFileSystemEntries(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathREAD, "*.tif");

            // OCR処理を起動します
            ocrMain(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathREAD,
                    Properties.Settings.Default.PathInst + Properties.Settings.Default.PathNG,
                    outPath,
                    formatFileName, 
                    sTif.Length);

            // 終了
            this.Close();
        }

        /// <summary>
        /// スキャナより勤務票をスキャンして画像を取得します
        /// </summary>
        private void ScanOcr()
        {
            //出力先パス初期化
            _ocrPath = string.Empty;

            try
            {
                RasterSaveDialogFileFormatsList saveDlgFormatList = new RasterSaveDialogFileFormatsList(RasterDialogFileFormatDataContent.User);

                string tifPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathSCAN;
                _fileName = tifPath + string.Format("{0:0000}", DateTime.Today.Year) +
                                                                    string.Format("{0:00}", DateTime.Today.Month) +
                                                                    string.Format("{0:00}", DateTime.Today.Day) +
                                                                    string.Format("{0:00}", DateTime.Now.Hour) +
                                                                    string.Format("{0:00}", DateTime.Now.Minute) +
                                                                    string.Format("{0:00}", DateTime.Now.Second) + ".tif";

                ///以下、TWAIN取得関連 //////////////////////////////////////////////////////////////////////

                _fileFormat = RasterImageFormat.CcittGroup4;
                _bitsPerPixel = 1;

                string pathName = System.IO.Path.GetDirectoryName(_fileName);
                if (System.IO.Directory.Exists(pathName))
                {
                    // ページカウンタを初期化します。
                    _pageNo = 1;

                    // 出力ファイルカウンタをインクリメントします。
                    _sFileNumber++;

                    // AcquirePageイベントハンドラを設定します。
                    _twainSession.AcquirePage += new EventHandler<TwainAcquirePageEventArgs>(_twain_AcquirePage);

                    // Acquire pages
                    _twainSession.Acquire(TwainUserInterfaceFlags.Show);

                    // AcquirePageイベントハンドラを削除します。
                    _twainSession.AcquirePage -= new EventHandler<TwainAcquirePageEventArgs>(_twain_AcquirePage);
                }
                else
                {
                    MessageBox.Show("ファイル名の書式が正しくありません。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                _twainSession.Shutdown();
                _twainSession.Startup(this, "GrapeCity Inc.", "LEADTOOLS", "Ver.16.5J", "LEADTOOLS TWAIN 取得 サンプル", TwainStartupFlags.None);
            }
            finally
            {
                UpdateStatusBarText();
            }
        }

        /// <summary>
        /// AcquirePageイベント処理
        /// </summary>
        private void _twain_AcquirePage(object sender, TwainAcquirePageEventArgs e)
        {
            try
            {
                if (e.Image != null)
                {
                    // 選択されているファイルフォーマットがマルチページに対応しているかどうかを確認します。
                    if ((_fileFormat == RasterImageFormat.Tif) || (_fileFormat == RasterImageFormat.Ccitt) ||
                       (_fileFormat == RasterImageFormat.CcittGroup31Dim) || (_fileFormat == RasterImageFormat.CcittGroup32Dim) ||
                       (_fileFormat == RasterImageFormat.CcittGroup4) || (_fileFormat == RasterImageFormat.TifCmp) ||
                       (_fileFormat == RasterImageFormat.TifCmw) || (_fileFormat == RasterImageFormat.TifCmyk) ||
                       (_fileFormat == RasterImageFormat.TifCustom) ||
                       (_fileFormat == RasterImageFormat.TifJ2k) || (_fileFormat == RasterImageFormat.TifJbig) ||
                       (_fileFormat == RasterImageFormat.TifJpeg) || (_fileFormat == RasterImageFormat.TifJpeg411) ||
                       (_fileFormat == RasterImageFormat.TifJpeg422) || (_fileFormat == RasterImageFormat.TifLead1Bit) ||
                       (_fileFormat == RasterImageFormat.TifLzw) || (_fileFormat == RasterImageFormat.TifLzwCmyk) ||
                       (_fileFormat == RasterImageFormat.TifLzwYcc) || (_fileFormat == RasterImageFormat.TifPackBits) ||
                       (_fileFormat == RasterImageFormat.TifPackBitsCmyk) || (_fileFormat == RasterImageFormat.TifPackbitsYcc) ||
                       (_fileFormat == RasterImageFormat.TifUnknown) || (_fileFormat == RasterImageFormat.TifYcc) ||
                       (_fileFormat == RasterImageFormat.Gif))
                    { 
                        // ファイル拡張子の保存変数を初期化します。
                        string ext = string.Empty;

                        // ファイル名に拡張子を含んでいない場合、拡張子を追加します。
                        if (System.IO.Path.HasExtension(_fileName))
                            ext = System.IO.Path.GetExtension(_fileName);

                        // 保存ファイル名に連番を付加します。
                        string tmpFileName = System.IO.Path.GetFileNameWithoutExtension(_fileName);
                        string tmpDirName = System.IO.Path.GetDirectoryName(_fileName);
                        string newFileName = string.Format("{0}\\{1}{2:000}{3}", tmpDirName, tmpFileName, _sFileNumber, ext);
                        // 取得したページを保存します。
                        _codecs.Save(e.Image, newFileName, _fileFormat, _bitsPerPixel, 1, 1, 1, CodecsSavePageMode.Append);
                    }
                    else
                    {

                        // マルチページに対応していないフォーマットの場合、ファイルに番号を付加して保存します。
                        // ファイル拡張子の保存変数を初期化します。
                        string ext = string.Empty;

                        // ファイル名に拡張子を含んでいない場合、拡張子を追加します。
                        if (System.IO.Path.HasExtension(_fileName))
                            ext = System.IO.Path.GetExtension(_fileName);

                        // 保存ファイル名にページ番号を付加します。
                        string tmpFileName = System.IO.Path.GetFileNameWithoutExtension(_fileName);
                        string tmpDirName = System.IO.Path.GetDirectoryName(_fileName);
                        string newFileName = string.Format("{0}\\{1}{2:000}{3}", tmpDirName, tmpFileName, _pageNo, ext);

                        _codecs.Save(e.Image, newFileName, _fileFormat, _bitsPerPixel);

                        // ページ数のカウンタをインクリメントします。
                        _pageNo++;
                    }

                    // 取得ページをビューアに表示します。
                    if (_viewer.Image == null)
                    {
                        _viewer.Image = e.Image;
                    }
                    else
                    {
                        _viewer.Image.AddPage(e.Image);
                        _viewer.Image.Page = _viewer.Image.PageCount;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CleanUp()
        {
            RasterCodecs.Shutdown();

            if (_twainSession != null)
            {
                try
                {
                    _twainSession.Shutdown();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        /// <summary>
        /// 解像度の調整
        /// </summary>
        /// <param name="imgPath">入力ファイルパス</param>
        private void ImageResize(string imgPath)
        {
        }

        /// <summary>
        /// ステータスバーのページ表示を更新します。
        /// </summary>
        private void UpdateStatusBarText()
        {
            if (_viewer.Image != null)
                this.label1.Text = string.Format("ページ {0} / {1}", _viewer.Image.Page, _viewer.Image.PageCount);
            else
                this.label1.Text = "準備済み";
        }

        /// <summary>
        /// マルチフレームの画像ファイルを頁ごとに分割する
        /// </summary>
        /// <param name="InPath">画像ファイルパス</param>
        private void MultiTif(int ocrSel, string InPath, string WrkPath)
        {
            //スキャン出力画像を確認
            string[] intif = System.IO.Directory.GetFiles(InPath, "*.tif");
            if (intif.Length == 0)
            {
                MessageBox.Show("ＯＣＲ変換処理対象の画像ファイルが指定フォルダ " + InPath + " に存在しません", "スキャナ画像確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // READフォルダがなければ作成する
            string rPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathREAD;
            if (System.IO.Directory.Exists(rPath) == false)
                System.IO.Directory.CreateDirectory(rPath);

            // READフォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            foreach (string files in System.IO.Directory.GetFiles(rPath, "*"))
            {
                System.IO.File.Delete(files);
            }

            // FAX画像処理時にWORKフォルダがなければ作成する
            if (ocrSel == global.FAX_SELECT)
            {
                if (System.IO.Directory.Exists(WrkPath) == false)
                    System.IO.Directory.CreateDirectory(WrkPath);

                // WORKフォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
                foreach (string files in System.IO.Directory.GetFiles(WrkPath, "*"))
                {
                    System.IO.File.Delete(files);
                }
            }

            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();
            _pageNo = 0;
            string fnm = string.Empty;

            // １．マルチTIFを分解して画像ファイルをREADフォルダかWorkフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                // 頁数を取得
                int _fd_count = leadImg.PageCount;

                // 頁ごとに読み出す
                for (int i = 1; i <= _fd_count; i++)
                {
                    // ファイル名（日付時間部分）
                    string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                            string.Format("{0:00}", DateTime.Today.Month) +
                            string.Format("{0:00}", DateTime.Today.Day) +
                            string.Format("{0:00}", DateTime.Now.Hour) +
                            string.Format("{0:00}", DateTime.Now.Minute) +
                            string.Format("{0:00}", DateTime.Now.Second);

                    // ファイル名設定
                    _pageNo++;

                    if (ocrSel == global.SCAN_SELECT)
                        fnm = rPath + fName + string.Format("{0:000}", _pageNo) + ".tif";   // ＳＣＡＮのときREADフォルダへ出力
                    else fnm = WrkPath + fName + string.Format("{0:000}", _pageNo) + ".tif";    // スキャナのときWORKフォルダへ出力

                    // 画像保存
                    cs.Save(leadImg, fnm, RasterImageFormat.CcittGroup4, 0, i, i, 1, CodecsSavePageMode.Insert);
                }
            }

            // 2．InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }

            // ３.ＦＡＸ画像解像度の調整
            if (_OcrSel == global.FAX_SELECT)
            {
                foreach (string files in System.IO.Directory.GetFiles(WrkPath, "*.tif"))
                {
                    ////// 描画時に使用される速度、品質、およびスタイルを制御します。 
                    ////RasterPaintProperties prop = new RasterPaintProperties();
                    ////prop = RasterPaintProperties.Default;
                    ////prop.PaintDisplayMode = RasterPaintDisplayModeFlags.Resample;
                    ////leadImg.PaintProperties = prop;

                    // 画像読み出す
                    RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                    // 画像リサイズ(DPI:X:200,Y:200)
                    ResizeCommand rs = new ResizeCommand();
                    rs.DestinationImage = leadImg;
                    rs.DestinationImage.XResolution = 200;
                    rs.DestinationImage.YResolution = 200;
                    rs.Flags = RasterSizeFlags.None;
                    rs.Run(leadImg);

                    // A4サイズへ変更
                    Leadtools.ImageProcessing.SizeCommand sz = new SizeCommand();
                    sz.Width = 1693;
                    sz.Height = 2338;
                    sz.Run(leadImg);

                    // 画像保存
                    cs.Save(leadImg, files, RasterImageFormat.Tif, 0, 1, -1, 1, CodecsSavePageMode.Overwrite);
                }

                // 変換済みtif画像をREADフォルダへ移動する            
                foreach (string files in System.IO.Directory.GetFiles(WrkPath, "*.tif"))
                { 
                    System.IO.File.Move(files, rPath + "\\" + System.IO.Path.GetFileName(files));
                }
            }
        }

        /// <summary>
        /// OCR処理を実施します
        /// </summary>
        /// <param name="InPath">入力パス</param>
        /// <param name="NgPath">NG出力パス</param>
        /// <param name="rePath">OCR変換結果出力パス</param>
        /// <param name="FormatName">書式ファイル名</param>
        /// <param name="fCnt">書式ファイルの件数</param>
        private void ocrMain(string InPath, string NgPath, string rePath, string FormatName, int fCnt)
        {
            IEngine en = null;		            // OCRエンジンのインスタンスを保持
            string ocr_csv = string.Empty;      // OCR変換出力CSVファイル
            int _okCount = 0;                   // OCR変換画像枚数
            int _ngCount = 0;                   // フォーマットアンマッチ画像枚数
            string fnm = string.Empty;          // ファイル名

            try
            {
                // 指定された出力先フォルダがなければ作成する
                if (System.IO.Directory.Exists(rePath) == false)
                    System.IO.Directory.CreateDirectory(rePath);

                // 指定されたNGの場合の出力先フォルダがなければ作成する
                if (System.IO.Directory.Exists(NgPath) == false)
                    System.IO.Directory.CreateDirectory(NgPath);

                // OCRエンジンのインスタンスの生成・取得
                en = EngineFactory.GetEngine();
                if (en == null)
                {
                    // エンジンが他で取得されている場合は、Release() されるまで取得できない
                    System.Console.WriteLine("SDKは使用中です");
                    return;
                }

                //オーナーフォームを無効にする
                this.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                IFormatList FormatList;
                IFormat Format;
                IField Field;
                int nPage;
                int ocrPage = 0;
                int fileCount = 0;

                // フォーマットのロード・設定
                FormatList = en.FormatList;
                FormatList.Add(FormatName);

                // tifファイルの認識
                foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
                {
                    nPage = 1;
                    while (true)
                    {
                        try
                        {
                            // 対象画像を設定する
                            en.SetBitmap(files, nPage);

                            //プログレスバー表示
                            fileCount++;
                            frmP.Text = "OCR変換処理実行中　" + fileCount.ToString() + "/" + fCnt.ToString();
                            frmP.progressValue = fileCount * 100 / fCnt;
                            frmP.ProgressStep();
                        }
                        catch (IDRException ex)
                        {
                            // ページ読み込みエラー
                            if (ex.No == ErrorCode.IDR_ERROR_FORM_FILEREAD)
                            {
                                // ページの終了
                                break;
                            }
                            else
                            {
                                // 例外のキャッチ
                                MessageBox.Show("例外が発生しました：Error No ={0:X}", ex.No.ToString());
                            }
                        }

                        //////Console.WriteLine("-----" + strImageFile + "の" + nPage + "ページ-----");
                        // 現在ロードされている画像を自動的に傾き補正する
                        en.AutoSkew();

                        // 傾き角度の取得
                        double angle = en.GetSkewAngle();
                        //////System.Console.WriteLine("時計回りに" + angle + "度傾き補正を行いました");

                        try
                        {
                            // 現在ロードされている画像を自動回転してマッチする番号を取得する
                            Format = en.MatchFormatRotate();
                            int direct = en.GetRotateAngle();

                            //画像ロード
                            RasterCodecs.Startup();
                            RasterCodecs cs = new RasterCodecs();
                            //RasterImage img;

                            // 描画時に使用される速度、品質、およびスタイルを制御します。 
                            //RasterPaintProperties prop = new RasterPaintProperties();
                            //prop = RasterPaintProperties.Default;
                            //prop.PaintDisplayMode = RasterPaintDisplayModeFlags.Resample;
                            //leadImg.PaintProperties = prop;

                            RasterImage img = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, 1);

                            RotateCommand rc = new RotateCommand();
                            rc.Angle = (direct) * 90 * 100;
                            rc.FillColor = new RasterColor(255, 255, 255);
                            rc.Flags = RotateCommandFlags.Resize;
                            rc.Run(img);
                            //rc.Run(leadImg.Image);

                            //cs.Save(leadImg.Image, files, RasterImageFormat.Tif, 0, 1, 1, 1, CodecsSavePageMode.Overwrite);
                            cs.Save(img, files, RasterImageFormat.CcittGroup4, 0, 1, 1, 1, CodecsSavePageMode.Overwrite);

                            // マッチしたフォーマットに登録されているフィールド数を取得
                            int fieldNum = Format.NumOfFields;
                            int matchNum = Format.FormatNo + 1;
                            //////System.Console.WriteLine(matchNum + "番目のフォーマットがマッチ");
                            int i = 1;
                            ocr_csv = string.Empty;

                            // ファイルの先頭フィールドにファイル番号をセットします
                            ocr_csv = System.IO.Path.GetFileNameWithoutExtension(files) + ",";

                            // ファイルに画像ファイル名フィールドを付加します
                            ocr_csv += System.IO.Path.GetFileName(files);

                            // 認識されたフィールドを順次読み出します
                            Field = Format.Begin();
                            while (Field != null)
                            {
                                //カンマ付加
                                if (ocr_csv != string.Empty)
                                    ocr_csv += ",";

                                // 指定フィールドを認識し、テキストを取得
                                string strText = Field.ExtractFieldText();
                                ocr_csv += strText;

                                // 次のフィールドの取得
                                Field = Format.Next();
                                i += 1;
                            }

                            //出力ファイル
                            System.IO.StreamWriter outFile = new System.IO.StreamWriter(InPath + System.IO.Path.GetFileNameWithoutExtension(files) + ".csv", false, System.Text.Encoding.GetEncoding(932));
                            outFile.WriteLine(ocr_csv);
                            outFile.Close();

                            //OCR変換枚数カウント
                            _okCount++;
                        }
                        catch (IDRWarning ex)
                        {
                            // Engine.MatchFormatRotate() で
                            // フォーマットにマッチしなかった場合の処理
                            if (ex.No == ErrorCode.IDR_WARN_FORM_NO_MATCH)
                            {
                                // NGフォルダへ移動する
                                System.IO.File.Move(files, NgPath + "E" + System.IO.Path.GetFileName(files));
                                
                                //NG枚数カウント
                                _ngCount++;
                            }
                        }

                        ocrPage++;
                        nPage += 1;
                    }
                }

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                string finMessage = string.Empty;
                StringBuilder sb = new StringBuilder();

                // NGメッセージ
                if (_ngCount > 0)
                    MessageBox.Show("OCR認識を正常に行うことが出来なかった画像があります。確認してください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 終了メッセージ
                sb.Clear();
                sb.Append("ＯＣＲ認識処理が終了しました。");
                sb.Append("引き続き修正確認＆受け渡しデータ作成を行ってください。");
                sb.Append(Environment.NewLine);
                sb.Append(Environment.NewLine);
                sb.Append("OK件数 : ");
                sb.Append(_okCount.ToString());
                sb.Append(Environment.NewLine);
                sb.Append("NG件数 : ");
                sb.Append(_ngCount.ToString());
                sb.Append(Environment.NewLine);

                MessageBox.Show(sb.ToString(), "処理終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // OCR変換画像とCSVデータをOCR結果出力フォルダへ移動する            
                foreach (string files in System.IO.Directory.GetFiles(InPath, "*.*"))
                {
                    System.IO.File.Move(files, rePath + System.IO.Path.GetFileName(files));
                }

                FormatList.Delete(0);
            }
            catch (System.Exception ex)
            {
                // 例外のキャッチ
                string errMessage = string.Empty;
                errMessage += "System例外が発生しました：" + Environment.NewLine;
                errMessage += "必要なDLL等が実行モジュールと同ディレクトリに存在するか確認してください。：" + Environment.NewLine;
                errMessage += ex.Message.ToString();
                MessageBox.Show(errMessage, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                en.Release();
            }
        }
    }
}

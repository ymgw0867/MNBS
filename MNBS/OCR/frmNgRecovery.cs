using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
    public partial class frmNgRecovery : Form
    {
        public frmNgRecovery()
        {
            InitializeComponent();
        }

        string _InPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.PathNG;
        string _OutPath = Utility.exisInFileDir();

        clsNG[] ngf;

        private void frmNgRecovery_Load(object sender, EventArgs e)
        {
            Common.Utility.WindowsMaxSize(this, this.Width, this.Height);
            Common.Utility.WindowsMinSize(this, this.Width, this.Height);

            // NGリスト
            GetNgList();

            // ボタン
            BtnEnabled_false();
        }

        private void BtnEnabled_false()
        {
            btnPlus.Enabled = false;
            btnMinus.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }

        private void BtnEnabled_true()
        {
            btnPlus.Enabled = true;
            btnMinus.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmNgRecovery_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        /// <summary>
        /// ＮＧ画像リストを表示する
        /// </summary>
        private void GetNgList()
        {
            checkedListBox1.Items.Clear();
            string[] f = System.IO.Directory.GetFiles(_InPath, "*.tif");

            if (f.Length == 0)
            {
                label1.Text = "NG画像はありませんでした";
                return;
            }

            ngf = new clsNG[f.Length];

            int Cnt = 0;

            foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.tif"))
            {
                ngf[Cnt] = new clsNG();
                ngf[Cnt].ngFileName = files;
                string fn = System.IO.Path.GetFileName(files);
                ngf[Cnt].ngRecDate = fn.Substring(1, 4) + "年" + fn.Substring(5, 2) + "月" + fn.Substring(7, 2) + "日" +
                                     fn.Substring(9, 2) + "時" + fn.Substring(11, 2) + "分" + fn.Substring(13, 2) + "秒";

                checkedListBox1.Items.Add(System.IO.Path.GetFileName(ngf[Cnt].ngRecDate));
                Cnt++;
            }

            label1.Text = "NG画像が" + f.Length.ToString() + "件あります";
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkedListBox1.SelectedItem == null) return;
            else BtnEnabled_true();

            //画像イメージ表示
            ShowImage(ngf[checkedListBox1.SelectedIndex].ngFileName);
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
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
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

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;
            }
            else
            {
                //画像ファイルがないとき
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private class clsNG
        {
            public string ngFileName { get; set; }
            public string ngRecDate { get; set; }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // ＮＧファイルリカバリ
            NgRecovery();
        }

        /// <summary>
        /// ＮＧファイルリカバリ
        /// </summary>
        private void NgRecovery()
        {
            if (ngFileCount() == 0)
            {
                MessageBox.Show("画像が選択されていません", "画像未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                if (MessageBox.Show(ngFileCount().ToString() + "件の画像を有効化します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                    return;
            }

            // ＮＧファイルリカバリ処理
            int fCnt = 0;
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    fCnt++;
                    NgToData(fCnt, i);
                }
            }

            // ＮＧ画像リスト再表示
            GetNgList();

            // イメージ表示初期化
            leadImg.Image = null;
            BtnEnabled_false();
        }


        /// <summary>
        /// ＣＳＶデータファイル作成・ＮＧ画像→データ画像へ
        /// </summary>
        /// <param name="fCnt">リカバリファイル番号</param>
        /// <param name="ind">リストボックスインデックス</param>
        private void NgToData(int fCnt, int ind)
        {
            // IDを取得します
            string _ID = string.Format("{0:0000}", DateTime.Today.Year) +
                         string.Format("{0:00}", DateTime.Today.Month) +
                         string.Format("{0:00}", DateTime.Today.Day) +
                         string.Format("{0:00}", DateTime.Now.Hour) +
                         string.Format("{0:00}", DateTime.Now.Minute) +
                         string.Format("{0:00}", DateTime.Now.Second) +
                         fCnt.ToString().PadLeft(3, '0');

            // 出力ファイルインスタンス作成
            StreamWriter outFile = new StreamWriter(_OutPath + _ID + ".csv", false, System.Text.Encoding.GetEncoding(932));

            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();

                // ヘッダ情報
                sb.Append(_ID).Append(",");             // ファイル番号
                sb.Append(_ID + ".tif").Append(",");    // 画像ファイル名
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

                // ＣＳＶファイル作成
                outFile.WriteLine(sb.ToString());

                // 画像ファイル移動
                File.Move(ngf[ind].ngFileName, _OutPath + _ID + ".tif");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ＮＧ画像リカバリ処理", MessageBoxButtons.OK);
            }
            finally
            {
                outFile.Close();
            }
        }

        /// <summary>
        ///チェックボックス選択数取得
        /// </summary>
        /// <returns>選択アイテム数</returns>
        private int ngFileCount()
        {
            return checkedListBox1.CheckedItems.Count;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // ＮＧ画像削除処理
            NgFileDelete();
        }

        /// <summary>
        /// ＮＧ画像削除処理
        /// </summary>
        private void NgFileDelete()
        {
            if (ngFileCount() == 0)
            {
                MessageBox.Show("画像が選択されていません", "画像未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                if (MessageBox.Show(ngFileCount().ToString() + "件の画像を削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                    return;
            }

            // ＮＧファイル削除
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    imgDelete(ngf[i].ngFileName);
                }
            }

            // ＮＧ画像リスト再表示
            GetNgList();

            // イメージ表示初期化
            leadImg.Image = null;
            BtnEnabled_false();
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

        private void button6_Click(object sender, EventArgs e)
        {
            // ＮＧ画像印刷
            NgImagePrint();
        }

        /// <summary>
        /// ＮＧ画像印刷
        /// </summary>
        private void NgImagePrint()
        {
            if (ngFileCount() == 0)
            {
                MessageBox.Show("画像が選択されていません", "画像未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                if (MessageBox.Show(ngFileCount().ToString() + "件の画像を印刷します。よろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                    return;
            }

            // ＮＧ画像印刷
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    cPrint pr = new cPrint();
                    pr.Image(ngf[i].ngFileName);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 右へ90°回転させる
            RotateCommand rc = new RotateCommand();
            rc.Angle = 90 * 100;
            rc.FillColor = new RasterColor(255, 255, 255);
            //rc.Flags = RotateCommandFlags.Bicubic;
            rc.Flags = RotateCommandFlags.Resize;
            rc.Run(leadImg.Image);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 左へ90°回転させる
            RotateCommand rc = new RotateCommand();
            rc.Angle = -90 * 100;
            rc.FillColor = new RasterColor(255, 255, 255);
            //rc.Flags = RotateCommandFlags.Bicubic;
            rc.Flags = RotateCommandFlags.Resize;
            rc.Run(leadImg.Image);
        }
    }
}

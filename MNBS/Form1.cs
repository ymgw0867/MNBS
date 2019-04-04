using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MNBS.Common;
using MNBS.OCR;

namespace MNBS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            //  ローカルフォルダ作成処理
            Utility.dirCreate(Properties.Settings.Default.DATAS1);
            Utility.dirCreate(Properties.Settings.Default.DATAS2);
            Utility.dirCreate(Properties.Settings.Default.DATAJ1);
            Utility.dirCreate(Properties.Settings.Default.DATAJ1);
            Utility.dirCreate(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathMDB);
            Utility.dirCreate(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathXLS);
            Utility.dirCreate(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathNG);
            Utility.dirCreate(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathREAD);
            Utility.dirCreate(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathOK);
            Utility.dirCreate(Properties.Settings.Default.PathInst + Properties.Settings.Default.PathTRAY);
            
            // 設定情報を取得します
            msConfig ms = new msConfig();
            ms.GetCommonYearMonth();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new Config.frmConfig();
            frm.ShowDialog();
            this.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DoOCR();
        }

        /// <summary>
        /// ＯＣＲ処理を実行します
        /// </summary>
        private void DoOCR()
        {
            int _pcSel = global.END_SELECT;
            int _OcrSel = global.END_SELECT;
            this.Hide();

            // PCを選択する
            frmPcSelect frmSel = new frmPcSelect();
            frmSel.ShowDialog();
            _pcSel = frmSel._pcSel;

            frmSel.Dispose();

            if (_pcSel == global.END_SELECT)
            {
                this.Show();
                return;
            }

            // ＯＣＲ処理方法を選択する
            frmOcrSelect frmOcrSel = new frmOcrSelect();
            frmOcrSel.ShowDialog();
            _OcrSel = frmOcrSel._OcrSel;

            frmOcrSel.Dispose();

            if (_OcrSel == global.END_SELECT)
            {
                this.Show();
                return;
            }

            // 確認メッセージ
            StringBuilder sb = new StringBuilder();

            if (_OcrSel == global.SCAN_SELECT)
                sb.Append("スキャナで画像を読み取りOCR認識処理を行います。").Append(Environment.NewLine).Append(Environment.NewLine);
            else sb.Append("受信済みのFAX画像を読み取りOCR認識処理を行います。").Append(Environment.NewLine).Append(Environment.NewLine);
            sb.Append("よろしいですか？中止する場合は「いいえ」をクリックしてください。");

            if ((MessageBox.Show(sb.ToString(), this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No))
            {
                this.Show();
                return;
            }
            else
            {
                // 勤務票スキャン、ＯＣＲ処理
                frmOCR frm = new frmOCR(_pcSel, _OcrSel);
                frm.ShowDialog();
                this.Show();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataEntry();
        }

        /// <summary>
        /// データ登録
        /// </summary>
        private void DataEntry()
        {
            // 2019/04/04 コメント化
            //string msg = "設定年月は " + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月です。よろしいですか？";

            // 環境設定年月の確認：西暦表示 2019/04/04
            string msg = "設定年月は " + (global.sYear + Properties.Settings.Default.RekiHosei) + "年 " + global.sMonth.ToString() + "月です。よろしいですか？";
            if (MessageBox.Show(msg, "勤務データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            this.Hide();

            // 勤怠データ登録画面表示
            frmCorrect frmData = new frmCorrect(global.sEDITMODE);
            frmData.ShowDialog();
            this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string outPath = string.Empty;

            // ＯＣＲ入力パス
            if (Properties.Settings.Default.PC == global.PC1SELECT.ToString())
                outPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.DATAS1;
            else outPath = Properties.Settings.Default.PathInst + Properties.Settings.Default.DATAS2;

            var s = System.IO.Directory.GetFiles(outPath, "*.CSV");
            if (s.Count() != 0)
            {
                string msg = string.Empty;
                msg += "受渡し勤怠データ作成途中のOCR認識データが残っています。" + Environment.NewLine + Environment.NewLine;
                msg += "先に受渡し勤怠データ作成を行ってください。";
                MessageBox.Show(msg, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 勤怠データ新規登録画面表示
            frmCorrect frmData = new frmCorrect(global.sADDMODE);
            frmData.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmNgRecovery frm = new frmNgRecovery();
            frm.ShowDialog();
            this.Show();
        }
    }
}

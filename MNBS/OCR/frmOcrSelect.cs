using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MNBS.OCR;
using MNBS.Common;

namespace MNBS.OCR
{
    public partial class frmOcrSelect : Form
    {
        public int _OcrSel { get; set; }

        public frmOcrSelect()
        {
            InitializeComponent();
        }

        private void frmPcStaffSelect_Load(object sender, EventArgs e)
        {
            // フォームの最大サイズ、最小サイズの設定
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // スキャナを既定値とする
            this.rPcBtn1.Checked = true;

            // 値選択既定値
            _OcrSel = global.END_SELECT;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            // OCR選択
            if (rPcBtn1.Checked) _OcrSel = global.SCAN_SELECT;
            else if (rPcBtn2.Checked) _OcrSel = global.FAX_SELECT;

            this.Close();
        }

        private void btnNo_Click(object sender, EventArgs e)
        {
            // 終了
            this.Close();
        }
    }
}

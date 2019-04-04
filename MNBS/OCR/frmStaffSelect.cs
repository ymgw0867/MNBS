using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MNBS.Common;

namespace MNBS.OCR
{
    public partial class frmStaffSelect : Form
    {
        public frmStaffSelect(msShain [] m)
        {
            InitializeComponent();
            _msCode = string.Empty;
            ms = m;
        }

        msShain[] ms;

        private void frmStaffSelect_Load(object sender, EventArgs e)
        {
            Common.Utility.WindowsMaxSize(this, this.Width, this.Height);
            Common.Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビュー定義
            GridViewStaffSetting(dataGridView1);

            // スタッフマスター表示
            GridViewStaffShow(dataGridView1, string.Empty, string.Empty);       
        }

        // 指定スタッフ
        private int _usrSel;

        // データグリッドビューカラム名
        private string cCode = "c1";
        private string cName = "c2";
        private string cKinmuCode = "c3";
        private string cKinmuName = "c4";
        private string cKinmuBusho = "c5";
        private string cKKikan = "c6";
        private string cKTime = "c7";
        private string cKOrderCode = "c8";
        private string cTcode = "c9";
        private string cStime = "c10";
        private string cEtime = "c11";
        private string cRest = "c12";

        /// <summary>
        /// スタッフマスターグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewStaffSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 380;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cKOrderCode, "オーダーコード");
                tempDGV.Columns.Add(cCode, "スタッフコード");
                tempDGV.Columns.Add(cName, "氏名");
                tempDGV.Columns.Add(cKinmuCode, "派遣先コード");
                tempDGV.Columns.Add(cKinmuName, "派遣先名");
                tempDGV.Columns.Add(cKinmuBusho, "部署");
                tempDGV.Columns.Add(cStime, "開始時刻");
                tempDGV.Columns.Add(cEtime, "終了時刻");
                tempDGV.Columns.Add(cRest, "休憩（分）");
                tempDGV.Columns.Add(cTcode, "コード");
                tempDGV.Columns[cTcode].Visible = false;

                tempDGV.Columns[cKOrderCode].Width = 100;
                tempDGV.Columns[cCode].Width = 80;
                tempDGV.Columns[cName].Width = 120;
                tempDGV.Columns[cKinmuCode].Width = 120;
                tempDGV.Columns[cKinmuName].Width = 200;
                tempDGV.Columns[cKinmuBusho].Width = 160;
                tempDGV.Columns[cStime].Width = 100;
                tempDGV.Columns[cEtime].Width = 100;
                tempDGV.Columns[cRest].Width = 110;

                tempDGV.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKinmuCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cStime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cEtime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKOrderCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cRest].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                
                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

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

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// スタッフマスタをグリッドビューへ表示します
        /// </summary>
        /// <param name="tempGrid">データグリッドビューオブジェクト</param>
        private void GridViewStaffShow(DataGridView tempGrid, string sCode, string sName)
        {
            tempGrid.Rows.Clear();
            tempGrid.RowCount = 0;

            // カーソルを待機状態にする
            this.Cursor = Cursors.WaitCursor;

            for (int i = 0; i < ms.Length; i++)
            {
                if (sCode == string.Empty || ms[i]._StaffCode.Contains(sCode))
                {
                    if (sName == string.Empty || ms[i]._StaffName.Contains(sName))
                    {
                        tempGrid.Rows.Add();

                        tempGrid[cKOrderCode, tempGrid.RowCount - 1].Value = ms[i]._OrderCode;
                        tempGrid[cCode, tempGrid.RowCount - 1].Value = ms[i]._StaffCode;
                        tempGrid[cName, tempGrid.RowCount - 1].Value = ms[i]._StaffName;
                        tempGrid[cKinmuCode, tempGrid.RowCount - 1].Value = ms[i]._HaCode;
                        tempGrid[cKinmuName, tempGrid.RowCount - 1].Value = ms[i]._HaName;
                        tempGrid[cKinmuBusho, tempGrid.RowCount - 1].Value = ms[i]._BuName;
                        tempGrid[cStime, tempGrid.RowCount - 1].Value = ms[i]._STime;
                        tempGrid[cEtime, tempGrid.RowCount - 1].Value = ms[i]._ETime;
                        tempGrid[cRest, tempGrid.RowCount - 1].Value = ms[i]._Kyukei;
                        tempGrid[cTcode, tempGrid.RowCount - 1].Value = ms[i]._HaCode;
                    }
                }
            }
            tempGrid.CurrentCell = null;

            // カーソルを戻す
            this.Cursor = Cursors.Default;
        }

        /// <summary>
        /// スタッフ情報を取得
        /// </summary>
        /// <param name="g">データグリッドビューオブジェクト</param>
        private void GetGridViewData(DataGridView g)
        {
            if (g.SelectedRows.Count == 0) return;

            int r = g.SelectedRows[0].Index;

            _msCode = g[cCode, r].Value.ToString();

            _msName = g[cName, r].Value.ToString();
            _msTenpoCode = g[cKinmuCode, r].Value.ToString();
            _msTenpoName = g[cKinmuName, r].Value.ToString();
            _msOrderCode = g[cKOrderCode, r].Value.ToString();
        }

        public string _msCode { get; set; }
        public string _msName { get; set; }
        public string _msTenpoCode { get; set; }
        public string _msTenpoName { get; set; }
        public string _msBusho { get; set; }
        public string _msTcode { get; set; }
        public string _msOrderCode { get; set; }
        public string _msStime { get; set; }
        public string _msEtime { get; set; }
        public string _msHStime1 { get; set; }
        public string _msHEtime1 { get; set; }
        public string _msHStime2 { get; set; }
        public string _msHEtime2 { get; set; }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GetGridViewData(dataGridView1);
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            GetGridViewData(dataGridView1);
            this.Close();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmStaffSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // スタッフマスター表示
            GridViewStaffShow(dataGridView1, txtsCode.Text, txtsName.Text);


            ////int r = -1;
            ////if (txtsCode.Text != string.Empty) 
            ////    r = GridViewFind(dataGridView1, cCode, txtsCode.Text);
            ////else if (txtsName.Text != string.Empty) 
            ////    r = GridViewFind(dataGridView1, cName, txtsName.Text);

            ////if (r != -1)
            ////{
            ////    dataGridView1.FirstDisplayedScrollingRowIndex = r;
            ////    dataGridView1.CurrentCell = dataGridView1[cCode, r];
            ////}
        }

        /// <summary>
        /// コードまたは氏名で該当するデータの行を返します
        /// </summary>
        /// <param name="d">データグリッドビューオブジェクト</param>
        /// <param name="ColName">対象のカラム名</param>
        /// <param name="Val">検索する値</param>
        /// <returns>行のIndex</returns>
        private int GridViewFind(DataGridView d, string ColName, string Val)
        {
            int rtn = -1;

            for (int i = 0; i < d.Rows.Count; i++)
            {
                if (d[ColName, i].Value.ToString().Contains(Val))
                {
                    rtn = i;
                    break;
                }
            }

            return rtn;
        }

        private void txtsCode_TextChanged(object sender, EventArgs e)
        {
            if (txtsCode.Text.Length > 0) txtsName.Text = string.Empty;
        }

        private void txtsName_TextChanged(object sender, EventArgs e)
        {
            if (txtsName.Text.Length > 0) txtsCode.Text = string.Empty;
        }
    }
}

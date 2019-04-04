namespace MNBS.OCR
{
    partial class frmCorrect
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCorrect));
            this.label5 = new System.Windows.Forms.Label();
            this.txtShozokuCode = new System.Windows.Forms.TextBox();
            this.lblShozoku = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.txtNo = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.leadImg = new Leadtools.WinForms.RasterImageViewer();
            this.lblErrMsg = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtOrderCode = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtShuTl = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.btnMinus = new System.Windows.Forms.Button();
            this.btnPlus = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnDataMake = new System.Windows.Forms.Button();
            this.btnErrCheck = new System.Windows.Forms.Button();
            this.btnEnd = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnBefore = new System.Windows.Forms.Button();
            this.btnFirst = new System.Windows.Forms.Button();
            this.lblNoImage = new System.Windows.Forms.Label();
            this.lblPage = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label19 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.dg1 = new MNBS.DataGridViewEx();
            this.panel2.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).BeginInit();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(12, 31);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 25);
            this.label5.TabIndex = 74;
            this.label5.Text = "派遣先：";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtShozokuCode
            // 
            this.txtShozokuCode.BackColor = System.Drawing.SystemColors.Window;
            this.txtShozokuCode.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtShozokuCode.Location = new System.Drawing.Point(91, 33);
            this.txtShozokuCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtShozokuCode.MaxLength = 8;
            this.txtShozokuCode.Name = "txtShozokuCode";
            this.txtShozokuCode.ReadOnly = true;
            this.txtShozokuCode.Size = new System.Drawing.Size(86, 24);
            this.txtShozokuCode.TabIndex = 3;
            this.txtShozokuCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtShozokuCode.WordWrap = false;
            this.txtShozokuCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // lblShozoku
            // 
            this.lblShozoku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblShozoku.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblShozoku.Location = new System.Drawing.Point(180, 33);
            this.lblShozoku.Name = "lblShozoku";
            this.lblShozoku.Size = new System.Drawing.Size(319, 23);
            this.lblShozoku.TabIndex = 73;
            this.lblShozoku.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblName
            // 
            this.lblName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblName.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblName.Location = new System.Drawing.Point(354, 5);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(145, 23);
            this.lblName.TabIndex = 72;
            this.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtNo
            // 
            this.txtNo.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtNo.Location = new System.Drawing.Point(280, 4);
            this.txtNo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtNo.MaxLength = 6;
            this.txtNo.Name = "txtNo";
            this.txtNo.Size = new System.Drawing.Size(68, 24);
            this.txtNo.TabIndex = 2;
            this.txtNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtNo.WordWrap = false;
            this.txtNo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNo_KeyDown);
            this.txtNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            this.txtNo.Leave += new System.EventHandler(this.txtNo_Leave);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(206, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 15);
            this.label4.TabIndex = 71;
            this.label4.Text = "スタッフコード：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(177, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(19, 15);
            this.label3.TabIndex = 70;
            this.label3.Text = "月";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(123, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(19, 15);
            this.label2.TabIndex = 69;
            this.label2.Text = "年";
            // 
            // txtMonth
            // 
            this.txtMonth.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMonth.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtMonth.Location = new System.Drawing.Point(145, 6);
            this.txtMonth.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtMonth.MaxLength = 2;
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(32, 24);
            this.txtMonth.TabIndex = 1;
            this.txtMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtMonth.WordWrap = false;
            this.txtMonth.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // txtYear
            // 
            this.txtYear.BackColor = System.Drawing.SystemColors.Window;
            this.txtYear.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtYear.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtYear.Location = new System.Drawing.Point(91, 6);
            this.txtYear.Margin = new System.Windows.Forms.Padding(0);
            this.txtYear.MaxLength = 2;
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(32, 24);
            this.txtYear.TabIndex = 0;
            this.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtYear.WordWrap = false;
            this.txtYear.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(29, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 25);
            this.label1.TabIndex = 68;
            this.label1.Text = "年月：20";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // leadImg
            // 
            this.leadImg.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.leadImg.Location = new System.Drawing.Point(-2, -2);
            this.leadImg.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.leadImg.Name = "leadImg";
            this.leadImg.Size = new System.Drawing.Size(484, 573);
            this.leadImg.TabIndex = 80;
            // 
            // lblErrMsg
            // 
            this.lblErrMsg.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblErrMsg.Font = new System.Drawing.Font("Meiryo UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblErrMsg.ForeColor = System.Drawing.Color.Red;
            this.lblErrMsg.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblErrMsg.Location = new System.Drawing.Point(0, 0);
            this.lblErrMsg.Name = "lblErrMsg";
            this.lblErrMsg.Size = new System.Drawing.Size(410, 22);
            this.lblErrMsg.TabIndex = 81;
            this.lblErrMsg.Text = "エラー内容を表示します";
            this.lblErrMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(5, 59);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(90, 25);
            this.label7.TabIndex = 83;
            this.label7.Text = "オーダーコード：";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtOrderCode
            // 
            this.txtOrderCode.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtOrderCode.Location = new System.Drawing.Point(91, 60);
            this.txtOrderCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtOrderCode.MaxLength = 8;
            this.txtOrderCode.Name = "txtOrderCode";
            this.txtOrderCode.Size = new System.Drawing.Size(86, 24);
            this.txtOrderCode.TabIndex = 5;
            this.txtOrderCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOrderCode.WordWrap = false;
            this.txtOrderCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label14
            // 
            this.label14.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label14.Location = new System.Drawing.Point(945, 193);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(23, 25);
            this.label14.TabIndex = 103;
            this.label14.Text = "日";
            this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtShuTl
            // 
            this.txtShuTl.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtShuTl.Location = new System.Drawing.Point(913, 193);
            this.txtShuTl.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtShuTl.MaxLength = 7;
            this.txtShuTl.Name = "txtShuTl";
            this.txtShuTl.Size = new System.Drawing.Size(34, 24);
            this.txtShuTl.TabIndex = 2;
            this.txtShuTl.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtShuTl.WordWrap = false;
            this.txtShuTl.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label15
            // 
            this.label15.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label15.Location = new System.Drawing.Point(809, 192);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(108, 25);
            this.label15.TabIndex = 101;
            this.label15.Text = "出勤日数合計：";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnMinus
            // 
            this.btnMinus.Image = global::MNBS.Properties.Resources.tvuZoomOut;
            this.btnMinus.Location = new System.Drawing.Point(471, 59);
            this.btnMinus.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(28, 28);
            this.btnMinus.TabIndex = 8;
            this.btnMinus.TabStop = false;
            this.btnMinus.UseVisualStyleBackColor = true;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // btnPlus
            // 
            this.btnPlus.Image = global::MNBS.Properties.Resources.tvuZoomIn;
            this.btnPlus.Location = new System.Drawing.Point(444, 59);
            this.btnPlus.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(28, 28);
            this.btnPlus.TabIndex = 7;
            this.btnPlus.TabStop = false;
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRtn.Location = new System.Drawing.Point(816, 631);
            this.btnRtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(152, 36);
            this.btnRtn.TabIndex = 19;
            this.btnRtn.Text = "　　終　了(&Q)";
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // btnDel
            // 
            this.btnDel.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDel.Location = new System.Drawing.Point(815, 96);
            this.btnDel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(152, 40);
            this.btnDel.TabIndex = 18;
            this.btnDel.Text = "表示中データ削除(&D)";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnDataMake
            // 
            this.btnDataMake.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDataMake.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDataMake.Location = new System.Drawing.Point(815, 141);
            this.btnDataMake.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnDataMake.Name = "btnDataMake";
            this.btnDataMake.Size = new System.Drawing.Size(152, 36);
            this.btnDataMake.TabIndex = 17;
            this.btnDataMake.Text = "受け渡しデータ作成(&U)";
            this.btnDataMake.UseVisualStyleBackColor = true;
            this.btnDataMake.Click += new System.EventHandler(this.btnDataMake_Click);
            // 
            // btnErrCheck
            // 
            this.btnErrCheck.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnErrCheck.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnErrCheck.Location = new System.Drawing.Point(815, 56);
            this.btnErrCheck.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnErrCheck.Name = "btnErrCheck";
            this.btnErrCheck.Size = new System.Drawing.Size(152, 35);
            this.btnErrCheck.TabIndex = 16;
            this.btnErrCheck.Text = "エラーチェック(&E)";
            this.btnErrCheck.UseVisualStyleBackColor = true;
            this.btnErrCheck.Click += new System.EventHandler(this.btnErrCheck_Click);
            // 
            // btnEnd
            // 
            this.btnEnd.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnEnd.Image = ((System.Drawing.Image)(resources.GetObject("btnEnd.Image")));
            this.btnEnd.Location = new System.Drawing.Point(736, 32);
            this.btnEnd.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(72, 24);
            this.btnEnd.TabIndex = 13;
            this.btnEnd.TabStop = false;
            this.btnEnd.UseVisualStyleBackColor = true;
            this.btnEnd.Click += new System.EventHandler(this.btnEnd_Click);
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.Location = new System.Drawing.Point(659, 32);
            this.btnNext.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(72, 24);
            this.btnNext.TabIndex = 12;
            this.btnNext.TabStop = false;
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnBefore
            // 
            this.btnBefore.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnBefore.Image = ((System.Drawing.Image)(resources.GetObject("btnBefore.Image")));
            this.btnBefore.Location = new System.Drawing.Point(582, 32);
            this.btnBefore.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnBefore.Name = "btnBefore";
            this.btnBefore.Size = new System.Drawing.Size(72, 24);
            this.btnBefore.TabIndex = 11;
            this.btnBefore.TabStop = false;
            this.btnBefore.UseVisualStyleBackColor = true;
            this.btnBefore.Click += new System.EventHandler(this.btnBefore_Click);
            // 
            // btnFirst
            // 
            this.btnFirst.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnFirst.Image = ((System.Drawing.Image)(resources.GetObject("btnFirst.Image")));
            this.btnFirst.Location = new System.Drawing.Point(505, 32);
            this.btnFirst.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(72, 24);
            this.btnFirst.TabIndex = 10;
            this.btnFirst.TabStop = false;
            this.btnFirst.UseVisualStyleBackColor = true;
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
            // 
            // lblNoImage
            // 
            this.lblNoImage.Font = new System.Drawing.Font("Meiryo UI", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNoImage.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.lblNoImage.Location = new System.Drawing.Point(26, 210);
            this.lblNoImage.Name = "lblNoImage";
            this.lblNoImage.Size = new System.Drawing.Size(425, 81);
            this.lblNoImage.TabIndex = 123;
            this.lblNoImage.Text = "勤務票イメージがありません";
            this.lblNoImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblPage
            // 
            this.lblPage.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblPage.Location = new System.Drawing.Point(814, 32);
            this.lblPage.Name = "lblPage";
            this.lblPage.Size = new System.Drawing.Size(73, 22);
            this.lblPage.TabIndex = 124;
            this.lblPage.Text = "1/1 件目";
            this.lblPage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.lblErrMsg);
            this.panel2.Location = new System.Drawing.Point(554, 2);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(414, 26);
            this.panel2.TabIndex = 125;
            // 
            // label19
            // 
            this.label19.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label19.Location = new System.Drawing.Point(502, 3);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(56, 25);
            this.label19.TabIndex = 126;
            this.label19.Text = "エラー：";
            this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel4.Controls.Add(this.lblNoImage);
            this.panel4.Controls.Add(this.leadImg);
            this.panel4.Location = new System.Drawing.Point(15, 88);
            this.panel4.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(484, 579);
            this.panel4.TabIndex = 130;
            // 
            // dg1
            // 
            this.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg1.Location = new System.Drawing.Point(505, 57);
            this.dg1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dg1.Name = "dg1";
            this.dg1.RowTemplate.Height = 21;
            this.dg1.Size = new System.Drawing.Size(304, 610);
            this.dg1.TabIndex = 0;
            this.dg1.TabStop = false;
            this.dg1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg1_CellClick);
            this.dg1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg1_CellContentClick);
            this.dg1.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg1_CellEnter);
            this.dg1.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dg1_CellPainting);
            this.dg1.CellParsing += new System.Windows.Forms.DataGridViewCellParsingEventHandler(this.dg1_CellParsing);
            this.dg1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg1_CellValueChanged);
            this.dg1.CurrentCellDirtyStateChanged += new System.EventHandler(this.dg1_CurrentCellDirtyStateChanged);
            this.dg1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dg1_EditingControlShowing);
            // 
            // frmCorrect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(976, 677);
            this.Controls.Add(this.txtShuTl);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.lblPage);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnDataMake);
            this.Controls.Add(this.btnErrCheck);
            this.Controls.Add(this.txtOrderCode);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtShozokuCode);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnBefore);
            this.Controls.Add(this.btnFirst);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblShozoku);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.txtNo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtMonth);
            this.Controls.Add(this.txtYear);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dg1);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.panel4);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "frmCorrect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "勤怠データ登録";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmCorrect_FormClosing);
            this.Load += new System.EventHandler(this.frmCorrect_Load);
            this.panel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DataGridViewEx dg1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtShozokuCode;
        private System.Windows.Forms.Label lblShozoku;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.TextBox txtNo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMonth;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnBefore;
        private System.Windows.Forms.Button btnFirst;
        private Leadtools.WinForms.RasterImageViewer leadImg;
        private System.Windows.Forms.Label lblErrMsg;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtOrderCode;
        private System.Windows.Forms.Button btnDataMake;
        private System.Windows.Forms.Button btnErrCheck;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txtShuTl;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.Label lblNoImage;
        private System.Windows.Forms.Label lblPage;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Panel panel4;
    }
}
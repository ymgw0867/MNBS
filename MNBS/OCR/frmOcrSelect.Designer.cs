namespace MNBS.OCR
{
    partial class frmOcrSelect
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmOcrSelect));
            this.panel1 = new System.Windows.Forms.Panel();
            this.rPcBtn2 = new System.Windows.Forms.RadioButton();
            this.rPcBtn1 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnNo = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.rPcBtn2);
            this.panel1.Controls.Add(this.rPcBtn1);
            this.panel1.Location = new System.Drawing.Point(26, 39);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(442, 78);
            this.panel1.TabIndex = 0;
            // 
            // rPcBtn2
            // 
            this.rPcBtn2.AutoSize = true;
            this.rPcBtn2.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.rPcBtn2.Location = new System.Drawing.Point(218, 24);
            this.rPcBtn2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rPcBtn2.Name = "rPcBtn2";
            this.rPcBtn2.Size = new System.Drawing.Size(206, 27);
            this.rPcBtn2.TabIndex = 1;
            this.rPcBtn2.TabStop = true;
            this.rPcBtn2.Text = "FAX受信データを読み込み";
            this.rPcBtn2.UseVisualStyleBackColor = true;
            // 
            // rPcBtn1
            // 
            this.rPcBtn1.AutoSize = true;
            this.rPcBtn1.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.rPcBtn1.Location = new System.Drawing.Point(15, 24);
            this.rPcBtn1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rPcBtn1.Name = "rPcBtn1";
            this.rPcBtn1.Size = new System.Drawing.Size(178, 27);
            this.rPcBtn1.TabIndex = 0;
            this.rPcBtn1.TabStop = true;
            this.rPcBtn1.Text = "スキャナから読み込み";
            this.rPcBtn1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(22, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(217, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "読み込みの方法を選択してください";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(208, 151);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(127, 33);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "ＯＫ(&O)";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnNo
            // 
            this.btnNo.Location = new System.Drawing.Point(341, 151);
            this.btnNo.Name = "btnNo";
            this.btnNo.Size = new System.Drawing.Size(127, 33);
            this.btnNo.TabIndex = 5;
            this.btnNo.Text = "キャンセル(&Q)";
            this.btnNo.UseVisualStyleBackColor = true;
            this.btnNo.Click += new System.EventHandler(this.btnNo_Click);
            // 
            // frmOcrSelect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(494, 196);
            this.Controls.Add(this.btnNo);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmOcrSelect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "OCR認識処理";
            this.Load += new System.EventHandler(this.frmPcStaffSelect_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton rPcBtn2;
        private System.Windows.Forms.RadioButton rPcBtn1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnNo;
    }
}
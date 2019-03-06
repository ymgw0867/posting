namespace posting
{
    partial class frmEiUriageRep
    {
        /// <summary>
        /// 必要なデザイナ変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナで生成されたコード

        /// <summary>
        /// デザイナ サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディタで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmEiUriageRep));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.tDate = new System.Windows.Forms.DateTimePicker();
            this.tDate2 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.btnPrn = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblGaichuhi = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.lblArari2 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblEgenka = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.lblArari1 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.lblArarisai = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.lblGaichuhi2 = new System.Windows.Forms.Label();
            this.lblGaichuhi3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(14, 42);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 18;
            this.dataGridView1.Size = new System.Drawing.Size(1142, 540);
            this.dataGridView1.TabIndex = 8;
            // 
            // btnRtn
            // 
            this.btnRtn.BackColor = System.Drawing.SystemColors.Control;
            this.btnRtn.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(1066, 633);
            this.btnRtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(90, 33);
            this.btnRtn.TabIndex = 5;
            this.btnRtn.Text = "戻る(&R)";
            this.btnRtn.UseCompatibleTextRendering = true;
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.SystemColors.Control;
            this.btnUpdate.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnUpdate.Location = new System.Drawing.Point(1066, 9);
            this.btnUpdate.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(90, 28);
            this.btnUpdate.TabIndex = 3;
            this.btnUpdate.Text = "選択(&S)";
            this.btnUpdate.UseCompatibleTextRendering = true;
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.SteelBlue;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(15, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 23);
            this.label3.TabIndex = 20;
            this.label3.Text = "入金日";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tDate
            // 
            this.tDate.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.tDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.tDate.Location = new System.Drawing.Point(76, 12);
            this.tDate.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tDate.Name = "tDate";
            this.tDate.ShowCheckBox = true;
            this.tDate.Size = new System.Drawing.Size(132, 23);
            this.tDate.TabIndex = 0;
            // 
            // tDate2
            // 
            this.tDate2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.tDate2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.tDate2.Location = new System.Drawing.Point(231, 12);
            this.tDate2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tDate2.Name = "tDate2";
            this.tDate2.ShowCheckBox = true;
            this.tDate2.Size = new System.Drawing.Size(132, 23);
            this.tDate2.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.SteelBlue;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(386, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 23);
            this.label1.TabIndex = 22;
            this.label1.Text = "担当者";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(448, 12);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(282, 23);
            this.comboBox1.TabIndex = 2;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // btnPrn
            // 
            this.btnPrn.BackColor = System.Drawing.SystemColors.Control;
            this.btnPrn.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPrn.Location = new System.Drawing.Point(970, 633);
            this.btnPrn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnPrn.Name = "btnPrn";
            this.btnPrn.Size = new System.Drawing.Size(90, 33);
            this.btnPrn.TabIndex = 4;
            this.btnPrn.Text = "印刷(&P)";
            this.btnPrn.UseVisualStyleBackColor = true;
            this.btnPrn.Click += new System.EventHandler(this.btnPrn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(210, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(19, 15);
            this.label2.TabIndex = 25;
            this.label2.Text = "～";
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.SteelBlue;
            this.label4.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(16, 592);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 25);
            this.label4.TabIndex = 26;
            this.label4.Text = "入金計";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Window;
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label5.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(74, 592);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 25);
            this.label5.TabIndex = 27;
            this.label5.Text = "12,000,000";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblGaichuhi
            // 
            this.lblGaichuhi.BackColor = System.Drawing.SystemColors.Window;
            this.lblGaichuhi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblGaichuhi.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblGaichuhi.Location = new System.Drawing.Point(541, 592);
            this.lblGaichuhi.Name = "lblGaichuhi";
            this.lblGaichuhi.Size = new System.Drawing.Size(83, 25);
            this.lblGaichuhi.TabIndex = 29;
            this.lblGaichuhi.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.SteelBlue;
            this.label7.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(477, 592);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(63, 25);
            this.label7.TabIndex = 28;
            this.label7.Text = "外注費計";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblArari2
            // 
            this.lblArari2.BackColor = System.Drawing.SystemColors.Window;
            this.lblArari2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblArari2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblArari2.Location = new System.Drawing.Point(855, 592);
            this.lblArari2.Name = "lblArari2";
            this.lblArari2.Size = new System.Drawing.Size(119, 25);
            this.lblArari2.TabIndex = 31;
            this.lblArari2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.Color.SteelBlue;
            this.label9.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(798, 592);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(55, 25);
            this.label9.TabIndex = 30;
            this.label9.Text = "粗利２";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblEgenka
            // 
            this.lblEgenka.BackColor = System.Drawing.SystemColors.Window;
            this.lblEgenka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblEgenka.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblEgenka.Location = new System.Drawing.Point(243, 592);
            this.lblEgenka.Name = "lblEgenka";
            this.lblEgenka.Size = new System.Drawing.Size(83, 25);
            this.lblEgenka.TabIndex = 33;
            this.lblEgenka.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label11
            // 
            this.label11.BackColor = System.Drawing.Color.SteelBlue;
            this.label11.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label11.ForeColor = System.Drawing.Color.White;
            this.label11.Location = new System.Drawing.Point(164, 592);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(78, 25);
            this.label11.TabIndex = 32;
            this.label11.Text = "営業原価計";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblArari1
            // 
            this.lblArari1.BackColor = System.Drawing.SystemColors.Window;
            this.lblArari1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblArari1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblArari1.Location = new System.Drawing.Point(388, 592);
            this.lblArari1.Name = "lblArari1";
            this.lblArari1.Size = new System.Drawing.Size(83, 25);
            this.lblArari1.TabIndex = 35;
            this.lblArari1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label13
            // 
            this.label13.BackColor = System.Drawing.Color.SteelBlue;
            this.label13.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label13.ForeColor = System.Drawing.Color.White;
            this.label13.Location = new System.Drawing.Point(332, 592);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(55, 25);
            this.label13.TabIndex = 34;
            this.label13.Text = "粗利１";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblArarisai
            // 
            this.lblArarisai.BackColor = System.Drawing.SystemColors.Window;
            this.lblArarisai.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblArarisai.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblArarisai.Location = new System.Drawing.Point(1042, 592);
            this.lblArarisai.Name = "lblArarisai";
            this.lblArarisai.Size = new System.Drawing.Size(113, 25);
            this.lblArarisai.TabIndex = 37;
            this.lblArarisai.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label15
            // 
            this.label15.BackColor = System.Drawing.Color.SteelBlue;
            this.label15.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label15.ForeColor = System.Drawing.Color.White;
            this.label15.Location = new System.Drawing.Point(979, 592);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(61, 25);
            this.label15.TabIndex = 36;
            this.label15.Text = "粗利差異";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblGaichuhi2
            // 
            this.lblGaichuhi2.BackColor = System.Drawing.SystemColors.Window;
            this.lblGaichuhi2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblGaichuhi2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblGaichuhi2.Location = new System.Drawing.Point(626, 592);
            this.lblGaichuhi2.Name = "lblGaichuhi2";
            this.lblGaichuhi2.Size = new System.Drawing.Size(83, 25);
            this.lblGaichuhi2.TabIndex = 39;
            this.lblGaichuhi2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblGaichuhi3
            // 
            this.lblGaichuhi3.BackColor = System.Drawing.SystemColors.Window;
            this.lblGaichuhi3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblGaichuhi3.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblGaichuhi3.Location = new System.Drawing.Point(711, 592);
            this.lblGaichuhi3.Name = "lblGaichuhi3";
            this.lblGaichuhi3.Size = new System.Drawing.Size(83, 25);
            this.lblGaichuhi3.TabIndex = 41;
            this.lblGaichuhi3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // frmEiUriageRep
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1163, 678);
            this.Controls.Add(this.lblGaichuhi3);
            this.Controls.Add(this.lblGaichuhi2);
            this.Controls.Add(this.lblArarisai);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.lblArari1);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.lblEgenka);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.lblArari2);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.lblGaichuhi);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnPrn);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tDate2);
            this.Controls.Add(this.tDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "frmEiUriageRep";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "営業別売上表";
            this.Load += new System.EventHandler(this.form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        internal System.Windows.Forms.Button btnRtn;
        internal System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker tDate;
        private System.Windows.Forms.DateTimePicker tDate2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBox1;
        internal System.Windows.Forms.Button btnPrn;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblGaichuhi;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblArari2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblEgenka;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label lblArari1;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label lblArarisai;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label lblGaichuhi2;
        private System.Windows.Forms.Label lblGaichuhi3;
    }
}
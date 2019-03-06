namespace posting
{
    partial class frmKouza
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmKouza));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.金融機関名DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.支店名DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.口座種別DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.口座番号DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.口座名義DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.登録年月日DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.変更年月日DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.振込口座BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.darwinDataSet = new posting.darwinDataSet();
            this.btnCsv = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClr = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.txtMeigi = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtName1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtName2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.振込口座TableAdapter = new posting.darwinDataSetTableAdapters.振込口座TableAdapter();
            this.label2 = new System.Windows.Forms.Label();
            this.txtNumber = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.振込口座BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.darwinDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.金融機関名DataGridViewTextBoxColumn,
            this.支店名DataGridViewTextBoxColumn,
            this.口座種別DataGridViewTextBoxColumn,
            this.口座番号DataGridViewTextBoxColumn,
            this.口座名義DataGridViewTextBoxColumn,
            this.登録年月日DataGridViewTextBoxColumn,
            this.変更年月日DataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.振込口座BindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 18;
            this.dataGridView1.Size = new System.Drawing.Size(643, 180);
            this.dataGridView1.TabIndex = 10;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
            // 
            // iDDataGridViewTextBoxColumn
            // 
            this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
            this.iDDataGridViewTextBoxColumn.HeaderText = "コード";
            this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
            this.iDDataGridViewTextBoxColumn.ReadOnly = true;
            this.iDDataGridViewTextBoxColumn.Visible = false;
            // 
            // 金融機関名DataGridViewTextBoxColumn
            // 
            this.金融機関名DataGridViewTextBoxColumn.DataPropertyName = "金融機関名";
            this.金融機関名DataGridViewTextBoxColumn.HeaderText = "金融機関名";
            this.金融機関名DataGridViewTextBoxColumn.Name = "金融機関名DataGridViewTextBoxColumn";
            this.金融機関名DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 支店名DataGridViewTextBoxColumn
            // 
            this.支店名DataGridViewTextBoxColumn.DataPropertyName = "支店名";
            this.支店名DataGridViewTextBoxColumn.HeaderText = "支店名";
            this.支店名DataGridViewTextBoxColumn.Name = "支店名DataGridViewTextBoxColumn";
            this.支店名DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 口座種別DataGridViewTextBoxColumn
            // 
            this.口座種別DataGridViewTextBoxColumn.DataPropertyName = "口座種別";
            this.口座種別DataGridViewTextBoxColumn.HeaderText = "口座種別";
            this.口座種別DataGridViewTextBoxColumn.Name = "口座種別DataGridViewTextBoxColumn";
            this.口座種別DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 口座番号DataGridViewTextBoxColumn
            // 
            this.口座番号DataGridViewTextBoxColumn.DataPropertyName = "口座番号";
            this.口座番号DataGridViewTextBoxColumn.HeaderText = "口座番号";
            this.口座番号DataGridViewTextBoxColumn.Name = "口座番号DataGridViewTextBoxColumn";
            this.口座番号DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 口座名義DataGridViewTextBoxColumn
            // 
            this.口座名義DataGridViewTextBoxColumn.DataPropertyName = "口座名義";
            this.口座名義DataGridViewTextBoxColumn.HeaderText = "口座名義";
            this.口座名義DataGridViewTextBoxColumn.Name = "口座名義DataGridViewTextBoxColumn";
            this.口座名義DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 登録年月日DataGridViewTextBoxColumn
            // 
            this.登録年月日DataGridViewTextBoxColumn.DataPropertyName = "登録年月日";
            this.登録年月日DataGridViewTextBoxColumn.HeaderText = "登録年月日";
            this.登録年月日DataGridViewTextBoxColumn.Name = "登録年月日DataGridViewTextBoxColumn";
            this.登録年月日DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 変更年月日DataGridViewTextBoxColumn
            // 
            this.変更年月日DataGridViewTextBoxColumn.DataPropertyName = "変更年月日";
            this.変更年月日DataGridViewTextBoxColumn.HeaderText = "変更年月日";
            this.変更年月日DataGridViewTextBoxColumn.Name = "変更年月日DataGridViewTextBoxColumn";
            this.変更年月日DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 振込口座BindingSource
            // 
            this.振込口座BindingSource.DataMember = "振込口座";
            this.振込口座BindingSource.DataSource = this.darwinDataSet;
            // 
            // darwinDataSet
            // 
            this.darwinDataSet.DataSetName = "darwinDataSet";
            this.darwinDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // btnCsv
            // 
            this.btnCsv.BackColor = System.Drawing.SystemColors.Control;
            this.btnCsv.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCsv.Location = new System.Drawing.Point(471, 362);
            this.btnCsv.Name = "btnCsv";
            this.btnCsv.Size = new System.Drawing.Size(89, 31);
            this.btnCsv.TabIndex = 8;
            this.btnCsv.Text = "CSV出力(&D)";
            this.btnCsv.UseVisualStyleBackColor = false;
            this.btnCsv.Click += new System.EventHandler(this.btnCsv_Click);
            // 
            // btnDel
            // 
            this.btnDel.BackColor = System.Drawing.SystemColors.Control;
            this.btnDel.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDel.Location = new System.Drawing.Point(281, 362);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(89, 31);
            this.btnDel.TabIndex = 6;
            this.btnDel.Text = "削除(&D)";
            this.btnDel.UseVisualStyleBackColor = false;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnClr
            // 
            this.btnClr.BackColor = System.Drawing.SystemColors.Control;
            this.btnClr.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnClr.Location = new System.Drawing.Point(376, 362);
            this.btnClr.Name = "btnClr";
            this.btnClr.Size = new System.Drawing.Size(89, 31);
            this.btnClr.TabIndex = 7;
            this.btnClr.Text = "取消(&C)";
            this.btnClr.UseVisualStyleBackColor = false;
            this.btnClr.Click += new System.EventHandler(this.btnClr_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.BackColor = System.Drawing.SystemColors.Control;
            this.btnRtn.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(566, 362);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(89, 31);
            this.btnRtn.TabIndex = 9;
            this.btnRtn.Text = "戻る(&R)";
            this.btnRtn.UseVisualStyleBackColor = false;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.SystemColors.Control;
            this.btnUpdate.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnUpdate.Location = new System.Drawing.Point(186, 362);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(89, 31);
            this.btnUpdate.TabIndex = 5;
            this.btnUpdate.Text = "更新(&U)";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // txtMeigi
            // 
            this.txtMeigi.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMeigi.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMeigi.Location = new System.Drawing.Point(80, 320);
            this.txtMeigi.MaxLength = 50;
            this.txtMeigi.Name = "txtMeigi";
            this.txtMeigi.Size = new System.Drawing.Size(575, 22);
            this.txtMeigi.TabIndex = 4;
            this.txtMeigi.Leave += new System.EventHandler(this.txtLeave);
            this.txtMeigi.Enter += new System.EventHandler(this.txtEnter);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label4.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label4.Location = new System.Drawing.Point(11, 320);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 22);
            this.label4.TabIndex = 18;
            this.label4.Text = "口座名義";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtName1
            // 
            this.txtName1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName1.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName1.Location = new System.Drawing.Point(80, 208);
            this.txtName1.MaxLength = 50;
            this.txtName1.Name = "txtName1";
            this.txtName1.Size = new System.Drawing.Size(575, 22);
            this.txtName1.TabIndex = 0;
            this.txtName1.Leave += new System.EventHandler(this.txtLeave);
            this.txtName1.Enter += new System.EventHandler(this.txtEnter);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label1.Location = new System.Drawing.Point(11, 208);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 22);
            this.label1.TabIndex = 20;
            this.label1.Text = "銀行名";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtName2
            // 
            this.txtName2.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName2.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName2.Location = new System.Drawing.Point(80, 236);
            this.txtName2.MaxLength = 50;
            this.txtName2.Name = "txtName2";
            this.txtName2.Size = new System.Drawing.Size(575, 22);
            this.txtName2.TabIndex = 1;
            this.txtName2.Leave += new System.EventHandler(this.txtLeave);
            this.txtName2.Enter += new System.EventHandler(this.txtEnter);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label3.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label3.Location = new System.Drawing.Point(11, 236);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 22);
            this.label3.TabIndex = 22;
            this.label3.Text = "支店名";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // 振込口座TableAdapter
            // 
            this.振込口座TableAdapter.ClearBeforeFill = true;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label2.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label2.Location = new System.Drawing.Point(11, 264);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 23);
            this.label2.TabIndex = 24;
            this.label2.Text = "口座種別";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtNumber
            // 
            this.txtNumber.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtNumber.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtNumber.Location = new System.Drawing.Point(80, 292);
            this.txtNumber.MaxLength = 7;
            this.txtNumber.Name = "txtNumber";
            this.txtNumber.Size = new System.Drawing.Size(80, 22);
            this.txtNumber.TabIndex = 3;
            this.txtNumber.Leave += new System.EventHandler(this.txtLeave);
            this.txtNumber.Enter += new System.EventHandler(this.txtEnter);
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label5.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label5.Location = new System.Drawing.Point(11, 292);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 22);
            this.label5.TabIndex = 26;
            this.label5.Text = "口座番号";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(80, 264);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(80, 23);
            this.comboBox1.TabIndex = 2;
            this.comboBox1.Leave += new System.EventHandler(this.txtLeave);
            this.comboBox1.Enter += new System.EventHandler(this.txtEnter);
            // 
            // frmKouza
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(668, 402);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.txtNumber);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtName2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtName1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtMeigi);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnCsv);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnClr);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmKouza";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "銀行口座マスター保守";
            this.Load += new System.EventHandler(this.form_Load);
            this.Enter += new System.EventHandler(this.txtEnter);
            this.Leave += new System.EventHandler(this.txtLeave);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.振込口座BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.darwinDataSet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        internal System.Windows.Forms.Button btnCsv;
        internal System.Windows.Forms.Button btnDel;
        internal System.Windows.Forms.Button btnClr;
        internal System.Windows.Forms.Button btnRtn;
        internal System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.TextBox txtMeigi;
        private System.Windows.Forms.Label label4;
        private darwinDataSet darwinDataSet;
        private System.Windows.Forms.TextBox txtName1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtName2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.BindingSource 振込口座BindingSource;
        private posting.darwinDataSetTableAdapters.振込口座TableAdapter 振込口座TableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 金融機関名DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 支店名DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 口座種別DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 口座番号DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 口座名義DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 登録年月日DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 変更年月日DataGridViewTextBoxColumn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtNumber;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}
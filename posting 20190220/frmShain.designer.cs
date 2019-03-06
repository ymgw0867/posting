namespace posting
{
    partial class frmShain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmShain));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.氏名DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.フリガナDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.所属名1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.役職DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.入社年月日DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.備考DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.登録年月日DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.変更年月日DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.社員BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.darwinDataSet = new posting.darwinDataSet();
            this.btnCsv = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClr = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFuri = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.iDate = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.txtYaku = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.社員TableAdapter = new posting.darwinDataSetTableAdapters.社員TableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.社員BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.darwinDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.氏名DataGridViewTextBoxColumn,
            this.フリガナDataGridViewTextBoxColumn,
            this.所属名1,
            this.役職DataGridViewTextBoxColumn,
            this.入社年月日DataGridViewTextBoxColumn,
            this.備考DataGridViewTextBoxColumn,
            this.登録年月日DataGridViewTextBoxColumn,
            this.変更年月日DataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.社員BindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(643, 180);
            this.dataGridView1.TabIndex = 12;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
            // 
            // iDDataGridViewTextBoxColumn
            // 
            this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
            this.iDDataGridViewTextBoxColumn.HeaderText = "コード";
            this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
            this.iDDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 氏名DataGridViewTextBoxColumn
            // 
            this.氏名DataGridViewTextBoxColumn.DataPropertyName = "氏名";
            this.氏名DataGridViewTextBoxColumn.HeaderText = "氏名";
            this.氏名DataGridViewTextBoxColumn.Name = "氏名DataGridViewTextBoxColumn";
            this.氏名DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // フリガナDataGridViewTextBoxColumn
            // 
            this.フリガナDataGridViewTextBoxColumn.DataPropertyName = "フリガナ";
            this.フリガナDataGridViewTextBoxColumn.HeaderText = "フリガナ";
            this.フリガナDataGridViewTextBoxColumn.Name = "フリガナDataGridViewTextBoxColumn";
            this.フリガナDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 所属名1
            // 
            this.所属名1.DataPropertyName = "所属名1";
            this.所属名1.HeaderText = "所属";
            this.所属名1.Name = "所属名1";
            this.所属名1.ReadOnly = true;
            // 
            // 役職DataGridViewTextBoxColumn
            // 
            this.役職DataGridViewTextBoxColumn.DataPropertyName = "役職";
            this.役職DataGridViewTextBoxColumn.HeaderText = "役職";
            this.役職DataGridViewTextBoxColumn.Name = "役職DataGridViewTextBoxColumn";
            this.役職DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 入社年月日DataGridViewTextBoxColumn
            // 
            this.入社年月日DataGridViewTextBoxColumn.DataPropertyName = "入社年月日";
            this.入社年月日DataGridViewTextBoxColumn.HeaderText = "入社年月日";
            this.入社年月日DataGridViewTextBoxColumn.Name = "入社年月日DataGridViewTextBoxColumn";
            this.入社年月日DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 備考DataGridViewTextBoxColumn
            // 
            this.備考DataGridViewTextBoxColumn.DataPropertyName = "備考";
            this.備考DataGridViewTextBoxColumn.HeaderText = "備考";
            this.備考DataGridViewTextBoxColumn.Name = "備考DataGridViewTextBoxColumn";
            this.備考DataGridViewTextBoxColumn.ReadOnly = true;
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
            // 社員BindingSource
            // 
            this.社員BindingSource.DataMember = "社員";
            this.社員BindingSource.DataSource = this.darwinDataSet;
            // 
            // darwinDataSet
            // 
            this.darwinDataSet.DataSetName = "darwinDataSet";
            this.darwinDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // btnCsv
            // 
            this.btnCsv.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCsv.BackColor = System.Drawing.SystemColors.Control;
            this.btnCsv.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCsv.Location = new System.Drawing.Point(471, 420);
            this.btnCsv.Name = "btnCsv";
            this.btnCsv.Size = new System.Drawing.Size(89, 31);
            this.btnCsv.TabIndex = 10;
            this.btnCsv.Text = "CSV出力(&D)";
            this.btnCsv.UseVisualStyleBackColor = false;
            this.btnCsv.Click += new System.EventHandler(this.btnCsv_Click);
            // 
            // btnDel
            // 
            this.btnDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDel.BackColor = System.Drawing.SystemColors.Control;
            this.btnDel.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDel.Location = new System.Drawing.Point(281, 420);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(89, 31);
            this.btnDel.TabIndex = 8;
            this.btnDel.Text = "削除(&D)";
            this.btnDel.UseVisualStyleBackColor = false;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnClr
            // 
            this.btnClr.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClr.BackColor = System.Drawing.SystemColors.Control;
            this.btnClr.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnClr.Location = new System.Drawing.Point(376, 420);
            this.btnClr.Name = "btnClr";
            this.btnClr.Size = new System.Drawing.Size(89, 31);
            this.btnClr.TabIndex = 9;
            this.btnClr.Text = "取消(&C)";
            this.btnClr.UseVisualStyleBackColor = false;
            this.btnClr.Click += new System.EventHandler(this.btnClr_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRtn.BackColor = System.Drawing.SystemColors.Control;
            this.btnRtn.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(566, 420);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(89, 31);
            this.btnRtn.TabIndex = 11;
            this.btnRtn.Text = "戻る(&R)";
            this.btnRtn.UseVisualStyleBackColor = false;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUpdate.BackColor = System.Drawing.SystemColors.Control;
            this.btnUpdate.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnUpdate.Location = new System.Drawing.Point(186, 420);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(89, 31);
            this.btnUpdate.TabIndex = 7;
            this.btnUpdate.Text = "更新(&U)";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label1.Location = new System.Drawing.Point(11, 208);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 22);
            this.label1.TabIndex = 12;
            this.label1.Text = "コード";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtCode
            // 
            this.txtCode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtCode.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtCode.Location = new System.Drawing.Point(80, 208);
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(51, 22);
            this.txtCode.TabIndex = 0;
            this.txtCode.Leave += new System.EventHandler(this.txtLeave);
            this.txtCode.Enter += new System.EventHandler(this.txtEnter);
            // 
            // txtName
            // 
            this.txtName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtName.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName.Location = new System.Drawing.Point(80, 236);
            this.txtName.MaxLength = 20;
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(575, 22);
            this.txtName.TabIndex = 1;
            this.txtName.TextChanged += new System.EventHandler(this.txtName_TextChanged);
            this.txtName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtName_KeyDown);
            this.txtName.Leave += new System.EventHandler(this.txtLeave);
            this.txtName.Enter += new System.EventHandler(this.txtEnter);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label2.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label2.Location = new System.Drawing.Point(11, 236);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 22);
            this.label2.TabIndex = 14;
            this.label2.Text = "氏名";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFuri
            // 
            this.txtFuri.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtFuri.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFuri.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtFuri.Location = new System.Drawing.Point(80, 264);
            this.txtFuri.MaxLength = 50;
            this.txtFuri.Name = "txtFuri";
            this.txtFuri.Size = new System.Drawing.Size(575, 22);
            this.txtFuri.TabIndex = 2;
            this.txtFuri.Leave += new System.EventHandler(this.txtLeave);
            this.txtFuri.Enter += new System.EventHandler(this.txtEnter);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label3.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label3.Location = new System.Drawing.Point(11, 264);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 22);
            this.label3.TabIndex = 16;
            this.label3.Text = "フリガナ";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtMemo
            // 
            this.txtMemo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtMemo.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMemo.Location = new System.Drawing.Point(80, 376);
            this.txtMemo.MaxLength = 50;
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(576, 22);
            this.txtMemo.TabIndex = 6;
            this.txtMemo.Leave += new System.EventHandler(this.txtLeave);
            this.txtMemo.Enter += new System.EventHandler(this.txtEnter);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label4.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label4.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label4.Location = new System.Drawing.Point(11, 376);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 22);
            this.label4.TabIndex = 18;
            this.label4.Text = "備考";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label5.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label5.Location = new System.Drawing.Point(11, 292);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 23);
            this.label5.TabIndex = 20;
            this.label5.Text = "所属";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label5.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.label5_MouseDoubleClick);
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label6.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label6.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label6.Location = new System.Drawing.Point(11, 348);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(70, 22);
            this.label6.TabIndex = 22;
            this.label6.Text = "入社日";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // iDate
            // 
            this.iDate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.iDate.CalendarFont = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.iDate.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.iDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.iDate.Location = new System.Drawing.Point(80, 348);
            this.iDate.Name = "iDate";
            this.iDate.ShowCheckBox = true;
            this.iDate.Size = new System.Drawing.Size(144, 22);
            this.iDate.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label7.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label7.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label7.Location = new System.Drawing.Point(11, 320);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 22);
            this.label7.TabIndex = 24;
            this.label7.Text = "役職";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtYaku
            // 
            this.txtYaku.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtYaku.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtYaku.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtYaku.Location = new System.Drawing.Point(80, 320);
            this.txtYaku.MaxLength = 50;
            this.txtYaku.Name = "txtYaku";
            this.txtYaku.Size = new System.Drawing.Size(575, 22);
            this.txtYaku.TabIndex = 4;
            this.txtYaku.Leave += new System.EventHandler(this.txtLeave);
            this.txtYaku.Enter += new System.EventHandler(this.txtEnter);
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.comboBox1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(80, 292);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(340, 23);
            this.comboBox1.TabIndex = 3;
            // 
            // 社員TableAdapter
            // 
            this.社員TableAdapter.ClearBeforeFill = true;
            // 
            // frmShain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(668, 464);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.iDate);
            this.Controls.Add(this.txtYaku);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtMemo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtFuri);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCode);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCsv);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnClr);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmShain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "社員マスター保守";
            this.Load += new System.EventHandler(this.Gengo_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.社員BindingSource)).EndInit();
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
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtFuri;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Windows.Forms.Label label4;
        private darwinDataSet darwinDataSet;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.BindingSource 社員BindingSource;
        private posting.darwinDataSetTableAdapters.社員TableAdapter 社員TableAdapter;
        private System.Windows.Forms.DateTimePicker iDate;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtYaku;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 氏名DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn フリガナDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 所属名1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 役職DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 入社年月日DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 備考DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 登録年月日DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 変更年月日DataGridViewTextBoxColumn;
    }
}
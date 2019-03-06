namespace posting
{
    partial class frmOffice
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmOffice));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.名称DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.郵便番号DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.住所1DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.住所2DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.電話番号DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fAX番号DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.備考DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.登録年月日DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.変更年月日DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.事業所BindingSource = new System.Windows.Forms.BindingSource(this.components);
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
            this.txtZipCode = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtAddress1 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtAddress2 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtTel = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtFax = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.事業所TableAdapter = new posting.darwinDataSetTableAdapters.事業所TableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.事業所BindingSource)).BeginInit();
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
            this.名称DataGridViewTextBoxColumn,
            this.郵便番号DataGridViewTextBoxColumn,
            this.住所1DataGridViewTextBoxColumn,
            this.住所2DataGridViewTextBoxColumn,
            this.電話番号DataGridViewTextBoxColumn,
            this.fAX番号DataGridViewTextBoxColumn,
            this.備考DataGridViewTextBoxColumn,
            this.登録年月日DataGridViewTextBoxColumn,
            this.変更年月日DataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.事業所BindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(643, 180);
            this.dataGridView1.TabIndex = 13;
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
            // 名称DataGridViewTextBoxColumn
            // 
            this.名称DataGridViewTextBoxColumn.DataPropertyName = "名称";
            this.名称DataGridViewTextBoxColumn.HeaderText = "名称";
            this.名称DataGridViewTextBoxColumn.Name = "名称DataGridViewTextBoxColumn";
            this.名称DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 郵便番号DataGridViewTextBoxColumn
            // 
            this.郵便番号DataGridViewTextBoxColumn.DataPropertyName = "郵便番号";
            this.郵便番号DataGridViewTextBoxColumn.HeaderText = "郵便番号";
            this.郵便番号DataGridViewTextBoxColumn.Name = "郵便番号DataGridViewTextBoxColumn";
            this.郵便番号DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 住所1DataGridViewTextBoxColumn
            // 
            this.住所1DataGridViewTextBoxColumn.DataPropertyName = "住所1";
            this.住所1DataGridViewTextBoxColumn.HeaderText = "住所1";
            this.住所1DataGridViewTextBoxColumn.Name = "住所1DataGridViewTextBoxColumn";
            this.住所1DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 住所2DataGridViewTextBoxColumn
            // 
            this.住所2DataGridViewTextBoxColumn.DataPropertyName = "住所2";
            this.住所2DataGridViewTextBoxColumn.HeaderText = "住所2";
            this.住所2DataGridViewTextBoxColumn.Name = "住所2DataGridViewTextBoxColumn";
            this.住所2DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // 電話番号DataGridViewTextBoxColumn
            // 
            this.電話番号DataGridViewTextBoxColumn.DataPropertyName = "電話番号";
            this.電話番号DataGridViewTextBoxColumn.HeaderText = "電話番号";
            this.電話番号DataGridViewTextBoxColumn.Name = "電話番号DataGridViewTextBoxColumn";
            this.電話番号DataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // fAX番号DataGridViewTextBoxColumn
            // 
            this.fAX番号DataGridViewTextBoxColumn.DataPropertyName = "FAX番号";
            this.fAX番号DataGridViewTextBoxColumn.HeaderText = "FAX番号";
            this.fAX番号DataGridViewTextBoxColumn.Name = "fAX番号DataGridViewTextBoxColumn";
            this.fAX番号DataGridViewTextBoxColumn.ReadOnly = true;
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
            // 事業所BindingSource
            // 
            this.事業所BindingSource.DataMember = "事業所";
            this.事業所BindingSource.DataSource = this.darwinDataSet;
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
            this.btnCsv.TabIndex = 11;
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
            this.btnDel.TabIndex = 9;
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
            this.btnClr.TabIndex = 10;
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
            this.btnRtn.TabIndex = 12;
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
            this.btnUpdate.TabIndex = 8;
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
            this.label1.Text = "№";
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
            this.txtCode.Enter += new System.EventHandler(this.txtEnter);
            this.txtCode.Leave += new System.EventHandler(this.txtLeave);
            // 
            // txtName
            // 
            this.txtName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtName.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName.Location = new System.Drawing.Point(80, 236);
            this.txtName.MaxLength = 50;
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(575, 22);
            this.txtName.TabIndex = 1;
            this.txtName.Enter += new System.EventHandler(this.txtEnter);
            this.txtName.Leave += new System.EventHandler(this.txtLeave);
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
            this.label2.Text = "名称";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtZipCode
            // 
            this.txtZipCode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtZipCode.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtZipCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtZipCode.Location = new System.Drawing.Point(80, 264);
            this.txtZipCode.MaxLength = 8;
            this.txtZipCode.Name = "txtZipCode";
            this.txtZipCode.Size = new System.Drawing.Size(77, 22);
            this.txtZipCode.TabIndex = 2;
            this.txtZipCode.TextChanged += new System.EventHandler(this.txtZipCode_TextChanged);
            this.txtZipCode.Enter += new System.EventHandler(this.txtEnter);
            this.txtZipCode.Leave += new System.EventHandler(this.txtLeave);
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
            this.label3.Text = "郵便番号";
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
            this.txtMemo.TabIndex = 7;
            this.txtMemo.Enter += new System.EventHandler(this.txtEnter);
            this.txtMemo.Leave += new System.EventHandler(this.txtLeave);
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
            // txtAddress1
            // 
            this.txtAddress1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtAddress1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress1.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress1.Location = new System.Drawing.Point(80, 292);
            this.txtAddress1.MaxLength = 50;
            this.txtAddress1.Name = "txtAddress1";
            this.txtAddress1.Size = new System.Drawing.Size(575, 22);
            this.txtAddress1.TabIndex = 3;
            this.txtAddress1.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress1.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label5.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label5.Location = new System.Drawing.Point(11, 292);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 22);
            this.label5.TabIndex = 20;
            this.label5.Text = "住所１";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtAddress2
            // 
            this.txtAddress2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtAddress2.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress2.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress2.Location = new System.Drawing.Point(80, 320);
            this.txtAddress2.MaxLength = 50;
            this.txtAddress2.Name = "txtAddress2";
            this.txtAddress2.Size = new System.Drawing.Size(576, 22);
            this.txtAddress2.TabIndex = 4;
            this.txtAddress2.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress2.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label6.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label6.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label6.Location = new System.Drawing.Point(11, 320);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(70, 22);
            this.label6.TabIndex = 22;
            this.label6.Text = "住所２";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtTel
            // 
            this.txtTel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtTel.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTel.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTel.Location = new System.Drawing.Point(80, 348);
            this.txtTel.MaxLength = 20;
            this.txtTel.Name = "txtTel";
            this.txtTel.Size = new System.Drawing.Size(144, 22);
            this.txtTel.TabIndex = 5;
            this.txtTel.Enter += new System.EventHandler(this.txtEnter);
            this.txtTel.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label7.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label7.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label7.Location = new System.Drawing.Point(11, 348);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 22);
            this.label7.TabIndex = 24;
            this.label7.Text = "電話番号";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFax
            // 
            this.txtFax.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtFax.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFax.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtFax.Location = new System.Drawing.Point(303, 348);
            this.txtFax.MaxLength = 20;
            this.txtFax.Name = "txtFax";
            this.txtFax.Size = new System.Drawing.Size(144, 22);
            this.txtFax.TabIndex = 6;
            this.txtFax.Enter += new System.EventHandler(this.txtEnter);
            this.txtFax.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label8.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label8.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label8.Location = new System.Drawing.Point(234, 348);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(70, 22);
            this.label8.TabIndex = 26;
            this.label8.Text = "FAX番号";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // 事業所TableAdapter
            // 
            this.事業所TableAdapter.ClearBeforeFill = true;
            // 
            // frmOffice
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(668, 464);
            this.Controls.Add(this.txtFax);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtTel);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtAddress2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtAddress1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtMemo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtZipCode);
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
            this.Name = "frmOffice";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "事業所マスター保守";
            this.Load += new System.EventHandler(this.Gengo_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.事業所BindingSource)).EndInit();
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
        private System.Windows.Forms.TextBox txtZipCode;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Windows.Forms.Label label4;
        private darwinDataSet darwinDataSet;
        private System.Windows.Forms.TextBox txtAddress1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtAddress2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtTel;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtFax;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.BindingSource 事業所BindingSource;
        private posting.darwinDataSetTableAdapters.事業所TableAdapter 事業所TableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 名称DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 郵便番号DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 住所1DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 住所2DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 電話番号DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fAX番号DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 備考DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 登録年月日DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 変更年月日DataGridViewTextBoxColumn;
    }
}
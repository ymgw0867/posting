namespace posting
{
    partial class frmClient
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmClient));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnCsv = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClr = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.txtAddress1 = new System.Windows.Forms.TextBox();
            this.txtName1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFuri = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtName2 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbKeisho = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtTantou = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtBusho = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.mtxtZipCode = new System.Windows.Forms.MaskedTextBox();
            this.txtAddress2 = new System.Windows.Forms.TextBox();
            this.txtTel = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtFax = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.cmbShain = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtShimebi = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.cmbTax = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbCity = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.txtAddress2S = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.mtxtZipCodeS = new System.Windows.Forms.MaskedTextBox();
            this.cmbCityS = new System.Windows.Forms.ComboBox();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.txtAddress1S = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.label27 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.txtNameSeikyu = new System.Windows.Forms.TextBox();
            this.label24 = new System.Windows.Forms.Label();
            this.得意先BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.darwinDataSet = new posting.darwinDataSet();
            this.得意先TableAdapter = new posting.darwinDataSetTableAdapters.得意先TableAdapter();
            this.label25 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label28 = new System.Windows.Forms.Label();
            this.txtsTel = new System.Windows.Forms.TextBox();
            this.label29 = new System.Windows.Forms.Label();
            this.txtsZip = new System.Windows.Forms.TextBox();
            this.txtTantouS = new System.Windows.Forms.TextBox();
            this.label30 = new System.Windows.Forms.Label();
            this.txtBushoS = new System.Windows.Forms.TextBox();
            this.label31 = new System.Windows.Forms.Label();
            this.label32 = new System.Windows.Forms.Label();
            this.cmbKeishoS = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.得意先BindingSource)).BeginInit();
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
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 18;
            this.dataGridView1.Size = new System.Drawing.Size(898, 180);
            this.dataGridView1.TabIndex = 26;
            this.dataGridView1.TabStop = false;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
            // 
            // btnCsv
            // 
            this.btnCsv.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCsv.BackColor = System.Drawing.SystemColors.Control;
            this.btnCsv.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCsv.Location = new System.Drawing.Point(732, 620);
            this.btnCsv.Name = "btnCsv";
            this.btnCsv.Size = new System.Drawing.Size(89, 31);
            this.btnCsv.TabIndex = 29;
            this.btnCsv.Text = "CSV出力(&D)";
            this.btnCsv.UseVisualStyleBackColor = true;
            this.btnCsv.Click += new System.EventHandler(this.btnCsv_Click);
            // 
            // btnDel
            // 
            this.btnDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDel.BackColor = System.Drawing.SystemColors.Control;
            this.btnDel.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDel.Location = new System.Drawing.Point(556, 620);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(82, 31);
            this.btnDel.TabIndex = 27;
            this.btnDel.Text = "削除(&D)";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnClr
            // 
            this.btnClr.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClr.BackColor = System.Drawing.SystemColors.Control;
            this.btnClr.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnClr.Location = new System.Drawing.Point(644, 620);
            this.btnClr.Name = "btnClr";
            this.btnClr.Size = new System.Drawing.Size(82, 31);
            this.btnClr.TabIndex = 28;
            this.btnClr.Text = "取消(&C)";
            this.btnClr.UseVisualStyleBackColor = true;
            this.btnClr.Click += new System.EventHandler(this.btnClr_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRtn.BackColor = System.Drawing.SystemColors.Control;
            this.btnRtn.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(827, 620);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(82, 31);
            this.btnRtn.TabIndex = 30;
            this.btnRtn.Text = "戻る(&R)";
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUpdate.BackColor = System.Drawing.SystemColors.Control;
            this.btnUpdate.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnUpdate.Location = new System.Drawing.Point(468, 620);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(82, 31);
            this.btnUpdate.TabIndex = 26;
            this.btnUpdate.Text = "更新(&U)";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // txtAddress1
            // 
            this.txtAddress1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtAddress1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress1.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress1.Location = new System.Drawing.Point(81, 459);
            this.txtAddress1.MaxLength = 50;
            this.txtAddress1.Name = "txtAddress1";
            this.txtAddress1.Size = new System.Drawing.Size(358, 23);
            this.txtAddress1.TabIndex = 8;
            this.txtAddress1.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress1.Leave += new System.EventHandler(this.txtLeave);
            // 
            // txtName1
            // 
            this.txtName1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtName1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName1.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName1.Location = new System.Drawing.Point(81, 258);
            this.txtName1.MaxLength = 50;
            this.txtName1.Name = "txtName1";
            this.txtName1.Size = new System.Drawing.Size(358, 23);
            this.txtName1.TabIndex = 0;
            this.txtName1.TextChanged += new System.EventHandler(this.txtName1_TextChanged);
            this.txtName1.Enter += new System.EventHandler(this.txtEnter);
            this.txtName1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtName1_KeyDown);
            this.txtName1.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label1.Location = new System.Drawing.Point(12, 258);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 23);
            this.label1.TabIndex = 20;
            this.label1.Text = "略称";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFuri
            // 
            this.txtFuri.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtFuri.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFuri.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtFuri.Location = new System.Drawing.Point(81, 283);
            this.txtFuri.MaxLength = 50;
            this.txtFuri.Name = "txtFuri";
            this.txtFuri.Size = new System.Drawing.Size(358, 23);
            this.txtFuri.TabIndex = 1;
            this.txtFuri.Enter += new System.EventHandler(this.txtEnter);
            this.txtFuri.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label3.Location = new System.Drawing.Point(12, 283);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 23);
            this.label3.TabIndex = 22;
            this.label3.Text = "フリガナ";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label5.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label5.Location = new System.Drawing.Point(12, 409);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 23);
            this.label5.TabIndex = 26;
            this.label5.Text = "郵便番号";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtName2
            // 
            this.txtName2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtName2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName2.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName2.Location = new System.Drawing.Point(81, 308);
            this.txtName2.MaxLength = 50;
            this.txtName2.Name = "txtName2";
            this.txtName2.Size = new System.Drawing.Size(358, 23);
            this.txtName2.TabIndex = 2;
            this.txtName2.Enter += new System.EventHandler(this.txtEnter);
            this.txtName2.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label6.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label6.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label6.Location = new System.Drawing.Point(12, 308);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(70, 23);
            this.label6.TabIndex = 28;
            this.label6.Text = "名称";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbKeisho
            // 
            this.cmbKeisho.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbKeisho.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbKeisho.FormattingEnabled = true;
            this.cmbKeisho.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.cmbKeisho.Location = new System.Drawing.Point(81, 358);
            this.cmbKeisho.Name = "cmbKeisho";
            this.cmbKeisho.Size = new System.Drawing.Size(80, 23);
            this.cmbKeisho.TabIndex = 4;
            this.cmbKeisho.SelectedIndexChanged += new System.EventHandler(this.cmbKeisho_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label7.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label7.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label7.Location = new System.Drawing.Point(12, 358);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 23);
            this.label7.TabIndex = 30;
            this.label7.Text = "敬称";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtTantou
            // 
            this.txtTantou.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtTantou.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTantou.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtTantou.Location = new System.Drawing.Point(81, 333);
            this.txtTantou.MaxLength = 50;
            this.txtTantou.Name = "txtTantou";
            this.txtTantou.Size = new System.Drawing.Size(358, 23);
            this.txtTantou.TabIndex = 3;
            this.txtTantou.Enter += new System.EventHandler(this.txtEnter);
            this.txtTantou.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label8.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label8.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label8.Location = new System.Drawing.Point(12, 333);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(70, 23);
            this.label8.TabIndex = 32;
            this.label8.Text = "ご担当者";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtBusho
            // 
            this.txtBusho.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtBusho.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtBusho.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtBusho.Location = new System.Drawing.Point(81, 384);
            this.txtBusho.MaxLength = 50;
            this.txtBusho.Name = "txtBusho";
            this.txtBusho.Size = new System.Drawing.Size(358, 23);
            this.txtBusho.TabIndex = 5;
            this.txtBusho.Enter += new System.EventHandler(this.txtEnter);
            this.txtBusho.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label9.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label9.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label9.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label9.Location = new System.Drawing.Point(12, 384);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(70, 23);
            this.label9.TabIndex = 34;
            this.label9.Text = "部署名";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // mtxtZipCode
            // 
            this.mtxtZipCode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.mtxtZipCode.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.mtxtZipCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.mtxtZipCode.Location = new System.Drawing.Point(81, 409);
            this.mtxtZipCode.Mask = "000-0000";
            this.mtxtZipCode.Name = "mtxtZipCode";
            this.mtxtZipCode.Size = new System.Drawing.Size(80, 23);
            this.mtxtZipCode.TabIndex = 6;
            this.mtxtZipCode.TextChanged += new System.EventHandler(this.mtxtZipCode_TextChanged);
            this.mtxtZipCode.Enter += new System.EventHandler(this.txtEnter);
            this.mtxtZipCode.Leave += new System.EventHandler(this.txtLeave);
            // 
            // txtAddress2
            // 
            this.txtAddress2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtAddress2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress2.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress2.Location = new System.Drawing.Point(81, 484);
            this.txtAddress2.MaxLength = 50;
            this.txtAddress2.Name = "txtAddress2";
            this.txtAddress2.Size = new System.Drawing.Size(358, 23);
            this.txtAddress2.TabIndex = 9;
            this.txtAddress2.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress2.Leave += new System.EventHandler(this.txtLeave);
            // 
            // txtTel
            // 
            this.txtTel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtTel.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTel.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTel.Location = new System.Drawing.Point(81, 509);
            this.txtTel.MaxLength = 50;
            this.txtTel.Name = "txtTel";
            this.txtTel.Size = new System.Drawing.Size(358, 23);
            this.txtTel.TabIndex = 10;
            this.txtTel.Enter += new System.EventHandler(this.txtEnter);
            this.txtTel.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label11.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label11.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label11.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label11.Location = new System.Drawing.Point(12, 509);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(70, 23);
            this.label11.TabIndex = 39;
            this.label11.Text = "電話番号";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFax
            // 
            this.txtFax.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtFax.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFax.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtFax.Location = new System.Drawing.Point(81, 534);
            this.txtFax.MaxLength = 50;
            this.txtFax.Name = "txtFax";
            this.txtFax.Size = new System.Drawing.Size(358, 23);
            this.txtFax.TabIndex = 11;
            this.txtFax.Enter += new System.EventHandler(this.txtEnter);
            this.txtFax.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label12
            // 
            this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label12.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label12.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label12.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label12.Location = new System.Drawing.Point(12, 534);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(70, 23);
            this.label12.TabIndex = 41;
            this.label12.Text = "FAX番号";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtEmail
            // 
            this.txtEmail.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtEmail.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtEmail.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtEmail.Location = new System.Drawing.Point(539, 258);
            this.txtEmail.MaxLength = 50;
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(358, 23);
            this.txtEmail.TabIndex = 12;
            this.txtEmail.Enter += new System.EventHandler(this.txtEnter);
            this.txtEmail.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label13.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label13.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label13.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label13.Location = new System.Drawing.Point(470, 258);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(70, 23);
            this.label13.TabIndex = 43;
            this.label13.Text = "e_mail";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbShain
            // 
            this.cmbShain.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbShain.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbShain.FormattingEnabled = true;
            this.cmbShain.Location = new System.Drawing.Point(539, 283);
            this.cmbShain.Name = "cmbShain";
            this.cmbShain.Size = new System.Drawing.Size(358, 23);
            this.cmbShain.TabIndex = 13;
            // 
            // label14
            // 
            this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label14.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label14.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label14.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label14.Location = new System.Drawing.Point(470, 283);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(70, 23);
            this.label14.TabIndex = 45;
            this.label14.Text = "担当社員";
            this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label14.Click += new System.EventHandler(this.label14_Click);
            // 
            // txtShimebi
            // 
            this.txtShimebi.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtShimebi.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtShimebi.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtShimebi.Location = new System.Drawing.Point(539, 308);
            this.txtShimebi.MaxLength = 2;
            this.txtShimebi.Name = "txtShimebi";
            this.txtShimebi.Size = new System.Drawing.Size(47, 23);
            this.txtShimebi.TabIndex = 14;
            this.txtShimebi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtShimebi.Enter += new System.EventHandler(this.txtEnter);
            this.txtShimebi.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtShimebi_KeyPress);
            this.txtShimebi.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label15
            // 
            this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label15.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label15.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label15.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label15.Location = new System.Drawing.Point(470, 308);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(70, 23);
            this.label15.TabIndex = 47;
            this.label15.Text = "締日";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbTax
            // 
            this.cmbTax.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbTax.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbTax.FormattingEnabled = true;
            this.cmbTax.Location = new System.Drawing.Point(539, 333);
            this.cmbTax.Name = "cmbTax";
            this.cmbTax.Size = new System.Drawing.Size(154, 23);
            this.cmbTax.TabIndex = 15;
            // 
            // label16
            // 
            this.label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label16.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label16.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label16.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label16.Location = new System.Drawing.Point(470, 333);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(70, 23);
            this.label16.TabIndex = 49;
            this.label16.Text = "税通知";
            this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label4.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label4.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label4.Location = new System.Drawing.Point(12, 459);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 23);
            this.label4.TabIndex = 18;
            this.label4.Text = "住所１";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label2.Location = new System.Drawing.Point(12, 434);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 23);
            this.label2.TabIndex = 24;
            this.label2.Text = "都道府県";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbCity
            // 
            this.cmbCity.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbCity.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbCity.FormattingEnabled = true;
            this.cmbCity.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.cmbCity.Location = new System.Drawing.Point(81, 434);
            this.cmbCity.Name = "cmbCity";
            this.cmbCity.Size = new System.Drawing.Size(117, 23);
            this.cmbCity.TabIndex = 7;
            this.cmbCity.Enter += new System.EventHandler(this.txtEnter);
            this.cmbCity.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label10.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label10.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label10.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label10.Location = new System.Drawing.Point(12, 484);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(70, 23);
            this.label10.TabIndex = 37;
            this.label10.Text = "住所２";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label17
            // 
            this.label17.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label17.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label17.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label17.Location = new System.Drawing.Point(461, 369);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(449, 189);
            this.label17.TabIndex = 16;
            // 
            // txtAddress2S
            // 
            this.txtAddress2S.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtAddress2S.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress2S.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress2S.Location = new System.Drawing.Point(539, 455);
            this.txtAddress2S.MaxLength = 50;
            this.txtAddress2S.Name = "txtAddress2S";
            this.txtAddress2S.Size = new System.Drawing.Size(358, 23);
            this.txtAddress2S.TabIndex = 21;
            this.txtAddress2S.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress2S.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label18
            // 
            this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label18.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label18.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label18.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label18.Location = new System.Drawing.Point(470, 455);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(70, 23);
            this.label18.TabIndex = 58;
            this.label18.Text = "住所２";
            this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // mtxtZipCodeS
            // 
            this.mtxtZipCodeS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.mtxtZipCodeS.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.mtxtZipCodeS.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.mtxtZipCodeS.Location = new System.Drawing.Point(539, 405);
            this.mtxtZipCodeS.Mask = "000-0000";
            this.mtxtZipCodeS.Name = "mtxtZipCodeS";
            this.mtxtZipCodeS.Size = new System.Drawing.Size(90, 23);
            this.mtxtZipCodeS.TabIndex = 18;
            this.mtxtZipCodeS.TextChanged += new System.EventHandler(this.mtxtZipCode_TextChanged);
            this.mtxtZipCodeS.Enter += new System.EventHandler(this.txtEnter);
            this.mtxtZipCodeS.Leave += new System.EventHandler(this.txtLeave);
            // 
            // cmbCityS
            // 
            this.cmbCityS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbCityS.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbCityS.FormattingEnabled = true;
            this.cmbCityS.Location = new System.Drawing.Point(704, 405);
            this.cmbCityS.Name = "cmbCityS";
            this.cmbCityS.Size = new System.Drawing.Size(193, 23);
            this.cmbCityS.TabIndex = 19;
            // 
            // label19
            // 
            this.label19.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label19.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label19.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label19.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label19.Location = new System.Drawing.Point(470, 405);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(70, 23);
            this.label19.TabIndex = 55;
            this.label19.Text = "郵便番号";
            this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label20
            // 
            this.label20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label20.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label20.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label20.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label20.Location = new System.Drawing.Point(635, 405);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(70, 23);
            this.label20.TabIndex = 54;
            this.label20.Text = "都道府県";
            this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtAddress1S
            // 
            this.txtAddress1S.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtAddress1S.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress1S.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress1S.Location = new System.Drawing.Point(539, 430);
            this.txtAddress1S.MaxLength = 50;
            this.txtAddress1S.Name = "txtAddress1S";
            this.txtAddress1S.Size = new System.Drawing.Size(358, 23);
            this.txtAddress1S.TabIndex = 20;
            this.txtAddress1S.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress1S.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label21
            // 
            this.label21.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label21.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label21.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label21.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label21.Location = new System.Drawing.Point(470, 430);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(70, 23);
            this.label21.TabIndex = 53;
            this.label21.Text = "住所１";
            this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label22
            // 
            this.label22.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label22.Location = new System.Drawing.Point(470, 361);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(79, 15);
            this.label22.TabIndex = 59;
            this.label22.Text = "請求書送付先";
            // 
            // txtMemo
            // 
            this.txtMemo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtMemo.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMemo.Location = new System.Drawing.Point(538, 565);
            this.txtMemo.MaxLength = 50;
            this.txtMemo.Multiline = true;
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(371, 38);
            this.txtMemo.TabIndex = 25;
            this.txtMemo.Enter += new System.EventHandler(this.txtEnter);
            this.txtMemo.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label23
            // 
            this.label23.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label23.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label23.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label23.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label23.Location = new System.Drawing.Point(469, 565);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(70, 38);
            this.label23.TabIndex = 61;
            this.label23.Text = "備考";
            this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label27
            // 
            this.label27.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label27.AutoSize = true;
            this.label27.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label27.Location = new System.Drawing.Point(316, 212);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(31, 15);
            this.label27.TabIndex = 76;
            this.label27.Text = "略称";
            this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(828, 204);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(69, 30);
            this.button1.TabIndex = 33;
            this.button1.Text = "検索";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox1.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.textBox1.Location = new System.Drawing.Point(353, 209);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(196, 23);
            this.textBox1.TabIndex = 31;
            this.textBox1.Enter += new System.EventHandler(this.txtEnter);
            this.textBox1.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label26
            // 
            this.label26.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label26.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label26.Location = new System.Drawing.Point(12, 199);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(898, 43);
            this.label26.TabIndex = 28;
            // 
            // txtNameSeikyu
            // 
            this.txtNameSeikyu.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtNameSeikyu.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtNameSeikyu.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtNameSeikyu.Location = new System.Drawing.Point(539, 380);
            this.txtNameSeikyu.MaxLength = 50;
            this.txtNameSeikyu.Name = "txtNameSeikyu";
            this.txtNameSeikyu.Size = new System.Drawing.Size(358, 23);
            this.txtNameSeikyu.TabIndex = 17;
            this.txtNameSeikyu.Enter += new System.EventHandler(this.txtEnter);
            this.txtNameSeikyu.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label24
            // 
            this.label24.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label24.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label24.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label24.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label24.Location = new System.Drawing.Point(470, 380);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(70, 23);
            this.label24.TabIndex = 78;
            this.label24.Text = "請求先名称";
            this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // 得意先BindingSource
            // 
            this.得意先BindingSource.DataMember = "得意先";
            this.得意先BindingSource.DataSource = this.darwinDataSet;
            // 
            // darwinDataSet
            // 
            this.darwinDataSet.DataSetName = "darwinDataSet";
            this.darwinDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // 得意先TableAdapter
            // 
            this.得意先TableAdapter.ClearBeforeFill = true;
            // 
            // label25
            // 
            this.label25.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label25.Location = new System.Drawing.Point(557, 212);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(43, 15);
            this.label25.TabIndex = 80;
            this.label25.Text = "請求先";
            this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox2.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.textBox2.Location = new System.Drawing.Point(602, 209);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(215, 23);
            this.textBox2.TabIndex = 32;
            this.textBox2.Enter += new System.EventHandler(this.txtEnter);
            this.textBox2.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label28
            // 
            this.label28.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label28.AutoSize = true;
            this.label28.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label28.Location = new System.Drawing.Point(19, 212);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(29, 15);
            this.label28.TabIndex = 82;
            this.label28.Text = "TEL";
            this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtsTel
            // 
            this.txtsTel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtsTel.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtsTel.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtsTel.Location = new System.Drawing.Point(49, 209);
            this.txtsTel.MaxLength = 14;
            this.txtsTel.Name = "txtsTel";
            this.txtsTel.Size = new System.Drawing.Size(114, 23);
            this.txtsTel.TabIndex = 29;
            this.txtsTel.Enter += new System.EventHandler(this.txtEnter);
            this.txtsTel.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label29
            // 
            this.label29.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label29.AutoSize = true;
            this.label29.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label29.Location = new System.Drawing.Point(173, 212);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(55, 15);
            this.label29.TabIndex = 84;
            this.label29.Text = "郵便番号";
            this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtsZip
            // 
            this.txtsZip.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtsZip.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtsZip.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtsZip.Location = new System.Drawing.Point(230, 209);
            this.txtsZip.MaxLength = 8;
            this.txtsZip.Name = "txtsZip";
            this.txtsZip.Size = new System.Drawing.Size(73, 23);
            this.txtsZip.TabIndex = 30;
            this.txtsZip.Enter += new System.EventHandler(this.txtEnter);
            this.txtsZip.Leave += new System.EventHandler(this.txtLeave);
            // 
            // txtTantouS
            // 
            this.txtTantouS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtTantouS.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTantouS.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtTantouS.Location = new System.Drawing.Point(539, 480);
            this.txtTantouS.MaxLength = 50;
            this.txtTantouS.Name = "txtTantouS";
            this.txtTantouS.Size = new System.Drawing.Size(358, 23);
            this.txtTantouS.TabIndex = 22;
            this.txtTantouS.Enter += new System.EventHandler(this.txtEnter);
            this.txtTantouS.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label30
            // 
            this.label30.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label30.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label30.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label30.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label30.Location = new System.Drawing.Point(470, 480);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(70, 23);
            this.label30.TabIndex = 86;
            this.label30.Text = "担当者名";
            this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtBushoS
            // 
            this.txtBushoS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtBushoS.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtBushoS.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtBushoS.Location = new System.Drawing.Point(539, 530);
            this.txtBushoS.MaxLength = 50;
            this.txtBushoS.Name = "txtBushoS";
            this.txtBushoS.Size = new System.Drawing.Size(358, 23);
            this.txtBushoS.TabIndex = 24;
            this.txtBushoS.Enter += new System.EventHandler(this.txtEnter);
            this.txtBushoS.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label31
            // 
            this.label31.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label31.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label31.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label31.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label31.Location = new System.Drawing.Point(470, 530);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(70, 23);
            this.label31.TabIndex = 88;
            this.label31.Text = "部署名";
            this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label32
            // 
            this.label32.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label32.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label32.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label32.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label32.Location = new System.Drawing.Point(470, 505);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(70, 22);
            this.label32.TabIndex = 90;
            this.label32.Text = "敬称";
            this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbKeishoS
            // 
            this.cmbKeishoS.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbKeishoS.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbKeishoS.FormattingEnabled = true;
            this.cmbKeishoS.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.cmbKeishoS.Location = new System.Drawing.Point(539, 504);
            this.cmbKeishoS.Name = "cmbKeishoS";
            this.cmbKeishoS.Size = new System.Drawing.Size(80, 23);
            this.cmbKeishoS.TabIndex = 23;
            // 
            // frmClient
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(924, 660);
            this.Controls.Add(this.cmbKeishoS);
            this.Controls.Add(this.label32);
            this.Controls.Add(this.txtBushoS);
            this.Controls.Add(this.label31);
            this.Controls.Add(this.txtTantouS);
            this.Controls.Add(this.label30);
            this.Controls.Add(this.label29);
            this.Controls.Add(this.txtsZip);
            this.Controls.Add(this.label28);
            this.Controls.Add(this.txtsTel);
            this.Controls.Add(this.label25);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.txtNameSeikyu);
            this.Controls.Add(this.label24);
            this.Controls.Add(this.label27);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label26);
            this.Controls.Add(this.txtMemo);
            this.Controls.Add(this.label23);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.txtAddress2S);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.mtxtZipCodeS);
            this.Controls.Add(this.cmbCityS);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.label20);
            this.Controls.Add(this.txtAddress1S);
            this.Controls.Add(this.label21);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.cmbTax);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.txtShimebi);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.cmbShain);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.txtEmail);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.txtFax);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.txtTel);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.txtAddress2);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.mtxtZipCode);
            this.Controls.Add(this.txtBusho);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.txtTantou);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.cmbKeisho);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtName2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cmbCity);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtFuri);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtName1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtAddress1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnCsv);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnClr);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmClient";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "得意先マスター保守";
            this.Load += new System.EventHandler(this.form_Load);
            this.Enter += new System.EventHandler(this.txtEnter);
            this.Leave += new System.EventHandler(this.txtLeave);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.得意先BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.darwinDataSet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox txtAddress1;
        private darwinDataSet darwinDataSet;
        private System.Windows.Forms.TextBox txtName1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFuri;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.BindingSource 得意先BindingSource;
        private posting.darwinDataSetTableAdapters.得意先TableAdapter 得意先TableAdapter;
        private System.Windows.Forms.TextBox txtName2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbKeisho;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtTantou;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtBusho;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.MaskedTextBox mtxtZipCode;
        private System.Windows.Forms.TextBox txtAddress2;
        private System.Windows.Forms.TextBox txtTel;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtFax;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtEmail;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ComboBox cmbShain;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txtShimebi;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.ComboBox cmbTax;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbCity;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtAddress2S;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.MaskedTextBox mtxtZipCodeS;
        private System.Windows.Forms.ComboBox cmbCityS;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox txtAddress1S;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label26;
        internal System.Windows.Forms.Button btnCsv;
        internal System.Windows.Forms.Button btnDel;
        internal System.Windows.Forms.Button btnClr;
        internal System.Windows.Forms.Button btnRtn;
        internal System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.TextBox txtNameSeikyu;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.TextBox txtsTel;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.TextBox txtsZip;
        private System.Windows.Forms.TextBox txtTantouS;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.TextBox txtBushoS;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.Label label32;
        private System.Windows.Forms.ComboBox cmbKeishoS;
    }
}
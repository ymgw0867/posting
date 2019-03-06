namespace posting
{
    partial class frmSystem
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSystem));
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.txtMemo1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtBankName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtShitenName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtNumber = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtDaihyo = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtYaku = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtTel = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtFax = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtAddress1 = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtAddress2 = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.mtxtZipCode = new System.Windows.Forms.MaskedTextBox();
            this.txtBusho = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.txtTantou = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.txtIraiCode = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.txtIraiName = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.txtBankCode = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.txtShitenCode = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.txtMemo2 = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.txtFlg = new System.Windows.Forms.TextBox();
            this.label22 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txtZipPath = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // btnRtn
            // 
            this.btnRtn.BackColor = System.Drawing.SystemColors.Control;
            this.btnRtn.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(591, 498);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(89, 31);
            this.btnRtn.TabIndex = 23;
            this.btnRtn.Text = "戻る(&R)";
            this.btnRtn.UseCompatibleTextRendering = true;
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.SystemColors.Control;
            this.btnUpdate.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnUpdate.Location = new System.Drawing.Point(496, 498);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(89, 31);
            this.btnUpdate.TabIndex = 22;
            this.btnUpdate.Text = "更新(&U)";
            this.btnUpdate.UseCompatibleTextRendering = true;
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // txtMemo1
            // 
            this.txtMemo1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo1.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMemo1.Location = new System.Drawing.Point(101, 376);
            this.txtMemo1.MaxLength = 50;
            this.txtMemo1.Name = "txtMemo1";
            this.txtMemo1.Size = new System.Drawing.Size(577, 23);
            this.txtMemo1.TabIndex = 19;
            this.txtMemo1.Enter += new System.EventHandler(this.txtEnter);
            this.txtMemo1.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label4.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label4.Location = new System.Drawing.Point(10, 376);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 22);
            this.label4.TabIndex = 18;
            this.label4.Text = "特記事項１";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtBankName
            // 
            this.txtBankName.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtBankName.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtBankName.Location = new System.Drawing.Point(281, 275);
            this.txtBankName.MaxLength = 15;
            this.txtBankName.Name = "txtBankName";
            this.txtBankName.Size = new System.Drawing.Size(397, 23);
            this.txtBankName.TabIndex = 14;
            this.txtBankName.Enter += new System.EventHandler(this.txtEnter);
            this.txtBankName.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label1.Location = new System.Drawing.Point(212, 275);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 22);
            this.label1.TabIndex = 20;
            this.label1.Text = "銀行名";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtShitenName
            // 
            this.txtShitenName.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtShitenName.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtShitenName.Location = new System.Drawing.Point(281, 300);
            this.txtShitenName.MaxLength = 15;
            this.txtShitenName.Name = "txtShitenName";
            this.txtShitenName.Size = new System.Drawing.Size(397, 23);
            this.txtShitenName.TabIndex = 16;
            this.txtShitenName.Enter += new System.EventHandler(this.txtEnter);
            this.txtShitenName.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label3.Location = new System.Drawing.Point(212, 300);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 22);
            this.label3.TabIndex = 22;
            this.label3.Text = "支店名";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label2.Location = new System.Drawing.Point(10, 325);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 23);
            this.label2.TabIndex = 24;
            this.label2.Text = "口座種別";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtNumber
            // 
            this.txtNumber.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtNumber.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtNumber.Location = new System.Drawing.Point(101, 351);
            this.txtNumber.MaxLength = 7;
            this.txtNumber.Name = "txtNumber";
            this.txtNumber.Size = new System.Drawing.Size(80, 23);
            this.txtNumber.TabIndex = 18;
            this.txtNumber.Enter += new System.EventHandler(this.txtEnter);
            this.txtNumber.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label5.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label5.Location = new System.Drawing.Point(10, 351);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 22);
            this.label5.TabIndex = 26;
            this.label5.Text = "口座番号";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(101, 325);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(80, 23);
            this.comboBox1.TabIndex = 17;
            this.comboBox1.Enter += new System.EventHandler(this.txtEnter);
            this.comboBox1.Leave += new System.EventHandler(this.txtLeave);
            // 
            // txtName
            // 
            this.txtName.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName.Location = new System.Drawing.Point(101, 25);
            this.txtName.MaxLength = 50;
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(577, 23);
            this.txtName.TabIndex = 0;
            this.txtName.Enter += new System.EventHandler(this.txtEnter);
            this.txtName.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label6.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label6.Location = new System.Drawing.Point(10, 25);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(92, 22);
            this.label6.TabIndex = 28;
            this.label6.Text = "会社名";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtDaihyo
            // 
            this.txtDaihyo.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtDaihyo.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtDaihyo.Location = new System.Drawing.Point(101, 50);
            this.txtDaihyo.MaxLength = 50;
            this.txtDaihyo.Name = "txtDaihyo";
            this.txtDaihyo.Size = new System.Drawing.Size(268, 23);
            this.txtDaihyo.TabIndex = 1;
            this.txtDaihyo.Enter += new System.EventHandler(this.txtEnter);
            this.txtDaihyo.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label7.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label7.Location = new System.Drawing.Point(10, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(92, 22);
            this.label7.TabIndex = 30;
            this.label7.Text = "代表者名";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtYaku
            // 
            this.txtYaku.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtYaku.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtYaku.Location = new System.Drawing.Point(443, 50);
            this.txtYaku.MaxLength = 50;
            this.txtYaku.Name = "txtYaku";
            this.txtYaku.Size = new System.Drawing.Size(235, 23);
            this.txtYaku.TabIndex = 2;
            this.txtYaku.Enter += new System.EventHandler(this.txtEnter);
            this.txtYaku.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label8.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label8.Location = new System.Drawing.Point(374, 50);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(70, 22);
            this.label8.TabIndex = 32;
            this.label8.Text = "役職名";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtTel
            // 
            this.txtTel.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTel.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTel.Location = new System.Drawing.Point(101, 75);
            this.txtTel.MaxLength = 20;
            this.txtTel.Name = "txtTel";
            this.txtTel.Size = new System.Drawing.Size(268, 23);
            this.txtTel.TabIndex = 3;
            this.txtTel.Enter += new System.EventHandler(this.txtEnter);
            this.txtTel.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label9.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label9.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label9.Location = new System.Drawing.Point(10, 75);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(92, 22);
            this.label9.TabIndex = 34;
            this.label9.Text = "電話番号";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFax
            // 
            this.txtFax.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFax.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtFax.Location = new System.Drawing.Point(443, 75);
            this.txtFax.MaxLength = 20;
            this.txtFax.Name = "txtFax";
            this.txtFax.Size = new System.Drawing.Size(235, 23);
            this.txtFax.TabIndex = 4;
            this.txtFax.Enter += new System.EventHandler(this.txtEnter);
            this.txtFax.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label10.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label10.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label10.Location = new System.Drawing.Point(374, 75);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(70, 22);
            this.label10.TabIndex = 36;
            this.label10.Text = "FAX番号";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtAddress1
            // 
            this.txtAddress1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress1.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress1.Location = new System.Drawing.Point(101, 125);
            this.txtAddress1.MaxLength = 50;
            this.txtAddress1.Name = "txtAddress1";
            this.txtAddress1.Size = new System.Drawing.Size(577, 23);
            this.txtAddress1.TabIndex = 6;
            this.txtAddress1.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress1.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label11
            // 
            this.label11.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label11.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label11.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label11.Location = new System.Drawing.Point(10, 125);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(92, 22);
            this.label11.TabIndex = 38;
            this.label11.Text = "住所１";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtAddress2
            // 
            this.txtAddress2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddress2.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAddress2.Location = new System.Drawing.Point(101, 150);
            this.txtAddress2.MaxLength = 50;
            this.txtAddress2.Name = "txtAddress2";
            this.txtAddress2.Size = new System.Drawing.Size(577, 23);
            this.txtAddress2.TabIndex = 7;
            this.txtAddress2.Enter += new System.EventHandler(this.txtEnter);
            this.txtAddress2.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label12
            // 
            this.label12.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label12.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label12.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label12.Location = new System.Drawing.Point(10, 150);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(92, 22);
            this.label12.TabIndex = 40;
            this.label12.Text = "住所２";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label13
            // 
            this.label13.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label13.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label13.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label13.Location = new System.Drawing.Point(10, 100);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(92, 22);
            this.label13.TabIndex = 42;
            this.label13.Text = "郵便番号";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtEmail
            // 
            this.txtEmail.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtEmail.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtEmail.Location = new System.Drawing.Point(101, 175);
            this.txtEmail.MaxLength = 50;
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(577, 23);
            this.txtEmail.TabIndex = 8;
            this.txtEmail.Enter += new System.EventHandler(this.txtEnter);
            this.txtEmail.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label14
            // 
            this.label14.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label14.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label14.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label14.Location = new System.Drawing.Point(10, 175);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(92, 22);
            this.label14.TabIndex = 44;
            this.label14.Text = "eMail";
            this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // mtxtZipCode
            // 
            this.mtxtZipCode.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.mtxtZipCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.mtxtZipCode.Location = new System.Drawing.Point(101, 100);
            this.mtxtZipCode.Mask = "000-0000";
            this.mtxtZipCode.Name = "mtxtZipCode";
            this.mtxtZipCode.Size = new System.Drawing.Size(81, 23);
            this.mtxtZipCode.TabIndex = 5;
            this.mtxtZipCode.Enter += new System.EventHandler(this.txtEnter);
            this.mtxtZipCode.Leave += new System.EventHandler(this.txtLeave);
            // 
            // txtBusho
            // 
            this.txtBusho.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtBusho.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtBusho.Location = new System.Drawing.Point(101, 200);
            this.txtBusho.MaxLength = 50;
            this.txtBusho.Name = "txtBusho";
            this.txtBusho.Size = new System.Drawing.Size(577, 23);
            this.txtBusho.TabIndex = 9;
            this.txtBusho.Enter += new System.EventHandler(this.txtEnter);
            this.txtBusho.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label15
            // 
            this.label15.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label15.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label15.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label15.Location = new System.Drawing.Point(10, 200);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(92, 22);
            this.label15.TabIndex = 47;
            this.label15.Text = "部署名";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtTantou
            // 
            this.txtTantou.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTantou.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtTantou.Location = new System.Drawing.Point(101, 225);
            this.txtTantou.MaxLength = 50;
            this.txtTantou.Name = "txtTantou";
            this.txtTantou.Size = new System.Drawing.Size(577, 23);
            this.txtTantou.TabIndex = 10;
            this.txtTantou.Enter += new System.EventHandler(this.txtEnter);
            this.txtTantou.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label16
            // 
            this.label16.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label16.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label16.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label16.Location = new System.Drawing.Point(10, 225);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(92, 22);
            this.label16.TabIndex = 49;
            this.label16.Text = "担当者名";
            this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtIraiCode
            // 
            this.txtIraiCode.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtIraiCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtIraiCode.Location = new System.Drawing.Point(101, 250);
            this.txtIraiCode.MaxLength = 10;
            this.txtIraiCode.Name = "txtIraiCode";
            this.txtIraiCode.Size = new System.Drawing.Size(109, 23);
            this.txtIraiCode.TabIndex = 11;
            this.txtIraiCode.Enter += new System.EventHandler(this.txtEnter);
            this.txtIraiCode.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label17
            // 
            this.label17.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label17.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label17.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label17.Location = new System.Drawing.Point(10, 250);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(92, 22);
            this.label17.TabIndex = 51;
            this.label17.Text = "依頼人コード";
            this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtIraiName
            // 
            this.txtIraiName.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtIraiName.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtIraiName.Location = new System.Drawing.Point(281, 250);
            this.txtIraiName.MaxLength = 40;
            this.txtIraiName.Name = "txtIraiName";
            this.txtIraiName.Size = new System.Drawing.Size(397, 23);
            this.txtIraiName.TabIndex = 12;
            this.txtIraiName.Enter += new System.EventHandler(this.txtEnter);
            this.txtIraiName.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label18
            // 
            this.label18.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label18.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label18.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label18.Location = new System.Drawing.Point(212, 250);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(70, 22);
            this.label18.TabIndex = 53;
            this.label18.Text = "依頼人名";
            this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtBankCode
            // 
            this.txtBankCode.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtBankCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtBankCode.Location = new System.Drawing.Point(101, 275);
            this.txtBankCode.MaxLength = 4;
            this.txtBankCode.Name = "txtBankCode";
            this.txtBankCode.Size = new System.Drawing.Size(80, 23);
            this.txtBankCode.TabIndex = 13;
            this.txtBankCode.Enter += new System.EventHandler(this.txtEnter);
            this.txtBankCode.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label19
            // 
            this.label19.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label19.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label19.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label19.Location = new System.Drawing.Point(10, 275);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(92, 22);
            this.label19.TabIndex = 55;
            this.label19.Text = "銀行コード";
            this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtShitenCode
            // 
            this.txtShitenCode.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtShitenCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtShitenCode.Location = new System.Drawing.Point(101, 300);
            this.txtShitenCode.MaxLength = 3;
            this.txtShitenCode.Name = "txtShitenCode";
            this.txtShitenCode.Size = new System.Drawing.Size(80, 23);
            this.txtShitenCode.TabIndex = 15;
            this.txtShitenCode.Enter += new System.EventHandler(this.txtEnter);
            this.txtShitenCode.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label20
            // 
            this.label20.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label20.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label20.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label20.Location = new System.Drawing.Point(10, 300);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(92, 22);
            this.label20.TabIndex = 57;
            this.label20.Text = "支店コード";
            this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtMemo2
            // 
            this.txtMemo2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo2.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMemo2.Location = new System.Drawing.Point(101, 401);
            this.txtMemo2.MaxLength = 50;
            this.txtMemo2.Name = "txtMemo2";
            this.txtMemo2.Size = new System.Drawing.Size(577, 23);
            this.txtMemo2.TabIndex = 20;
            this.txtMemo2.Enter += new System.EventHandler(this.txtEnter);
            this.txtMemo2.Leave += new System.EventHandler(this.txtLeave);
            // 
            // label21
            // 
            this.label21.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label21.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label21.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label21.Location = new System.Drawing.Point(10, 401);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(92, 22);
            this.label21.TabIndex = 59;
            this.label21.Text = "特記事項２";
            this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFlg
            // 
            this.txtFlg.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFlg.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtFlg.Location = new System.Drawing.Point(101, 449);
            this.txtFlg.MaxLength = 1;
            this.txtFlg.Name = "txtFlg";
            this.txtFlg.Size = new System.Drawing.Size(80, 23);
            this.txtFlg.TabIndex = 22;
            this.txtFlg.Enter += new System.EventHandler(this.txtEnter);
            this.txtFlg.Leave += new System.EventHandler(this.txtLeave);
            this.txtFlg.Validating += new System.ComponentModel.CancelEventHandler(this.txtFlg_Validating);
            // 
            // label22
            // 
            this.label22.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label22.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label22.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label22.Location = new System.Drawing.Point(10, 450);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(92, 22);
            this.label22.TabIndex = 61;
            this.label22.Text = "配布フラグ";
            this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(12, 498);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 26);
            this.button1.TabIndex = 62;
            this.button1.Text = "配布フラグ(&F)";
            this.button1.UseCompatibleTextRendering = true;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtZipPath
            // 
            this.txtZipPath.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtZipPath.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtZipPath.Location = new System.Drawing.Point(101, 425);
            this.txtZipPath.MaxLength = 100;
            this.txtZipPath.Name = "txtZipPath";
            this.txtZipPath.Size = new System.Drawing.Size(528, 23);
            this.txtZipPath.TabIndex = 21;
            // 
            // label23
            // 
            this.label23.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.label23.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label23.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label23.Location = new System.Drawing.Point(10, 425);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(92, 22);
            this.label23.TabIndex = 64;
            this.label23.Text = "郵便番号CSV";
            this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.Control;
            this.button2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.Location = new System.Drawing.Point(630, 425);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(50, 26);
            this.button2.TabIndex = 22;
            this.button2.Text = "参照";
            this.button2.UseCompatibleTextRendering = true;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // frmSystem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(687, 541);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.txtZipPath);
            this.Controls.Add(this.label23);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtFlg);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.txtMemo2);
            this.Controls.Add(this.label21);
            this.Controls.Add(this.txtShitenCode);
            this.Controls.Add(this.label20);
            this.Controls.Add(this.txtBankCode);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.txtIraiName);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.txtIraiCode);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.txtTantou);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.txtBusho);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.mtxtZipCode);
            this.Controls.Add(this.txtEmail);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.txtAddress2);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.txtAddress1);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.txtFax);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.txtTel);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.txtYaku);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtDaihyo);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.txtNumber);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtShitenName);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtBankName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtMemo1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnUpdate);
            this.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmSystem";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "会社情報保守";
            this.Load += new System.EventHandler(this.form_Load);
            this.Enter += new System.EventHandler(this.txtEnter);
            this.Leave += new System.EventHandler(this.txtLeave);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Button btnRtn;
        internal System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.TextBox txtMemo1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtBankName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtShitenName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtNumber;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtDaihyo;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtYaku;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtTel;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtFax;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtAddress1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtAddress2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtEmail;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.MaskedTextBox mtxtZipCode;
        private System.Windows.Forms.TextBox txtBusho;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtTantou;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox txtIraiCode;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtIraiName;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox txtBankCode;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox txtShitenCode;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox txtMemo2;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox txtFlg;
        private System.Windows.Forms.Label label22;
        internal System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtZipPath;
        private System.Windows.Forms.Label label23;
        internal System.Windows.Forms.Button button2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
namespace posting
{
    partial class frmOrderExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmOrderExcel));
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cmbNewRep = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtEiTantou = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.dtOrder = new System.Windows.Forms.DateTimePicker();
            this.txtsAdd = new System.Windows.Forms.TextBox();
            this.txtsTantou = new System.Windows.Forms.TextBox();
            this.txtsBusho = new System.Windows.Forms.TextBox();
            this.txtsName = new System.Windows.Forms.TextBox();
            this.txtsZipCode = new System.Windows.Forms.TextBox();
            this.txtTantou = new System.Windows.Forms.TextBox();
            this.txtBusho = new System.Windows.Forms.TextBox();
            this.txtFax = new System.Windows.Forms.TextBox();
            this.txtTel = new System.Windows.Forms.TextBox();
            this.txtAdd = new System.Windows.Forms.TextBox();
            this.txtZipCode = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtClient = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.label29 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dtSeikyuShime = new System.Windows.Forms.DateTimePicker();
            this.label22 = new System.Windows.Forms.Label();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.cmbJizenseikyusho = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtHaifuhoukoku = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.dtNouhin = new System.Windows.Forms.DateTimePicker();
            this.dtNyukin = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbHaifukeitai = new System.Windows.Forms.ComboBox();
            this.cmbHaifuJyoken = new System.Windows.Forms.ComboBox();
            this.cmbNouhinBasho = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtArari = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtZeinuki = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.btnRep = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.txtArari2 = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(60, 33);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 15);
            this.label3.TabIndex = 2;
            this.label3.Text = "クライアント：";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(75, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "受注日：";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(274, 11);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 15);
            this.label5.TabIndex = 4;
            this.label5.Text = "新規／リピート：";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbNewRep
            // 
            this.cmbNewRep.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbNewRep.FormattingEnabled = true;
            this.cmbNewRep.Items.AddRange(new object[] {
            "新規 (OUT)",
            "新規 (紹介)",
            "新規 (IN)",
            "リピート"});
            this.cmbNewRep.Location = new System.Drawing.Point(362, 5);
            this.cmbNewRep.Name = "cmbNewRep";
            this.cmbNewRep.Size = new System.Drawing.Size(154, 26);
            this.cmbNewRep.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(540, 33);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(55, 15);
            this.label6.TabIndex = 6;
            this.label6.Text = "商品名：";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(99, 56);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(31, 15);
            this.label8.TabIndex = 8;
            this.label8.Text = "〒：";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(686, 60);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(43, 15);
            this.label11.TabIndex = 11;
            this.label11.Text = "電話：";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(874, 60);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(55, 15);
            this.label14.TabIndex = 13;
            this.label14.Text = "ＦＡＸ：";
            this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(88, 84);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(43, 15);
            this.label16.TabIndex = 15;
            this.label16.Text = "部署：";
            this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(320, 84);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(43, 15);
            this.label18.TabIndex = 17;
            this.label18.Text = "担当：";
            this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label19
            // 
            this.label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label19.Font = new System.Drawing.Font("Meiryo UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label19.Location = new System.Drawing.Point(10, 102);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(65, 49);
            this.label19.TabIndex = 19;
            this.label19.Text = "請求先";
            this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(485, 132);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(31, 15);
            this.label20.TabIndex = 20;
            this.label20.Text = "〒：";
            this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(320, 132);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(43, 15);
            this.label24.TabIndex = 25;
            this.label24.Text = "担当：";
            this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(87, 132);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(43, 15);
            this.label26.TabIndex = 23;
            this.label26.Text = "部署：";
            this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(1086, 161);
            this.dataGridView1.TabIndex = 31;
            this.dataGridView1.TabStop = false;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(4, 330);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 21;
            this.dataGridView2.Size = new System.Drawing.Size(1076, 160);
            this.dataGridView2.TabIndex = 33;
            this.dataGridView2.TabStop = false;
            this.dataGridView2.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);
            this.dataGridView2.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellEnter);
            this.dataGridView2.CellLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellLeave);
            this.dataGridView2.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellValueChanged);
            this.dataGridView2.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.dataGridView2_RowsRemoved);
            this.dataGridView2.Leave += new System.EventHandler(this.dataGridView2_Leave);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.txtEiTantou);
            this.panel1.Controls.Add(this.label21);
            this.panel1.Controls.Add(this.dtOrder);
            this.panel1.Controls.Add(this.txtsAdd);
            this.panel1.Controls.Add(this.txtsTantou);
            this.panel1.Controls.Add(this.txtsBusho);
            this.panel1.Controls.Add(this.txtsName);
            this.panel1.Controls.Add(this.txtsZipCode);
            this.panel1.Controls.Add(this.txtTantou);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.txtBusho);
            this.panel1.Controls.Add(this.txtFax);
            this.panel1.Controls.Add(this.txtTel);
            this.panel1.Controls.Add(this.txtAdd);
            this.panel1.Controls.Add(this.txtZipCode);
            this.panel1.Controls.Add(this.txtName);
            this.panel1.Controls.Add(this.txtClient);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.label19);
            this.panel1.Controls.Add(this.label20);
            this.panel1.Controls.Add(this.label16);
            this.panel1.Controls.Add(this.label26);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label24);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.cmbNewRep);
            this.panel1.Controls.Add(this.label17);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Location = new System.Drawing.Point(4, 169);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1076, 155);
            this.panel1.TabIndex = 0;
            // 
            // txtEiTantou
            // 
            this.txtEiTantou.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtEiTantou.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtEiTantou.Location = new System.Drawing.Point(595, 6);
            this.txtEiTantou.Name = "txtEiTantou";
            this.txtEiTantou.Size = new System.Drawing.Size(474, 25);
            this.txtEiTantou.TabIndex = 2;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(528, 12);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(67, 15);
            this.label21.TabIndex = 49;
            this.label21.Text = "営業担当：";
            this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtOrder
            // 
            this.dtOrder.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtOrder.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtOrder.Location = new System.Drawing.Point(128, 6);
            this.dtOrder.Name = "dtOrder";
            this.dtOrder.ShowCheckBox = true;
            this.dtOrder.Size = new System.Drawing.Size(143, 25);
            this.dtOrder.TabIndex = 0;
            // 
            // txtsAdd
            // 
            this.txtsAdd.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtsAdd.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtsAdd.Location = new System.Drawing.Point(594, 126);
            this.txtsAdd.Name = "txtsAdd";
            this.txtsAdd.Size = new System.Drawing.Size(475, 25);
            this.txtsAdd.TabIndex = 15;
            // 
            // txtsTantou
            // 
            this.txtsTantou.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtsTantou.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtsTantou.Location = new System.Drawing.Point(362, 126);
            this.txtsTantou.Name = "txtsTantou";
            this.txtsTantou.Size = new System.Drawing.Size(108, 25);
            this.txtsTantou.TabIndex = 13;
            // 
            // txtsBusho
            // 
            this.txtsBusho.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtsBusho.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtsBusho.Location = new System.Drawing.Point(128, 126);
            this.txtsBusho.Name = "txtsBusho";
            this.txtsBusho.Size = new System.Drawing.Size(182, 25);
            this.txtsBusho.TabIndex = 12;
            // 
            // txtsName
            // 
            this.txtsName.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtsName.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtsName.Location = new System.Drawing.Point(128, 102);
            this.txtsName.Name = "txtsName";
            this.txtsName.Size = new System.Drawing.Size(941, 25);
            this.txtsName.TabIndex = 11;
            // 
            // txtsZipCode
            // 
            this.txtsZipCode.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtsZipCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtsZipCode.Location = new System.Drawing.Point(513, 126);
            this.txtsZipCode.Name = "txtsZipCode";
            this.txtsZipCode.Size = new System.Drawing.Size(82, 25);
            this.txtsZipCode.TabIndex = 14;
            this.txtsZipCode.Text = "123-0098";
            this.txtsZipCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtTantou
            // 
            this.txtTantou.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTantou.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtTantou.Location = new System.Drawing.Point(362, 78);
            this.txtTantou.Name = "txtTantou";
            this.txtTantou.Size = new System.Drawing.Size(108, 25);
            this.txtTantou.TabIndex = 10;
            // 
            // txtBusho
            // 
            this.txtBusho.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtBusho.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtBusho.Location = new System.Drawing.Point(128, 78);
            this.txtBusho.Name = "txtBusho";
            this.txtBusho.Size = new System.Drawing.Size(182, 25);
            this.txtBusho.TabIndex = 9;
            // 
            // txtFax
            // 
            this.txtFax.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFax.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtFax.Location = new System.Drawing.Point(927, 54);
            this.txtFax.Name = "txtFax";
            this.txtFax.Size = new System.Drawing.Size(142, 25);
            this.txtFax.TabIndex = 8;
            this.txtFax.Text = "03-8765-9272";
            this.txtFax.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtTel
            // 
            this.txtTel.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTel.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTel.Location = new System.Drawing.Point(726, 54);
            this.txtTel.Name = "txtTel";
            this.txtTel.Size = new System.Drawing.Size(142, 25);
            this.txtTel.TabIndex = 7;
            this.txtTel.Text = "03-8765-9272";
            this.txtTel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtAdd
            // 
            this.txtAdd.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAdd.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAdd.Location = new System.Drawing.Point(213, 54);
            this.txtAdd.Name = "txtAdd";
            this.txtAdd.Size = new System.Drawing.Size(467, 25);
            this.txtAdd.TabIndex = 6;
            // 
            // txtZipCode
            // 
            this.txtZipCode.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtZipCode.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtZipCode.Location = new System.Drawing.Point(128, 54);
            this.txtZipCode.Name = "txtZipCode";
            this.txtZipCode.Size = new System.Drawing.Size(86, 25);
            this.txtZipCode.TabIndex = 5;
            this.txtZipCode.Text = "123-0098";
            this.txtZipCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtName
            // 
            this.txtName.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtName.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtName.Location = new System.Drawing.Point(595, 30);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(474, 25);
            this.txtName.TabIndex = 4;
            // 
            // txtClient
            // 
            this.txtClient.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtClient.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtClient.Location = new System.Drawing.Point(128, 30);
            this.txtClient.Name = "txtClient";
            this.txtClient.Size = new System.Drawing.Size(388, 25);
            this.txtClient.TabIndex = 3;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(76, 108);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(55, 15);
            this.label17.TabIndex = 48;
            this.label17.Text = "御社名：";
            this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(10, 9);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(120, 15);
            this.label29.TabIndex = 35;
            this.label29.Text = "ポスティング配布形態：";
            this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.dtSeikyuShime);
            this.panel2.Controls.Add(this.label22);
            this.panel2.Controls.Add(this.txtMemo);
            this.panel2.Controls.Add(this.label12);
            this.panel2.Controls.Add(this.cmbJizenseikyusho);
            this.panel2.Controls.Add(this.label10);
            this.panel2.Controls.Add(this.txtHaifuhoukoku);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.dtNouhin);
            this.panel2.Controls.Add(this.dtNyukin);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.cmbHaifukeitai);
            this.panel2.Controls.Add(this.label29);
            this.panel2.Controls.Add(this.cmbHaifuJyoken);
            this.panel2.Controls.Add(this.cmbNouhinBasho);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Location = new System.Drawing.Point(4, 522);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1076, 109);
            this.panel2.TabIndex = 1;
            // 
            // dtSeikyuShime
            // 
            this.dtSeikyuShime.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtSeikyuShime.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtSeikyuShime.Location = new System.Drawing.Point(128, 53);
            this.dtSeikyuShime.Name = "dtSeikyuShime";
            this.dtSeikyuShime.ShowCheckBox = true;
            this.dtSeikyuShime.Size = new System.Drawing.Size(182, 25);
            this.dtSeikyuShime.TabIndex = 4;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(63, 57);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(67, 15);
            this.label22.TabIndex = 53;
            this.label22.Text = "請求締日：";
            this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtMemo
            // 
            this.txtMemo.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMemo.Location = new System.Drawing.Point(638, 6);
            this.txtMemo.Multiline = true;
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(431, 93);
            this.txtMemo.TabIndex = 7;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(574, 9);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(67, 15);
            this.label12.TabIndex = 51;
            this.label12.Text = "特記事項：";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbJizenseikyusho
            // 
            this.cmbJizenseikyusho.FormattingEnabled = true;
            this.cmbJizenseikyusho.Location = new System.Drawing.Point(595, 76);
            this.cmbJizenseikyusho.Name = "cmbJizenseikyusho";
            this.cmbJizenseikyusho.Size = new System.Drawing.Size(37, 23);
            this.cmbJizenseikyusho.TabIndex = 6;
            this.cmbJizenseikyusho.TabStop = false;
            this.cmbJizenseikyusho.Visible = false;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(516, 80);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(79, 15);
            this.label10.TabIndex = 49;
            this.label10.Text = "事前請求書：";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.label10.Visible = false;
            // 
            // txtHaifuhoukoku
            // 
            this.txtHaifuhoukoku.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtHaifuhoukoku.Location = new System.Drawing.Point(489, 77);
            this.txtHaifuhoukoku.Name = "txtHaifuhoukoku";
            this.txtHaifuhoukoku.Size = new System.Drawing.Size(21, 23);
            this.txtHaifuhoukoku.TabIndex = 5;
            this.txtHaifuhoukoku.TabStop = false;
            this.txtHaifuhoukoku.Visible = false;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(416, 80);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(67, 15);
            this.label9.TabIndex = 47;
            this.label9.Text = "配布報告：";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.label9.Visible = false;
            // 
            // dtNouhin
            // 
            this.dtNouhin.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtNouhin.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtNouhin.Location = new System.Drawing.Point(128, 30);
            this.dtNouhin.Name = "dtNouhin";
            this.dtNouhin.ShowCheckBox = true;
            this.dtNouhin.Size = new System.Drawing.Size(182, 25);
            this.dtNouhin.TabIndex = 2;
            // 
            // dtNyukin
            // 
            this.dtNyukin.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtNyukin.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtNyukin.Location = new System.Drawing.Point(128, 77);
            this.dtNyukin.Name = "dtNyukin";
            this.dtNyukin.ShowCheckBox = true;
            this.dtNyukin.Size = new System.Drawing.Size(182, 25);
            this.dtNyukin.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(63, 81);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(67, 15);
            this.label7.TabIndex = 44;
            this.label7.Text = "支払期日：";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(75, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 15);
            this.label2.TabIndex = 40;
            this.label2.Text = "納品日：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbHaifukeitai
            // 
            this.cmbHaifukeitai.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbHaifukeitai.FormattingEnabled = true;
            this.cmbHaifukeitai.Location = new System.Drawing.Point(128, 5);
            this.cmbHaifukeitai.Name = "cmbHaifukeitai";
            this.cmbHaifukeitai.Size = new System.Drawing.Size(182, 26);
            this.cmbHaifukeitai.TabIndex = 0;
            // 
            // cmbHaifuJyoken
            // 
            this.cmbHaifuJyoken.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbHaifuJyoken.FormattingEnabled = true;
            this.cmbHaifuJyoken.Location = new System.Drawing.Point(309, 5);
            this.cmbHaifuJyoken.Name = "cmbHaifuJyoken";
            this.cmbHaifuJyoken.Size = new System.Drawing.Size(223, 26);
            this.cmbHaifuJyoken.TabIndex = 1;
            // 
            // cmbNouhinBasho
            // 
            this.cmbNouhinBasho.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbNouhinBasho.FormattingEnabled = true;
            this.cmbNouhinBasho.Location = new System.Drawing.Point(381, 29);
            this.cmbNouhinBasho.Name = "cmbNouhinBasho";
            this.cmbNouhinBasho.Size = new System.Drawing.Size(151, 26);
            this.cmbNouhinBasho.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(316, 35);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 15);
            this.label4.TabIndex = 42;
            this.label4.Text = "納品場所：";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtArari
            // 
            this.txtArari.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtArari.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtArari.Location = new System.Drawing.Point(743, 495);
            this.txtArari.Name = "txtArari";
            this.txtArari.ReadOnly = true;
            this.txtArari.Size = new System.Drawing.Size(119, 25);
            this.txtArari.TabIndex = 39;
            this.txtArari.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label13.Location = new System.Drawing.Point(654, 498);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(92, 18);
            this.label13.TabIndex = 38;
            this.label13.Text = "合計粗利①：";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtZeinuki
            // 
            this.txtZeinuki.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtZeinuki.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtZeinuki.Location = new System.Drawing.Point(525, 495);
            this.txtZeinuki.Name = "txtZeinuki";
            this.txtZeinuki.ReadOnly = true;
            this.txtZeinuki.Size = new System.Drawing.Size(119, 25);
            this.txtZeinuki.TabIndex = 41;
            this.txtZeinuki.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label15.Location = new System.Drawing.Point(423, 498);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(106, 18);
            this.label15.TabIndex = 40;
            this.label15.Text = "合計（税抜）：";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnRep
            // 
            this.btnRep.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRep.Location = new System.Drawing.Point(742, 637);
            this.btnRep.Name = "btnRep";
            this.btnRep.Size = new System.Drawing.Size(168, 29);
            this.btnRep.TabIndex = 2;
            this.btnRep.Text = "受注確定書発行(&F)";
            this.btnRep.UseVisualStyleBackColor = true;
            this.btnRep.Click += new System.EventHandler(this.btnRep_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.Location = new System.Drawing.Point(912, 637);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(168, 29);
            this.button2.TabIndex = 3;
            this.button2.Text = "終了(&E)";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // txtArari2
            // 
            this.txtArari2.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtArari2.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtArari2.Location = new System.Drawing.Point(956, 495);
            this.txtArari2.Name = "txtArari2";
            this.txtArari2.ReadOnly = true;
            this.txtArari2.Size = new System.Drawing.Size(119, 25);
            this.txtArari2.TabIndex = 43;
            this.txtArari2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Font = new System.Drawing.Font("Meiryo UI", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label23.Location = new System.Drawing.Point(868, 498);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(92, 18);
            this.label23.TabIndex = 42;
            this.label23.Text = "合計粗利②：";
            this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // frmOrderExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1086, 672);
            this.Controls.Add(this.txtArari2);
            this.Controls.Add(this.label23);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnRep);
            this.Controls.Add(this.txtZeinuki);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.txtArari);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "frmOrderExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "エクセル受注確定書作成";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmOrderExcel_FormClosing);
            this.Load += new System.EventHandler(this.frmOrderExcel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cmbNewRep;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.ComboBox cmbHaifuJyoken;
        private System.Windows.Forms.ComboBox cmbHaifukeitai;
        private System.Windows.Forms.TextBox txtsAdd;
        private System.Windows.Forms.TextBox txtsTantou;
        private System.Windows.Forms.TextBox txtsBusho;
        private System.Windows.Forms.TextBox txtsName;
        private System.Windows.Forms.TextBox txtsZipCode;
        private System.Windows.Forms.TextBox txtTantou;
        private System.Windows.Forms.TextBox txtBusho;
        private System.Windows.Forms.TextBox txtFax;
        private System.Windows.Forms.TextBox txtTel;
        private System.Windows.Forms.TextBox txtAdd;
        private System.Windows.Forms.TextBox txtZipCode;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtClient;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtOrder;
        private System.Windows.Forms.TextBox txtHaifuhoukoku;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.DateTimePicker dtNouhin;
        private System.Windows.Forms.DateTimePicker dtNyukin;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbNouhinBasho;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox cmbJizenseikyusho;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtArari;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtZeinuki;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button btnRep;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtEiTantou;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.DateTimePicker dtSeikyuShime;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox txtArari2;
        private System.Windows.Forms.Label label23;
    }
}
namespace _1113354_陳冠瑋_房貸計算器
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.pnlMain = new System.Windows.Forms.TableLayoutPanel();
            this.gbInput = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanelInput = new System.Windows.Forms.TableLayoutPanel();
            this.lblPrice = new System.Windows.Forms.Label();
            this.txtPrice = new System.Windows.Forms.TextBox();
            this.lblDownPayment = new System.Windows.Forms.Label();
            this.pnlDown = new System.Windows.Forms.FlowLayoutPanel();
            this.txtDownPayment = new System.Windows.Forms.TextBox();
            this.cmbDownPaymentType = new System.Windows.Forms.ComboBox();
            this.lblRate = new System.Windows.Forms.Label();
            this.txtRate = new System.Windows.Forms.TextBox();
            this.lblTerm = new System.Windows.Forms.Label();
            this.cmbTerm = new System.Windows.Forms.ComboBox();
            this.lblGrace = new System.Windows.Forms.Label();
            this.txtGrace = new System.Windows.Forms.TextBox();
            this.pnlButtons = new System.Windows.Forms.FlowLayoutPanel();
            this.btnCalc = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.lblValidationHint = new System.Windows.Forms.Label();
            this.tabControlHelper = new System.Windows.Forms.TabControl();
            this.tabSummary = new System.Windows.Forms.TabPage();
            this.tableLayoutPanelOutput = new System.Windows.Forms.TableLayoutPanel();
            this.lblTitleLoan = new System.Windows.Forms.Label();
            this.lblResultTotalLoan = new System.Windows.Forms.Label();
            this.lblTitleMonthly = new System.Windows.Forms.Label();
            this.lblResultMonthly = new System.Windows.Forms.Label();
            this.lblTitleFirstInt = new System.Windows.Forms.Label();
            this.lblResultFirstInt = new System.Windows.Forms.Label();
            this.lblTitleFirstPrin = new System.Windows.Forms.Label();
            this.lblResultFirstPrin = new System.Windows.Forms.Label();
            this.lblTitleTotalInt = new System.Windows.Forms.Label();
            this.lblResultTotalInt = new System.Windows.Forms.Label();
            this.lblTitleTotalRepay = new System.Windows.Forms.Label();
            this.lblResultTotalRepay = new System.Windows.Forms.Label();
            this.picChart = new System.Windows.Forms.PictureBox();
            this.tabSchedule = new System.Windows.Forms.TabPage();
            this.dgvSchedule = new System.Windows.Forms.DataGridView();
            this.tabAI = new System.Windows.Forms.TabPage();
            this.rtbAI = new System.Windows.Forms.RichTextBox();
            this.lblMainTitle = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Label();
            this.btnMinimize = new System.Windows.Forms.Label();
            
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.pnlMain.SuspendLayout();
            this.gbInput.SuspendLayout();
            this.tableLayoutPanelInput.SuspendLayout();
            this.pnlDown.SuspendLayout();
            this.pnlButtons.SuspendLayout();
            this.tabControlHelper.SuspendLayout();
            this.tabSummary.SuspendLayout();
            this.tableLayoutPanelOutput.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picChart)).BeginInit();
            this.tabSchedule.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSchedule)).BeginInit();
            this.tabAI.SuspendLayout();
            this.SuspendLayout();
            
            // errorProvider1
            this.errorProvider1.ContainerControl = this;
            
            // lblMainTitle
            this.lblMainTitle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(62)))), ((int)(((byte)(80)))));
            this.lblMainTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblMainTitle.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblMainTitle.ForeColor = System.Drawing.Color.White;
            this.lblMainTitle.Location = new System.Drawing.Point(0, 0);
            this.lblMainTitle.Name = "lblMainTitle";
            this.lblMainTitle.Size = new System.Drawing.Size(984, 60);
            this.lblMainTitle.TabIndex = 0;
            this.lblMainTitle.Text = "  🏡 企業級個人房貸試算中心";
            this.lblMainTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            
            // btnClose / btnMinimize Control Box
            this.btnClose.BackColor = System.Drawing.Color.Transparent;
            this.btnClose.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnClose.Font = new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold);
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Location = new System.Drawing.Point(940, 15);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(30, 30);
            this.btnClose.Text = "✕";
            this.btnClose.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            this.btnMinimize.BackColor = System.Drawing.Color.Transparent;
            this.btnMinimize.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnMinimize.Font = new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold);
            this.btnMinimize.ForeColor = System.Drawing.Color.White;
            this.btnMinimize.Location = new System.Drawing.Point(910, 15);
            this.btnMinimize.Name = "btnMinimize";
            this.btnMinimize.Size = new System.Drawing.Size(30, 30);
            this.btnMinimize.Text = "－";
            this.btnMinimize.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // pnlMain
            this.pnlMain.ColumnCount = 2;
            this.pnlMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40F));
            this.pnlMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60F));
            this.pnlMain.Controls.Add(this.gbInput, 0, 0);
            this.pnlMain.Controls.Add(this.tabControlHelper, 1, 0);
            this.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlMain.Location = new System.Drawing.Point(0, 60);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Padding = new System.Windows.Forms.Padding(20, 20, 20, 20);
            this.pnlMain.RowCount = 1;
            this.pnlMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.pnlMain.Size = new System.Drawing.Size(984, 600);
            this.pnlMain.TabIndex = 1;
            
            // gbInput
            this.gbInput.Controls.Add(this.tableLayoutPanelInput);
            this.gbInput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gbInput.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.gbInput.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(73)))), ((int)(((byte)(94)))));
            this.gbInput.Location = new System.Drawing.Point(23, 23);
            this.gbInput.Name = "gbInput";
            this.gbInput.Padding = new System.Windows.Forms.Padding(15);
            this.gbInput.Size = new System.Drawing.Size(371, 554);
            this.gbInput.TabIndex = 0;
            this.gbInput.TabStop = false;
            this.gbInput.Text = "⚙️ 參數設定區域";
            
            // tableLayoutPanelInput
            this.tableLayoutPanelInput.ColumnCount = 2;
            this.tableLayoutPanelInput.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.tableLayoutPanelInput.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.tableLayoutPanelInput.Controls.Add(this.lblPrice, 0, 0);
            this.tableLayoutPanelInput.Controls.Add(this.txtPrice, 1, 0);
            this.tableLayoutPanelInput.Controls.Add(this.lblDownPayment, 0, 1);
            this.tableLayoutPanelInput.Controls.Add(this.pnlDown, 1, 1);
            this.tableLayoutPanelInput.Controls.Add(this.lblRate, 0, 2);
            this.tableLayoutPanelInput.Controls.Add(this.txtRate, 1, 2);
            this.tableLayoutPanelInput.Controls.Add(this.lblTerm, 0, 3);
            this.tableLayoutPanelInput.Controls.Add(this.cmbTerm, 1, 3);
            this.tableLayoutPanelInput.Controls.Add(this.lblGrace, 0, 4);
            this.tableLayoutPanelInput.Controls.Add(this.txtGrace, 1, 4);
            this.tableLayoutPanelInput.Controls.Add(this.pnlButtons, 1, 5);
            this.tableLayoutPanelInput.Controls.Add(this.lblValidationHint, 0, 6);
            this.tableLayoutPanelInput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanelInput.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tableLayoutPanelInput.Location = new System.Drawing.Point(15, 37);
            this.tableLayoutPanelInput.Name = "tableLayoutPanelInput";
            this.tableLayoutPanelInput.RowCount = 7;
            this.tableLayoutPanelInput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanelInput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanelInput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanelInput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanelInput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanelInput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanelInput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanelInput.Size = new System.Drawing.Size(341, 502);
            this.tableLayoutPanelInput.TabIndex = 0;
            
            // lblPrice
            this.lblPrice.AutoSize = true;
            this.lblPrice.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblPrice.Location = new System.Drawing.Point(3, 0);
            this.lblPrice.Name = "lblPrice";
            this.lblPrice.Size = new System.Drawing.Size(113, 50);
            this.lblPrice.Text = "房屋總價($)";
            this.lblPrice.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            
            // txtPrice
            this.txtPrice.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPrice.Location = new System.Drawing.Point(122, 11);
            this.txtPrice.Name = "txtPrice";
            this.txtPrice.Size = new System.Drawing.Size(216, 27);
            
            // lblDownPayment
            this.lblDownPayment.AutoSize = true;
            this.lblDownPayment.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblDownPayment.Location = new System.Drawing.Point(3, 50);
            this.lblDownPayment.Name = "lblDownPayment";
            this.lblDownPayment.Size = new System.Drawing.Size(113, 50);
            this.lblDownPayment.Text = "初期自備款";
            this.lblDownPayment.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            
            // pnlDown
            this.pnlDown.Controls.Add(this.txtDownPayment);
            this.pnlDown.Controls.Add(this.cmbDownPaymentType);
            this.pnlDown.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDown.Location = new System.Drawing.Point(119, 50);
            this.pnlDown.Margin = new System.Windows.Forms.Padding(0);
            this.pnlDown.Name = "pnlDown";
            this.pnlDown.Size = new System.Drawing.Size(222, 50);
            
            // txtDownPayment
            this.txtDownPayment.Location = new System.Drawing.Point(3, 11);
            this.txtDownPayment.Margin = new System.Windows.Forms.Padding(3, 11, 3, 3);
            this.txtDownPayment.Name = "txtDownPayment";
            this.txtDownPayment.Size = new System.Drawing.Size(120, 27);
            
            // cmbDownPaymentType
            this.cmbDownPaymentType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDownPaymentType.Items.AddRange(new object[] { "%", "元" });
            this.cmbDownPaymentType.Location = new System.Drawing.Point(129, 11);
            this.cmbDownPaymentType.Margin = new System.Windows.Forms.Padding(3, 11, 3, 3);
            this.cmbDownPaymentType.Name = "cmbDownPaymentType";
            this.cmbDownPaymentType.Size = new System.Drawing.Size(60, 27);
            
            // lblRate
            this.lblRate.AutoSize = true;
            this.lblRate.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblRate.Location = new System.Drawing.Point(3, 100);
            this.lblRate.Name = "lblRate";
            this.lblRate.Size = new System.Drawing.Size(113, 50);
            this.lblRate.Text = "貸款年利率(%)";
            this.lblRate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            
            // txtRate
            this.txtRate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtRate.Location = new System.Drawing.Point(122, 111);
            this.txtRate.Name = "txtRate";
            this.txtRate.Size = new System.Drawing.Size(216, 27);
            
            // lblTerm
            this.lblTerm.AutoSize = true;
            this.lblTerm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTerm.Location = new System.Drawing.Point(3, 150);
            this.lblTerm.Name = "lblTerm";
            this.lblTerm.Size = new System.Drawing.Size(113, 50);
            this.lblTerm.Text = "貸款年限(年)";
            this.lblTerm.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            
            // cmbTerm
            this.cmbTerm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTerm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbTerm.Items.AddRange(new object[] { "10", "15", "20", "30", "40" });
            this.cmbTerm.Location = new System.Drawing.Point(122, 161);
            this.cmbTerm.Name = "cmbTerm";
            this.cmbTerm.Size = new System.Drawing.Size(216, 27);
            
            // lblGrace
            this.lblGrace.AutoSize = true;
            this.lblGrace.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblGrace.Location = new System.Drawing.Point(3, 200);
            this.lblGrace.Name = "lblGrace";
            this.lblGrace.Size = new System.Drawing.Size(113, 50);
            this.lblGrace.Text = "使用寬限期(年)";
            this.lblGrace.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            
            // txtGrace
            this.txtGrace.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtGrace.Location = new System.Drawing.Point(122, 211);
            this.txtGrace.Name = "txtGrace";
            this.txtGrace.Size = new System.Drawing.Size(216, 27);
            
            // pnlButtons
            this.pnlButtons.Controls.Add(this.btnCalc);
            this.pnlButtons.Controls.Add(this.btnReset);
            this.pnlButtons.Controls.Add(this.btnExport);
            this.pnlButtons.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlButtons.Location = new System.Drawing.Point(119, 250);
            this.pnlButtons.Margin = new System.Windows.Forms.Padding(0);
            this.pnlButtons.Name = "pnlButtons";
            this.pnlButtons.Size = new System.Drawing.Size(222, 100);
            
            // btnCalc
            this.btnCalc.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(174)))), ((int)(((byte)(96)))));
            this.btnCalc.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCalc.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCalc.FlatAppearance.BorderSize = 0;
            this.btnCalc.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btnCalc.ForeColor = System.Drawing.Color.White;
            this.btnCalc.Location = new System.Drawing.Point(3, 10);
            this.btnCalc.Margin = new System.Windows.Forms.Padding(3, 10, 3, 3);
            this.btnCalc.Name = "btnCalc";
            this.btnCalc.Size = new System.Drawing.Size(200, 40);
            this.btnCalc.Text = "➜ 開始深度試算";
            this.btnCalc.Click += new System.EventHandler(this.btnCalc_Click);
            
            // btnReset
            this.btnReset.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(149)))), ((int)(((byte)(165)))), ((int)(((byte)(166)))));
            this.btnReset.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnReset.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReset.FlatAppearance.BorderSize = 0;
            this.btnReset.Font = new System.Drawing.Font("微軟正黑體", 11F, System.Drawing.FontStyle.Bold);
            this.btnReset.ForeColor = System.Drawing.Color.White;
            this.btnReset.Location = new System.Drawing.Point(3, 56);
            this.btnReset.Margin = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(90, 35);
            this.btnReset.Text = "清除重置";
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);

            // btnExport
            this.btnExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.btnExport.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExport.FlatAppearance.BorderSize = 0;
            this.btnExport.Font = new System.Drawing.Font("微軟正黑體", 11F, System.Drawing.FontStyle.Bold);
            this.btnExport.ForeColor = System.Drawing.Color.White;
            this.btnExport.Location = new System.Drawing.Point(100, 56);
            this.btnExport.Margin = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(103, 35);
            this.btnExport.Text = "匯出 CSV";
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);

            // lblValidationHint
            this.lblValidationHint.AutoSize = true;
            this.tableLayoutPanelInput.SetColumnSpan(this.lblValidationHint, 2);
            this.lblValidationHint.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblValidationHint.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(57)))), ((int)(((byte)(43)))));
            this.lblValidationHint.Location = new System.Drawing.Point(3, 350);
            this.lblValidationHint.Name = "lblValidationHint";
            this.lblValidationHint.Size = new System.Drawing.Size(335, 19);
            this.lblValidationHint.TabIndex = 11;
            this.lblValidationHint.Text = "";
            
            // tabControlHelper
            this.tabControlHelper.Controls.Add(this.tabSummary);
            this.tabControlHelper.Controls.Add(this.tabAI);
            this.tabControlHelper.Controls.Add(this.tabSchedule);
            this.tabControlHelper.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlHelper.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tabControlHelper.Location = new System.Drawing.Point(400, 23);
            this.tabControlHelper.Name = "tabControlHelper";
            this.tabControlHelper.SelectedIndex = 0;
            this.tabControlHelper.Size = new System.Drawing.Size(561, 554);
            
            // tabSummary
            this.tabSummary.Controls.Add(this.tableLayoutPanelOutput);
            this.tabSummary.Controls.Add(this.picChart);
            this.tabSummary.Location = new System.Drawing.Point(4, 29);
            this.tabSummary.Name = "tabSummary";
            this.tabSummary.Padding = new System.Windows.Forms.Padding(15);
            this.tabSummary.Size = new System.Drawing.Size(553, 521);
            this.tabSummary.Text = "📊 全局試算結果";
            this.tabSummary.UseVisualStyleBackColor = true;
            
            // tableLayoutPanelOutput
            this.tableLayoutPanelOutput.ColumnCount = 2;
            this.tableLayoutPanelOutput.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.tableLayoutPanelOutput.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.tableLayoutPanelOutput.Controls.Add(this.lblTitleLoan, 0, 0);
            this.tableLayoutPanelOutput.Controls.Add(this.lblResultTotalLoan, 1, 0);
            this.tableLayoutPanelOutput.Controls.Add(this.lblTitleMonthly, 0, 1);
            this.tableLayoutPanelOutput.Controls.Add(this.lblResultMonthly, 1, 1);
            this.tableLayoutPanelOutput.Controls.Add(this.lblTitleFirstInt, 0, 2);
            this.tableLayoutPanelOutput.Controls.Add(this.lblResultFirstInt, 1, 2);
            this.tableLayoutPanelOutput.Controls.Add(this.lblTitleFirstPrin, 0, 3);
            this.tableLayoutPanelOutput.Controls.Add(this.lblResultFirstPrin, 1, 3);
            this.tableLayoutPanelOutput.Controls.Add(this.lblTitleTotalInt, 0, 4);
            this.tableLayoutPanelOutput.Controls.Add(this.lblResultTotalInt, 1, 4);
            this.tableLayoutPanelOutput.Controls.Add(this.lblTitleTotalRepay, 0, 5);
            this.tableLayoutPanelOutput.Controls.Add(this.lblResultTotalRepay, 1, 5);
            this.tableLayoutPanelOutput.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanelOutput.Location = new System.Drawing.Point(15, 15);
            this.tableLayoutPanelOutput.Name = "tableLayoutPanelOutput";
            this.tableLayoutPanelOutput.RowCount = 6;
            this.tableLayoutPanelOutput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            this.tableLayoutPanelOutput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            this.tableLayoutPanelOutput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            this.tableLayoutPanelOutput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            this.tableLayoutPanelOutput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            this.tableLayoutPanelOutput.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            this.tableLayoutPanelOutput.Size = new System.Drawing.Size(523, 271);
            
            // Output Labels Template
            System.Drawing.Font titleFont = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            System.Drawing.Font valueFont = new System.Drawing.Font("Consolas", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            System.Drawing.Color titleColor = System.Drawing.Color.Gray;
            System.Drawing.Color valueColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(62)))), ((int)(((byte)(80)))));

            this.lblTitleLoan.Text = "總核貸金額：";
            this.lblTitleLoan.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitleLoan.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblTitleLoan.Font = titleFont;
            this.lblTitleLoan.ForeColor = titleColor;

            this.lblResultTotalLoan.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblResultTotalLoan.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblResultTotalLoan.Font = valueFont;
            this.lblResultTotalLoan.ForeColor = valueColor;

            this.lblTitleMonthly.Text = "每月還款負擔：";
            this.lblTitleMonthly.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitleMonthly.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblTitleMonthly.Font = titleFont;
            this.lblTitleMonthly.ForeColor = titleColor;

            this.lblResultMonthly.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblResultMonthly.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblResultMonthly.Font = new System.Drawing.Font("Consolas", 16F, System.Drawing.FontStyle.Bold);
            this.lblResultMonthly.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(57)))), ((int)(((byte)(43))))); // Red for monthly

            this.lblTitleFirstInt.Text = "首期衍生利息：";
            this.lblTitleFirstInt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitleFirstInt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblTitleFirstInt.Font = titleFont;
            this.lblTitleFirstInt.ForeColor = titleColor;

            this.lblResultFirstInt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblResultFirstInt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblResultFirstInt.Font = valueFont;

            this.lblTitleFirstPrin.Text = "首期實還本金：";
            this.lblTitleFirstPrin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitleFirstPrin.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblTitleFirstPrin.Font = titleFont;
            this.lblTitleFirstPrin.ForeColor = titleColor;

            this.lblResultFirstPrin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblResultFirstPrin.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblResultFirstPrin.Font = valueFont;

            this.lblTitleTotalInt.Text = "利息累計支出：";
            this.lblTitleTotalInt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitleTotalInt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblTitleTotalInt.Font = titleFont;
            this.lblTitleTotalInt.ForeColor = titleColor;

            this.lblResultTotalInt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblResultTotalInt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblResultTotalInt.Font = valueFont;
            this.lblResultTotalInt.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(126)))), ((int)(((byte)(34))))); // Orange

            this.lblTitleTotalRepay.Text = "終期總還款數：";
            this.lblTitleTotalRepay.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitleTotalRepay.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblTitleTotalRepay.Font = titleFont;
            this.lblTitleTotalRepay.ForeColor = titleColor;

            this.lblResultTotalRepay.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblResultTotalRepay.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblResultTotalRepay.Font = valueFont;
            
            // picChart
            this.picChart.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.picChart.Location = new System.Drawing.Point(15, 292);
            this.picChart.Name = "picChart";
            this.picChart.Size = new System.Drawing.Size(523, 214);
            this.picChart.Paint += new System.Windows.Forms.PaintEventHandler(this.picChart_Paint);
            
            // tabAI
            this.tabAI.Controls.Add(this.rtbAI);
            this.tabAI.Location = new System.Drawing.Point(4, 29);
            this.tabAI.Name = "tabAI";
            this.tabAI.Padding = new System.Windows.Forms.Padding(15);
            this.tabAI.Size = new System.Drawing.Size(553, 521);
            this.tabAI.Text = "🧠 AI 財務分析報告";
            this.tabAI.UseVisualStyleBackColor = true;
            
            // rtbAI
            this.rtbAI.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rtbAI.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.rtbAI.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.rtbAI.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(62)))), ((int)(((byte)(80)))));
            this.rtbAI.Location = new System.Drawing.Point(15, 15);
            this.rtbAI.Name = "rtbAI";
            this.rtbAI.ReadOnly = true;
            this.rtbAI.Size = new System.Drawing.Size(523, 491);
            this.rtbAI.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtbAI.Text = "未執行分析...";

            // tabSchedule
            this.tabSchedule.Controls.Add(this.dgvSchedule);
            this.tabSchedule.Location = new System.Drawing.Point(4, 29);
            this.tabSchedule.Name = "tabSchedule";
            this.tabSchedule.Padding = new System.Windows.Forms.Padding(0);
            this.tabSchedule.Size = new System.Drawing.Size(553, 521);
            this.tabSchedule.Text = "📅 各期攤還報表";
            this.tabSchedule.UseVisualStyleBackColor = true;
            
            // dgvSchedule
            this.dgvSchedule.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSchedule.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvSchedule.Location = new System.Drawing.Point(0, 0);
            this.dgvSchedule.Name = "dgvSchedule";
            this.dgvSchedule.RowTemplate.Height = 30; // Larger row height for readability
            this.dgvSchedule.Size = new System.Drawing.Size(553, 521);
            
            // Form1
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 660);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnMinimize);
            this.Controls.Add(this.pnlMain);
            this.Controls.Add(this.lblMainTitle);
            this.Name = "Form1";
            this.Text = "個人房貸計算器";
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.pnlMain.ResumeLayout(false);
            this.gbInput.ResumeLayout(false);
            this.tableLayoutPanelInput.ResumeLayout(false);
            this.tableLayoutPanelInput.PerformLayout();
            this.pnlDown.ResumeLayout(false);
            this.pnlDown.PerformLayout();
            this.pnlButtons.ResumeLayout(false);
            this.tabControlHelper.ResumeLayout(false);
            this.tabSummary.ResumeLayout(false);
            this.tableLayoutPanelOutput.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picChart)).EndInit();
            this.tabSchedule.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSchedule)).EndInit();
            this.tabAI.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.Label lblMainTitle;
        private System.Windows.Forms.Label btnClose;
        private System.Windows.Forms.Label btnMinimize;
        private System.Windows.Forms.TableLayoutPanel pnlMain;
        private System.Windows.Forms.GroupBox gbInput;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelInput;
        private System.Windows.Forms.Label lblPrice;
        private System.Windows.Forms.TextBox txtPrice;
        private System.Windows.Forms.Label lblDownPayment;
        private System.Windows.Forms.FlowLayoutPanel pnlDown;
        private System.Windows.Forms.TextBox txtDownPayment;
        private System.Windows.Forms.ComboBox cmbDownPaymentType;
        private System.Windows.Forms.Label lblRate;
        private System.Windows.Forms.TextBox txtRate;
        private System.Windows.Forms.Label lblTerm;
        private System.Windows.Forms.ComboBox cmbTerm;
        private System.Windows.Forms.Label lblGrace;
        private System.Windows.Forms.TextBox txtGrace;
        private System.Windows.Forms.FlowLayoutPanel pnlButtons;
        private System.Windows.Forms.Button btnCalc;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Label lblValidationHint;
        private System.Windows.Forms.TabControl tabControlHelper;
        private System.Windows.Forms.TabPage tabSummary;
        private System.Windows.Forms.TabPage tabSchedule;
        private System.Windows.Forms.TabPage tabAI;
        private System.Windows.Forms.RichTextBox rtbAI;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelOutput;
        private System.Windows.Forms.Label lblTitleLoan;
        private System.Windows.Forms.Label lblResultTotalLoan;
        private System.Windows.Forms.Label lblTitleMonthly;
        private System.Windows.Forms.Label lblResultMonthly;
        private System.Windows.Forms.Label lblTitleFirstInt;
        private System.Windows.Forms.Label lblResultFirstInt;
        private System.Windows.Forms.Label lblTitleFirstPrin;
        private System.Windows.Forms.Label lblResultFirstPrin;
        private System.Windows.Forms.Label lblTitleTotalInt;
        private System.Windows.Forms.Label lblResultTotalInt;
        private System.Windows.Forms.Label lblTitleTotalRepay;
        private System.Windows.Forms.Label lblResultTotalRepay;
        private System.Windows.Forms.PictureBox picChart;
        private System.Windows.Forms.DataGridView dgvSchedule;
    }
}
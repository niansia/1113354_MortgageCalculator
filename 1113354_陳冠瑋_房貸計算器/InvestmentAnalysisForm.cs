using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace _1113354_陳冠瑋_房貸計算器
{
    public class InvestmentAnalysisForm : Form
    {
        private double _propertyPrice;
        private double _downPayment;
        private double _loanAmount;
        private double _monthlyPayment;
        private double _loanRate;
        private int _totalMonths;

        // Input controls
        private NumericUpDown numGrossRent;
        private NumericUpDown numVacancyRate;
        private NumericUpDown numOpexRate;
        private NumericUpDown numPropApp;
        private NumericUpDown numSellingCost;
        private NumericUpDown numHoldingYears;

        private Label lblKpi;
        private DataGridView dgvCashFlow;

        public InvestmentAnalysisForm(double propertyPrice, double downPayment, double monthlyPayment, int totalMonths, double loanRate)
        {
            _propertyPrice = Math.Max(1, propertyPrice);
            _downPayment = downPayment;
            _loanAmount = _propertyPrice - _downPayment;
            _monthlyPayment = monthlyPayment;
            _totalMonths = totalMonths;
            _loanRate = loanRate;

            InitializeUI();
            CalculateMetrics();
        }

        private void InitializeUI()
        {
            this.Text = "📈 不動產投資報酬分析 (IRR, Cap Rate, Cash-on-Cash)";
            this.Size = new Size(1000, 680);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("微軟正黑體", 10F);
            this.ShowIcon = false;

            var lblTitle = new Label
            {
                Text = "🏢 商業不動產投資評估 (Commercial RE Investment Pro Forma)",
                Dock = DockStyle.Top,
                Font = new Font("微軟正黑體", 16F, FontStyle.Bold),
                BackColor = Color.FromArgb(39, 174, 96),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 60
            };
            this.Controls.Add(lblTitle);

            var pnlTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 85,
                Padding = new Padding(10, 10, 10, 5),
                BackColor = Color.White,
                WrapContents = true,
                AutoScroll = true
            };
            this.Controls.Add(pnlTop);
            pnlTop.BringToFront();

            // Default defaults
            decimal defRent = (decimal)Math.Round((_propertyPrice * 0.035) / 12.0 / 100) * 100; // 3.5% basic yield

            numGrossRent = AddInput(pnlTop, "預估月租金($):", defRent, 0, 10000000, 1000);
            numVacancyRate = AddInput(pnlTop, "空置率(Vacancy %):", 5.0M, 0, 50, 0.5M);
            numOpexRate = AddInput(pnlTop, "營運雜支(OpEx %):", 15.0M, 0, 80, 1M);
            numPropApp = AddInput(pnlTop, "年房價增值(%):", 2.5M, -10, 30, 0.5M);
            numSellingCost = AddInput(pnlTop, "處分賣房成本(%):", 4.0M, 0, 15, 0.5M);
            numHoldingYears = AddInput(pnlTop, "投資持有期(年):", 10M, 1, 50, 1M);

            var btnCalc = new Button
            {
                Text = "計算報酬率",
                BackColor = Color.FromArgb(41, 128, 185),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 100,
                Height = 32,
                Margin = new Padding(15, 8, 0, 0),
                Font = new Font("微軟正黑體", 10F, FontStyle.Bold)
            };
            btnCalc.FlatAppearance.BorderSize = 0;
            btnCalc.Click += (s, e) => CalculateMetrics();
            pnlTop.Controls.Add(btnCalc);

            lblKpi = new Label
            {
                Dock = DockStyle.Top,
                Height = 120,
                Padding = new Padding(15),
                Font = new Font("Consolas", 11.5F, FontStyle.Bold),
                BackColor = Color.FromArgb(235, 245, 251),
                ForeColor = Color.FromArgb(44, 62, 80),
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(lblKpi);
            lblKpi.BringToFront();

            dgvCashFlow = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                EnableHeadersVisualStyles = false
            };
            dgvCashFlow.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 73, 94);
            dgvCashFlow.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvCashFlow.ColumnHeadersDefaultCellStyle.Padding = new Padding(4);
            dgvCashFlow.ColumnHeadersHeight = 36;
            dgvCashFlow.DefaultCellStyle.Padding = new Padding(4);
            dgvCashFlow.RowTemplate.Height = 30;

            this.Controls.Add(dgvCashFlow);
            dgvCashFlow.BringToFront();
        }

        private NumericUpDown AddInput(FlowLayoutPanel pnl, string label, decimal val, decimal min, decimal max, decimal inc)
        {
            pnl.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(10, 12, 0, 0) });
            
            if (val < min) val = min;
            if (val > max) val = max;

            var num = new NumericUpDown
            {
                Minimum = min, Maximum = max, Value = val, Increment = inc, Width = 80,
                DecimalPlaces = inc < 1 ? 1 : 0, Margin = new Padding(2, 10, 8, 0)
            };
            pnl.Controls.Add(num);
            return num;
        }

        private void CalculateMetrics()
        {
            double grossRent = (double)numGrossRent.Value * 12; // GI
            double vacancy = (double)numVacancyRate.Value / 100.0;
            double opexRate = (double)numOpexRate.Value / 100.0;
            double appRate = (double)numPropApp.Value / 100.0;
            double sellCost = (double)numSellingCost.Value / 100.0;
            int hYears = (int)numHoldingYears.Value;

            double egi = grossRent * (1 - vacancy); // Effective Gross Income
            double opex = egi * opexRate; // Operating Expenses
            double noi = egi - opex; // Net Operating Income (首年NOI)

            double capRate = _propertyPrice > 0 ? (noi / _propertyPrice) : 0;
            
            double annualDebtService = _monthlyPayment * 12;
            double preTaxCashFlow = noi - annualDebtService; // PTCF
            
            double cashOnCash = _downPayment > 0 ? (preTaxCashFlow / _downPayment) : 0;

            // Compute Cash Flows over Holding Period
            List<double> cashFlows = new List<double>();
            cashFlows.Add(-_downPayment); // Year 0 initial equity

            dgvCashFlow.Rows.Clear();
            dgvCashFlow.Columns.Clear();
            dgvCashFlow.Columns.Add("Year", "年度");
            dgvCashFlow.Columns.Add("Value", "推估房價");
            dgvCashFlow.Columns.Add("NOI", "NOI (淨營運收入)");
            dgvCashFlow.Columns.Add("Debt", "房貸支出");
            dgvCashFlow.Columns.Add("CashFlow", "淨現金流");
            dgvCashFlow.Columns.Add("Balance", "剩餘貸款");

            double currentVal = _propertyPrice;
            double balance = _loanAmount;
            double mRate = (_loanRate / 100.0) / 12.0;

            for (int y = 1; y <= hYears; y++)
            {
                currentVal *= (1 + appRate);

                // Assuming NOI increases with property value appreciation roughly
                double currentNoi = noi * Math.Pow(1 + appRate, y - 1);
                
                double annualPrin = 0;
                for (int m = 1; m <= 12; m++)
                {
                    if (balance > 0)
                    {
                        double interest = balance * mRate;
                        double prin = _monthlyPayment - interest;
                        if (prin > balance) prin = balance;
                        balance -= prin;
                        annualPrin += prin;
                    }
                }

                double cf = currentNoi - annualDebtService;

                // If it's the last year, add sale proceeds to cashflow
                if (y == hYears)
                {
                    double netSaleProceeds = (currentVal * (1 - sellCost)) - balance;
                    cf += netSaleProceeds;
                    dgvCashFlow.Rows.Add(y, $"NT$ {currentVal:N0}", $"NT$ {currentNoi:N0}", $"NT$ {annualDebtService:N0}", $"NT$ {cf:N0} (含售出)", $"NT$ {balance:N0}");
                }
                else
                {
                    dgvCashFlow.Rows.Add(y, $"NT$ {currentVal:N0}", $"NT$ {currentNoi:N0}", $"NT$ {annualDebtService:N0}", $"NT$ {cf:N0}", $"NT$ {balance:N0}");
                }

                cashFlows.Add(cf);
            }

            double irr = ComputeIRR(cashFlows.ToArray());

            lblKpi.Text = 
                $"【核心投資回報指標 (KPIs)】\n" +
                $"首年資本化率 (Cap Rate):      {capRate:P2}   (NOI / 房產總價 => 評估租金收益與資產定價的基準)\n" +
                $"現金回報率 (Cash-on-Cash) :  {cashOnCash:P2}   (年度淨現金流 / 投入的自備款 => 觀察實質槓桿效率)\n" +
                $"內部報酬率 (IRR @ {hYears}年)  :  {(irr > -1 ? irr.ToString("P2") : "無限/低於-100%")}   (包含長期租金與最終房屋售出淨值)\n" +
                $"首年淨營運收入 (NOI)      : NT$ {noi:N0} / 房產售出稅後淨值(Year {hYears}): NT$ {cashFlows.Last():N0}";
        }

        // Simple Newton-Raphson method for IRR
        private double ComputeIRR(double[] values, double guess = 0.1)
        {
            double irr = guess;
            double step = 0.0001;
            int maxIter = 1000;
            
            for (int i = 0; i < maxIter; i++)
            {
                double npv = 0;
                double dNpv = 0;
                
                for (int j = 0; j < values.Length; j++)
                {
                    npv += values[j] / Math.Pow(1 + irr, j);
                    dNpv -= j * values[j] / Math.Pow(1 + irr, j + 1);
                }
                
                if (Math.Abs(npv) < 0.0001) return irr;
                if (Math.Abs(dNpv) < 1e-10) break;
                
                double nextIrr = irr - npv / dNpv;
                if (Math.Abs(nextIrr - irr) < 1e-6) return nextIrr;
                irr = nextIrr;
            }
            return irr;
        }
    }
}
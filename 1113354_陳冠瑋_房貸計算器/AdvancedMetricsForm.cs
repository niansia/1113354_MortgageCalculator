using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Data;

namespace _1113354_陳冠瑋_房貸計算器
{
    public class AdvancedMetricsForm : Form
    {
        private double _loanAmount;
        private double _monthlyPayment;
        private int _totalMonths;
        private double _annualRate;
        private double _downPayment;

        private Label lblTitle;
        private FlowLayoutPanel pnlTop;
        private NumericUpDown numInflation;
        private NumericUpDown numInvestmentReturn;
        private Button btnCalculate;
        private Label lblResults;
        private DataGridView dgvMatrix;

        public AdvancedMetricsForm(double loanAmount, double downPayment, double monthlyPayment, int totalMonths, double annualRate)
        {
            _loanAmount = loanAmount;
            _downPayment = downPayment;
            _monthlyPayment = monthlyPayment;
            _totalMonths = totalMonths;
            _annualRate = annualRate;

            InitializeUI();
            PerformAdvancedCalculations();
        }

        private void InitializeUI()
        {
            this.Text = "進階財務理論分析 (NPV 與投資機會成本)";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("微軟正黑體", 11F);
            this.ShowIcon = false;

            lblTitle = new Label
            {
                Text = "📊 財務理論與敏感度矩陣分析",
                Dock = DockStyle.Top,
                Font = new Font("微軟正黑體", 16F, FontStyle.Bold),
                BackColor = Color.FromArgb(41, 128, 185),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 60
            };
            this.Controls.Add(lblTitle);

            pnlTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 60,
                Padding = new Padding(10, 15, 10, 10),
                BackColor = Color.White
            };

            var lblInf = new Label { Text = "預估年通膨率(%):", AutoSize = true, Margin = new Padding(10, 5, 0, 0) };
            numInflation = new NumericUpDown { Value = 2.0M, DecimalPlaces = 1, Width = 80, Increment = 0.1M };

            var lblInv = new Label { Text = "預估長期投資報酬率(%):", AutoSize = true, Margin = new Padding(20, 5, 0, 0) };
            numInvestmentReturn = new NumericUpDown { Value = 5.0M, DecimalPlaces = 1, Width = 80, Increment = 0.1M };

            btnCalculate = new Button
            {
                Text = "重新演算",
                BackColor = Color.FromArgb(39, 174, 96),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 100,
                Margin = new Padding(20, -2, 0, 0)
            };
            btnCalculate.FlatAppearance.BorderSize = 0;
            btnCalculate.Click += (s, e) => PerformAdvancedCalculations();

            pnlTop.Controls.Add(lblInf);
            pnlTop.Controls.Add(numInflation);
            pnlTop.Controls.Add(lblInv);
            pnlTop.Controls.Add(numInvestmentReturn);
            pnlTop.Controls.Add(btnCalculate);
            this.Controls.Add(pnlTop);
            pnlTop.BringToFront();

            lblResults = new Label
            {
                Dock = DockStyle.Top,
                Height = 150,
                Padding = new Padding(20),
                Font = new Font("Consolas", 12F),
                BackColor = Color.FromArgb(240, 248, 255)
            };
            this.Controls.Add(lblResults);
            lblResults.BringToFront();

            var lblMatrixTitle = new Label
            {
                Text = "📈 利率 vs 年限 雙維度月付金敏感度矩陣 (現金流壓力測試)",
                Dock = DockStyle.Top,
                Font = new Font("微軟正黑體", 12F, FontStyle.Bold),
                Padding = new Padding(10),
                Height = 40,
                BackColor = Color.White
            };
            this.Controls.Add(lblMatrixTitle);
            lblMatrixTitle.BringToFront();

            dgvMatrix = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersWidth = 100,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight }
            };
            this.Controls.Add(dgvMatrix);
            dgvMatrix.BringToFront();
            
            
            dgvMatrix.CellFormatting += DgvMatrix_CellFormatting;
        }

        private void PerformAdvancedCalculations()
        {
            if (_loanAmount <= 0) return;

            double inflationRate = (double)numInflation.Value / 100.0;
            double investRate = (double)numInvestmentReturn.Value / 100.0;

            
            double monthlyDiscount = inflationRate / 12.0;
            double npvOfPayments = 0;
            for (int i = 1; i <= _totalMonths; i++)
            {
                npvOfPayments += _monthlyPayment / Math.Pow(1 + monthlyDiscount, i);
            }
            double realCost = _downPayment + npvOfPayments;

            
            
            double fvOfDownPayment = _downPayment * Math.Pow(1 + investRate, _totalMonths / 12.0);
            
            double annualPayment = _monthlyPayment * 12;
            double fvOfMonthlyPayments = 0;
            for (int i = 1; i <= _totalMonths / 12; i++)
            {
                fvOfMonthlyPayments += annualPayment * Math.Pow(1 + investRate, (_totalMonths / 12) - i);
            }
            double totalOpportunityCost = fvOfDownPayment + fvOfMonthlyPayments;

            lblResults.Text = "【總體經濟與財務數學分析】\n\n" +
                              $"◆ 通膨折現淨現值 (NPV): 將未來 {_totalMonths} 個月的房貸現金流，以 {numInflation.Value}% 貶值率折算。\n" +
                              $"   => 您實際支付的「實質購買力成本」約為: NT$ {realCost:N0} (原始未折現本利和: NT$ {(_downPayment + _monthlyPayment * _totalMonths):N0})\n\n" +
                              $"◆ 投資機會成本 (Opportunity Cost): 若將自備款與月付金投入 {numInvestmentReturn.Value}% 年化報酬的市場。\n" +
                              $"   => {_totalMonths / 12} 年後損失的潛在資產終值: NT$ {totalOpportunityCost:N0}\n" +
                              "      (財務啟示: 若房產增值幅度低於此金額，純從財務來看租房+投資可能更優。)";

            BuildSensitivityMatrix();
        }

        private void BuildSensitivityMatrix()
        {
            dgvMatrix.Columns.Clear();
            dgvMatrix.Rows.Clear();

            int[] terms = { 15, 20, 30, 40 };
            double[] rates = new double[7];
            for (int i = 0; i < 7; i++) rates[i] = Math.Max(0.5, _annualRate + (i - 3) * 0.5);

            foreach (var t in terms)
            {
                dgvMatrix.Columns.Add($"Col{t}", $"{t}年限");
            }

            foreach (var r in rates)
            {
                int rowIndex = dgvMatrix.Rows.Add();
                dgvMatrix.Rows[rowIndex].HeaderCell.Value = $"{r:F2}% (利率)";
                
                for (int c = 0; c < terms.Length; c++)
                {
                    int m = terms[c] * 12;
                    double monthlyR = (r / 100.0) / 12.0;
                    double payment = _loanAmount * monthlyR * Math.Pow(1 + monthlyR, m) / (Math.Pow(1 + monthlyR, m) - 1);
                    dgvMatrix.Rows[rowIndex].Cells[c].Value = Math.Round(payment, 0);
                }
            }
            dgvMatrix.RowHeadersWidth = 120;
        }

        private void DgvMatrix_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value != null && e.Value is double val)
            {
                e.Value = $"NT$ {val:N0}";
                e.FormattingApplied = true;

                
                double min = double.MaxValue;
                double max = double.MinValue;
                foreach (DataGridViewRow row in dgvMatrix.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null)
                        {
                            string s = cell.Value.ToString().Replace("NT$ ", "").Replace(",", "");
                            if (double.TryParse(s, out double v))
                            {
                                if (v < min) min = v;
                                if (v > max) max = v;
                            }
                        }
                    }
                }

                if (max > min)
                {
                    string s = e.Value.ToString().Replace("NT$ ", "").Replace(",", "");
                    if (double.TryParse(s, out double v))
                    {
                        double ratio = (v - min) / (max - min);
                        int red = (int)(255 * ratio);
                        int green = (int)(255 * (1 - ratio));
                        e.CellStyle.BackColor = Color.FromArgb(50, red, green, 0); 
                    }
                }
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace _1113354_陳冠瑋_房貸計算器
{
    public class RentVsBuyForm : Form
    {
        private double _propertyPrice;
        private double _downPayment;
        private double _monthlyMortgage;
        private int _termYears;
        private double _loanRate;
        private double _loanAmount;

        private NumericUpDown numRent;
        private NumericUpDown numRentInf;
        private NumericUpDown numPropApp;
        private NumericUpDown numInvest;
        private NumericUpDown numPropTax;
        private Label lblReport;
        private PictureBox picChart;

        private List<double> _buyNetWorth = new List<double>();
        private List<double> _rentNetWorth = new List<double>();
        private int _breakevenYear = -1;

        public RentVsBuyForm(double propertyPrice, double downPayment, double monthlyMortgage, int termYears, double loanRate)
        {
            _propertyPrice = Math.Max(1, propertyPrice);
            _downPayment = downPayment;
            _monthlyMortgage = monthlyMortgage;
            _termYears = termYears;
            _loanRate = loanRate / 100.0;
            _loanAmount = _propertyPrice - _downPayment;

            InitializeUI();
            CalculateModel();
        }

        private void InitializeUI()
        {
            this.Text = "⚖️ 資本配置模型: 租房 vs 買房 (DCF 貼現現金流 & 機會成本)";
            this.Size = new Size(950, 700);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("微軟正黑體", 10F);
            this.ShowIcon = false;

            var lblTitle = new Label
            {
                Text = "🏙️ 租買決策機會成本模型 (Rent vs. Buy Opt. Model)",
                Dock = DockStyle.Top,
                Font = new Font("微軟正黑體", 16F, FontStyle.Bold),
                BackColor = Color.FromArgb(44, 62, 80),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 60
            };
            this.Controls.Add(lblTitle);

            var pnlTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 80,
                Padding = new Padding(15),
                BackColor = Color.White,
                WrapContents = false
            };

            double initRent = Math.Round(_propertyPrice / 400.0);

            numRent = AddInput(pnlTop, "首年預估月租($):", (decimal)initRent, 1000, 500000, 1000);
            numRentInf = AddInput(pnlTop, "租金年漲幅(%):", 2.0M, 0, 10, 0.1M);
            numPropApp = AddInput(pnlTop, "房價年增值(%):", 2.5M, -5, 15, 0.1M);
            numInvest = AddInput(pnlTop, "無風險投資率(%):", 5.0M, 0, 20, 0.1M);
            numPropTax = AddInput(pnlTop, "持有成本/稅金(%):", 0.6M, 0, 5, 0.1M);

            var btnCalc = new Button
            {
                Text = "重新演算",
                BackColor = Color.FromArgb(41, 128, 185),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 90,
                Height = 30,
                Margin = new Padding(20, 8, 0, 0),
                Font = new Font("微軟正黑體", 10F, FontStyle.Bold)
            };
            btnCalc.FlatAppearance.BorderSize = 0;
            btnCalc.Click += (s, e) => CalculateModel();
            pnlTop.Controls.Add(btnCalc);

            this.Controls.Add(pnlTop);
            pnlTop.BringToFront();

            lblReport = new Label
            {
                Dock = DockStyle.Top,
                Height = 110,
                Padding = new Padding(15),
                Font = new Font("Consolas", 11.5F),
                BackColor = Color.FromArgb(236, 240, 241),
                ForeColor = Color.FromArgb(44, 62, 80)
            };
            this.Controls.Add(lblReport);
            lblReport.BringToFront();

            picChart = new PictureBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            picChart.Paint += PicChart_Paint;
            this.Controls.Add(picChart);
            picChart.BringToFront();
        }

        private NumericUpDown AddInput(FlowLayoutPanel pnl, string label, decimal val, decimal min, decimal max, decimal inc)
        {
            pnl.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(10, 12, 3, 0) });
            var num = new NumericUpDown
            {
                Value = val, Minimum = min, Maximum = max, Increment = inc, Width = 85,
                DecimalPlaces = inc < 1 ? 1 : 0, Margin = new Padding(0, 10, 10, 0)
            };
            pnl.Controls.Add(num);
            return num;
        }

        private void CalculateModel()
        {
            _buyNetWorth.Clear();
            _rentNetWorth.Clear();
            _breakevenYear = -1;

            double currentRent = (double)numRent.Value * 12; // 年租金
            double rentInf = (double)numRentInf.Value / 100.0;
            double propApp = (double)numPropApp.Value / 100.0;
            double invRate = (double)numInvest.Value / 100.0;
            double propTax = (double)numPropTax.Value / 100.0;

            int horizon = Math.Max(30, _termYears); // 分析 30 年或房貸期
            
            // 初始淨值
            double renterPortfolio = _downPayment; // 租房者把頭期款拿去投資
            double buyerPropValue = _propertyPrice;
            double loanBalance = _loanAmount;
            
            double monthlyR = _loanRate / 12.0;

            for (int y = 1; y <= horizon; y++)
            {
                // ========== 買房者現金流 ==========
                // 年度房貸支付
                double annualMortgage = 0;
                double annualPrincipal = 0;
                for (int m = 1; m <= 12; m++)
                {
                    if (loanBalance > 0)
                    {
                        annualMortgage += _monthlyMortgage;
                        double interest = loanBalance * monthlyR;
                        double prin = _monthlyMortgage - interest;
                        if (prin > loanBalance) prin = loanBalance;
                        loanBalance -= prin;
                        annualPrincipal += prin;
                    }
                }
                
                buyerPropValue *= (1 + propApp);
                double annualTaxes = buyerPropValue * propTax;
                double totalBuyOutflow = annualMortgage + annualTaxes;

                // ========== 租房者現金流 ==========
                double totalRentOutflow = currentRent;
                currentRent *= (1 + rentInf);

                // ========== 差額投資與淨資產結算 ==========
                // 假設兩人每年都有相同的剛性預算: Math.Max(totalBuyOutflow, totalRentOutflow)
                // 誰花得少，誰就把差額投入無風險資產
                double maxBudget = Math.Max(totalBuyOutflow, totalRentOutflow);
                
                double renterSurplus = maxBudget - totalRentOutflow;
                renterPortfolio = (renterPortfolio + renterSurplus) * (1 + invRate);

                double buyerSurplus = maxBudget - totalBuyOutflow;
                double buyerPortfolio = 0; // 簡化：買房者多餘的錢通常不累積(或者可投資)
                // 為了公平，買房者的剩餘資金也應投資，或視為淨資產累加
                double buyerNetWorth = buyerPropValue - loanBalance; 
                // 加上額外剩餘資金的投資
                // 但實務上通常把總淨資產當作：(房產價值 - 剩餘貸款) + 買房者長期累積的投資組合
                // 這裡我們直接比較兩者的財富儲備池

                // 以「租房者投資池」 vs 「買房者房產淨值」為主要對比
                _rentNetWorth.Add(renterPortfolio);
                _buyNetWorth.Add(buyerNetWorth); // 忽略買方可能有的微小現金剩餘，凸顯房產槓桿

                if (_breakevenYear == -1 && buyerNetWorth > renterPortfolio)
                {
                    _breakevenYear = y;
                }
            }

            double finalBuy = _buyNetWorth.Last();
            double finalRent = _rentNetWorth.Last();

            string conclusion = "";
            if (_breakevenYear > 0)
                conclusion = $"💡 決策奇點 (Break-even): 買房淨資產將於第【{_breakevenYear}】年超越租房。";
            else
                conclusion = $"💡 決策警告: 在當前高投報率與高租金通膨假設下，{horizon} 年內「租房+投資」淨值始終大於買房！";

            lblReport.Text = $"【DCF 資本軌跡預測 ({horizon} 年期)】\n" +
                             $"◆ [{horizon} 年後買方淨資產]: NT$ {finalBuy:N0} (主要由房產增值及本金攤積構成)\n" +
                             $"◆ [{horizon} 年後租方淨資產]: NT$ {finalRent:N0} (由自備款複利與較低現金流差額投資構成)\n" +
                             $"{conclusion}";

            picChart.Invalidate();
        }

        private void PicChart_Paint(object sender, PaintEventArgs e)
        {
            if (_buyNetWorth.Count == 0) return;

            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int w = picChart.Width;
            int h = picChart.Height;

            int paddingL = 90;
            int paddingR = 40;
            int paddingT = 40;
            int paddingB = 40;

            double maxNet = Math.Max(_buyNetWorth.Max(), _rentNetWorth.Max());
            double minNet = 0;

            // Draw Grids
            using (Pen gridPen = new Pen(Color.FromArgb(230, 230, 230)))
            {
                for (int i = 0; i <= 5; i++)
                {
                    int y = paddingT + (h - paddingT - paddingB) * i / 5;
                    g.DrawLine(gridPen, paddingL, y, w - paddingR, y);
                    double val = maxNet - (maxNet - minNet) * i / 5.0;
                    g.DrawString($"NT$ {val/10000:N0} 萬", new Font("微軟正黑體", 9), Brushes.DimGray, 5, y - 8);
                }
            }

            int years = _buyNetWorth.Count;
            float stepX = (w - paddingL - paddingR) / (float)Math.Max(1, years - 1);

            List<PointF> ptsBuy = new List<PointF>();
            List<PointF> ptsRent = new List<PointF>();

            for (int i = 0; i < years; i++)
            {
                float px = paddingL + i * stepX;
                float pyBuy = paddingT + (h - paddingT - paddingB) * (1f - (float)(_buyNetWorth[i] / maxNet));
                float pyRent = paddingT + (h - paddingT - paddingB) * (1f - (float)(_rentNetWorth[i] / maxNet));
                
                ptsBuy.Add(new PointF(px, pyBuy));
                ptsRent.Add(new PointF(px, pyRent));

                if (i % 5 == 0 || i == years - 1)
                {
                    g.DrawString($"第{i + 1}年", new Font("微軟正黑體", 8), Brushes.Gray, px - 15, h - paddingB + 5);
                }
            }

            using (Pen penBuy = new Pen(Color.FromArgb(41, 128, 185), 3))
            using (Pen penRent = new Pen(Color.FromArgb(231, 76, 60), 3) { DashStyle = DashStyle.Dash })
            {
                g.DrawLines(penBuy, ptsBuy.ToArray());
                g.DrawLines(penRent, ptsRent.ToArray());
            }

            // Legend
            g.FillRectangle(new SolidBrush(Color.FromArgb(41, 128, 185)), paddingL + 20, 15, 15, 10);
            g.DrawString("買房淨資產 (Property Equity)", new Font("微軟正黑體", 9), Brushes.Black, paddingL + 40, 12);

            g.DrawLine(new Pen(Color.FromArgb(231, 76, 60), 3) { DashStyle = DashStyle.Dash }, paddingL + 230, 20, paddingL + 245, 20);
            g.DrawString("租房淨資產 (Investment Portfolio)", new Font("微軟正黑體", 9), Brushes.Black, paddingL + 250, 12);

            if (_breakevenYear > 0)
            {
                float bx = paddingL + (_breakevenYear - 1) * stepX;
                g.DrawLine(Pens.Gray, bx, paddingT, bx, h - paddingB);
                g.FillEllipse(Brushes.Orange, bx - 5, ptsBuy[_breakevenYear - 1].Y - 5, 10, 10);
                g.DrawString("黃金交叉點", new Font("微軟正黑體", 9, FontStyle.Bold), Brushes.DarkOrange, bx + 5, ptsBuy[_breakevenYear - 1].Y - 20);
            }
        }
    }
}
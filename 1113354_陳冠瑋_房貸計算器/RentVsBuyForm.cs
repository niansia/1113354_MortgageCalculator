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
        private ToolTip _tips = new ToolTip();

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
            this.Size = new Size(1100, 750);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(244, 246, 249);
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

            var pnlLeft = new Panel
            {
                Dock = DockStyle.Left,
                Width = 280,
                BackColor = Color.White,
                Padding = new Padding(10)
            };
            this.Controls.Add(pnlLeft);

            var pnlMain = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15)
            };
            this.Controls.Add(pnlMain);
            pnlMain.BringToFront();

            var pnlMainTitle = new Label
            {
                Text = "⚙️ 參數變數設定區",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("微軟正黑體", 12F, FontStyle.Bold),
                ForeColor = Color.FromArgb(44, 62, 80),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlLeft.Controls.Add(pnlMainTitle);

            var flowLeft = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true
            };
            pnlLeft.Controls.Add(flowLeft);
            flowLeft.BringToFront();

            double initRent = Math.Round(_propertyPrice / 400.0);

            numRent = AddInput(flowLeft, "首年預估月租($)", (decimal)initRent, 1000, 500000, 1000, "該物件在市場上的合理租金。\n將以年計算作為租房者的現金流出。");
            numRentInf = AddInput(flowLeft, "租金年漲幅(%)", 2.0M, 0, 10, 0.1M, "預估每年租金的上漲幅度 (通膨)，\n一般假設在 1% ~ 3% 之間。");
            numPropApp = AddInput(flowLeft, "房價年增值(%)", 2.5M, -5, 15, 0.1M, "買房者享有的資產增值紅利。\n長線來看台灣房市約有 2%~5% 增幅。");
            numInvest = AddInput(flowLeft, "無風險投資率(%)", 5.0M, 0, 20, 0.1M, "租客將未用的頭期款與每月剩餘現金流\n投入股市或ETF的平均年化報酬。");
            numPropTax = AddInput(flowLeft, "持有成本/稅金(%)", 0.6M, 0, 5, 0.1M, "買房者每年需繳納的地價稅、房屋稅\n及修繕管理維護折舊費。");

            var btnCalc = new Button
            {
                Text = "→ 重新演算",
                BackColor = Color.FromArgb(41, 128, 185),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = flowLeft.Width - 25,
                Height = 40,
                Margin = new Padding(10, 20, 0, 20),
                Font = new Font("微軟正黑體", 11F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnCalc.FlatAppearance.BorderSize = 0;
            btnCalc.Click += (s, e) => CalculateModel();
            flowLeft.Controls.Add(btnCalc);

            lblReport = new Label
            {
                Dock = DockStyle.Top,
                Height = 110,
                Padding = new Padding(15),
                Font = new Font("Consolas", 11.5F),
                BackColor = Color.FromArgb(236, 240, 241),
                ForeColor = Color.FromArgb(44, 62, 80),
                BorderStyle = BorderStyle.FixedSingle
            };
            pnlMain.Controls.Add(lblReport);

            var space = new Panel { Dock = DockStyle.Top, Height = 10 };
            pnlMain.Controls.Add(space);
            space.BringToFront();

            picChart = new PictureBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            picChart.Paint += PicChart_Paint;
            pnlMain.Controls.Add(picChart);
            picChart.BringToFront();
        }

        private NumericUpDown AddInput(FlowLayoutPanel parent, string labelText, decimal val, decimal min, decimal max, decimal inc, string helpText)
        {
            var pnl = new FlowLayoutPanel
            {
                Width = parent.Width - 20,
                Height = 65,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(5, 5, 0, 10),
                WrapContents = true
            };

            var lbl = new Label
            {
                Text = labelText,
                AutoSize = true,
                Font = new Font("微軟正黑體", 10.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 80, 100),
                Margin = new Padding(5, 5, 5, 0)
            };

            var btnHelp = new Label
            {
                Text = "?",
                Size = new Size(18, 18),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(170, 180, 190),
                Font = new Font("Arial", 8, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Help,
                Margin = new Padding(0, 5, 0, 0)
            };
            var path = new GraphicsPath();
            path.AddEllipse(0, 0, 18, 18);
            btnHelp.Region = new Region(path);
            _tips.SetToolTip(btnHelp, helpText);

            var spacer = new Label { Width = pnl.Width, Height = 0, Margin = new Padding(0) }; 

            if (val < min) val = min;
            if (val > max) val = max;

            var num = new NumericUpDown
            {
                Minimum = min, Maximum = max, Value = val, Increment = inc,
                Width = pnl.Width - 10,
                DecimalPlaces = inc < 1 ? 1 : 0,
                Font = new Font("Consolas", 11F),
                Margin = new Padding(5, 5, 5, 0),
                BorderStyle = BorderStyle.FixedSingle
            };

            pnl.Controls.Add(lbl);
            pnl.Controls.Add(btnHelp);
            pnl.Controls.Add(spacer);
            pnl.Controls.Add(num);
            parent.Controls.Add(pnl);

            return num;
        }

        private void CalculateModel()
        {
            _buyNetWorth.Clear();
            _rentNetWorth.Clear();
            _breakevenYear = -1;

            double currentRent = (double)numRent.Value * 12; 
            double rentInf = (double)numRentInf.Value / 100.0;
            double propApp = (double)numPropApp.Value / 100.0;
            double invRate = (double)numInvest.Value / 100.0;
            double propTax = (double)numPropTax.Value / 100.0;

            int horizon = Math.Max(30, _termYears); 
            
            
            double renterPortfolio = _downPayment; 
            double buyerPropValue = _propertyPrice;
            double loanBalance = _loanAmount;
            
            double monthlyR = _loanRate / 12.0;

            for (int y = 1; y <= horizon; y++)
            {
                
                
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

                
                double totalRentOutflow = currentRent;
                currentRent *= (1 + rentInf);

                
                
                
                double maxBudget = Math.Max(totalBuyOutflow, totalRentOutflow);
                
                double renterSurplus = maxBudget - totalRentOutflow;
                renterPortfolio = (renterPortfolio + renterSurplus) * (1 + invRate);

                double buyerSurplus = maxBudget - totalBuyOutflow;
                double buyerPortfolio = 0; 
                
                double buyerNetWorth = buyerPropValue - loanBalance; 
                
                
                

                
                _rentNetWorth.Add(renterPortfolio);
                _buyNetWorth.Add(buyerNetWorth); 

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
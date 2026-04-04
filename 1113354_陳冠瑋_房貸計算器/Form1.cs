using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _1113354_陳冠瑋_房貸計算器
{
    public partial class Form1 : Form
    {
        // Interop for form dragging
        [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")] public static extern bool ReleaseCapture();
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        private const int WM_NCHITTEST = 0x84;
        private const int HTLEFT = 10;
        private const int HTRIGHT = 11;
        private const int HTTOP = 12;
        private const int HTTOPLEFT = 13;
        private const int HTTOPRIGHT = 14;
        private const int HTBOTTOM = 15;
        private const int HTBOTTOMLEFT = 16;
        private const int HTBOTTOMRIGHT = 17;
        private const int ResizeBorderWidth = 8;

        private double totalLoan = 0;
        private double totalRepayment = 0;
        private double totalInterest = 0;
        private int _lastGraceMonths = 0;
        private int _lastTotalMonths = 0;
        private List<AmortizationItem> _schedule = new List<AmortizationItem>();
        private BindingSource _bs = new BindingSource();
        private ToolTip _tips = new ToolTip();
        private float _currentScale = 1f;
        private Color _customAccent = Color.FromArgb(41, 128, 185);
        private ComboBox _cmbTheme;
        private ComboBox _cmbScale;
        private Button _btnCustomTheme;
        private Button _btnStressTest;
        private Button _btnCopySummary;
        private Button _btnExportPdf;
        private Button _btnMonteCarlo;
        private Button _btnInputMode;
        private Button _btnAnimStrength;
        private Label _btnMaximize;
        private ComboBox _cmbRegion;
        private ComboBox _cmbDistrict;
        private ComboBox _cmbPresetType;
        private NumericUpDown _numPing;
        private Button _btnEstimatePrice;
        private Button _btnApplyMarketPreset;
        private TextBox _txtDistrictSearch;
        private Label _lblPresetSource;
        private Color _windowBtnNormalBack;
        private Color _windowBtnHoverBack;
        private Color _windowBtnFore;
        private Timer _resultAnimTimer;
        private int _animTick = 0;
        private const int AnimTicks = 18;
        private Dictionary<Label, double> _animFrom = new Dictionary<Label, double>();
        private Dictionary<Label, double> _animTo = new Dictionary<Label, double>();
        private int _currentAnimTicks = AnimTicks;
        private int _animStrengthLevel = 2; // 2:高 1:低 0:關
        private bool _advancedInputVisible = true;
        private bool _monthlyHasGrace = false;
        private double _monthlyNormalValue = 0;
        private double _monthlyGraceValue = 0;
        private CheckBox _chkNewYouth;
        private NumericUpDown _numMonthlyIncome;
        private NumericUpDown _numAnnualPrepay;
        private TextBox _txtReportStudentId;
        private TextBox _txtReportStudentName;
        private TextBox _txtReportCourse;
        private TextBox _txtReportSchool;
        private TextBox _txtReportDepartment;
        private TextBox _txtReportAdvisor;
        private DateTimePicker _dtReportDate;
        private Button _btnPdfTemplate;
        private int _effectivePayoffMonths = 0;
        private readonly List<PointF> _balanceCurvePoints = new List<PointF>();
        private readonly List<int> _balanceCurveMonths = new List<int>();
        private readonly ToolTip _chartTip = new ToolTip();
        private readonly Dictionary<string, List<string>> _districtNameCache = new Dictionary<string, List<string>>();
        private int _chartZoomMonths = 0;
        private bool _showPrincipalShareLine = true;

        private TabPage _tabCompare;
        private DataGridView _dgvCompare;
        private Button _btnAddScenario;
        private Button _btnClearScenario;
        private Button _btnAutoScenario;
        private BindingSource _compareBs = new BindingSource();
        private Timer _uiMotionTimer;
        private float _chartDashOffset = 0f;

        private enum PdfTemplateMode
        {
            AcademicZh,
            BusinessBrief
        }

        private PdfTemplateMode _pdfTemplateMode = PdfTemplateMode.AcademicZh;

        public class ScenarioResult
        {
            public string 方案 { get; set; }
            public string 房價 { get; set; }
            public string 利率 { get; set; }
            public string 年限 { get; set; }
            public string 寬限 { get; set; }
            public string 年提前還款 { get; set; }
            public string 每月還款 { get; set; }
            public string 總利息 { get; set; }
            public string 總還款 { get; set; }
            public string 清償月數 { get; set; }
        }
        private List<ScenarioResult> _scenarios = new List<ScenarioResult>();

        private class MarketPreset
        {
            public double AvgTotalPrice { get; set; }
            public double AvgRate { get; set; }
            public int SuggestedTerm { get; set; }
        }

        private Dictionary<string, MarketPreset> _marketPresets = new Dictionary<string, MarketPreset>
        {
            { "台北市-大樓", new MarketPreset { AvgTotalPrice = 28000000, AvgRate = 2.18, SuggestedTerm = 30 } },
            { "新北市-大樓", new MarketPreset { AvgTotalPrice = 21000000, AvgRate = 2.16, SuggestedTerm = 30 } },
            { "桃園市-大樓", new MarketPreset { AvgTotalPrice = 16000000, AvgRate = 2.14, SuggestedTerm = 30 } },
            { "台中市-大樓", new MarketPreset { AvgTotalPrice = 17500000, AvgRate = 2.15, SuggestedTerm = 30 } },
            { "台南市-大樓", new MarketPreset { AvgTotalPrice = 13800000, AvgRate = 2.13, SuggestedTerm = 30 } },
            { "高雄市-大樓", new MarketPreset { AvgTotalPrice = 14500000, AvgRate = 2.14, SuggestedTerm = 30 } },
            { "台北市-公寓", new MarketPreset { AvgTotalPrice = 22000000, AvgRate = 2.17, SuggestedTerm = 25 } },
            { "新北市-公寓", new MarketPreset { AvgTotalPrice = 16800000, AvgRate = 2.15, SuggestedTerm = 25 } },
            { "台中市-透天", new MarketPreset { AvgTotalPrice = 19800000, AvgRate = 2.16, SuggestedTerm = 30 } },
            { "台南市-透天", new MarketPreset { AvgTotalPrice = 16500000, AvgRate = 2.14, SuggestedTerm = 30 } }
        };

        private readonly Dictionary<string, double> _countyUnitPriceWanPerPing = new Dictionary<string, double>
        {
            { "台北市", 95 }, { "新北市", 58 }, { "基隆市", 28 }, { "桃園市", 42 }, { "新竹市", 50 },
            { "新竹縣", 40 }, { "苗栗縣", 26 }, { "台中市", 42 }, { "彰化縣", 25 }, { "南投縣", 23 },
            { "雲林縣", 20 }, { "嘉義市", 24 }, { "嘉義縣", 18 }, { "台南市", 34 }, { "高雄市", 33 },
            { "屏東縣", 19 }, { "宜蘭縣", 27 }, { "花蓮縣", 23 }, { "台東縣", 18 }, { "澎湖縣", 21 },
            { "金門縣", 22 }, { "連江縣", 20 }
        };

        private readonly Dictionary<string, Dictionary<string, double>> _districtUnitPriceWanPerPing = new Dictionary<string, Dictionary<string, double>>
        {
            { "台北市", new Dictionary<string, double> { { "全市平均", 95 }, { "大安區", 122 }, { "信義區", 118 }, { "松山區", 110 }, { "中山區", 105 }, { "中正區", 102 }, { "內湖區", 93 }, { "文山區", 78 }, { "北投區", 82 }, { "萬華區", 88 } } },
            { "新北市", new Dictionary<string, double> { { "全市平均", 58 }, { "板橋區", 72 }, { "新店區", 66 }, { "永和區", 68 }, { "中和區", 61 }, { "新莊區", 58 }, { "三重區", 60 }, { "土城區", 55 }, { "林口區", 54 }, { "淡水區", 45 } } },
            { "基隆市", new Dictionary<string, double> { { "全市平均", 28 }, { "仁愛區", 30 }, { "中正區", 29 }, { "安樂區", 27 }, { "暖暖區", 24 }, { "七堵區", 23 } } },
            { "桃園市", new Dictionary<string, double> { { "全市平均", 42 }, { "桃園區", 48 }, { "中壢區", 46 }, { "蘆竹區", 44 }, { "龜山區", 43 }, { "八德區", 40 }, { "平鎮區", 39 }, { "楊梅區", 34 }, { "大園區", 36 } } },
            { "新竹市", new Dictionary<string, double> { { "全市平均", 50 }, { "東區", 56 }, { "北區", 46 }, { "香山區", 38 } } },
            { "新竹縣", new Dictionary<string, double> { { "全縣平均", 40 }, { "竹北市", 49 }, { "竹東鎮", 33 }, { "湖口鄉", 30 }, { "新豐鄉", 31 }, { "關西鎮", 24 } } },
            { "苗栗縣", new Dictionary<string, double> { { "全縣平均", 26 }, { "竹南鎮", 31 }, { "頭份市", 33 }, { "苗栗市", 27 }, { "苑裡鎮", 22 }, { "後龍鎮", 21 } } },
            { "台中市", new Dictionary<string, double> { { "全市平均", 42 }, { "西屯區", 56 }, { "南屯區", 51 }, { "北屯區", 46 }, { "西區", 48 }, { "南區", 40 }, { "東區", 37 }, { "太平區", 34 }, { "大里區", 35 }, { "沙鹿區", 30 } } },
            { "彰化縣", new Dictionary<string, double> { { "全縣平均", 25 }, { "彰化市", 30 }, { "員林市", 28 }, { "和美鎮", 24 }, { "鹿港鎮", 24 }, { "北斗鎮", 22 } } },
            { "南投縣", new Dictionary<string, double> { { "全縣平均", 23 }, { "南投市", 25 }, { "草屯鎮", 27 }, { "埔里鎮", 24 }, { "竹山鎮", 20 } } },
            { "雲林縣", new Dictionary<string, double> { { "全縣平均", 20 }, { "斗六市", 24 }, { "虎尾鎮", 22 }, { "北港鎮", 20 }, { "西螺鎮", 19 } } },
            { "嘉義市", new Dictionary<string, double> { { "全市平均", 24 }, { "東區", 25 }, { "西區", 23 } } },
            { "嘉義縣", new Dictionary<string, double> { { "全縣平均", 18 }, { "太保市", 20 }, { "朴子市", 19 }, { "民雄鄉", 21 }, { "大林鎮", 17 } } },
            { "台南市", new Dictionary<string, double> { { "全市平均", 34 }, { "東區", 41 }, { "永康區", 38 }, { "安平區", 39 }, { "北區", 35 }, { "中西區", 37 }, { "善化區", 35 }, { "新營區", 28 }, { "仁德區", 30 } } },
            { "高雄市", new Dictionary<string, double> { { "全市平均", 33 }, { "鼓山區", 45 }, { "左營區", 42 }, { "三民區", 34 }, { "楠梓區", 31 }, { "前鎮區", 36 }, { "苓雅區", 39 }, { "鳳山區", 32 }, { "岡山區", 29 } } },
            { "屏東縣", new Dictionary<string, double> { { "全縣平均", 19 }, { "屏東市", 23 }, { "潮州鎮", 19 }, { "東港鎮", 20 }, { "恆春鎮", 18 } } },
            { "宜蘭縣", new Dictionary<string, double> { { "全縣平均", 27 }, { "宜蘭市", 30 }, { "羅東鎮", 31 }, { "礁溪鄉", 29 }, { "頭城鎮", 24 } } },
            { "花蓮縣", new Dictionary<string, double> { { "全縣平均", 23 }, { "花蓮市", 27 }, { "吉安鄉", 22 }, { "新城鄉", 20 } } },
            { "台東縣", new Dictionary<string, double> { { "全縣平均", 18 }, { "台東市", 21 }, { "成功鎮", 16 }, { "關山鎮", 15 } } },
            { "澎湖縣", new Dictionary<string, double> { { "全縣平均", 21 }, { "馬公市", 24 }, { "湖西鄉", 19 }, { "白沙鄉", 18 } } },
            { "金門縣", new Dictionary<string, double> { { "全縣平均", 22 }, { "金城鎮", 25 }, { "金湖鎮", 21 }, { "金沙鎮", 19 } } },
            { "連江縣", new Dictionary<string, double> { { "全縣平均", 20 }, { "南竿鄉", 22 }, { "北竿鄉", 18 }, { "莒光鄉", 16 } } }
        };

        public Form1()
        {
            InitializeComponent();
            SetupCustomUI();
            InitializeAnimationEngine();
            InitializeDefaults();
        }

        private void TabControlHelper_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabControl tabControl = sender as TabControl;
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            TabPage page = tabControl.TabPages[e.Index];
            Rectangle r = e.Bounds;

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color bg = isSelected ? Color.White : Color.FromArgb(240, 240, 240);
            Color text = isSelected ? _customAccent : Color.Gray;
            Font f = isSelected ? new Font("微軟正黑體", 11.5f, FontStyle.Bold) : new Font("微軟正黑體", 11f, FontStyle.Regular);

            using (SolidBrush brush = new SolidBrush(bg))
            {
                g.FillRectangle(brush, r);
            }

            if (isSelected)
            {
                using (Pen p = new Pen(_customAccent, 3))
                {
                    g.DrawLine(p, r.X, r.Top + 1, r.Right - 1, r.Top + 1);
                }
            }

            TextRenderer.DrawText(g, page.Text, f, r, text, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private void InitializeAnimationEngine()
        {
            _resultAnimTimer = new Timer();
            _resultAnimTimer.Interval = 20;
            _resultAnimTimer.Tick += (s, e) =>
            {
                _animTick++;
                int ticks = Math.Max(1, _currentAnimTicks);
                double t = Math.Min(1.0, (double)_animTick / ticks);
                double ease = 1 - Math.Pow(1 - t, 3);

                foreach (var kv in _animTo)
                {
                    var label = kv.Key;
                    double from = _animFrom[label];
                    double to = kv.Value;
                    double value = from + ((to - from) * ease);
                    label.Text = "NT$ " + value.ToString("N2");
                }

                if (!_monthlyHasGrace)
                {
                    lblResultMonthly.Text = "NT$ " + (_monthlyNormalValue * ease).ToString("N2");
                }
                else
                {
                    lblResultMonthly.Text = "NT$ " + (_monthlyGraceValue * ease).ToString("N2") + " (寬限期後: " + _monthlyNormalValue.ToString("N2") + ")";
                }

                if (t >= 1.0)
                {
                    _resultAnimTimer.Stop();
                }
            };

            _uiMotionTimer = new Timer();
            _uiMotionTimer.Interval = 80;
            _uiMotionTimer.Tick += (s, e) =>
            {
                _chartDashOffset += 0.7f;
                if (_chartDashOffset > 12f) _chartDashOffset = 0f;
                if (picChart != null && picChart.Visible) picChart.Invalidate();
            };
            _uiMotionTimer.Start();
        }

        private void SetupCustomUI()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(244, 246, 249);
            this.MinimumSize = new Size(980, 640);
            this.Font = new Font("微軟正黑體", 11F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));
            this.AcceptButton = btnCalc;

            lblMainTitle.MouseDown += TitleBar_MouseDown;
            lblMainTitle.DoubleClick += (s, e) => ToggleWindowState();

            txtPrice.KeyPress += FloatOnly_KeyPress;
            txtDownPayment.KeyPress += FloatOnly_KeyPress;
            txtRate.KeyPress += FloatOnly_KeyPress;
            txtGrace.KeyPress += NumberOnly_KeyPress;

            // Auto-format for Price text
            txtPrice.Leave += (s, e) => { if(double.TryParse(txtPrice.Text.Replace(",", ""), out double d)) txtPrice.Text = d.ToString("N0"); };
            txtPrice.Enter += (s, e) => { txtPrice.Text = txtPrice.Text.Replace(",", ""); };

            cmbDownPaymentType.SelectedIndexChanged += (s, e) => {
                if(cmbDownPaymentType.SelectedIndex == 0) // %
                    txtDownPayment.Text = "20";
                else
                    txtDownPayment.Text = (GetDouble(txtPrice.Text) * 0.2).ToString("0");
            };

            dgvSchedule.DataSource = _bs;
            
            // Custom button paint
            btnCalc.Paint += FlatButton_Paint;
            btnReset.Paint += FlatButton_Paint;
            btnExport.Paint += FlatButton_Paint;
            picChart.MouseMove += picChart_MouseMove;
            picChart.MouseLeave += picChart_MouseLeave;
            picChart.MouseWheel += picChart_MouseWheel;
            picChart.MouseClick += picChart_MouseClick;
            picChart.MouseEnter += (s, e) => picChart.Focus();
            picChart.TabStop = true;
            ConfigureWindowControlButtons();
            btnClose.Click += (s,e) => Application.Exit();
            btnMinimize.Click += (s,e) => this.WindowState = FormWindowState.Minimized;
            this.Resize += (s, e) => LayoutTitleButtons();
            WireWindowButtonHover(btnClose, true);
            WireWindowButtonHover(btnMinimize, false);

            InitializeAdvancedToolbar();
            InitializeInputEnhancements();
            InitializeScenarioComparisonTab();
            ApplyModernLeftPanelDesign();
            LayoutTitleButtons();
            ApplyTheme("明亮");

            // Tab control modern aesthetics
            tabControlHelper.ItemSize = new Size(140, 36);
            tabControlHelper.SizeMode = TabSizeMode.Fixed;
            tabControlHelper.DrawMode = TabDrawMode.OwnerDrawFixed;
            tabControlHelper.DrawItem += TabControlHelper_DrawItem;

            _tips.SetToolTip(txtGrace, "寬限期必須小於貸款年限，例如年限20年，寬限期最多19年。");
            _tips.SetToolTip(btnExport, "匯出 CSV 並自動產生 SHA256 驗證檔。\n可用於資料完整性驗證。\n");
            _tips.SetToolTip(cmbTerm, "可直接輸入年限（5~50年），不受固定選項限制。");
            if (_btnApplyMarketPreset != null)
            {
                _tips.SetToolTip(_btnApplyMarketPreset, "套用台灣地區示意房價/利率預設值，方便快速試算。");
            }
            if (_numAnnualPrepay != null)
            {
                _tips.SetToolTip(_numAnnualPrepay, "每年固定多還本金，用於提前還款模擬。\n0 代表不啟用。\n");
            }
            if (_txtReportStudentId != null)
            {
                _tips.SetToolTip(_txtReportStudentId, "報告封面學號");
                _tips.SetToolTip(_txtReportStudentName, "報告封面姓名");
                _tips.SetToolTip(_txtReportCourse, "報告封面課程名稱");
                _tips.SetToolTip(_txtReportSchool, "報告封面學校");
                _tips.SetToolTip(_txtReportDepartment, "報告封面系所");
                _tips.SetToolTip(_txtReportAdvisor, "報告封面指導老師");
            }
            _tips.SetToolTip(_btnPdfTemplate, "目前：中文學術格式（IEEE風格章節）");
            _tips.SetToolTip(_btnMonteCarlo, "一鍵隨機壓測（蒙地卡羅），輸出平均/分位數月付結果。");
            _tips.SetToolTip(_btnInputMode, "切換基本/進階輸入面板。");
            _tips.SetToolTip(_btnAnimStrength, "切換動畫強度：高 / 低 / 關（低配可關）。");
            if (_txtDistrictSearch != null)
            {
                _tips.SetToolTip(_txtDistrictSearch, "快速搜尋行政區，例如：中壢、永和、東區");
            }

            rtbAI.Text = "未執行分析...\n提示：按「開始深度試算」後將顯示 AI 財務建議。";

            ApplyModernLeftPanelDesign();
        }

        private void ApplyModernLeftPanelDesign()
        {
            gbInput.Text = "⚙️ 參數設定與智慧決策區";

            var flatControls = new Control[] { txtPrice, txtDownPayment, txtRate, txtGrace, cmbTerm, cmbDownPaymentType, 
                _cmbRegion, _cmbDistrict, _cmbPresetType, _txtDistrictSearch, _numPing, _numAnnualPrepay, _numMonthlyIncome, 
                _txtReportStudentId, _txtReportStudentName, _txtReportCourse, _txtReportSchool, _txtReportDepartment, _txtReportAdvisor };

            foreach (var c in flatControls)
            {
                if (c == null) continue;
                if (c is TextBox tb)
                {
                    tb.BorderStyle = BorderStyle.FixedSingle;
                    tb.BackColor = Color.FromArgb(248, 250, 252);
                    tb.Font = new Font("Consolas", 11.5F, FontStyle.Bold);
                }
                else if (c is ComboBox cb)
                {
                    cb.FlatStyle = FlatStyle.Flat;
                    cb.BackColor = Color.FromArgb(248, 250, 252);
                    cb.Font = new Font("微軟正黑體", 10.5F, FontStyle.Bold);
                }
                else if (c is NumericUpDown nud)
                {
                    nud.BorderStyle = BorderStyle.FixedSingle;
                    nud.BackColor = Color.FromArgb(248, 250, 252);
                    nud.Font = new Font("Consolas", 10.5F, FontStyle.Bold);
                }
            }

            if (_btnEstimatePrice != null)
            {
                _btnEstimatePrice.BackColor = Color.FromArgb(52, 152, 219);
                _btnEstimatePrice.ForeColor = Color.White;
                _btnEstimatePrice.FlatAppearance.BorderSize = 0;
                _btnEstimatePrice.Font = new Font("微軟正黑體", 9F, FontStyle.Bold);
                _btnEstimatePrice.Paint += FlatButton_Paint;
            }
            if (_btnApplyMarketPreset != null)
            {
                _btnApplyMarketPreset.BackColor = Color.FromArgb(46, 204, 113);
                _btnApplyMarketPreset.ForeColor = Color.White;
                _btnApplyMarketPreset.FlatAppearance.BorderSize = 0;
                _btnApplyMarketPreset.Font = new Font("微軟正黑體", 9F, FontStyle.Bold);
                _btnApplyMarketPreset.Paint += FlatButton_Paint;
            }

            for (int i = 0; i < tableLayoutPanelInput.RowStyles.Count; i++)
            {
                if (tableLayoutPanelInput.RowStyles[i].SizeType == SizeType.Absolute && tableLayoutPanelInput.RowStyles[i].Height == 50F)
                {
                    tableLayoutPanelInput.RowStyles[i].Height = 58F;
                }
            }

            foreach (Control c in tableLayoutPanelInput.Controls)
            {
                if (c is Label lbl && lbl.TextAlign == ContentAlignment.MiddleRight)
                {
                    lbl.Font = new Font("微軟正黑體", 10.5F, FontStyle.Bold);
                    lbl.ForeColor = Color.FromArgb(60, 80, 100);
                }
            }

            btnCalc.Font = new Font("微軟正黑體", 14F, FontStyle.Bold);
            btnReset.Font = new Font("微軟正黑體", 11F, FontStyle.Bold);
            btnExport.Font = new Font("微軟正黑體", 11F, FontStyle.Bold);

            pnlButtons.Padding = new Padding(0, 5, 0, 0);
            int btnW = 210;
            btnCalc.Size = new Size(btnW, 46);
            btnCalc.Margin = new Padding(3, 5, 3, 8);

            btnReset.Size = new Size((btnW / 2) - 4, 36);
            btnReset.Margin = new Padding(3, 0, 2, 5);

            btnExport.Size = new Size((btnW / 2) - 3, 36);
            btnExport.Margin = new Padding(1, 0, 3, 5);

            WireActionBtnHover(btnCalc, btnCalc.BackColor);
            WireActionBtnHover(btnReset, btnReset.BackColor);
            WireActionBtnHover(btnExport, btnExport.BackColor);

            lblValidationHint.Padding = new Padding(5, 10, 0, 0);
        }

        private void WireActionBtnHover(Button btn, Color baseColor)
        {
            btn.MouseEnter += (s, e) => { btn.BackColor = ControlPaint.Light(baseColor, 0.15f); btn.Invalidate(); };
            btn.MouseLeave += (s, e) => { btn.BackColor = baseColor; btn.Invalidate(); };
            btn.MouseDown += (s, e) => { btn.BackColor = ControlPaint.Dark(baseColor, 0.1f); btn.Invalidate(); };
            btn.MouseUp += (s, e) => { btn.BackColor = ControlPaint.Light(baseColor, 0.15f); btn.Invalidate(); };
        }

        private void InitializeScenarioComparisonTab()
        {
            _tabCompare = new TabPage("📌 情境比較(最多3套)");
            _tabCompare.Padding = new Padding(12);
            _tabCompare.BackColor = Color.White;

            var topPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 48,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 5, 0, 5)
            };

            _btnAddScenario = new Button { Text = "＋ 加入目前方案", Width = 140, Height = 36, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            _btnAddScenario.FlatAppearance.BorderSize = 0;
            _btnAddScenario.BackColor = Color.FromArgb(41, 128, 185);
            _btnAddScenario.ForeColor = Color.White;
            _btnAddScenario.Font = new Font("微軟正黑體", 10F, FontStyle.Bold);
            _btnAddScenario.Paint += FlatButton_Paint;

            _btnAutoScenario = new Button { Text = "⚡ 自動生成 3 方案", Width = 160, Height = 36, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            _btnAutoScenario.FlatAppearance.BorderSize = 0;
            _btnAutoScenario.BackColor = Color.FromArgb(39, 174, 96);
            _btnAutoScenario.ForeColor = Color.White;
            _btnAutoScenario.Font = new Font("微軟正黑體", 10F, FontStyle.Bold);
            _btnAutoScenario.Paint += FlatButton_Paint;

            _btnClearScenario = new Button { Text = "🗑 清空比較", Width = 120, Height = 36, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            _btnClearScenario.FlatAppearance.BorderSize = 0;
            _btnClearScenario.BackColor = Color.FromArgb(149, 165, 166);
            _btnClearScenario.ForeColor = Color.White;
            _btnClearScenario.Font = new Font("微軟正黑體", 10F, FontStyle.Bold);
            _btnClearScenario.Paint += FlatButton_Paint;

            _btnAddScenario.Margin = new Padding(0, 0, 10, 0);
            _btnAutoScenario.Margin = new Padding(0, 0, 10, 0);

            _btnAddScenario.Click += (s, e) => AddCurrentScenario();
            _btnAutoScenario.Click += (s, e) => AutoGenerateScenarios();
            _btnClearScenario.Click += (s, e) =>
            {
                _scenarios.Clear();
                _compareBs.ResetBindings(false);
            };

            topPanel.Controls.Add(_btnAddScenario);
            topPanel.Controls.Add(_btnAutoScenario);
            topPanel.Controls.Add(_btnClearScenario);

            _dgvCompare = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(235, 235, 235),
                RowTemplate = { Height = 38 },
                ColumnHeadersHeight = 45,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            };

            _compareBs.DataSource = _scenarios;
            _dgvCompare.DataSource = _compareBs;

            _tabCompare.Controls.Add(_dgvCompare);
            _tabCompare.Controls.Add(topPanel);
            tabControlHelper.TabPages.Add(_tabCompare);

            // Add custom cell formatting to style values and highlight best ones if needed
            _dgvCompare.CellFormatting += (s, e) =>
            {
                if (e.Value != null && e.RowIndex % 2 == 0) e.CellStyle.BackColor = Color.FromArgb(250, 251, 252);
            };
        }

        private void InitializeInputEnhancements()
        {
            cmbTerm.DropDownStyle = ComboBoxStyle.DropDown;
            cmbTerm.Items.Clear();
            for (int year = 5; year <= 50; year++)
            {
                cmbTerm.Items.Add(year.ToString());
            }

            // 在輸入區新增「台灣區域預設」(程式內建樣本資料)
            var controls = tableLayoutPanelInput.Controls.Cast<Control>().ToList();
            foreach (var c in controls)
            {
                int row = tableLayoutPanelInput.GetRow(c);
                tableLayoutPanelInput.SetRow(c, row + 1);
            }

            tableLayoutPanelInput.RowCount += 1;
            tableLayoutPanelInput.RowStyles.Insert(0, new RowStyle(SizeType.Absolute, 100F));

            var lblPreset = new Label
            {
                Text = "區域預設",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight
            };

            var presetPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true
            };

            // (1) 重新調整寬度以減少換行太擠
            _cmbRegion = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 72, Font = new Font("微軟正黑體", 9F) };
            _cmbDistrict = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 72, Font = new Font("微軟正黑體", 9F) };
            _txtDistrictSearch = new TextBox { Width = 64, Font = new Font("微軟正黑體", 9F) };
            _cmbPresetType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 52, Font = new Font("微軟正黑體", 9F) };
            _numPing = new NumericUpDown { Width = 48, Height = 24, DecimalPlaces = 1, Minimum = 8, Maximum = 250, Value = 30, Font = new Font("微軟正黑體", 9F), Increment = 0.5M };
            _btnEstimatePrice = new Button { Text = "估價", Width = 40, Height = 25, FlatStyle = FlatStyle.Flat, Font = new Font("微軟正黑體", 8.5F, FontStyle.Bold) };
            _btnApplyMarketPreset = new Button { Text = "套用", Width = 40, Height = 25, FlatStyle = FlatStyle.Flat, Font = new Font("微軟正黑體", 8.5F, FontStyle.Bold) };
            _lblPresetSource = new Label { AutoSize = true, Margin = new Padding(5, 6, 3, 3), Font = new Font("微軟正黑體", 8.5F), ForeColor = Color.Gray, Text = "資料來源：台灣公開統計樣本(示意)" };

            _cmbRegion.Items.AddRange(_countyUnitPriceWanPerPing.Keys.OrderBy(x => x).ToArray());
            _cmbPresetType.Items.AddRange(new object[] { "大樓", "公寓", "透天", "華廈" });
            _cmbRegion.SelectedIndex = 1;
            _cmbPresetType.SelectedIndex = 0;

            _cmbRegion.SelectedIndexChanged += (s, e) => PopulateDistricts();
            _btnEstimatePrice.Click += (s, e) => EstimateByDistrictAndPing();
            _btnApplyMarketPreset.Click += (s, e) => ApplyMarketPreset();
            _txtDistrictSearch.TextChanged += (s, e) => ApplyDistrictFilter();
            EnsureAllTaiwanDistricts();
            PopulateDistricts();

            presetPanel.Controls.Add(_cmbRegion);
            presetPanel.Controls.Add(_cmbDistrict);
            presetPanel.Controls.Add(_txtDistrictSearch);
            presetPanel.Controls.Add(_cmbPresetType);
            presetPanel.Controls.Add(_numPing);
            presetPanel.Controls.Add(_btnEstimatePrice);
            presetPanel.Controls.Add(_btnApplyMarketPreset);
            presetPanel.Controls.Add(_lblPresetSource);

            tableLayoutPanelInput.Controls.Add(lblPreset, 0, 0);
            tableLayoutPanelInput.Controls.Add(presetPanel, 1, 0);

            // 新增：提前還款模擬器（每年固定多還X元）
            var controls2 = tableLayoutPanelInput.Controls.Cast<Control>().ToList();
            foreach (var c2 in controls2)
            {
                int r = tableLayoutPanelInput.GetRow(c2);
                if (r >= 6) tableLayoutPanelInput.SetRow(c2, r + 1);
            }
            tableLayoutPanelInput.RowCount += 1;
            tableLayoutPanelInput.RowStyles.Insert(6, new RowStyle(SizeType.Absolute, 50F));

            var lblPrepay = new Label
            {
                Text = "年提前還款($)",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight
            };

            var prepayPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };

            _numAnnualPrepay = new NumericUpDown
            {
                Width = 120,
                Height = 27,
                Font = new Font("微軟正黑體", 10F),
                Maximum = 5000000,
                Minimum = 0,
                Increment = 10000,
                ThousandsSeparator = true,
                Value = 0
            };
            var lblUnit = new Label { Text = "元 / 年", AutoSize = true, Margin = new Padding(8, 6, 3, 3), ForeColor = Color.Gray };
            prepayPanel.Controls.Add(_numAnnualPrepay);
            prepayPanel.Controls.Add(lblUnit);

            tableLayoutPanelInput.Controls.Add(lblPrepay, 0, 6);
            tableLayoutPanelInput.Controls.Add(prepayPanel, 1, 6);

            // 新增：進階設定 (新青安與月收入)
            var controlsDti = tableLayoutPanelInput.Controls.Cast<Control>().ToList();
            foreach (var c in controlsDti)
            {
                int r = tableLayoutPanelInput.GetRow(c);
                if (r >= 7) tableLayoutPanelInput.SetRow(c, r + 1);
            }
            tableLayoutPanelInput.RowCount += 1;
            tableLayoutPanelInput.RowStyles.Insert(7, new RowStyle(SizeType.Absolute, 50F));

            var lblAdvFeatures = new Label
            {
                Text = "額外設定",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight
            };

            var advPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true
            };

            _chkNewYouth = new CheckBox { Text = "新青安試算", AutoSize = true, Margin = new Padding(3, 4, 10, 3) };

            var lblMonthlyIncome = new Label { Text = "月薪:", AutoSize = true, Margin = new Padding(3, 5, 2, 3) };
            _numMonthlyIncome = new NumericUpDown
            {
                Width = 100, Height = 27, Font = new Font("微軟正黑體", 9F),
                Maximum = 10000000, Minimum = 0, Increment = 5000, ThousandsSeparator = true, Value = 0
            };

            advPanel.Controls.Add(_chkNewYouth);
            advPanel.Controls.Add(lblMonthlyIncome);
            advPanel.Controls.Add(_numMonthlyIncome);

            tableLayoutPanelInput.Controls.Add(lblAdvFeatures, 0, 7);
            tableLayoutPanelInput.Controls.Add(advPanel, 1, 7);

            // 新增：報告封面資訊（學號/姓名/課程）
            var controls3 = tableLayoutPanelInput.Controls.Cast<Control>().ToList();
            foreach (var c3 in controls3)
            {
                int r3 = tableLayoutPanelInput.GetRow(c3);
                if (r3 >= 7) tableLayoutPanelInput.SetRow(c3, r3 + 1);
            }
            tableLayoutPanelInput.RowCount += 1;
            tableLayoutPanelInput.RowStyles.Insert(7, new RowStyle(SizeType.Absolute, 94F));

            var lblReportInfo = new Label
            {
                Text = "報告資訊",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight
            };

            var reportPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true
            };

            _txtReportStudentId = new TextBox { Width = 72, Font = new Font("微軟正黑體", 9F), Text = "1113354" };
            _txtReportStudentName = new TextBox { Width = 82, Font = new Font("微軟正黑體", 9F), Text = "陳冠瑋" };
            _txtReportCourse = new TextBox { Width = 116, Font = new Font("微軟正黑體", 9F), Text = "程式設計實務" };
            _txtReportSchool = new TextBox { Width = 104, Font = new Font("微軟正黑體", 9F), Text = "國立OO大學" };
            _txtReportDepartment = new TextBox { Width = 104, Font = new Font("微軟正黑體", 9F), Text = "資訊工程學系" };
            _txtReportAdvisor = new TextBox { Width = 88, Font = new Font("微軟正黑體", 9F), Text = "王OO教授" };
            _dtReportDate = new DateTimePicker { Width = 116, Font = new Font("微軟正黑體", 9F), Format = DateTimePickerFormat.Short, Value = DateTime.Now };

            reportPanel.Controls.Add(_txtReportStudentId);
            reportPanel.Controls.Add(_txtReportStudentName);
            reportPanel.Controls.Add(_txtReportCourse);
            reportPanel.Controls.Add(_txtReportSchool);
            reportPanel.Controls.Add(_txtReportDepartment);
            reportPanel.Controls.Add(_txtReportAdvisor);
            reportPanel.Controls.Add(_dtReportDate);

            tableLayoutPanelInput.Controls.Add(lblReportInfo, 0, 7);
            tableLayoutPanelInput.Controls.Add(reportPanel, 1, 7);
        }

        private void PopulateDistricts()
        {
            if (_cmbRegion == null || _cmbDistrict == null) return;
            string county = _cmbRegion.SelectedItem?.ToString();
            _cmbDistrict.Items.Clear();

            if (string.IsNullOrWhiteSpace(county)) return;

            if (_districtUnitPriceWanPerPing.TryGetValue(county, out var districts))
            {
                var names = districts.Keys.OrderBy(x => x).ToList();
                _districtNameCache[county] = names;
                _cmbDistrict.Items.AddRange(names.ToArray());
            }
            else
            {
                _cmbDistrict.Items.Add("全縣市平均");
            }

            if (_cmbDistrict.Items.Count > 0)
                _cmbDistrict.SelectedIndex = 0;

            if (_txtDistrictSearch != null)
            {
                _txtDistrictSearch.Text = string.Empty;
            }
        }

        private void ApplyDistrictFilter()
        {
            if (_cmbRegion == null || _cmbDistrict == null || _txtDistrictSearch == null) return;
            string county = _cmbRegion.SelectedItem?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(county)) return;

            if (!_districtNameCache.TryGetValue(county, out var all))
            {
                if (_districtUnitPriceWanPerPing.TryGetValue(county, out var districts))
                    all = districts.Keys.OrderBy(x => x).ToList();
                else
                    all = new List<string>();
                _districtNameCache[county] = all;
            }

            string keyword = (_txtDistrictSearch.Text ?? string.Empty).Trim();
            var filtered = string.IsNullOrWhiteSpace(keyword)
                ? all
                : all.Where(x => x.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

            string current = _cmbDistrict.SelectedItem?.ToString();
            _cmbDistrict.BeginUpdate();
            _cmbDistrict.Items.Clear();
            _cmbDistrict.Items.AddRange(filtered.ToArray());
            if (_cmbDistrict.Items.Count > 0)
            {
                int idx = !string.IsNullOrWhiteSpace(current) ? _cmbDistrict.Items.IndexOf(current) : -1;
                _cmbDistrict.SelectedIndex = idx >= 0 ? idx : 0;
            }
            _cmbDistrict.EndUpdate();
        }

        private void EnsureAllTaiwanDistricts()
        {
            var full = new Dictionary<string, string>
            {
                { "台北市", "中正區、中山區、松山區、大安區、萬華區、信義區、士林區、北投區、內湖區、南港區、文山區、大同區" },
                { "新北市", "板橋區、三重區、中和區、永和區、新莊區、新店區、樹林區、鶯歌區、三峽區、淡水區、汐止區、瑞芳區、土城區、蘆洲區、五股區、泰山區、林口區、深坑區、石碇區、坪林區、三芝區、石門區、八里區、平溪區、雙溪區、貢寮區、金山區、萬里區、烏來區" },
                { "基隆市", "仁愛區、信義區、中正區、中山區、安樂區、暖暖區、七堵區" },
                { "桃園市", "桃園區、中壢區、大溪區、楊梅區、蘆竹區、大園區、龜山區、八德區、龍潭區、平鎮區、新屋區、觀音區、復興區" },
                { "新竹市", "東區、北區、香山區" },
                { "新竹縣", "竹北市、竹東鎮、新埔鎮、關西鎮、湖口鄉、新豐鄉、芎林鄉、橫山鄉、北埔鄉、寶山鄉、峨眉鄉、尖石鄉、五峰鄉" },
                { "苗栗縣", "苗栗市、苑裡鎮、通霄鎮、竹南鎮、頭份市、後龍鎮、卓蘭鎮、大湖鄉、公館鄉、銅鑼鄉、南庄鄉、頭屋鄉、三義鄉、西湖鄉、造橋鄉、三灣鄉、獅潭鄉、泰安鄉" },
                { "台中市", "中區、東區、南區、西區、北區、西屯區、南屯區、北屯區、豐原區、東勢區、大甲區、清水區、沙鹿區、梧棲區、后里區、神岡區、潭子區、大雅區、新社區、石岡區、外埔區、大安區、烏日區、大肚區、龍井區、霧峰區、太平區、大里區、和平區" },
                { "彰化縣", "彰化市、鹿港鎮、和美鎮、線西鄉、伸港鄉、福興鄉、秀水鄉、花壇鄉、芬園鄉、員林市、溪湖鎮、田中鎮、大村鄉、埔鹽鄉、埔心鄉、永靖鄉、社頭鄉、二水鄉、北斗鎮、二林鎮、田尾鄉、埤頭鄉、芳苑鄉、大城鄉、竹塘鄉、溪州鄉" },
                { "南投縣", "南投市、埔里鎮、草屯鎮、竹山鎮、集集鎮、名間鄉、鹿谷鄉、中寮鄉、魚池鄉、國姓鄉、水里鄉、信義鄉、仁愛鄉" },
                { "雲林縣", "斗六市、斗南鎮、虎尾鎮、西螺鎮、土庫鎮、北港鎮、古坑鄉、大埤鄉、莿桐鄉、林內鄉、二崙鄉、崙背鄉、麥寮鄉、東勢鄉、褒忠鄉、臺西鄉、元長鄉、四湖鄉、口湖鄉、水林鄉" },
                { "嘉義市", "東區、西區" },
                { "嘉義縣", "太保市、朴子市、布袋鎮、大林鎮、民雄鄉、溪口鄉、新港鄉、六腳鄉、東石鄉、義竹鄉、鹿草鄉、水上鄉、中埔鄉、竹崎鄉、梅山鄉、番路鄉、大埔鄉、阿里山鄉" },
                { "台南市", "中西區、東區、南區、北區、安平區、安南區、永康區、歸仁區、新化區、左鎮區、玉井區、楠西區、南化區、仁德區、關廟區、龍崎區、官田區、麻豆區、佳里區、西港區、七股區、將軍區、學甲區、北門區、新營區、後壁區、白河區、東山區、六甲區、下營區、柳營區、鹽水區、善化區、大內區、山上區、新市區、安定區" },
                { "高雄市", "新興區、前金區、苓雅區、鹽埕區、鼓山區、旗津區、前鎮區、三民區、楠梓區、小港區、左營區、仁武區、大社區、東沙群島、南沙群島、岡山區、路竹區、阿蓮區、田寮區、燕巢區、橋頭區、梓官區、彌陀區、永安區、湖內區、鳳山區、大寮區、林園區、鳥松區、大樹區、旗山區、美濃區、六龜區、內門區、杉林區、甲仙區、桃源區、那瑪夏區、茂林區、茄萣區" },
                { "屏東縣", "屏東市、潮州鎮、東港鎮、恆春鎮、萬丹鄉、長治鄉、麟洛鄉、九如鄉、里港鄉、鹽埔鄉、高樹鄉、萬巒鄉、內埔鄉、竹田鄉、新埤鄉、枋寮鄉、新園鄉、崁頂鄉、林邊鄉、南州鄉、佳冬鄉、琉球鄉、車城鄉、滿州鄉、枋山鄉、三地門鄉、霧臺鄉、瑪家鄉、泰武鄉、來義鄉、春日鄉、獅子鄉、牡丹鄉" },
                { "宜蘭縣", "宜蘭市、羅東鎮、蘇澳鎮、頭城鎮、礁溪鄉、壯圍鄉、員山鄉、冬山鄉、五結鄉、三星鄉、大同鄉、南澳鄉" },
                { "花蓮縣", "花蓮市、鳳林鎮、玉里鎮、新城鄉、吉安鄉、壽豐鄉、光復鄉、豐濱鄉、瑞穗鄉、富里鄉、秀林鄉、萬榮鄉、卓溪鄉" },
                { "台東縣", "臺東市、成功鎮、關山鎮、卑南鄉、大武鄉、太麻里鄉、東河鄉、長濱鄉、鹿野鄉、池上鄉、綠島鄉、延平鄉、海端鄉、達仁鄉、金峰鄉、蘭嶼鄉" },
                { "澎湖縣", "馬公市、湖西鄉、白沙鄉、西嶼鄉、望安鄉、七美鄉" },
                { "金門縣", "金城鎮、金湖鎮、金沙鎮、金寧鄉、烈嶼鄉、烏坵鄉" },
                { "連江縣", "南竿鄉、北竿鄉、莒光鄉、東引鄉" }
            };

            foreach (var kv in full)
            {
                string county = kv.Key;
                if (!_districtUnitPriceWanPerPing.TryGetValue(county, out var districtMap))
                {
                    districtMap = new Dictionary<string, double>();
                    _districtUnitPriceWanPerPing[county] = districtMap;
                }

                double countyAvg = _countyUnitPriceWanPerPing.TryGetValue(county, out var cAvg) ? cAvg : 35;
                string avgKey = county.EndsWith("市") ? "全市平均" : "全縣平均";
                if (!districtMap.ContainsKey(avgKey)) districtMap[avgKey] = countyAvg;

                string[] names = kv.Value.Split(new[] { '、' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < names.Length; i++)
                {
                    string d = names[i].Trim();
                    if (string.IsNullOrWhiteSpace(d)) continue;
                    if (!districtMap.ContainsKey(d))
                    {
                        districtMap[d] = EstimateDistrictUnitPrice(d, countyAvg);
                    }
                }
            }
        }

        private double EstimateDistrictUnitPrice(string district, double countyAvg)
        {
            double factor = 1.0;
            if (district.EndsWith("市")) factor = 1.06;
            else if (district.EndsWith("區")) factor = 1.00;
            else if (district.EndsWith("鎮")) factor = 0.95;
            else if (district.EndsWith("鄉")) factor = 0.88;

            if (district.Contains("東") || district.Contains("南") || district.Contains("北") || district.Contains("西"))
                factor += 0.02;
            if (district.Contains("山") || district.Contains("海") || district.Contains("蘭嶼") || district.Contains("綠島"))
                factor -= 0.05;

            return Math.Round(Math.Max(12, countyAvg * factor), 1);
        }

        private double ResolveUnitPriceBySelection()
        {
            string county = _cmbRegion.SelectedItem?.ToString() ?? "";
            string district = _cmbDistrict.SelectedItem?.ToString() ?? "全縣市平均";

            double unit = _countyUnitPriceWanPerPing.TryGetValue(county, out var countyPrice) ? countyPrice : 35;
            if (_districtUnitPriceWanPerPing.TryGetValue(county, out var districts) && districts.TryGetValue(district, out var districtPrice))
            {
                unit = districtPrice;
            }

            string type = _cmbPresetType.SelectedItem?.ToString() ?? "大樓";
            double factor = 1.0;
            if (type == "公寓") factor = 0.90;
            else if (type == "透天") factor = 1.15;
            else if (type == "華廈") factor = 0.96;

            return unit * factor;
        }

        private void EstimateByDistrictAndPing()
        {
            double unit = ResolveUnitPriceBySelection(); // 萬/坪
            double ping = (double)_numPing.Value;
            double total = unit * ping * 10000;

            txtPrice.Text = total.ToString("N0");
            _lblPresetSource.Text = string.Format("估價：{0}/{1}/{2}坪 ≈ {3:N1}萬", _cmbRegion.Text, _cmbDistrict.Text, ping.ToString("0.0"), unit * ping);
        }

        private void ApplyMarketPreset()
        {
            string key = _cmbRegion.SelectedItem + "-" + _cmbPresetType.SelectedItem;
            if (_marketPresets.TryGetValue(key, out MarketPreset preset))
            {
                txtPrice.Text = preset.AvgTotalPrice.ToString("N0");
                txtRate.Text = preset.AvgRate.ToString("0.00");
                cmbTerm.Text = preset.SuggestedTerm.ToString();
                _lblPresetSource.Text = string.Format("已套用：{0}（示意統計）", key);
            }
            else
            {
                EstimateByDistrictAndPing();
                txtRate.Text = "2.15";
                if (string.IsNullOrWhiteSpace(cmbTerm.Text)) cmbTerm.Text = "30";
                _lblPresetSource.Text += "｜已套用房貸參數";
            }
        }

        private void InitializeAdvancedToolbar()
        {
            _btnMaximize = new Label
            {
                Size = new Size(34, 30),
                Text = "□",
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(20, 0, 0, 0),
                BorderStyle = BorderStyle.FixedSingle
            };
            _btnMaximize.Click += (s, e) => ToggleWindowState();
            this.Controls.Add(_btnMaximize);
            _btnMaximize.BringToFront();
            WireWindowButtonHover(_btnMaximize, false);

            _cmbTheme = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(90, 24)
            };
            _cmbTheme.Items.AddRange(new object[] { "明亮", "夜間", "自定" });
            _cmbTheme.SelectedIndex = 0;
            _cmbTheme.SelectedIndexChanged += (s, e) => ApplyTheme(_cmbTheme.SelectedItem.ToString());

            _cmbScale = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(80, 24)
            };
            _cmbScale.Items.AddRange(new object[] { "100%", "110%", "125%", "140%" });
            _cmbScale.SelectedIndex = 0;
            _cmbScale.SelectedIndexChanged += (s, e) =>
            {
                float target = 1f;
                string text = _cmbScale.SelectedItem.ToString().Replace("%", "");
                if (float.TryParse(text, out float percent))
                {
                    target = percent / 100f;
                }
                ApplyScale(target);
            };

            _btnCustomTheme = new Button
            {
                Text = "色彩",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(52, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnCustomTheme.FlatAppearance.BorderColor = Color.LightGray;
            _btnCustomTheme.Click += (s, e) =>
            {
                using (var dlg = new ColorDialog())
                {
                    dlg.Color = _customAccent;
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        _customAccent = dlg.Color;
                        _cmbTheme.SelectedItem = "自定";
                        ApplyTheme("自定");
                    }
                }
            };

            _btnStressTest = new Button
            {
                Text = "壓測",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(52, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnStressTest.FlatAppearance.BorderColor = Color.LightGray;
            _btnStressTest.Click += (s, e) => RunStressTest();

            _btnMonteCarlo = new Button
            {
                Text = "隨機壓測",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(72, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnMonteCarlo.FlatAppearance.BorderColor = Color.LightGray;
            _btnMonteCarlo.Click += (s, e) => RunMonteCarloStressTest();

            _btnCopySummary = new Button
            {
                Text = "複製",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(52, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnCopySummary.FlatAppearance.BorderColor = Color.LightGray;
            _btnCopySummary.Click += (s, e) => CopySummaryToClipboard();

            _btnExportPdf = new Button
            {
                Text = "PDF",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(52, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnExportPdf.FlatAppearance.BorderColor = Color.LightGray;
            _btnExportPdf.Click += (s, e) => ExportPdfReport();

            _btnPdfTemplate = new Button
            {
                Text = "學術",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(60, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnPdfTemplate.FlatAppearance.BorderColor = Color.LightGray;
            _btnPdfTemplate.Click += (s, e) => TogglePdfTemplateMode();

            _btnInputMode = new Button
            {
                Text = "進階",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(56, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnInputMode.FlatAppearance.BorderColor = Color.LightGray;
            _btnInputMode.Click += (s, e) => ToggleInputMode();

            _btnAnimStrength = new Button
            {
                Text = "動畫:高",
                Font = new Font("微軟正黑體", 9F),
                Size = new Size(78, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White
            };
            _btnAnimStrength.FlatAppearance.BorderColor = Color.LightGray;
            _btnAnimStrength.Click += (s, e) => ToggleAnimationStrength();

            lblMainTitle.Controls.Add(_cmbTheme);
            lblMainTitle.Controls.Add(_cmbScale);
            lblMainTitle.Controls.Add(_btnCustomTheme);
            lblMainTitle.Controls.Add(_btnStressTest);
            lblMainTitle.Controls.Add(_btnMonteCarlo);
            lblMainTitle.Controls.Add(_btnCopySummary);
            lblMainTitle.Controls.Add(_btnInputMode);
            lblMainTitle.Controls.Add(_btnAnimStrength);
            lblMainTitle.Controls.Add(_btnPdfTemplate);
            lblMainTitle.Controls.Add(_btnExportPdf);
        }

        private void WireWindowButtonHover(Label button, bool isClose)
        {
            button.MouseEnter += (s, e) =>
            {
                if (isClose)
                {
                    button.BackColor = Color.FromArgb(220, 231, 76, 60);
                    button.ForeColor = Color.White;
                }
                else
                {
                    button.BackColor = _windowBtnHoverBack;
                    button.ForeColor = _windowBtnFore;
                }
            };

            button.MouseLeave += (s, e) =>
            {
                button.BackColor = _windowBtnNormalBack;
                button.ForeColor = _windowBtnFore;
            };
        }

        private void LayoutTitleButtons()
        {
            int top = 14;
            int right = this.ClientSize.Width - 8;

            btnClose.Location = new Point(right - btnClose.Width, top);
            right -= (btnClose.Width + 6);

            _btnMaximize.Location = new Point(right - _btnMaximize.Width, top);
            right -= (_btnMaximize.Width + 6);

            btnMinimize.Location = new Point(right - btnMinimize.Width, top);

            _btnExportPdf.Location = new Point(this.ClientSize.Width - 420, 18);
            _btnPdfTemplate.Location = new Point(this.ClientSize.Width - 486, 18);
            _btnAnimStrength.Location = new Point(this.ClientSize.Width - 570, 18);
            _btnInputMode.Location = new Point(this.ClientSize.Width - 632, 18);
            _btnCopySummary.Location = new Point(this.ClientSize.Width - 690, 18);
            _btnMonteCarlo.Location = new Point(this.ClientSize.Width - 768, 18);
            _btnStressTest.Location = new Point(this.ClientSize.Width - 826, 18);
            _btnCustomTheme.Location = new Point(this.ClientSize.Width - 884, 18);
            _cmbScale.Location = new Point(this.ClientSize.Width - 969, 18);
            _cmbTheme.Location = new Point(this.ClientSize.Width - 1064, 18);
        }

        private void ToggleWindowState()
        {
            this.WindowState = this.WindowState == FormWindowState.Maximized
                ? FormWindowState.Normal
                : FormWindowState.Maximized;
            _btnMaximize.Text = this.WindowState == FormWindowState.Maximized ? "❐" : "□";
        }

        private void ConfigureWindowControlButtons()
        {
            btnClose.Size = new Size(34, 30);
            btnMinimize.Size = new Size(34, 30);
            btnClose.Text = "×";
            btnMinimize.Text = "—";
            btnClose.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            btnMinimize.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            btnClose.BorderStyle = BorderStyle.FixedSingle;
            btnMinimize.BorderStyle = BorderStyle.FixedSingle;
            btnClose.BackColor = Color.FromArgb(20, 0, 0, 0);
            btnMinimize.BackColor = Color.FromArgb(20, 0, 0, 0);
            btnClose.ForeColor = Color.White;
            btnMinimize.ForeColor = Color.White;
        }

        private void TogglePdfTemplateMode()
        {
            _pdfTemplateMode = _pdfTemplateMode == PdfTemplateMode.AcademicZh
                ? PdfTemplateMode.BusinessBrief
                : PdfTemplateMode.AcademicZh;

            _btnPdfTemplate.Text = _pdfTemplateMode == PdfTemplateMode.AcademicZh ? "學術" : "商務";
            _tips.SetToolTip(_btnPdfTemplate, _pdfTemplateMode == PdfTemplateMode.AcademicZh
                ? "目前：中文學術格式（IEEE風格章節）"
                : "目前：商務簡報格式");
        }

        private void ApplyScale(float targetScale)
        {
            if (targetScale <= 0.5f || targetScale > 2.0f) return;
            float ratio = targetScale / _currentScale;
            this.SuspendLayout();
            this.Scale(new SizeF(ratio, ratio));
            this.ResumeLayout();
            _currentScale = targetScale;
            LayoutTitleButtons();
        }

        private void ApplyTheme(string mode)
        {
            Color bg;
            Color surface;
            Color text;
            Color accent;

            if (mode == "夜間")
            {
                bg = Color.FromArgb(28, 32, 38);
                surface = Color.FromArgb(38, 44, 52);
                text = Color.Gainsboro;
                accent = Color.FromArgb(88, 166, 255);
            }
            else if (mode == "自定")
            {
                bg = Color.FromArgb(245, 247, 250);
                surface = Color.White;
                text = Color.FromArgb(44, 62, 80);
                accent = _customAccent;
            }
            else
            {
                bg = Color.FromArgb(244, 246, 249);
                surface = Color.White;
                text = Color.FromArgb(44, 62, 80);
                accent = Color.FromArgb(41, 128, 185);
            }

            this.BackColor = bg;
            gbInput.BackColor = surface;
            gbInput.ForeColor = text;
            tabSummary.BackColor = surface;
            tabAI.BackColor = surface;
            tabSchedule.BackColor = surface;
            if (_tabCompare != null) _tabCompare.BackColor = surface;
            rtbAI.BackColor = surface;
            rtbAI.ForeColor = text;

            Color titleBack = Color.FromArgb(Math.Max(10, accent.R - 20), Math.Max(10, accent.G - 20), Math.Max(10, accent.B - 20));
            lblMainTitle.BackColor = titleBack;

            _windowBtnFore = GetContrastColor(titleBack);
            _windowBtnNormalBack = Color.FromArgb(72, 0, 0, 0);
            _windowBtnHoverBack = Color.FromArgb(110, 0, 0, 0);

            btnClose.ForeColor = _windowBtnFore;
            btnMinimize.ForeColor = _windowBtnFore;
            _btnMaximize.ForeColor = _windowBtnFore;
            btnClose.BackColor = _windowBtnNormalBack;
            btnMinimize.BackColor = _windowBtnNormalBack;
            _btnMaximize.BackColor = _windowBtnNormalBack;

            _cmbTheme.BackColor = Color.White;
            _cmbTheme.ForeColor = Color.FromArgb(35, 35, 35);
            _cmbScale.BackColor = Color.White;
            _cmbScale.ForeColor = Color.FromArgb(35, 35, 35);
            _btnCustomTheme.BackColor = Color.White;
            _btnCustomTheme.ForeColor = Color.FromArgb(35, 35, 35);
            _btnCustomTheme.FlatAppearance.BorderColor = Color.FromArgb(120, accent);
            _btnStressTest.BackColor = Color.White;
            _btnStressTest.ForeColor = Color.FromArgb(35, 35, 35);
            _btnStressTest.FlatAppearance.BorderColor = Color.FromArgb(120, accent);
            _btnMonteCarlo.BackColor = Color.White;
            _btnMonteCarlo.ForeColor = Color.FromArgb(35, 35, 35);
            _btnMonteCarlo.FlatAppearance.BorderColor = Color.FromArgb(120, accent);
            _btnCopySummary.BackColor = Color.White;
            _btnCopySummary.ForeColor = Color.FromArgb(35, 35, 35);
            _btnCopySummary.FlatAppearance.BorderColor = Color.FromArgb(120, accent);
            _btnInputMode.BackColor = Color.White;
            _btnInputMode.ForeColor = Color.FromArgb(35, 35, 35);
            _btnInputMode.FlatAppearance.BorderColor = Color.FromArgb(120, accent);
            _btnAnimStrength.BackColor = Color.White;
            _btnAnimStrength.ForeColor = Color.FromArgb(35, 35, 35);
            _btnAnimStrength.FlatAppearance.BorderColor = Color.FromArgb(120, accent);
            _btnPdfTemplate.BackColor = Color.White;
            _btnPdfTemplate.ForeColor = Color.FromArgb(35, 35, 35);
            _btnPdfTemplate.FlatAppearance.BorderColor = Color.FromArgb(120, accent);
            _btnExportPdf.BackColor = Color.White;
            _btnExportPdf.ForeColor = Color.FromArgb(35, 35, 35);
            _btnExportPdf.FlatAppearance.BorderColor = Color.FromArgb(120, accent);

            btnCalc.BackColor = Color.FromArgb(39, 174, 96);
            btnExport.BackColor = accent;
            btnReset.BackColor = Color.FromArgb(149, 165, 166);
            btnCalc.ForeColor = Color.White;
            btnExport.ForeColor = Color.White;
            btnReset.ForeColor = Color.White;

            dgvSchedule.BackgroundColor = surface;
            dgvSchedule.DefaultCellStyle.BackColor = surface;
            dgvSchedule.DefaultCellStyle.ForeColor = text;
            dgvSchedule.ColumnHeadersDefaultCellStyle.BackColor = accent;
            dgvSchedule.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvSchedule.EnableHeadersVisualStyles = false;

            // Modernize Schedule Table styling
            dgvSchedule.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgvSchedule.GridColor = Color.FromArgb(235, 235, 235);
            dgvSchedule.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgvSchedule.ColumnHeadersHeight = 45;
            dgvSchedule.RowTemplate.Height = 36;
            dgvSchedule.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 251, 252);
            dgvSchedule.DefaultCellStyle.SelectionBackColor = Color.FromArgb(235, 245, 251);
            dgvSchedule.DefaultCellStyle.SelectionForeColor = text;
            dgvSchedule.DefaultCellStyle.Padding = new Padding(6, 0, 6, 0);
            dgvSchedule.ColumnHeadersDefaultCellStyle.Padding = new Padding(6, 0, 6, 0);
            dgvSchedule.ColumnHeadersDefaultCellStyle.Font = new Font("微軟正黑體", 11.5F, FontStyle.Bold);

            if (_dgvCompare != null)
            {
                _dgvCompare.BackgroundColor = surface;
                _dgvCompare.DefaultCellStyle.BackColor = surface;
                _dgvCompare.DefaultCellStyle.ForeColor = text;
                _dgvCompare.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 73, 94);
                _dgvCompare.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                _dgvCompare.EnableHeadersVisualStyles = false;
                _dgvCompare.GridColor = Color.FromArgb(235, 235, 235);
                _dgvCompare.DefaultCellStyle.SelectionBackColor = Color.FromArgb(235, 245, 251);
                _dgvCompare.DefaultCellStyle.SelectionForeColor = text;
                _dgvCompare.DefaultCellStyle.Padding = new Padding(6, 0, 6, 0);
                _dgvCompare.ColumnHeadersDefaultCellStyle.Padding = new Padding(6, 0, 6, 0);
                _dgvCompare.ColumnHeadersDefaultCellStyle.Font = new Font("微軟正黑體", 11.5F, FontStyle.Bold);
            }
        }

        private Color GetContrastColor(Color c)
        {
            int yiq = ((c.R * 299) + (c.G * 587) + (c.B * 114)) / 1000;
            return yiq >= 150 ? Color.FromArgb(20, 20, 20) : Color.White;
        }

        private void TitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void FlatButton_Paint(object sender, PaintEventArgs e)
        {
            Button btn = sender as Button;
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle rect = new Rectangle(0, 0, btn.Width, btn.Height);
            
            using (GraphicsPath path = new GraphicsPath())
            {
                int radius = 10;
                path.AddArc(0, 0, radius, radius, 180, 90);
                path.AddArc(rect.Right - radius, 0, radius, radius, 270, 90);
                path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90);
                path.AddArc(0, rect.Bottom - radius, radius, radius, 90, 90);
                path.CloseAllFigures();
                btn.Region = new Region(path);
                
                using (SolidBrush brush = new SolidBrush(btn.BackColor))
                    g.FillPath(brush, path);
            }
            TextRenderer.DrawText(g, btn.Text, btn.Font, rect, btn.ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            using (Pen borderPen = new Pen(Color.FromArgb(52, 73, 94), 2))
            {
                g.DrawRectangle(borderPen, 1, 1, this.Width - 2, this.Height - 2);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_NCHITTEST && this.WindowState == FormWindowState.Normal)
            {
                base.WndProc(ref m);
                if ((int)m.Result == HT_CAPTION) return;

                Point p = this.PointToClient(Cursor.Position);
                bool left = p.X <= ResizeBorderWidth;
                bool right = p.X >= this.ClientSize.Width - ResizeBorderWidth;
                bool top = p.Y <= ResizeBorderWidth;
                bool bottom = p.Y >= this.ClientSize.Height - ResizeBorderWidth;

                if (left && top) m.Result = (IntPtr)HTTOPLEFT;
                else if (right && top) m.Result = (IntPtr)HTTOPRIGHT;
                else if (left && bottom) m.Result = (IntPtr)HTBOTTOMLEFT;
                else if (right && bottom) m.Result = (IntPtr)HTBOTTOMRIGHT;
                else if (left) m.Result = (IntPtr)HTLEFT;
                else if (right) m.Result = (IntPtr)HTRIGHT;
                else if (top) m.Result = (IntPtr)HTTOP;
                else if (bottom) m.Result = (IntPtr)HTBOTTOM;
                return;
            }
            base.WndProc(ref m);
        }

        private void InitializeDefaults()
        {
            txtPrice.Text = "15,000,000";
            cmbDownPaymentType.SelectedIndex = 0; // %
            txtDownPayment.Text = "20";
            txtRate.Text = "2.15";
            cmbTerm.Text = "20";
            txtGrace.Text = "0";
            if (_txtReportStudentId != null)
            {
                _txtReportStudentId.Text = "1113354";
                _txtReportStudentName.Text = "陳冠瑋";
                _txtReportCourse.Text = "程式設計實務";
                _txtReportSchool.Text = "國立OO大學";
                _txtReportDepartment.Text = "資訊工程學系";
                _txtReportAdvisor.Text = "王OO教授";
                _dtReportDate.Value = DateTime.Now;
            }

            lblResultTotalLoan.Text = "NT$ 0";
            lblResultMonthly.Text = "NT$ 0";
            lblResultFirstInt.Text = "NT$ 0";
            lblResultFirstPrin.Text = "NT$ 0";
            lblResultTotalInt.Text = "NT$ 0";
            lblResultTotalRepay.Text = "NT$ 0";
            lblValidationHint.Text = "";

            dgvSchedule.ReadOnly = true;
            dgvSchedule.AllowUserToAddRows = false;
            dgvSchedule.RowHeadersVisible = false;
            dgvSchedule.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvSchedule.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvSchedule.BackgroundColor = Color.White;
            dgvSchedule.BorderStyle = BorderStyle.None;
            dgvSchedule.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        }

        private void NumberOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true;
        }

        private void FloatOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.')) e.Handled = true;
            if ((e.KeyChar == '.') && (((TextBox)sender).Text.IndexOf('.') > -1)) e.Handled = true;
        }

        private double GetDouble(string text)
        {
            double.TryParse(text.Replace(",", ""), out double result);
            return result;
        }

        private async void btnCalc_Click(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            bool hasError = false;
            var errors = new List<string>();

            double price = GetDouble(txtPrice.Text);
            if (price <= 0)
            {
                errorProvider1.SetError(txtPrice, "請輸入有效的房屋總價");
                errors.Add("房屋總價必須大於 0。建議輸入如：15000000。");
                hasError = true;
            }

            double downPaymentVal = GetDouble(txtDownPayment.Text);
            double downPaymentAmount = 0;
            if (cmbDownPaymentType.SelectedIndex == 0) // %
            {
                if (downPaymentVal < 0 || downPaymentVal >= 100)
                {
                    errorProvider1.SetError(txtDownPayment, "比例需介於 0 到 99 之間");
                    errors.Add("自備款比例需介於 0~99%。例如 20 代表 20%。");
                    hasError = true;
                }
                downPaymentAmount = price * (downPaymentVal / 100.0);
            }
            else // 元
            {
                if (downPaymentVal < 0 || downPaymentVal >= price)
                {
                    errorProvider1.SetError(txtDownPayment, "自備款需合理");
                    errors.Add("自備款金額不可小於 0，且必須小於房屋總價。");
                    hasError = true;
                }
                downPaymentAmount = downPaymentVal;
            }

            double annualRate = GetDouble(txtRate.Text);
            if (annualRate <= 0 || annualRate > 100)
            {
                errorProvider1.SetError(txtRate, "有效年利率(0~100)");
                errors.Add("貸款年利率需介於 0~100。通常可輸入 2.15。\n");
                hasError = true;
            }

            int termYears;
            if (!int.TryParse(cmbTerm.Text, out termYears) || termYears < 5 || termYears > 50)
            {
                errorProvider1.SetError(cmbTerm, "貸款年限需介於 5~50 年");
                errors.Add("貸款年限請輸入 5~50 年之間，可自行輸入如 22、35。\n");
                hasError = true;
                termYears = 20;
            }
            int graceYears = (int)GetDouble(txtGrace.Text);

            if (graceYears < 0)
            {
                errorProvider1.SetError(txtGrace, "寬限期不可為負值");
                errors.Add("寬限期不可為負數。\n");
                hasError = true;
            }

            if (graceYears >= termYears)
            {
                errorProvider1.SetError(txtGrace, "寬限期必須小於貸款年限");
                errors.Add("寬限期設定不合邏輯：使用寬限期必須小於貸款年限。\n例如年限 20 年，寬限期最多 19 年。");
                hasError = true;
            }

            if (hasError)
            {
                lblValidationHint.Text = "⚠ 參數不正確：\n- " + string.Join("\n- ", errors);
                rtbAI.Text = "目前無法產生分析，請先修正左側輸入參數。\n\n" + string.Join("\n", errors);
                return;
            }

            lblValidationHint.Text = "";

            btnCalc.Text = "計算中...";
            btnCalc.Enabled = false;

            totalLoan = price - downPaymentAmount;
            double monthlyRate = (annualRate / 100.0) / 12.0;
            int totalMonths = termYears * 12;
            int graceMonths = graceYears * 12;
            _lastTotalMonths = totalMonths;
            _lastGraceMonths = graceMonths;

            bool isNewYouth = _chkNewYouth != null && _chkNewYouth.Checked;
            double youthRate = (1.775 / 100.0) / 12.0;
            int youthMonths = 36; // 優惠 3 年

            double monthlyPaymentNormal = 0;
            int remainingMonths = totalMonths - graceMonths;
            double annualPrepay = _numAnnualPrepay != null ? (double)_numAnnualPrepay.Value : 0;

            double firstMonthInt = 0;
            double firstMonthPrin = 0;

            _schedule.Clear();

            // Generate scheduling in background task to show enterprise asynchronous pattern
            await Task.Run(() => {
                double tempInterest = 0;
                double tempRepayment = 0;
                int monthsUsed = totalMonths;

                double loanA = (isNewYouth && totalLoan > 10000000) ? 10000000 : totalLoan;
                double loanB = (isNewYouth && totalLoan > 10000000) ? totalLoan - 10000000 : 0;

                double balanceA = loanA;
                double balanceB = loanB;

                double currentPmtA = 0;
                double currentPmtB = 0;

                if (remainingMonths > 0)
                {
                    double rA = isNewYouth ? youthRate : monthlyRate;
                    currentPmtA = balanceA * rA * Math.Pow(1 + rA, remainingMonths) / (Math.Pow(1 + rA, remainingMonths) - 1);
                    if (loanB > 0) {
                        currentPmtB = balanceB * monthlyRate * Math.Pow(1 + monthlyRate, remainingMonths) / (Math.Pow(1 + monthlyRate, remainingMonths) - 1);
                    }
                }

                monthlyPaymentNormal = currentPmtA + currentPmtB;
                firstMonthInt = (loanA * (isNewYouth ? youthRate : monthlyRate)) + (loanB * monthlyRate);
                firstMonthPrin = (graceMonths > 0) ? 0 : (monthlyPaymentNormal - firstMonthInt);

                for (int i = 1; i <= totalMonths; i++)
                {
                    double rateA = (isNewYouth && i <= youthMonths) ? youthRate : monthlyRate;
                    double rateB = monthlyRate;

                    if (isNewYouth && i == youthMonths + 1 && i > graceMonths && balanceA > 0)
                    {
                        int remain = totalMonths - i + 1;
                        if (remain > 0) {
                            currentPmtA = balanceA * rateA * Math.Pow(1 + rateA, remain) / (Math.Pow(1 + rateA, remain) - 1);
                        }
                    }

                    double intA = balanceA * rateA;
                    double intB = balanceB * rateB;
                    double interest = intA + intB;

                    double prinA = 0, prinB = 0;
                    double pmtA = 0, pmtB = 0;

                    if (i <= graceMonths) { 
                        pmtA = intA; pmtB = intB;
                    } else { 
                        pmtA = currentPmtA; prinA = pmtA - intA;
                        pmtB = currentPmtB; prinB = pmtB - intB;
                    }

                    double payment = pmtA + pmtB;
                    double principal = prinA + prinB;

                    if (annualPrepay > 0 && i > graceMonths && i % 12 == 0 && (balanceA + balanceB) > 0)
                    {
                        double extra = Math.Min(annualPrepay, Math.Max(0, (balanceA + balanceB) - principal));
                        principal += extra;
                        payment += extra;

                        if (balanceB > 0) {
                            double extraB = Math.Min(extra, balanceB - prinB);
                            prinB += extraB;
                            extra -= extraB;
                        }
                        if (extra > 0 && balanceA > 0) {
                            prinA += extra;
                        }
                    }

                    tempInterest += interest;
                    tempRepayment += payment;

                    balanceA -= prinA;
                    balanceB -= prinB;

                    if (balanceA < 0) {
                        tempRepayment += balanceA; payment += balanceA;
                        balanceA = 0;
                    }
                    if (balanceB < 0) {
                        tempRepayment += balanceB; payment += balanceB;
                        balanceB = 0;
                    }
                    double balance = balanceA + balanceB;

                    _schedule.Add(new AmortizationItem {
                        Month = i,
                        Principal = principal.ToString("N2"),
                        Interest = interest.ToString("N2"),
                        Payment = payment.ToString("N2"),
                        Balance = balance.ToString("N2")
                    });

                    if (balance <= 0)
                    {
                        monthsUsed = i;
                        break;
                    }
                }
                totalInterest = tempInterest;
                totalRepayment = tempRepayment;
                _effectivePayoffMonths = monthsUsed;
            });

            _bs.DataSource = typeof(AmortizationItem);
            _bs.DataSource = _schedule;
            _bs.ResetBindings(false);

            lblResultTotalLoan.Text = "NT$ " + totalLoan.ToString("N2");
            StartResultAnimation(firstMonthInt, firstMonthPrin, totalInterest, totalRepayment, monthlyPaymentNormal, graceMonths > 0, firstMonthInt);

            rtbAI.Text = "AI 分析引擎處理中...";
            rtbAI.Text = await BuildAiReportAsync(price, downPaymentAmount, termYears, graceYears, annualRate, monthlyPaymentNormal);

            picChart.Invalidate();
            
            btnCalc.Text = "開始試算";
            btnCalc.Enabled = true;
        }

        private void StartResultAnimation(double firstMonthInt, double firstMonthPrin, double totalInt, double totalRepay, double monthlyNormal, bool hasGrace, double monthlyGrace)
        {
            _resultAnimTimer.Stop();
            _animTick = 0;

            _monthlyHasGrace = hasGrace;
            _monthlyNormalValue = monthlyNormal;
            _monthlyGraceValue = monthlyGrace;

            _animFrom = new Dictionary<Label, double>
            {
                { lblResultTotalLoan, 0 },
                { lblResultFirstInt, 0 },
                { lblResultFirstPrin, 0 },
                { lblResultTotalInt, 0 },
                { lblResultTotalRepay, 0 }
            };

            _animTo = new Dictionary<Label, double>
            {
                { lblResultTotalLoan, totalLoan },
                { lblResultFirstInt, firstMonthInt },
                { lblResultFirstPrin, firstMonthPrin },
                { lblResultTotalInt, totalInt },
                { lblResultTotalRepay, totalRepay }
            };

            if (_animStrengthLevel == 0)
            {
                lblResultTotalLoan.Text = "NT$ " + totalLoan.ToString("N2");
                lblResultFirstInt.Text = "NT$ " + firstMonthInt.ToString("N2");
                lblResultFirstPrin.Text = "NT$ " + firstMonthPrin.ToString("N2");
                lblResultTotalInt.Text = "NT$ " + totalInt.ToString("N2");
                lblResultTotalRepay.Text = "NT$ " + totalRepay.ToString("N2");
                lblResultMonthly.Text = hasGrace
                    ? "NT$ " + monthlyGrace.ToString("N2") + " (寬限期後: " + monthlyNormal.ToString("N2") + ")"
                    : "NT$ " + monthlyNormal.ToString("N2");
                return;
            }

            _resultAnimTimer.Start();
        }

        private void RunStressTest()
        {
            double loan = totalLoan;
            if (loan <= 0)
            {
                MessageBox.Show("請先完成一次試算，再執行壓力測試。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!double.TryParse(txtRate.Text, out double baseRate)) return;
            if (!int.TryParse(cmbTerm.Text, out int termYears)) return;
            if (!int.TryParse(txtGrace.Text, out int graceYears)) graceYears = 0;

            int totalMonths = termYears * 12;
            int remainMonths = Math.Max(1, totalMonths - (graceYears * 12));

            var sb = new StringBuilder();
            sb.AppendLine("【利率壓力測試】");
            for (double d = -1.0; d <= 1.0001; d += 0.5)
            {
                double testRate = Math.Max(0.01, baseRate + d);
                double m = (testRate / 100.0) / 12.0;
                double p = loan * m * Math.Pow(1 + m, remainMonths) / (Math.Pow(1 + m, remainMonths) - 1);
                sb.AppendLine(string.Format("年利率 {0:F2}% => 月付約 NT$ {1:N0}", testRate, p));
            }

            rtbAI.Text = sb.ToString() + "\n\n" + rtbAI.Text;
            tabControlHelper.SelectedTab = tabAI;
        }

        private void RunMonteCarloStressTest()
        {
            if (totalLoan <= 0)
            {
                MessageBox.Show("請先完成一次試算，再執行隨機壓測。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!double.TryParse(txtRate.Text, out double baseRate)) return;
            if (!int.TryParse(cmbTerm.Text, out int termYears)) return;
            if (!int.TryParse(txtGrace.Text, out int graceYears)) graceYears = 0;

            int totalMonths = termYears * 12;
            int remainMonths = Math.Max(1, totalMonths - (graceYears * 12));

            const int samples = 200;
            const double sigma = 0.45;
            var random = new Random();
            var pays = new List<double>(samples);

            for (int i = 0; i < samples; i++)
            {
                double z = NextGaussian(random);
                double rate = Math.Max(0.2, baseRate + (z * sigma));
                double m = (rate / 100.0) / 12.0;
                double p = totalLoan * m * Math.Pow(1 + m, remainMonths) / (Math.Pow(1 + m, remainMonths) - 1);
                if (!double.IsNaN(p) && !double.IsInfinity(p)) pays.Add(p);
            }

            if (pays.Count == 0) return;
            pays.Sort();
            double p10 = pays[(int)(pays.Count * 0.10)];
            double p50 = pays[(int)(pays.Count * 0.50)];
            double p90 = pays[(int)(pays.Count * 0.90)];
            double avg = pays.Average();

            var sb = new StringBuilder();
            sb.AppendLine("【蒙地卡羅隨機壓測】");
            sb.AppendLine("樣本數: " + pays.Count + "（利率基準 " + baseRate.ToString("0.00") + "%，σ=" + sigma.ToString("0.00") + "）");
            sb.AppendLine("平均月付: NT$ " + avg.ToString("N0"));
            sb.AppendLine("P10(月付較低): NT$ " + p10.ToString("N0"));
            sb.AppendLine("P50(中位數): NT$ " + p50.ToString("N0"));
            sb.AppendLine("P90(月付較高): NT$ " + p90.ToString("N0"));
            sb.AppendLine("風險帶寬(P90-P10): NT$ " + (p90 - p10).ToString("N0"));

            rtbAI.Text = sb.ToString() + "\n\n" + rtbAI.Text;
            tabControlHelper.SelectedTab = tabAI;
        }

        private static double NextGaussian(Random random)
        {
            double u1 = 1.0 - random.NextDouble();
            double u2 = 1.0 - random.NextDouble();
            return Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2);
        }

        private void ToggleInputMode()
        {
            _advancedInputVisible = !_advancedInputVisible;

            int[] advancedRows = { 0, 6, 7, 8 };
            foreach (Control c in tableLayoutPanelInput.Controls)
            {
                int row = tableLayoutPanelInput.GetRow(c);
                if (advancedRows.Contains(row)) c.Visible = _advancedInputVisible;
            }

            if (tableLayoutPanelInput.RowStyles.Count > 8)
            {
                tableLayoutPanelInput.RowStyles[0].Height = _advancedInputVisible ? 100F : 0F;
                tableLayoutPanelInput.RowStyles[6].Height = _advancedInputVisible ? 58F : 0F;
                tableLayoutPanelInput.RowStyles[7].Height = _advancedInputVisible ? 58F : 0F;
                tableLayoutPanelInput.RowStyles[8].Height = _advancedInputVisible ? 94F : 0F;
            }

            _btnInputMode.Text = _advancedInputVisible ? "進階" : "基本";
        }

        private void ToggleAnimationStrength()
        {
            _animStrengthLevel--;
            if (_animStrengthLevel < 0) _animStrengthLevel = 2;

            if (_animStrengthLevel == 2)
            {
                _currentAnimTicks = AnimTicks;
                _uiMotionTimer.Interval = 80;
                _uiMotionTimer.Start();
                _btnAnimStrength.Text = "動畫:高";
            }
            else if (_animStrengthLevel == 1)
            {
                _currentAnimTicks = 8;
                _uiMotionTimer.Interval = 170;
                _uiMotionTimer.Start();
                _btnAnimStrength.Text = "動畫:低";
            }
            else
            {
                _currentAnimTicks = 1;
                _uiMotionTimer.Stop();
                _btnAnimStrength.Text = "動畫:關";
            }
        }

        private void CopySummaryToClipboard()
        {
            if (totalLoan <= 0)
            {
                MessageBox.Show("請先執行試算，才可複製摘要。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var text = new StringBuilder();
            text.AppendLine("房貸試算摘要");
            text.AppendLine("----------------");
            text.AppendLine("貸款總額: " + lblResultTotalLoan.Text);
            text.AppendLine("每月還款: " + lblResultMonthly.Text);
            text.AppendLine("首期利息: " + lblResultFirstInt.Text);
            text.AppendLine("首期本金: " + lblResultFirstPrin.Text);
            text.AppendLine("總利息: " + lblResultTotalInt.Text);
            text.AppendLine("總還款: " + lblResultTotalRepay.Text);

            Clipboard.SetText(text.ToString());
            MessageBox.Show("已複製摘要到剪貼簿。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void AddCurrentScenario()
        {
            if (totalLoan <= 0)
            {
                MessageBox.Show("請先完成試算，再加入比較。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_scenarios.Count >= 3)
            {
                MessageBox.Show("最多只能比較 3 套方案。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var s = new ScenarioResult
            {
                方案 = "方案" + (_scenarios.Count + 1),
                房價 = txtPrice.Text,
                利率 = txtRate.Text + "%",
                年限 = cmbTerm.Text + "年",
                寬限 = txtGrace.Text + "年",
                年提前還款 = (_numAnnualPrepay != null ? _numAnnualPrepay.Value.ToString("N0") : "0") + "元",
                每月還款 = lblResultMonthly.Text,
                總利息 = lblResultTotalInt.Text,
                總還款 = lblResultTotalRepay.Text,
                清償月數 = _effectivePayoffMonths + "月"
            };
            _scenarios.Add(s);
            _compareBs.ResetBindings(false);
            tabControlHelper.SelectedTab = _tabCompare;
        }

        private void AutoGenerateScenarios()
        {
            double price = GetDouble(txtPrice.Text);
            if (price <= 0)
            {
                MessageBox.Show("請先輸入有效房價。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!int.TryParse(cmbTerm.Text, out int termYears)) termYears = 30;
            if (!double.TryParse(txtRate.Text, out double baseRate)) baseRate = 2.15;
            int graceYears = (int)GetDouble(txtGrace.Text);
            if (graceYears < 0) graceYears = 0;
            if (graceYears >= termYears) graceYears = Math.Max(0, termYears - 1);
            double annualPrepay = _numAnnualPrepay != null ? (double)_numAnnualPrepay.Value : 0;

            double downRatio = cmbDownPaymentType.SelectedIndex == 0
                ? GetDouble(txtDownPayment.Text) / 100.0
                : (price <= 0 ? 0.2 : GetDouble(txtDownPayment.Text) / price);
            downRatio = Math.Max(0.05, Math.Min(0.95, downRatio));

            _scenarios.Clear();
            _scenarios.Add(BuildScenarioResult("保守型", price, downRatio, Math.Min(50, termYears + 5), Math.Max(0.5, baseRate + 0.5), graceYears, 0));
            _scenarios.Add(BuildScenarioResult("平衡型", price, downRatio, termYears, baseRate, graceYears, annualPrepay));
            _scenarios.Add(BuildScenarioResult("積極型", price, Math.Min(0.95, downRatio + 0.05), Math.Max(5, termYears - 5), Math.Max(0.3, baseRate - 0.3), Math.Max(0, graceYears - 1), annualPrepay + 50000));

            _compareBs.ResetBindings(false);
            tabControlHelper.SelectedTab = _tabCompare;
        }

        private ScenarioResult BuildScenarioResult(string name, double price, double downRatio, int termYears, double annualRate, int graceYears, double annualPrepay)
        {
            double downAmount = price * downRatio;
            double loan = Math.Max(0, price - downAmount);
            double monthlyRate = (annualRate / 100.0) / 12.0;
            int totalMonths = Math.Max(1, termYears * 12);
            int graceMonths = Math.Max(0, Math.Min(totalMonths - 1, graceYears * 12));
            int repayMonths = Math.Max(1, totalMonths - graceMonths);

            double monthlyNormal = loan * monthlyRate * Math.Pow(1 + monthlyRate, repayMonths) / (Math.Pow(1 + monthlyRate, repayMonths) - 1);
            if (double.IsNaN(monthlyNormal) || double.IsInfinity(monthlyNormal)) monthlyNormal = loan / repayMonths;

            double balance = loan;
            double totalInt = 0;
            double totalRepay = 0;
            int payoff = totalMonths;

            for (int i = 1; i <= totalMonths; i++)
            {
                double interest = balance * monthlyRate;
                double payment;
                double principal;
                if (i <= graceMonths)
                {
                    payment = interest;
                    principal = 0;
                }
                else
                {
                    payment = monthlyNormal;
                    principal = payment - interest;
                }

                if (annualPrepay > 0 && i > graceMonths && i % 12 == 0 && balance > 0)
                {
                    double extra = Math.Min(annualPrepay, Math.Max(0, balance - principal));
                    principal += extra;
                    payment += extra;
                }

                totalInt += interest;
                totalRepay += payment;
                balance -= principal;
                if (balance <= 0)
                {
                    payoff = i;
                    break;
                }
            }

            return new ScenarioResult
            {
                方案 = name,
                房價 = price.ToString("N0"),
                利率 = annualRate.ToString("0.00") + "%",
                年限 = termYears + "年",
                寬限 = graceYears + "年",
                年提前還款 = annualPrepay.ToString("N0") + "元",
                每月還款 = "NT$ " + (graceMonths > 0 ? (loan * monthlyRate).ToString("N0") + " (後: " + monthlyNormal.ToString("N0") + ")" : monthlyNormal.ToString("N0")),
                總利息 = "NT$ " + totalInt.ToString("N0"),
                總還款 = "NT$ " + totalRepay.ToString("N0"),
                清償月數 = payoff + "月"
            };
        }

        private void ExportPdfReport()
        {
            if (totalLoan <= 0)
            {
                MessageBox.Show("請先執行試算，再匯出 PDF。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = "房貸試算報告.pdf" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                List<byte[]> pageImages = BuildReportBookPages();
                WritePdfFromJpegPages(sfd.FileName, pageImages);
                MessageBox.Show("PDF 匯出完成（內建引擎，不需安裝 Pandoc/LaTeX）。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private bool TryGeneratePandocAcademicPdf(string pdfPath, out string error)
        {
            error = string.Empty;
            try
            {
                string dir = Path.GetDirectoryName(pdfPath);
                string name = Path.GetFileNameWithoutExtension(pdfPath);
                string mdPath = Path.Combine(dir, name + ".md");
                string texPath = Path.Combine(dir, name + ".tex");
                string templatePath = Path.Combine(dir, "academic_ieee_style_template.tex");

                File.WriteAllText(mdPath, BuildAcademicMarkdown(), Encoding.UTF8);
                EnsureAcademicLatexTemplate(templatePath);

                string toolchainError;
                if (!ValidatePandocToolchain(out toolchainError))
                {
                    error = toolchainError + "\n\n已先輸出：\n- " + mdPath + "\n- " + templatePath;
                    return false;
                }

                var toTex = RunExternalProcess("pandoc", string.Format("\"{0}\" -f markdown -t latex -o \"{1}\"", mdPath, texPath));
                if (toTex.ExitCode != 0)
                {
                    error = "Markdown 轉 LaTeX 失敗：\n" + toTex.Output;
                    return false;
                }

                var toPdf = RunExternalProcess("pandoc", string.Format("\"{0}\" --from markdown --template \"{1}\" --pdf-engine=xelatex --toc --number-sections -V papersize=a4 -o \"{2}\"", mdPath, templatePath, pdfPath));
                if (toPdf.ExitCode != 0)
                {
                    error = "LaTeX 轉 PDF 失敗：\n" + toPdf.Output;
                    return false;
                }

                return File.Exists(pdfPath);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private bool ValidatePandocToolchain(out string message)
        {
            message = string.Empty;
            string pandocPath = FindExecutable("pandoc.exe");
            if (string.IsNullOrWhiteSpace(pandocPath))
            {
                message = "找不到 pandoc.exe。\n請先安裝 Pandoc 並重開 Visual Studio。\n下載：https://pandoc.org/installing.html";
                return false;
            }

            string xelatexPath = FindExecutable("xelatex.exe");
            if (string.IsNullOrWhiteSpace(xelatexPath))
            {
                message = "找不到 xelatex.exe。\n請安裝 MiKTeX 或 TeX Live，並將 XeLaTeX 加入 PATH。";
                return false;
            }

            return true;
        }

        private string FindExecutable(string exeName)
        {
            try
            {
                string local = Path.Combine(Application.StartupPath, exeName);
                if (File.Exists(local)) return local;

                string pathVar = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
                string[] paths = pathVar.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < paths.Length; i++)
                {
                    string p = paths[i].Trim();
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    string full = Path.Combine(p, exeName);
                    if (File.Exists(full)) return full;
                }
            }
            catch
            {
            }
            return string.Empty;
        }

        private string BuildAcademicMarkdown()
        {
            string sid = string.IsNullOrWhiteSpace(_txtReportStudentId?.Text) ? "未填寫" : _txtReportStudentId.Text.Trim();
            string sname = string.IsNullOrWhiteSpace(_txtReportStudentName?.Text) ? "未填寫" : _txtReportStudentName.Text.Trim();
            string course = string.IsNullOrWhiteSpace(_txtReportCourse?.Text) ? "未填寫" : _txtReportCourse.Text.Trim();
            string school = string.IsNullOrWhiteSpace(_txtReportSchool?.Text) ? "未填寫" : _txtReportSchool.Text.Trim();
            string dept = string.IsNullOrWhiteSpace(_txtReportDepartment?.Text) ? "未填寫" : _txtReportDepartment.Text.Trim();
            string advisor = string.IsNullOrWhiteSpace(_txtReportAdvisor?.Text) ? "未填寫" : _txtReportAdvisor.Text.Trim();
            string reportDate = (_dtReportDate != null ? _dtReportDate.Value : DateTime.Now).ToString("yyyy-MM-dd");

            var sb = new StringBuilder();
            sb.AppendLine("---");
            sb.AppendLine("title: \"個人房貸試算研究報告（IEEE風格）\"");
            sb.AppendLine("author: \"" + sname + "（" + sid + "）\"");
            sb.AppendLine("date: \"" + reportDate + "\"");
            sb.AppendLine("---");
            sb.AppendLine();
            sb.AppendLine("**學校**：" + school + "  ");
            sb.AppendLine("**系所**：" + dept + "  ");
            sb.AppendLine("**課程**：" + course + "  ");
            sb.AppendLine("**指導老師**：" + advisor + "  ");
            sb.AppendLine();
            sb.AppendLine("# Abstract");
            sb.AppendLine("本研究提出一套整合房貸試算、提前還款模擬、利率壓力測試與視覺化分析的決策支援系統，協助使用者在不同利率與年限條件下評估整體財務負擔。系統同時提供情境比較、雷達圖與熱力矩陣，並可輸出正式報告文件。實驗結果顯示，提前還款策略可有效縮短清償期並降低累積利息。\n");
            sb.AppendLine("**Keywords**—Mortgage; Amortization; Prepayment; Stress Test; Decision Support");
            sb.AppendLine();
            sb.AppendLine("# 1. Introduction");
            sb.AppendLine("房貸為長期財務決策，除月付能力外，亦需評估利率波動、寬限期風險與提前償還策略。本系統以互動試算與報告輸出方式，提升使用者對長期還款結構的理解。\n");
            sb.AppendLine("# 2. Methodology");
            sb.AppendLine("本研究採用本息平均攤還公式，並支援每年固定提前還款。另建立利率 ±1% 壓力測試，計算對月付與總利息的敏感度。\n");
            sb.AppendLine("# 3. Experimental Results");
            sb.AppendLine("- 貸款總額：" + lblResultTotalLoan.Text);
            sb.AppendLine("- 每月還款：" + lblResultMonthly.Text);
            sb.AppendLine("- 總利息：" + lblResultTotalInt.Text);
            sb.AppendLine("- 總還款：" + lblResultTotalRepay.Text);
            sb.AppendLine("- 有效清償月數：" + _effectivePayoffMonths + " 月\n");
            sb.AppendLine("## 3.1 Stress Test");
            foreach (var row in BuildStressRowsForReport())
                sb.AppendLine("- " + row);
            sb.AppendLine();
            sb.AppendLine("## 3.2 Amortization Appendix (Month 12-24)");
            sb.AppendLine("| 月份 | 償還本金 | 償還利息 | 本息合計 | 剩餘本金 |");
            sb.AppendLine("|---:|---:|---:|---:|---:|");
            for (int m = 12; m <= Math.Min(24, _schedule.Count); m++)
            {
                var r = _schedule[m - 1];
                sb.AppendLine(string.Format("| {0} | {1} | {2} | {3} | {4} |", r.Month, r.Principal, r.Interest, r.Payment, r.Balance));
            }
            sb.AppendLine();
            sb.AppendLine("# 4. Discussion and Recommendation");
            sb.AppendLine("分析顯示：寬限期雖可降低前期負擔，但後期支付壓力明顯提高。建議透過年度提前還款策略與定期利率重估，維持財務韌性。\n");
            sb.AppendLine("# 5. Conclusion");
            sb.AppendLine("本研究完成一套可視化房貸分析系統，具備教學展示與實務應用潛力。未來可整合外部 API，提升資料即時性。\n");
            sb.AppendLine("# References");
            sb.AppendLine("[1] 內政部不動產資訊平台，住宅價格統計。  ");
            sb.AppendLine("[2] 金融監督管理委員會，房貸與利率相關公開資料。  ");
            sb.AppendLine("[3] Local experimental results generated by this system.");
            return sb.ToString();
        }

        private void EnsureAcademicLatexTemplate(string path)
        {
            if (File.Exists(path)) return;
            var t = new StringBuilder();
            t.AppendLine("\\documentclass[10pt,twocolumn,a4paper]{article}");
            t.AppendLine("\\usepackage[UTF8]{ctex}");
            t.AppendLine("\\usepackage{geometry}");
            t.AppendLine("\\geometry{margin=1.8cm}");
            t.AppendLine("\\usepackage{fancyhdr}");
            t.AppendLine("\\usepackage{graphicx}");
            t.AppendLine("\\usepackage{booktabs}");
            t.AppendLine("\\usepackage{longtable}");
            t.AppendLine("\\usepackage{hyperref}");
            t.AppendLine("\\pagestyle{fancy}");
            t.AppendLine("\\fancyhf{}");
            t.AppendLine("\\fancyfoot[C]{\\thepage}");
            t.AppendLine("\\setlength{\\parskip}{0.4em}");
            t.AppendLine("\\setlength{\\parindent}{2em}");
            t.AppendLine("\\begin{document}");
            t.AppendLine("$if(title)$\\title{$title$}$endif$");
            t.AppendLine("$if(author)$\\author{$author$}$endif$");
            t.AppendLine("$if(date)$\\date{$date$}$endif$");
            t.AppendLine("\\maketitle");
            t.AppendLine("$if(toc)$\\tableofcontents\\newpage$endif$");
            t.AppendLine("$body$");
            t.AppendLine("\\end{document}");
            File.WriteAllText(path, t.ToString(), Encoding.UTF8);
        }

        private class ProcessResult
        {
            public int ExitCode { get; set; }
            public string Output { get; set; }
        }

        private ProcessResult RunExternalProcess(string fileName, string args)
        {
            var r = new ProcessResult();
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                using (var p = Process.Start(psi))
                {
                    string o = p.StandardOutput.ReadToEnd();
                    string e = p.StandardError.ReadToEnd();
                    p.WaitForExit(20000);
                    r.ExitCode = p.ExitCode;
                    r.Output = (o + "\n" + e).Trim();
                }
            }
            catch (Exception ex)
            {
                r.ExitCode = -1;
                r.Output = ex.Message;
            }
            return r;
        }

        private List<byte[]> BuildReportBookPages()
        {
            var pages = new List<byte[]>();
            int w = 1240;
            int h = 1754;

            int p = 1;
            int pCover = p++;
            int pToc = p++;
            int pChAbstract = p++;
            int pChMethodResult = p++;
            int pChMainChart = p++;
            int pChScenario = p++;
            int pAppendix = p++;
            int pFinal = p++;
            int totalPages = p - 1;

            string sid = string.IsNullOrWhiteSpace(_txtReportStudentId?.Text) ? "未填寫" : _txtReportStudentId.Text.Trim();
            string sname = string.IsNullOrWhiteSpace(_txtReportStudentName?.Text) ? "未填寫" : _txtReportStudentName.Text.Trim();
            string course = string.IsNullOrWhiteSpace(_txtReportCourse?.Text) ? "未填寫" : _txtReportCourse.Text.Trim();
            string school = string.IsNullOrWhiteSpace(_txtReportSchool?.Text) ? "未填寫" : _txtReportSchool.Text.Trim();
            string dept = string.IsNullOrWhiteSpace(_txtReportDepartment?.Text) ? "未填寫" : _txtReportDepartment.Text.Trim();
            string advisor = string.IsNullOrWhiteSpace(_txtReportAdvisor?.Text) ? "未填寫" : _txtReportAdvisor.Text.Trim();
            DateTime reportDate = _dtReportDate != null ? _dtReportDate.Value.Date : DateTime.Now.Date;

            bool academicMode = _pdfTemplateMode == PdfTemplateMode.AcademicZh;

            // 封面
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.FromArgb(32, 41, 128, 185)), 0, 0, w, 240);
                string coverTitle = academicMode ? "個人房貸試算研究報告" : "個人房貸策略企劃簡報";
                string coverSub = academicMode ? "IEEE Access-Style Structured Report" : "Business Planning Brief";
                g.DrawString(coverTitle, new Font("微軟正黑體", 46, FontStyle.Bold), Brushes.Black, 110, 420);
                g.DrawString(coverSub, new Font("Segoe UI", 24, FontStyle.Italic), Brushes.DimGray, 170, 520);
                g.DrawString("學校：" + school, new Font("微軟正黑體", 20), Brushes.Black, 280, 650);
                g.DrawString("系所：" + dept, new Font("微軟正黑體", 20), Brushes.Black, 280, 695);
                g.DrawString("課程：" + course, new Font("微軟正黑體", 20), Brushes.Black, 280, 740);
                g.DrawString("學號：" + sid, new Font("微軟正黑體", 20), Brushes.Black, 280, 785);
                g.DrawString("姓名：" + sname, new Font("微軟正黑體", 20), Brushes.Black, 280, 830);
                g.DrawString("指導老師：" + advisor, new Font("微軟正黑體", 20), Brushes.Black, 280, 875);
                g.DrawString("日期：" + reportDate.ToString("yyyy/MM/dd"), new Font("微軟正黑體", 20), Brushes.Black, 280, 920);
                g.DrawString("第 " + pCover + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            // 目錄
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                g.DrawString("目錄", new Font("微軟正黑體", 38, FontStyle.Bold), Brushes.Black, 100, 90);
                var f = new Font("微軟正黑體", 20);
                g.DrawString("1. 摘要（Abstract）................................ " + pChAbstract, f, Brushes.Black, 120, 220);
                g.DrawString("2. 方法與結果（Method & Result）.................. " + pChMethodResult, f, Brushes.Black, 120, 280);
                g.DrawString("3. 圖表主視覺分析................................ " + pChMainChart, f, Brushes.Black, 120, 340);
                g.DrawString("4. 情境比較雷達圖與熱力矩陣...................... " + pChScenario, f, Brushes.Black, 120, 400);
                g.DrawString("5. 附錄（12-24期攤還表）......................... " + pAppendix, f, Brushes.Black, 120, 460);
                g.DrawString("6. 結語與參考資料................................ " + pFinal, f, Brushes.Black, 120, 520);
                g.DrawString("第 " + pToc + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            // 第1章：摘要
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                DrawChapterTitle(g, academicMode ? "第一章 摘要（Abstract）" : "第一章 Executive Summary", 80);
                int y = 180;
                if (academicMode)
                {
                    y = DrawSection(g, "摘要", "本研究建構一套房貸試算與決策支援系統，整合區域估價、攤還模擬、利率壓力測試與可視化分析。", y, w);
                    y = DrawSection(g, "關鍵詞", "Mortgage Calculator, Taiwan Housing, Stress Test, Amortization, Prepayment Strategy", y, w);
                    y = DrawSection(g, "研究背景", "以房貸為核心的長期負債管理，需同時考量利率波動、寬限期策略與現金流風險。", y, w);
                }
                else
                {
                    y = DrawSection(g, "摘要", "本企劃聚焦房貸成本、還款負擔與風險視覺化，支援快速比較與策略決策。", y, w);
                    y = DrawSection(g, "輸入假設", "房價 NT$ " + txtPrice.Text + "、自備款 " + txtDownPayment.Text + (cmbDownPaymentType.SelectedIndex == 0 ? "%" : "元") + "、利率 " + txtRate.Text + "%、年限 " + cmbTerm.Text + " 年、寬限 " + txtGrace.Text + " 年。", y, w);
                    y = DrawSection(g, "核心觀察", "目前方案月付 " + lblResultMonthly.Text + "，總利息 " + lblResultTotalInt.Text + "，有效清償月數約 " + _effectivePayoffMonths + " 月。", y, w);
                }
                g.DrawString("第 " + pChAbstract + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            // 第2章：方法、結果、結論、建議
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                DrawChapterTitle(g, academicMode ? "第二章 方法、結果、結論與建議" : "第二章 Method & Result", 80);
                int y = 180;
                y = DrawSection(g, "方法", "採用本息平均攤還公式，並於每年期末套用固定提前還款額；同時建立利率±1%壓力情境，觀察現金流變化。", y, w);
                y = DrawSection(g, "結果", "貸款總額 " + lblResultTotalLoan.Text + "，總還款 " + lblResultTotalRepay.Text + "，總利息 " + lblResultTotalInt.Text + "。", y, w);
                y = DrawSection(g, "結論", "當寬限期偏長時，前期負擔較低但後期月付跳升；提前還款可顯著縮短清償期並降低總利息。", y, w);
                y = DrawSection(g, "建議", "可搭配情境比較分頁建立 2~3 套策略，並每季更新利率後重跑壓力測試，以維持財務韌性。", y, w);

                g.DrawString("【利率壓力測試摘錄】", new Font("微軟正黑體", 18, FontStyle.Bold), Brushes.Black, 90, y); y += 44;
                foreach (var row in BuildStressRowsForReport().Take(4))
                {
                    g.DrawString("- " + row, new Font("Consolas", 12), Brushes.Black, 110, y);
                    y += 24;
                }

                g.DrawString("第 " + pChMethodResult + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            // 圖表頁：主圖
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                DrawChapterTitle(g, "第三章 圖表分析（主視覺）", 80);
                using (Bitmap chart = CaptureControlBitmap(picChart))
                {
                    Rectangle target = new Rectangle(90, 180, w - 180, 720);
                    g.DrawImage(chart, target);
                    g.DrawRectangle(Pens.LightGray, target);
                }
                DrawWrappedText(g, "互動功能：滑鼠滾輪可縮放時間區間，移動可檢視期數明細，點擊可切換輔助曲線。", new Font("微軟正黑體", 16), Brushes.DimGray, new RectangleF(90, 940, w - 180, 100));
                g.DrawString("第 " + pChMainChart + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            // 圖表頁：雷達圖/熱力矩陣
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                DrawChapterTitle(g, "第四章 情境比較視覺化", 80);
                DrawScenarioRadar(g, new Rectangle(80, 180, 500, 500));
                DrawScenarioHeatmap(g, new Rectangle(620, 180, 520, 500));
                g.DrawString("左圖：雷達圖（月付、總利息、總還款、清償月數）", new Font("微軟正黑體", 13), Brushes.DimGray, 90, 710);
                g.DrawString("右圖：熱力矩陣（數值越高顏色越深）", new Font("微軟正黑體", 13), Brushes.DimGray, 90, 740);
                g.DrawString("第 " + pChScenario + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            // 附錄：12~24 期完整表
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                DrawChapterTitle(g, "附錄 A：12~24 期攤還明細", 80);

                int y = 190;
                g.DrawString("期數    償還本金       償還利息       本息合計       剩餘本金", new Font("Consolas", 14, FontStyle.Bold), Brushes.Black, 90, y);
                y += 36;
                int startMonth = Math.Min(12, Math.Max(1, _schedule.Count));
                int endMonth = Math.Min(24, _schedule.Count);

                for (int m = startMonth; m <= endMonth; m++)
                {
                    var r = _schedule[m - 1];
                    string line = string.Format("{0,4}    {1,10}    {2,10}    {3,10}    {4,12}", r.Month, r.Principal, r.Interest, r.Payment, r.Balance);
                    g.DrawString(line, new Font("Consolas", 13), Brushes.Black, 90, y);
                    y += 28;
                    if (y > h - 120) break;
                }

                g.DrawString("第 " + pAppendix + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            // 最末頁：結語與參考資料模板
            using (Bitmap bmp = new Bitmap(w, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                DrawChapterTitle(g, "最終章 結語與參考資料", 80);

                int y = 190;
                y = DrawSection(g, "結語", "本系統已完成試算、可視化、情境比較與報告輸出，具備作業展示與實務延伸價值。", y, w);
                y = DrawSection(g, "後續工作", "可進一步串接官方開放資料 API、加入多人協作與雲端版本管理。", y, w);

                g.DrawString("【參考資料（模板）】", new Font("微軟正黑體", 20, FontStyle.Bold), Brushes.Black, 90, y); y += 44;
                g.DrawString("[1] 內政部不動產資訊平台，住宅價格統計。", new Font("微軟正黑體", 15), Brushes.Black, 110, y); y += 30;
                g.DrawString("[2] 金融監督管理委員會，房貸與利率相關公開資訊。", new Font("微軟正黑體", 15), Brushes.Black, 110, y); y += 30;
                g.DrawString("[3] Project Source Code and Local Experimental Results.", new Font("Segoe UI", 13), Brushes.Black, 110, y); y += 30;

                if (academicMode)
                {
                    y += 24;
                    g.DrawString("IEEE-Style Citation Notes:", new Font("Segoe UI", 14, FontStyle.Bold), Brushes.DimGray, 110, y); y += 28;
                    g.DrawString("- Use bracketed references [n] and maintain consistent bibliography order.", new Font("Segoe UI", 12), Brushes.DimGray, 130, y); y += 22;
                    g.DrawString("- Include DOI/URL and access date for all online resources.", new Font("Segoe UI", 12), Brushes.DimGray, 130, y);
                }

                g.DrawString("第 " + pFinal + " 頁 / 共 " + totalPages + " 頁", new Font("微軟正黑體", 14), Brushes.Gray, w - 240, h - 56);
                pages.Add(BitmapToJpegBytes(bmp, 90L));
            }

            return pages;
        }

        private byte[] BitmapToJpegBytes(Bitmap bmp, long quality)
        {
            using (var ms = new MemoryStream())
            {
                var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.MimeType == "image/jpeg");
                if (codec == null)
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                }
                else
                {
                    var ep = new EncoderParameters(1);
                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                    bmp.Save(ms, codec, ep);
                }
                return ms.ToArray();
            }
        }

        private Bitmap CaptureControlBitmap(Control c)
        {
            Bitmap bmp = new Bitmap(Math.Max(1, c.Width), Math.Max(1, c.Height));
            c.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
            return bmp;
        }

        private void DrawChapterTitle(Graphics g, string title, int y)
        {
            g.DrawString(title, new Font("微軟正黑體", 32, FontStyle.Bold), Brushes.Black, 80, y);
            g.DrawLine(new Pen(Color.FromArgb(41, 128, 185), 3), 80, y + 64, 1160, y + 64);
        }

        private void DrawWrappedText(Graphics g, string text, Font font, Brush brush, RectangleF rect)
        {
            StringFormat sf = new StringFormat();
            sf.Trimming = StringTrimming.Word;
            sf.FormatFlags = StringFormatFlags.LineLimit;
            g.DrawString(text, font, brush, rect, sf);
        }

        private int DrawSection(Graphics g, string title, string body, int y, int pageWidth)
        {
            g.DrawString("【" + title + "】", new Font("微軟正黑體", 20, FontStyle.Bold), Brushes.Black, 90, y);
            y += 42;
            RectangleF rect = new RectangleF(110, y, pageWidth - 220, 110);
            DrawWrappedText(g, body, new Font("微軟正黑體", 16), Brushes.Black, rect);
            return y + 112;
        }

        private List<string> BuildStressRowsForReport()
        {
            var rows = new List<string>();
            if (!double.TryParse(txtRate.Text, out double baseRate)) return rows;
            if (!int.TryParse(cmbTerm.Text, out int termYears)) return rows;
            if (!int.TryParse(txtGrace.Text, out int graceYears)) graceYears = 0;

            int totalMonths = termYears * 12;
            int remainMonths = Math.Max(1, totalMonths - (graceYears * 12));

            for (double d = -1.0; d <= 1.0001; d += 0.5)
            {
                double testRate = Math.Max(0.01, baseRate + d);
                double m = (testRate / 100.0) / 12.0;
                double p = totalLoan * m * Math.Pow(1 + m, remainMonths) / (Math.Pow(1 + m, remainMonths) - 1);
                rows.Add(string.Format("Rate {0:F2}% => Monthly NT$ {1:N0}", testRate, p));
            }
            return rows;
        }

        private void WritePdfFromJpegPages(string path, List<byte[]> pages)
        {
            var sb = new StringBuilder();
            var offsets = new List<int> { 0 };
            var raw = new List<byte[]>();

            int pageCount = pages.Count;
            int fontObj = 3 + pageCount * 3 + 1;

            raw.Add(Encoding.ASCII.GetBytes("1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"));

            StringBuilder kids = new StringBuilder();
            for (int i = 0; i < pageCount; i++)
            {
                int pageObj = 3 + (i * 3);
                kids.Append(pageObj).Append(" 0 R ");
            }
            raw.Add(Encoding.ASCII.GetBytes("2 0 obj << /Type /Pages /Kids [" + kids + "] /Count " + pageCount + " >> endobj\n"));

            for (int i = 0; i < pageCount; i++)
            {
                int pageObj = 3 + (i * 3);
                int contentObj = pageObj + 1;
                int imageObj = pageObj + 2;
                byte[] jpg = pages[i];

                raw.Add(Encoding.ASCII.GetBytes(pageObj + " 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /XObject << /Im" + (i + 1) + " " + imageObj + " 0 R >> /Font << /F1 " + fontObj + " 0 R >> >> /Contents " + contentObj + " 0 R >> endobj\n"));

                string content = "q\n595 0 0 842 0 0 cm\n/Im" + (i + 1) + " Do\nQ\nBT\n/F1 10 Tf\n500 20 Td\n(" + (i + 1) + ") Tj\nET\n";
                raw.Add(Encoding.ASCII.GetBytes(contentObj + " 0 obj << /Length " + Encoding.ASCII.GetByteCount(content) + " >> stream\n" + content + "endstream endobj\n"));

                var head = Encoding.ASCII.GetBytes(imageObj + " 0 obj << /Type /XObject /Subtype /Image /Width 1240 /Height 1754 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpg.Length + " >> stream\n");
                var tail = Encoding.ASCII.GetBytes("\nendstream endobj\n");
                byte[] imgObj = new byte[head.Length + jpg.Length + tail.Length];
                Buffer.BlockCopy(head, 0, imgObj, 0, head.Length);
                Buffer.BlockCopy(jpg, 0, imgObj, head.Length, jpg.Length);
                Buffer.BlockCopy(tail, 0, imgObj, head.Length + jpg.Length, tail.Length);
                raw.Add(imgObj);
            }

            raw.Add(Encoding.ASCII.GetBytes(fontObj + " 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n"));

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                byte[] header = Encoding.ASCII.GetBytes("%PDF-1.4\n");
                fs.Write(header, 0, header.Length);

                for (int i = 0; i < raw.Count; i++)
                {
                    offsets.Add((int)fs.Position);
                    fs.Write(raw[i], 0, raw[i].Length);
                }

                int xrefPos = (int)fs.Position;
                byte[] xrefHead = Encoding.ASCII.GetBytes("xref\n0 " + (raw.Count + 1) + "\n0000000000 65535 f \n");
                fs.Write(xrefHead, 0, xrefHead.Length);
                for (int i = 1; i <= raw.Count; i++)
                {
                    byte[] row = Encoding.ASCII.GetBytes(offsets[i].ToString("D10") + " 00000 n \n");
                    fs.Write(row, 0, row.Length);
                }
                byte[] trailer = Encoding.ASCII.GetBytes("trailer << /Size " + (raw.Count + 1) + " /Root 1 0 R >>\nstartxref\n" + xrefPos + "\n%%EOF");
                fs.Write(trailer, 0, trailer.Length);
            }
        }

        private void DrawScenarioRadar(Graphics g, Rectangle rect)
        {
            var data = BuildScenarioMetricMatrix();
            if (data.Count == 0) return;

            PointF c = new PointF(rect.X + rect.Width / 2f, rect.Y + rect.Height / 2f + 5);
            float minDim = Math.Min(rect.Width, rect.Height);
            float padding = (minDim > 150) ? 30 : 15;
            float r = minDim / 2f - padding;
            int axes = 4;
            string[] labels = { "月付", "總利息", "總還款", "清償月" };

            using (Pen p = new Pen(Color.LightGray))
            {
                for (int k = 1; k <= 4; k++)
                {
                    float rr = r * k / 4f;
                    PointF[] poly = new PointF[axes];
                    for (int i = 0; i < axes; i++)
                    {
                        double a = -Math.PI / 2 + i * 2 * Math.PI / axes;
                        poly[i] = new PointF(c.X + (float)(Math.Cos(a) * rr), c.Y + (float)(Math.Sin(a) * rr));
                    }
                    g.DrawPolygon(p, poly);
                }
            }

            for (int i = 0; i < axes; i++)
            {
                double a = -Math.PI / 2 + i * 2 * Math.PI / axes;
                PointF end = new PointF(c.X + (float)(Math.Cos(a) * r), c.Y + (float)(Math.Sin(a) * r));
                g.DrawLine(Pens.Gray, c, end);
                g.DrawString(labels[i], new Font("微軟正黑體", 8), Brushes.DimGray, end.X - 16, end.Y - 12);
            }

            Color[] colors = { Color.FromArgb(90, 41, 128, 185), Color.FromArgb(90, 231, 76, 60), Color.FromArgb(90, 46, 204, 113) };
            for (int s = 0; s < Math.Min(3, data.Count); s++)
            {
                float[] m = data[s];
                PointF[] poly = new PointF[axes];
                for (int i = 0; i < axes; i++)
                {
                    double a = -Math.PI / 2 + i * 2 * Math.PI / axes;
                    float rr = r * m[i];
                    poly[i] = new PointF(c.X + (float)(Math.Cos(a) * rr), c.Y + (float)(Math.Sin(a) * rr));
                }
                using (SolidBrush b = new SolidBrush(colors[s])) g.FillPolygon(b, poly);
                g.DrawPolygon(new Pen(Color.FromArgb(160, colors[s].R, colors[s].G, colors[s].B), 2), poly);
            }
        }

        private void DrawScenarioHeatmap(Graphics g, Rectangle rect)
        {
            var data = BuildScenarioMetricMatrix();
            if (data.Count == 0) return;

            int rows = data.Count;
            int cols = 4;
            float cw = rect.Width / (float)cols;
            float rh = rect.Height / (float)rows;
            string[] colName = { "月付", "總利息", "總還款", "清償月" };

            for (int c = 0; c < cols; c++)
                g.DrawString(colName[c], new Font("微軟正黑體", 8), Brushes.Gray, rect.X + c * cw + 2, rect.Y - 14);

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    float v = data[r][c];
                    Color color = Color.FromArgb(50 + (int)(v * 180), 231, 76, 60);
                    using (SolidBrush b = new SolidBrush(color))
                    {
                        g.FillRectangle(b, rect.X + c * cw + 1, rect.Y + r * rh + 1, cw - 2, rh - 2);
                    }
                    g.DrawString((v * 100).ToString("F0"), new Font("Consolas", 8, FontStyle.Bold), Brushes.White, rect.X + c * cw + 4, rect.Y + r * rh + 4);
                }
                g.DrawString("S" + (r + 1), new Font("Consolas", 8), Brushes.Gray, rect.X - 20, rect.Y + r * rh + 4);
            }
        }

        private List<float[]> BuildScenarioMetricMatrix()
        {
            var rows = new List<float[]>();
            var source = _scenarios.ToList();
            if (source.Count == 0 && totalLoan > 0)
            {
                source.Add(new ScenarioResult
                {
                    每月還款 = lblResultMonthly.Text,
                    總利息 = lblResultTotalInt.Text,
                    總還款 = lblResultTotalRepay.Text,
                    清償月數 = _effectivePayoffMonths + "月"
                });
            }

            if (source.Count == 0) return rows;

            var raw = source.Select(s => new[]
            {
                ParseMoney(s.每月還款),
                ParseMoney(s.總利息),
                ParseMoney(s.總還款),
                ParseMoney(s.清償月數)
            }).ToList();

            for (int c = 0; c < 4; c++)
            {
                double max = raw.Max(r => r[c]);
                if (max <= 0) max = 1;
                for (int r = 0; r < raw.Count; r++)
                {
                    if (rows.Count <= r) rows.Add(new float[4]);
                    rows[r][c] = (float)(raw[r][c] / max);
                }
            }
            return rows;
        }

        private double ParseMoney(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            var chars = text.Where(c => char.IsDigit(c) || c == '.').ToArray();
            if (chars.Length == 0) return 0;
            double.TryParse(new string(chars), out double v);
            return v;
        }

        private void picChart_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _showPrincipalShareLine = !_showPrincipalShareLine;
                picChart.Invalidate();
            }
        }

        private void picChart_MouseWheel(object sender, MouseEventArgs e)
        {
            if (_schedule == null || _schedule.Count == 0) return;

            int step = 24;
            int max = _schedule.Count;
            if (e.Delta > 0)
            {
                if (_chartZoomMonths == 0) _chartZoomMonths = Math.Max(24, max / 2);
                else _chartZoomMonths = Math.Max(24, _chartZoomMonths - step);
            }
            else
            {
                if (_chartZoomMonths == 0) return;
                _chartZoomMonths = Math.Min(max, _chartZoomMonths + step);
                if (_chartZoomMonths >= max) _chartZoomMonths = 0;
            }

            picChart.Invalidate();
        }

        private void picChart_MouseLeave(object sender, EventArgs e)
        {
            _chartTip.Hide(picChart);
        }

        private void picChart_MouseMove(object sender, MouseEventArgs e)
        {
            if (_balanceCurvePoints.Count == 0) return;

            int nearest = 0;
            float best = float.MaxValue;
            for (int i = 0; i < _balanceCurvePoints.Count; i++)
            {
                float d = Math.Abs(_balanceCurvePoints[i].X - e.X);
                if (d < best)
                {
                    best = d;
                    nearest = i;
                }
            }

            if (nearest >= 0 && nearest < _balanceCurveMonths.Count)
            {
                int month = _balanceCurveMonths[nearest];
                if (month >= 1 && month <= _schedule.Count)
                {
                    var row = _schedule[month - 1];
                    string text = string.Format("第 {0} 期\n剩餘本金: NT$ {1}\n本息合計: NT$ {2}", month, row.Balance, row.Payment);
                    _chartTip.Show(text, picChart, e.X + 12, e.Y + 12, 900);
                }
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            InitializeDefaults();
            _schedule.Clear();
            _bs.ResetBindings(false);
            totalLoan = 0; totalRepayment = 0; totalInterest = 0;
            errorProvider1.Clear();
            lblValidationHint.Text = "";
            rtbAI.Text = "未執行分析...\n提示：按「開始深度試算」後將顯示 AI 財務建議。";
            picChart.Invalidate();
            txtPrice.Focus();
        }

        private async void btnExport_Click(object sender, EventArgs e)
        {
            if (!_schedule.Any()) {
                MessageBox.Show("無資料匯出，請先執行試算。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "CSV檔案|*.csv", FileName = "房貸試算表.csv" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("期數(月),償還本金,償還利息,本息合計,剩餘本金");
                    foreach (var row in _schedule)
                    {
                        sb.AppendLine($"{row.Month},{row.Principal.Replace(",","")},{row.Interest.Replace(",","")},{row.Payment.Replace(",","")},{row.Balance.Replace(",","")}");
                    }
                    byte[] csvBytes = Encoding.UTF8.GetBytes(sb.ToString());
                    await Task.Run(() => File.WriteAllBytes(sfd.FileName, csvBytes));

                    string hashFile = sfd.FileName + ".sha256";
                    using (var sha = SHA256.Create())
                    {
                        byte[] hash = sha.ComputeHash(csvBytes);
                        string hashText = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                        File.WriteAllText(hashFile, hashText, Encoding.UTF8);
                    }

                    MessageBox.Show("匯出成功！\n已建立 SHA256 驗證檔：" + Path.GetFileName(hashFile), "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private async Task<string> BuildAiReportAsync(double price, double downPaymentAmount, int termYears, int graceYears, double annualRate, double monthlyPaymentNormal)
        {
            string scriptPath = Path.Combine(Application.StartupPath, "ai_advisor.py");
            EnsurePythonScript(scriptPath);

            string pythonReport = await TryRunPythonAi(scriptPath, price, downPaymentAmount, termYears, graceYears, annualRate, monthlyPaymentNormal, totalInterest);
            if (!string.IsNullOrWhiteSpace(pythonReport))
            {
                return "[AI 引擎：Python 外部模組]\n\n"
                    + pythonReport
                    + "\n\n"
                    + BuildAnalysisMethodSection()
                    + "\n"
                    + BuildRewardLearningSection(price, downPaymentAmount, termYears, graceYears, annualRate, monthlyPaymentNormal);
            }

            return BuildLocalAiReport(price, downPaymentAmount, termYears, graceYears, annualRate, monthlyPaymentNormal)
                + "\n"
                + BuildAnalysisMethodSection()
                + "\n"
                + BuildRewardLearningSection(price, downPaymentAmount, termYears, graceYears, annualRate, monthlyPaymentNormal);
        }

        private void EnsurePythonScript(string scriptPath)
        {
            var py = new StringBuilder();
            py.AppendLine("# -*- coding: utf-8 -*-");
            py.AppendLine("import sys");
            py.AppendLine("price=float(sys.argv[1])");
            py.AppendLine("down=float(sys.argv[2])");
            py.AppendLine("term=int(sys.argv[3])");
            py.AppendLine("grace=int(sys.argv[4])");
            py.AppendLine("rate=float(sys.argv[5])");
            py.AppendLine("monthly=float(sys.argv[6])");
            py.AppendLine("interest=float(sys.argv[7])");
            py.AppendLine("loan=price-down");
            py.AppendLine("ltv=loan/price*100 if price>0 else 0");
            py.AppendLine("print('【Python AI 財務建議】')");
            py.AppendLine("print(f'貸款成數(LTV): {ltv:.1f}%')");
            py.AppendLine("print(f'寬限期: {grace} 年 / 貸款年限: {term} 年')");
            py.AppendLine("if ltv > 80: print('風險提醒：貸款成數偏高，建議提高預備金。')");
            py.AppendLine("if grace > 0: print('提醒：寬限期後月付金通常會上升，請預留現金流。')");
            py.AppendLine("print(f'月付金約: NT$ {monthly:,.0f}')");
            py.AppendLine("print(f'總利息約: NT$ {interest:,.0f}')");
            py.AppendLine("print('建議：若有額外獎金，可優先提前償還本金以降低總利息。')");
            py.AppendLine("print('分析方式：Python 規則模型 + 風險權重評分。')");

            File.WriteAllText(scriptPath, py.ToString(), Encoding.UTF8);
        }

        private async Task<string> TryRunPythonAi(string scriptPath, double price, double downPaymentAmount, int termYears, int graceYears, double annualRate, double monthlyPaymentNormal, double interest)
        {
            try
            {
                string args = string.Format("\"{0}\" {1} {2} {3} {4} {5} {6} {7}",
                    scriptPath,
                    price.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    downPaymentAmount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    termYears,
                    graceYears,
                    annualRate.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    monthlyPaymentNormal.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    interest.ToString(System.Globalization.CultureInfo.InvariantCulture));

                var psi = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = "-X utf8 " + args,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
                psi.EnvironmentVariables["PYTHONUTF8"] = "1";

                string stdout = string.Empty;
                string stderr = string.Empty;

                await Task.Run(() =>
                {
                    using (var p = Process.Start(psi))
                    {
                        stdout = p.StandardOutput.ReadToEnd();
                        stderr = p.StandardError.ReadToEnd();
                        p.WaitForExit(2500);
                    }
                });

                if (!string.IsNullOrWhiteSpace(stdout))
                {
                    string clean = stdout.Trim();
                    if (!clean.Contains("�"))
                    {
                        return clean;
                    }
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private string BuildLocalAiReport(double price, double downPaymentAmount, int termYears, int graceYears, double annualRate, double monthlyPaymentNormal)
        {
            double loan = price - downPaymentAmount;
            double ltv = price <= 0 ? 0 : loan / price * 100;

            var sb = new StringBuilder();
            sb.AppendLine("[AI 引擎：內建 NLP 規則模型]");
            sb.AppendLine();
            sb.AppendLine("1) 貸款結構分析");
            sb.AppendLine(string.Format("- 貸款成數(LTV)：{0:F1}%", ltv));
            sb.AppendLine(ltv > 80
                ? "- 解讀：成數偏高，建議提高自備款或保留至少 6 個月緊急預備金。"
                : "- 解讀：成數健康，整體風險相對可控。\n");
            sb.AppendLine();
            sb.AppendLine("2) 現金流壓力測試");
            sb.AppendLine(string.Format("- 寬限期：{0} 年，貸款年限：{1} 年", graceYears, termYears));
            sb.AppendLine(string.Format("- 預估月付金：NT$ {0:N0}", monthlyPaymentNormal));
            sb.AppendLine(string.Format("- 建議家庭月收入至少：NT$ {0:N0}（以 33% 還款占比估算）", monthlyPaymentNormal * 3));
            sb.AppendLine();
            sb.AppendLine("3) 利率敏感度提醒");
            sb.AppendLine(string.Format("- 目前年利率：{0:F2}%", annualRate));
            sb.AppendLine("- 若市場利率變動，建議每季重新試算並比較提前還款效益。\n");
            sb.AppendLine();
            sb.AppendLine("補充：若電腦有安裝 Python，本系統會自動切換為 Python 外部 AI 模組分析。\n");

            return sb.ToString();
        }

        private string BuildAnalysisMethodSection()
        {
            var sb = new StringBuilder();
            sb.AppendLine("【分析方式說明】");
            sb.AppendLine("- 模型A：財務規則引擎（LTV、還款占比、寬限期壓力）");
            sb.AppendLine("- 模型B：NLP模板生成（轉換為可讀建議）");
            sb.AppendLine("- 模型C：跨語言協作（可用時呼叫 Python 分析模組）");
            sb.AppendLine("- 安全：匯出支援 SHA256 完整性驗證");
            return sb.ToString();
        }

        private string BuildRewardLearningSection(double price, double downPaymentAmount, int termYears, int graceYears, double annualRate, double monthlyPaymentNormal)
        {
            double loan = price - downPaymentAmount;
            double riskPenalty = (graceYears * 1.8) + (annualRate * 1.2);

            double rewardKeep = 65 - riskPenalty;
            double rewardRaiseDown = 78 - (loan / price * 35) - (annualRate * 0.8);
            double rewardShortenTerm = 82 - (termYears * 1.2) - (graceYears * 1.3);
            double rewardPrepay = 80 - (annualRate * 1.1) - (termYears * 0.4);

            var options = new List<Tuple<string, double>>
            {
                Tuple.Create("策略1：維持現況", rewardKeep),
                Tuple.Create("策略2：提高自備款 5%", rewardRaiseDown),
                Tuple.Create("策略3：年限縮短 5 年", rewardShortenTerm),
                Tuple.Create("策略4：每年固定提前還款", rewardPrepay)
            };

            var best = options.OrderByDescending(x => x.Item2).First();

            var sb = new StringBuilder();
            sb.AppendLine("【Reward-based 建議引擎（類獎勵式學習）】");
            sb.AppendLine("- 系統以『風險最小化 + 現金流穩定 + 利息壓低』做獎勵函數。\n");
            foreach (var item in options)
            {
                sb.AppendLine(string.Format("  {0} => Reward Score: {1:F1}", item.Item1, item.Item2));
            }
            sb.AppendLine();
            sb.AppendLine("- 最佳策略預估：" + best.Item1);
            sb.AppendLine("- 建議：先執行最佳策略，再每季重新試算並更新策略分數。\n");
            sb.AppendLine(string.Format("- 預估舒適月收入門檻：NT$ {0:N0}", monthlyPaymentNormal * 3));
            return sb.ToString();
        }

        private void picChart_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.White);

            if (totalRepayment <= 0)
            {
                StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString("暫無圖表數據\n請點擊計算", new Font("微軟正黑體", 12F), Brushes.Gray, picChart.ClientRectangle, sf);
                return;
            }

            // Dual Charts - Left Donut, Right Area Chart
            Rectangle rectPie = new Rectangle(10, 20, 160, 160);
            float principalAngle = (float)(totalLoan / totalRepayment * 360.0);
            float interestAngle = 360f - principalAngle;

            using (SolidBrush pBrush = new SolidBrush(Color.FromArgb(41, 128, 185)))
            using (SolidBrush iBrush = new SolidBrush(Color.FromArgb(192, 57, 43)))
            {
                g.FillPie(pBrush, rectPie, -90, principalAngle);
                g.FillPie(iBrush, rectPie, -90 + principalAngle, interestAngle);
                
                // Donut Hole
                using (SolidBrush whiteBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(whiteBrush, new Rectangle(rectPie.X + 40, rectPie.Y + 40, rectPie.Width - 80, rectPie.Height - 80));
                }

                // Legend
                int legendX = rectPie.Right + 20;
                int legendY = 60;
                g.FillRectangle(pBrush, legendX, legendY, 15, 15);
                g.DrawString($"本金: {(totalLoan/totalRepayment*100):F1}%", this.Font, Brushes.Black, legendX + 20, legendY - 2);

                g.FillRectangle(iBrush, legendX, legendY + 30, 15, 15);
                g.DrawString($"利息: {(totalInterest/totalRepayment*100):F1}%", this.Font, Brushes.Black, legendX + 20, legendY + 28);
            }

            // Draw enhanced + interactive charts
            int chartX = rectPie.Right + 120;
            int chartWidth = picChart.Width - chartX - 20;
            int chartY = 28;
            int chartHeight = 92;
            int miniY = chartY + chartHeight + 28;
            int miniH = 50;

            if (_schedule.Count > 0 && chartWidth > 0)
            {
                _balanceCurvePoints.Clear();
                _balanceCurveMonths.Clear();
                g.DrawString("本金餘額趨勢（滑鼠滾輪縮放 / 點擊切線）", new Font("微軟正黑體", 8.5F), Brushes.Gray, chartX, 8);

                // Grid
                using (Pen gridPen = new Pen(Color.FromArgb(225, 225, 225)))
                {
                    for (int i = 0; i <= 4; i++)
                    {
                        int gy = chartY + (chartHeight * i / 4);
                        g.DrawLine(gridPen, chartX, gy, chartX + chartWidth, gy);
                    }
                }

                int startIdx = 0;
                int endIdx = _schedule.Count - 1;
                if (_chartZoomMonths > 0)
                {
                    startIdx = Math.Max(0, _schedule.Count - _chartZoomMonths);
                }
                int visibleCount = endIdx - startIdx + 1;

                if (_lastGraceMonths > 0 && _lastTotalMonths > 0 && visibleCount > 1)
                {
                    int graceEnd = Math.Min(_lastGraceMonths, _schedule.Count);
                    if (graceEnd > startIdx)
                    {
                        float graceX = chartX + (float)(graceEnd - startIdx) / visibleCount * chartWidth;
                        using (SolidBrush graceBrush = new SolidBrush(Color.FromArgb(32, 241, 196, 15)))
                        {
                            g.FillRectangle(graceBrush, chartX, chartY, Math.Max(1, graceX - chartX), chartHeight);
                        }
                        g.DrawString("寬限期", new Font("微軟正黑體", 8F), Brushes.Gray, chartX + 2, chartY + 2);
                    }
                }

                int sampleCount = Math.Min(160, visibleCount);
                List<PointF> balanceLine = new List<PointF>(sampleCount);
                List<PointF> paidLine = new List<PointF>(sampleCount);

                for (int s = 0; s < sampleCount; s++)
                {
                    int idx = startIdx + (int)Math.Round((double)s * (visibleCount - 1) / Math.Max(1, sampleCount - 1));
                    idx = Math.Min(endIdx, Math.Max(startIdx, idx));
                    double bal = GetDouble(_schedule[idx].Balance);
                    double remainRatio = totalLoan <= 0 ? 0 : (bal / totalLoan);
                    double principal = GetDouble(_schedule[idx].Principal);
                    double payment = GetDouble(_schedule[idx].Payment);
                    double principalShare = payment <= 0 ? 0 : principal / payment;

                    float x = chartX + (float)s / Math.Max(1, sampleCount - 1) * chartWidth;
                    float yBalance = chartY + (float)((1 - remainRatio) * chartHeight);
                    float yPaid = chartY + (float)((1 - principalShare) * chartHeight);

                    balanceLine.Add(new PointF(x, yBalance));
                    if (_showPrincipalShareLine)
                        paidLine.Add(new PointF(x, yPaid));

                    _balanceCurvePoints.Add(new PointF(x, yBalance));
                    _balanceCurveMonths.Add(idx + 1);
                }

                // area fill under balance line
                List<PointF> area = new List<PointF>();
                area.Add(new PointF(chartX, chartY + chartHeight));
                area.AddRange(balanceLine);
                area.Add(new PointF(chartX + chartWidth, chartY + chartHeight));

                using (GraphicsPath areaPath = new GraphicsPath())
                {
                    areaPath.AddLines(area.ToArray());
                    using (LinearGradientBrush lgb = new LinearGradientBrush(new Rectangle(chartX, chartY, chartWidth, chartHeight + 1), Color.FromArgb(100, 41, 128, 185), Color.Transparent, 90f))
                    {
                        g.FillPath(lgb, areaPath);
                    }
                }

                using (Pen balancePen = new Pen(Color.FromArgb(41, 128, 185), 2.2f))
                using (Pen paidPen = new Pen(Color.FromArgb(230, 126, 34), 1.8f))
                {
                    paidPen.DashStyle = DashStyle.Dash;
                    paidPen.DashOffset = _chartDashOffset;
                    if (balanceLine.Count > 1) g.DrawLines(balancePen, balanceLine.ToArray());
                    if (_showPrincipalShareLine && paidLine.Count > 1) g.DrawLines(paidPen, paidLine.ToArray());
                }

                // 近12期本息構成 mini bars
                int bars = Math.Min(12, visibleCount);
                float barsTotalW = Math.Max(20, chartWidth - 380);
                if (bars > 0)
                {
                    g.DrawString("近12期本息構成", new Font("微軟正黑體", 8F), Brushes.Gray, chartX, miniY - 12);
                    float bw = barsTotalW / (float)bars;
                    for (int b = 0; b < bars; b++)
                    {
                        int idx = endIdx - bars + 1 + b;
                        if (idx < 0 || idx >= _schedule.Count) continue;
                        double pmt = Math.Max(1, GetDouble(_schedule[idx].Payment));
                        double prin = Math.Max(0, GetDouble(_schedule[idx].Principal));
                        double intr = Math.Max(0, GetDouble(_schedule[idx].Interest));

                        float x = chartX + b * bw + 2;
                        float hP = (float)(miniH * (prin / pmt));
                        float hI = (float)(miniH * (intr / pmt));

                        using (SolidBrush bp = new SolidBrush(Color.FromArgb(70, 46, 204, 113)))
                        using (SolidBrush bi = new SolidBrush(Color.FromArgb(90, 231, 76, 60)))
                        {
                            g.FillRectangle(bp, x, miniY + miniH - hP, bw - 4, hP);
                            g.FillRectangle(bi, x, miniY + miniH - hP - hI, bw - 4, hI);
                        }
                    }
                }

                // 右下角再加情境雷達 + 熱力矩陣 (加寬避免文字重疊)
                int radarW = 100;
                int radarH = 75;
                DrawScenarioRadar(g, new Rectangle((int)(chartX + chartWidth - 260), miniY - 4, radarW, radarH));
                DrawScenarioHeatmap(g, new Rectangle((int)(chartX + chartWidth - 140), miniY - 4, 130, radarH));

                // 加入 DTI 儀表板 (月付金 / 月收入)
                double monthlyIncome = _numMonthlyIncome != null ? (double)_numMonthlyIncome.Value : 0;
                if (monthlyIncome > 0)
                {
                    double currentMonthly = GetDouble(lblResultMonthly.Text);
                    double dti = currentMonthly / Math.Max(1, monthlyIncome);

                    int dtiX = (int)(chartX + chartWidth - 360);
                    int dtiY = miniY + 10;
                    g.DrawString("收支比 (DTI)", new Font("微軟正黑體", 8F), Brushes.DimGray, dtiX, dtiY - 15);

                    Color dtiColor = Color.FromArgb(46, 204, 113);
                    string riskLvl = "正常";
                    if (dti > 0.6) { dtiColor = Color.FromArgb(231, 76, 60); riskLvl = "危險"; }
                    else if (dti > 0.4) { dtiColor = Color.FromArgb(241, 196, 15); riskLvl = "偏高"; }

                    using (Pen bgPen = new Pen(Color.FromArgb(230, 230, 230), 6))
                    using (Pen fgPen = new Pen(dtiColor, 6))
                    {
                        g.DrawArc(bgPen, dtiX + 10, dtiY, 40, 40, 135, 270);
                        g.DrawArc(fgPen, dtiX + 10, dtiY, 40, 40, 135, (float)Math.Min(270, dti * 270));
                    }
                    StringFormat dtiSf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    g.DrawString($"{(dti*100):F0}%", new Font("Consolas", 8F, FontStyle.Bold), Brushes.Black, new Rectangle(dtiX + 10, dtiY, 40, 40), dtiSf);
                    g.DrawString(riskLvl, new Font("微軟正黑體", 8F), new SolidBrush(dtiColor), dtiX + 18, dtiY + 45);
                }

                string zoomText = _chartZoomMonths == 0 ? "全期間" : ("最近 " + _chartZoomMonths + " 期");
                g.DrawString("藍:剩餘本金 橘虛線:本金占比 | 視窗: " + zoomText, new Font("微軟正黑體", 8F), Brushes.Gray, chartX, chartY + chartHeight + 2);
            }
        }

        public class AmortizationItem
        {
            [DisplayName("期數(月)")] public int Month { get; set; }
            [DisplayName("償還本金")] public string Principal { get; set; }
            [DisplayName("償還利息")] public string Interest { get; set; }
            [DisplayName("本息合計")] public string Payment { get; set; }
            [DisplayName("剩餘本金")] public string Balance { get; set; }
        }
    }
}
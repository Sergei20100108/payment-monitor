using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PaymentMonitor.Ui
{
    /// <summary>
    /// Шапка окна: подпись слева и круглые кнопки (свернуть, развернуть/восстановить, закрыть).
    /// Окно перетаскивается мышью за шапку, двойной щелчок по шапке разворачивает окно.
    /// </summary>
    public sealed class TitleBar : Panel
    {
        private const int WM_NCLBUTTONDOWN = 0x00A1;
        private const int HTCAPTION = 2;

        private readonly Label _caption;
        private readonly RoundButton _btnMin;
        private readonly RoundButton _btnMax;
        private readonly RoundButton _btnClose;

        private DateTime _lastClick = DateTime.MinValue;
        private bool _maximized;

        public event EventHandler MinimizeClicked;
        public event EventHandler MaximizeClicked;
        public event EventHandler CloseClicked;

        public TitleBar()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            DoubleBuffered = true;
            Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            Padding = new Padding(0);
            Margin = new Padding(0);

            _caption = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(45, 50, 58),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Padding = new Padding(14, 0, 0, 0),
                Margin = new Padding(0),
                Text = ""
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 112,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            _btnClose = new RoundButton
            {
                Glyph = RoundButton.GlyphKind.Close,
                IsAccent = true,
                AccessibleName = "Закрыть окно"
            };
            _btnMax = new RoundButton
            {
                Glyph = RoundButton.GlyphKind.Maximize,
                AccessibleName = "Развернуть окно"
            };
            _btnMin = new RoundButton
            {
                Glyph = RoundButton.GlyphKind.Minimize,
                AccessibleName = "Свернуть окно"
            };

            _btnClose.Click += delegate { if (CloseClicked != null) CloseClicked(this, EventArgs.Empty); };
            _btnMax.Click += delegate { if (MaximizeClicked != null) MaximizeClicked(this, EventArgs.Empty); };
            _btnMin.Click += delegate { if (MinimizeClicked != null) MinimizeClicked(this, EventArgs.Empty); };

            buttons.Controls.Add(_btnClose);
            buttons.Controls.Add(_btnMax);
            buttons.Controls.Add(_btnMin);

            // порядок докирования: кнопки прижаты к правому краю, подпись занимает остаток
            Controls.Add(_caption);
            Controls.Add(buttons);

            _caption.MouseDown += CaptionMouseDown;
        }

        /// <summary>Подпись в шапке окна.</summary>
        public string Caption
        {
            get { return _caption.Text; }
            set { _caption.Text = value ?? ""; }
        }

        /// <summary>Окно развёрнуто — на кнопке рисуется «восстановить».</summary>
        public bool IsMaximized
        {
            get { return _maximized; }
            set
            {
                if (_maximized == value) return;
                _maximized = value;
                _btnMax.Glyph = value ? RoundButton.GlyphKind.Restore : RoundButton.GlyphKind.Maximize;
                _btnMax.AccessibleName = value ? "Восстановить окно" : "Развернуть окно";
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left) HandleTitleBarClick(e.Clicks);
        }

        private void CaptionMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) HandleTitleBarClick(e.Clicks);
        }

        /// <summary>Первый щелчок — перетаскивание, второй — разворачивание окна.</summary>
        private void HandleTitleBarClick(int clicks)
        {
            DateTime now = DateTime.Now;
            bool isDouble = clicks >= 2 ||
                            (now - _lastClick).TotalMilliseconds <= SystemInformation.DoubleClickTime;
            _lastClick = now;
            if (isDouble)
            {
                ToggleMaximize();
                return;
            }
            BeginDrag();
        }

        private void ToggleMaximize()
        {
            if (MaximizeClicked != null) MaximizeClicked(this, EventArgs.Empty);
        }

        /// <summary>Передать перетаскивание системе (работает и «прилипание» к краям экрана).</summary>
        private void BeginDrag()
        {
            Control top = FindTopForm();
            if (top == null) return;
            ReleaseCapture();
            SendMessage(top.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
        }

        private Control FindTopForm()
        {
            Control c = this;
            while (c != null && !(c is Form)) c = c.Parent;
            return c;
        }

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    }
}

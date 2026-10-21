using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PaymentMonitor.Ui
{
    /// <summary>
    /// Круглая кнопка с рисунком: сворачивание, разворачивание, закрытие окна.
    /// Всё рисуется вручную — стандартный вид кнопок не подходит.
    /// </summary>
    public sealed class RoundButton : Control
    {
        /// <summary>Рисунок на кнопке.</summary>
        public enum GlyphKind { Minimize, Maximize, Restore, Close }

        private GlyphKind _glyph = GlyphKind.Minimize;
        private bool _hover;
        private bool _pressed;

        public RoundButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            TabStop = false;
            Size = new Size(28, 28);
            Margin = new Padding(4);
            Cursor = Cursors.Hand;

            BackColor = Color.White;                          // фон вокруг круга
            FaceColor = Color.FromArgb(233, 236, 241);        // кружок в покое
            HoverFaceColor = Color.FromArgb(211, 217, 226);   // кружок под курсором
            ForeColor = Color.FromArgb(56, 62, 72);           // рисунок
            AccentFaceColor = Color.FromArgb(214, 45, 32);    // кружок «закрыть» под курсором
            AccentForeColor = Color.White;
        }

        /// <summary>Рисунок на кнопке.</summary>
        public GlyphKind Glyph
        {
            get { return _glyph; }
            set { if (_glyph != value) { _glyph = value; Invalidate(); } }
        }

        /// <summary>Кружок под курсором красится акцентным цветом (кнопка «закрыть»).</summary>
        public bool IsAccent { get; set; }

        /// <summary>Цвет кружка в покое.</summary>
        public Color FaceColor { get; set; }

        /// <summary>Цвет кружка под курсором.</summary>
        public Color HoverFaceColor { get; set; }

        /// <summary>Акцентный цвет кружка (кнопка «закрыть»).</summary>
        public Color AccentFaceColor { get; set; }

        /// <summary>Цвет рисунка на акцентном кружке.</summary>
        public Color AccentForeColor { get; set; }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            _hover = true;
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hover = false;
            _pressed = false;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left)
            {
                _pressed = true;
                Invalidate();
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (_pressed)
            {
                _pressed = false;
                Invalidate();
            }
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            _hover = false;
            _pressed = false;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(BackColor);

            bool accent = IsAccent && _hover && Enabled;
            Color face;
            Color fore;
            if (accent)
            {
                face = AccentFaceColor;
                fore = AccentForeColor;
            }
            else
            {
                face = _hover ? HoverFaceColor : FaceColor;
                fore = Enabled ? ForeColor : Color.FromArgb(160, 160, 160);
            }
            if (_pressed) face = ControlPaint.Dark(face, 0.08F);

            int d = Math.Min(Width, Height);
            var circle = new RectangleF((Width - d) / 2F, (Height - d) / 2F, d, d);
            using (var brush = new SolidBrush(face))
                g.FillEllipse(brush, circle);

            using (var pen = new Pen(fore, 1.5F))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                DrawGlyph(g, pen);
            }
        }

        private void DrawGlyph(Graphics g, Pen pen)
        {
            float cx = Width / 2F;
            float cy = Height / 2F;
            const float s = 5F;

            switch (_glyph)
            {
                case GlyphKind.Minimize:
                    g.DrawLine(pen, cx - s, cy + s - 2F, cx + s, cy + s - 2F);
                    break;

                case GlyphKind.Maximize:
                    g.DrawRectangle(pen, cx - s, cy - s, s * 2F, s * 2F);
                    break;

                case GlyphKind.Restore:
                    g.DrawRectangle(pen, cx - s + 2F, cy - s, s * 2F - 2F, s * 2F - 2F);
                    g.DrawRectangle(pen, cx - s, cy - s + 2F, s * 2F - 2F, s * 2F - 2F);
                    break;

                case GlyphKind.Close:
                    g.DrawLine(pen, cx - s, cy - s, cx + s, cy + s);
                    g.DrawLine(pen, cx + s, cy - s, cx - s, cy + s);
                    break;
            }
        }
    }
}

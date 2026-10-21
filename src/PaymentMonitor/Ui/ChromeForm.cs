using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PaymentMonitor.Ui
{
    /// <summary>
    /// Форма без системной рамки: своя шапка с круглыми кнопками окна,
    /// рамка для растягивания мышью и корректное разворачивание на весь рабочий стол.
    /// </summary>
    public abstract class ChromeForm : Form
    {
        /// <summary>Ширина невидимой рамки, за которую окно тянется мышью.</summary>
        private const int ResizeBorder = 6;

        private const int WM_NCHITTEST = 0x0084;
        private const int WM_GETMINMAXINFO = 0x0024;
        private const int WM_SIZE = 0x0005;

        private const int SIZE_RESTORED = 0;
        private const int SIZE_MAXIMIZED = 2;

        private const int HTCLIENT = 1;
        private const int HTLEFT = 10;
        private const int HTRIGHT = 11;
        private const int HTTOP = 12;
        private const int HTTOPLEFT = 13;
        private const int HTTOPRIGHT = 14;
        private const int HTBOTTOM = 15;
        private const int HTBOTTOMLEFT = 16;
        private const int HTBOTTOMRIGHT = 17;

        private const int MONITOR_DEFAULTTONEAREST = 2;

        private readonly Panel _frame;
        private readonly TableLayoutPanel _root;

        /// <summary>Шапка окна с круглыми кнопками.</summary>
        protected readonly TitleBar TitleBar;

        /// <summary>Область, куда форма размещает своё содержимое (под шапкой).</summary>
        protected readonly Panel Content;

        protected ChromeForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            MaximizeBox = true;
            MinimizeBox = true;
            Padding = new Padding(ResizeBorder);
            BackColor = Color.FromArgb(226, 230, 236);
            DoubleBuffered = true;

            _frame = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Padding = new Padding(1),
                Margin = new Padding(0)
            };

            _root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            _root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            TitleBar = new TitleBar();
            Content = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            _root.Controls.Add(TitleBar, 0, 0);
            _root.Controls.Add(Content, 0, 1);
            _frame.Controls.Add(_root);
            Controls.Add(_frame);

            TitleBar.MinimizeClicked += delegate { WindowState = FormWindowState.Minimized; };
            TitleBar.MaximizeClicked += delegate { ToggleMaximize(); };
            TitleBar.CloseClicked += delegate { Close(); };

            TextChanged += delegate { TitleBar.Caption = Text; };
        }

        /// <summary>Развернуть окно или вернуть его в обычный размер.</summary>
        protected void ToggleMaximize()
        {
            if (WindowState == FormWindowState.Maximized)
            {
                WindowState = FormWindowState.Normal;
                TitleBar.IsMaximized = false;
                return;
            }
            WindowState = FormWindowState.Normal;   // чтобы запомнить обычные размеры
            WindowState = FormWindowState.Maximized;
            TitleBar.IsMaximized = true;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_NCHITTEST)
            {
                base.WndProc(ref m);
                if (WindowState == FormWindowState.Normal)
                {
                    int lp = m.LParam.ToInt32();
                    var screen = new Point((short)(lp & 0xFFFF), (short)((lp >> 16) & 0xFFFF));
                    m.Result = HitTest(PointToClient(screen));
                }
                return;
            }

            if (m.Msg == WM_GETMINMAXINFO)
            {
                base.WndProc(ref m);
                ApplyMaximizeBounds(ref m);
                return;
            }

            if (m.Msg == WM_SIZE)
            {
                base.WndProc(ref m);
                // окно развернули или свернули другим способом (например, клавишами Win+↑)
                int sizeType = m.WParam.ToInt32();
                TitleBar.IsMaximized = sizeType == SIZE_MAXIMIZED;
                return;
            }

            base.WndProc(ref m);
        }

        /// <summary>Края окна: там возвращаем системе коды растягивания.</summary>
        private IntPtr HitTest(Point p)
        {
            int w = ClientSize.Width;
            int h = ClientSize.Height;
            bool left = p.X < ResizeBorder;
            bool right = p.X >= w - ResizeBorder;
            bool top = p.Y < ResizeBorder;
            bool bottom = p.Y >= h - ResizeBorder;

            if (top && left) return (IntPtr)HTTOPLEFT;
            if (top && right) return (IntPtr)HTTOPRIGHT;
            if (bottom && left) return (IntPtr)HTBOTTOMLEFT;
            if (bottom && right) return (IntPtr)HTBOTTOMRIGHT;
            if (top) return (IntPtr)HTTOP;
            if (bottom) return (IntPtr)HTBOTTOM;
            if (left) return (IntPtr)HTLEFT;
            if (right) return (IntPtr)HTRIGHT;
            return (IntPtr)HTCLIENT;
        }

        /// <summary>Развёрнутое окно занимает рабочий стол, а не весь экран.</summary>
        private void ApplyMaximizeBounds(ref Message m)
        {
            IntPtr monitor = MonitorFromWindow(Handle, MONITOR_DEFAULTTONEAREST);
            if (monitor == IntPtr.Zero) return;

            var info = new MONITORINFO();
            info.cbSize = Marshal.SizeOf(typeof(MONITORINFO));
            if (!GetMonitorInfo(monitor, ref info)) return;

            var mmi = (MINMAXINFO)Marshal.PtrToStructure(m.LParam, typeof(MINMAXINFO));
            mmi.ptMaxPosition.x = info.rcWork.Left - info.rcMonitor.Left;
            mmi.ptMaxPosition.y = info.rcWork.Top - info.rcMonitor.Top;
            mmi.ptMaxSize.x = info.rcWork.Right - info.rcWork.Left;
            mmi.ptMaxSize.y = info.rcWork.Bottom - info.rcWork.Top;
            mmi.ptMaxTrackSize.x = mmi.ptMaxSize.x;
            mmi.ptMaxTrackSize.y = mmi.ptMaxSize.y;
            Marshal.StructureToPtr(mmi, m.LParam, true);
            m.Result = (IntPtr)0;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MINMAXINFO
        {
            public POINT ptReserved;
            public POINT ptMaxSize;
            public POINT ptMaxPosition;
            public POINT ptMinTrackSize;
            public POINT ptMaxTrackSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public int dwFlags;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, int dwFlags);

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);
    }
}

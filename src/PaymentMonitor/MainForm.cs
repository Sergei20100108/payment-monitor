using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PaymentMonitor.Core;
using PaymentMonitor.Ui;

namespace PaymentMonitor
{
    /// <summary>Главное окно: импорт выписок, мониторинг папки, таблица платежей, история, журнал.</summary>
    public sealed class MainForm : ChromeForm
    {
        private AppSettings _settings;
        private string _dbPath;
        private Database _db;
        private Importer _importer;
        private bool _busy;

        private readonly Timer _watchTimer = new Timer();
        private PaymentStatus[] _allStatuses;

        // элементы управления
        private TableLayoutPanel _row1;
        private TableLayoutPanel _row2;
        private Button _btnImport;
        private Button _btnBrowse;
        private Button _btnWatch;
        private Button _btnRefresh;
        private Button _btnEditor;
        private TextBox _txtFolder;
        private TextBox _txtBank;
        private TextBox _txtSearch;
        private NumericUpDown _nudInterval;
        private ComboBox _cmbStatus;
        private SplitContainer _splitMain;
        private SplitContainer _splitBottom;
        private DataGridView _gridPayments;
        private DataGridView _gridHistory;
        private TextBox _txtLog;
        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _lblMsg;
        private ToolStripStatusLabel _lblCounts;
        private ToolStripStatusLabel _lblDb;

        private static readonly Font StatusFont =
            new Font(SystemFonts.MessageBoxFont, FontStyle.Bold);

        public MainForm()
        {
            try
            {
                _settings = SettingsStore.Load();
                _dbPath = SettingsStore.ResolveDbPath(_settings);
                _db = new Database(_dbPath);
                _importer = new Importer(_dbPath, _settings.Bank);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось инициализировать базу данных:\r\n\r\n" + ex.Message,
                    "PaymentMonitor", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(3);
                return;
            }

            _allStatuses = StatusText.All();

            InitializeComponent();
            ApplySettingsToControls();
            AttachHandlers();

            Reload();
            AppendLog("Готово. База данных: " + _dbPath);
            if (!string.IsNullOrEmpty(_settings.WatchFolder))
                AppendLog("Папка мониторинга: " + _settings.WatchFolder);
        }

        // ── построение интерфейса ───────────────────────────────────────

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "Мониторинг платежей — выписки Excel";
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(1140, 720);
            MinimumSize = new Size(920, 540);
            StartPosition = FormStartPosition.CenterScreen;

            // корень: три строки — верхняя панель, содержимое, статус-строка
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 78F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));

            BuildTopPanel();
            root.Controls.Add(_row1.Parent, 0, 0);   // _row1 живёт внутри панели _rowHost
            BuildContent();
            root.Controls.Add(_splitMain, 0, 1);
            BuildStatusStrip();
            root.Controls.Add(_statusStrip, 0, 2);

            Content.Controls.Add(root);

            // таймер мониторинга папки
            _watchTimer.Tick += WatchTick;

            ResumeLayout(false);
            PerformLayout();
        }

        private Panel _rowHost;

        private void BuildTopPanel()
        {
            _rowHost = new Panel { Dock = DockStyle.Fill, Margin = new Padding(6, 6, 6, 2) };

            // ── строка 1: импорт, папка, мониторинг, интервал
            _row1 = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 32,
                ColumnCount = 6,
                RowCount = 1,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            _row1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            _row1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95F));
            _row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
            _row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 165F));

            _btnImport = new Button { Text = "Импорт файла…", Dock = DockStyle.Fill };
            _txtFolder = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(6, 4, 0, 3) };
            _btnBrowse = new Button { Text = "Обзор…", Dock = DockStyle.Fill, Margin = new Padding(6, 0, 0, 0) };
            _btnWatch = new Button { Text = "▶  Мониторинг: выкл", Dock = DockStyle.Fill, Margin = new Padding(6, 0, 0, 0) };
            _btnEditor = new Button { Text = "Редактор выписки…", Dock = DockStyle.Fill, Margin = new Padding(6, 0, 0, 0) };

            var intervalPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(6, 0, 0, 0),
                Padding = new Padding(0)
            };
            intervalPanel.Controls.Add(new Label
            {
                Text = "Интервал, сек:",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 8, 4, 0)
            });
            _nudInterval = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 3600,
                Value = 60,
                Width = 55,
                Margin = new Padding(0, 4, 0, 0)
            };
            intervalPanel.Controls.Add(_nudInterval);

            _row1.Controls.Add(_btnImport, 0, 0);
            _row1.Controls.Add(_txtFolder, 1, 0);
            _row1.Controls.Add(_btnBrowse, 2, 0);
            _row1.Controls.Add(_btnWatch, 3, 0);
            _row1.Controls.Add(intervalPanel, 4, 0);
            _row1.Controls.Add(_btnEditor, 5, 0);

            // ── строка 2: банк, статус, поиск, обновить
            _row2 = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 32,
                ColumnCount = 7,
                RowCount = 1,
                Margin = new Padding(0, 4, 0, 0),
                Padding = new Padding(0)
            };
            _row2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _row2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52F));
            _row2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170F));
            _row2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 66F));
            _row2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 165F));
            _row2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 62F));
            _row2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _row2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105F));

            _txtBank = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 3) };
            _cmbStatus = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(6, 4, 0, 3)
            };
            _txtSearch = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(6, 4, 0, 3) };
            _btnRefresh = new Button { Text = "Обновить", Dock = DockStyle.Fill, Margin = new Padding(6, 0, 0, 0) };

            _row2.Controls.Add(MakeLabel("Банк:"), 0, 0);
            _row2.Controls.Add(_txtBank, 1, 0);
            _row2.Controls.Add(MakeLabel("Статус:"), 2, 0);
            _row2.Controls.Add(_cmbStatus, 3, 0);
            _row2.Controls.Add(MakeLabel("Поиск:"), 4, 0);
            _row2.Controls.Add(_txtSearch, 5, 0);
            _row2.Controls.Add(_btnRefresh, 6, 0);

            // порядок докирования внутри _rowHost: без Dock-коллизий —
            // _row1 сверху (авт.высота), _row2 под ней
            _rowHost.Controls.Add(_row2);
            _rowHost.Controls.Add(_row1);

            // фильтр статусов
            _cmbStatus.Items.Add("Все статусы");
            foreach (PaymentStatus s in _allStatuses) _cmbStatus.Items.Add(StatusText.Russian(s));
            _cmbStatus.SelectedIndex = 0;

            // подсказки
            var tip = new ToolTip();
            tip.SetToolTip(_txtFolder, "Папка, за которой следит мониторинг (все новые xlsx импортируются автоматически)");
            tip.SetToolTip(_txtBank, "Название банка — группирует платежи из выписок разных банков");
            tip.SetToolTip(_txtSearch, "Поиск по номеру, контрагенту, назначению");
            tip.SetToolTip(_btnEditor, "Открыть выписку xlsx и изменить в ней данные (статус, сумму, дату и т.д.)");
        }

        private static Label MakeLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private void BuildContent()
        {
            _splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                Margin = new Padding(0)
            };
            _splitBottom = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Margin = new Padding(0)
            };

            // ── таблица платежей
            _gridPayments = MakeGrid();
            _gridPayments.Columns.Add(new DataGridViewTextBoxColumn
            { HeaderText = "Дата", Width = 82 });
            _gridPayments.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Сумма",
                Width = 110,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight }
            });
            _gridPayments.Columns.Add(new DataGridViewTextBoxColumn
            { HeaderText = "Статус", Width = 100 });
            _gridPayments.Columns.Add(new DataGridViewTextBoxColumn
            { HeaderText = "Номер", Width = 105 });
            _gridPayments.Columns.Add(new DataGridViewTextBoxColumn
            { HeaderText = "Контрагент", Width = 175 });
            _gridPayments.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Назначение",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 140
            });
            _gridPayments.Columns.Add(new DataGridViewTextBoxColumn
            { HeaderText = "Обновлён", Width = 110 });

            var menuPayments = new ContextMenuStrip();
            menuPayments.Items.Add(new ToolStripMenuItem("Открыть выписку в редакторе…",
                null, delegate { EditStatementOfCurrentPayment(); }));
            _gridPayments.ContextMenuStrip = menuPayments;

            // ── история выбранного платежа
            _gridHistory = MakeGrid();
            _gridHistory.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Время", Width = 125 });
            _gridHistory.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Было", Width = 95 });
            _gridHistory.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Стало", Width = 95 });
            _gridHistory.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Файл",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 120
            });

            var historyPanel = MakeTitledPanel("История изменений выбранного платежа", _gridHistory);

            // ── журнал
            _txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = false,
                BackColor = Color.White,
                Font = new Font("Consolas", 9F),
                BorderStyle = BorderStyle.FixedSingle
            };
            var logPanel = MakeTitledPanel("Журнал событий", _txtLog);

            _splitBottom.Panel1.Controls.Add(historyPanel);
            _splitBottom.Panel2.Controls.Add(logPanel);
            _splitMain.Panel1.Controls.Add(_gridPayments);
            _splitMain.Panel2.Controls.Add(_splitBottom);
        }

        /// <summary>Панель «заголовок сверху + содержимое», без коллизий докирования.</summary>
        private static TableLayoutPanel MakeTitledPanel(string title, Control content)
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var lbl = new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 60, 60),
                Margin = new Padding(2, 0, 0, 2)
            };
            content.Margin = new Padding(2, 0, 0, 2);

            t.Controls.Add(lbl, 0, 0);
            t.Controls.Add(content, 0, 1);
            return t;
        }

        private static DataGridView MakeGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 26,
                EnableHeadersVisualStyles = false,
                MultiSelect = false,
                RowHeadersVisible = false,
                RowTemplate = { Height = 22 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            };
        }

        private void BuildStatusStrip()
        {
            _statusStrip = new StatusStrip { SizingGrip = false };
            _lblMsg = new ToolStripStatusLabel
            {
                Text = "Готово",
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            _lblCounts = new ToolStripStatusLabel();
            _lblDb = new ToolStripStatusLabel();
            _statusStrip.Items.Add(_lblMsg);
            _statusStrip.Items.Add(_lblCounts);
            _statusStrip.Items.Add(_lblDb);
        }

        // ── обработчики ──────────────────────────────────────────────────

        private void AttachHandlers()
        {
            _btnImport.Click += delegate { ChooseAndImport(); };
            _btnBrowse.Click += delegate { BrowseFolder(); };
            _btnWatch.Click += delegate { ToggleWatch(); };
            _btnRefresh.Click += delegate { Reload(); };
            _btnEditor.Click += delegate { ChooseAndEdit(); };
            _nudInterval.ValueChanged += delegate
            {
                _settings.WatchIntervalSec = (int)_nudInterval.Value;
                if (_watchTimer.Enabled)
                    _watchTimer.Interval = Math.Max(1, (int)_nudInterval.Value) * 1000;
            };
            _txtFolder.TextChanged += delegate { _settings.WatchFolder = _txtFolder.Text.Trim(); };
            _txtFolder.Leave += delegate { SaveSettings(); };
            _txtBank.TextChanged += delegate
            {
                _settings.Bank = _txtBank.Text.Trim();
                _importer.Bank = string.IsNullOrEmpty(_settings.Bank) ? "Не указан" : _settings.Bank;
            };
            _txtBank.Leave += delegate { SaveSettings(); };
            _cmbStatus.SelectedIndexChanged += delegate { Reload(); };
            _txtSearch.TextChanged += delegate { Reload(); };
            _gridPayments.SelectionChanged += delegate { ShowHistoryForCurrent(); };
            FormClosing += delegate
            {
                _watchTimer.Stop();
                SaveSettings();
                _db.Dispose();
            };
        }

        private void ApplySettingsToControls()
        {
            _txtBank.Text = _settings.Bank ?? "";
            _txtFolder.Text = _settings.WatchFolder ?? "";
            int sec = _settings.WatchIntervalSec;
            if (sec < 1) sec = 10;
            _nudInterval.Value = Math.Min(Math.Max(sec, (int)_nudInterval.Minimum), (int)_nudInterval.Maximum);
            _lblDb.Text = "БД: " + _dbPath;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // размеры панелей задаём только здесь: при создании SplitContainer ещё не имеет размеров
            try
            {
                int mainH = _splitMain.ClientSize.Height;
                if (mainH > 400)
                {
                    _splitMain.Panel1MinSize = 120;
                    _splitMain.Panel2MinSize = 140;
                    _splitMain.SplitterDistance = Math.Max(140, (int)(mainH * 0.62));
                }

                int bottomW = _splitBottom.ClientSize.Width;
                if (bottomW > 520)
                {
                    _splitBottom.Panel1MinSize = 200;
                    _splitBottom.Panel2MinSize = 200;
                    _splitBottom.SplitterDistance = Math.Max(240, (int)(bottomW * 0.55));
                }
            }
            catch
            {
                // оставляем размеры по умолчанию
            }
        }

        private void SaveSettings()
        {
            SettingsStore.Save(_settings);
        }

        // ── импорт ───────────────────────────────────────────────────────

        private void ChooseAndImport()
        {
            using (var dlg = new OpenFileDialog
            {
                Title = "Выберите файлы выписки",
                Filter = "Excel (*.xlsx; *.xlsm)|*.xlsx;*.xlsm|Все файлы (*.*)|*.*",
                Multiselect = true
            })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                if (!string.IsNullOrEmpty(Path.GetDirectoryName(dlg.FileName)))
                {
                    _settings.WatchFolder = Path.GetDirectoryName(dlg.FileName);
                    _txtFolder.Text = _settings.WatchFolder;
                }
                ImportPaths(dlg.FileNames, false);
            }
        }

        private async void ImportPaths(string[] files, bool force)
        {
            if (_busy || files == null || files.Length == 0) return;
            _busy = true;
            SetBusy(true);
            try
            {
                foreach (string f in files)
                {
                    AppendLog("Импорт: " + Path.GetFileName(f));
                    ImportResult res = await Task.Run(() => _importer.ImportFile(f, force));
                    LogImport(res);
                }
                Reload();
            }
            finally
            {
                _busy = false;
                SetBusy(false);
            }
        }

        // ── редактор выписки ─────────────────────────────────────────────

        /// <summary>Выбрать файл и открыть его в редакторе.</summary>
        private void ChooseAndEdit()
        {
            using (var dlg = new OpenFileDialog
            {
                Title = "Выберите выписку для правки",
                Filter = "Excel (*.xlsx; *.xlsm)|*.xlsx;*.xlsm|Все файлы (*.*)|*.*"
            })
            {
                string folder = _txtFolder.Text.Trim();
                if (folder.Length > 0 && Directory.Exists(folder)) dlg.InitialDirectory = folder;

                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                OpenEditorFor(dlg.FileName);
            }
        }

        /// <summary>Открыть редактор для конкретного файла и импортировать правки.</summary>
        private void OpenEditorFor(string path)
        {
            try
            {
                using (var form = new StatementEditorForm(path))
                {
                    form.ShowDialog(this);
                    if (form.DialogResult == DialogResult.OK)
                    {
                        AppendLog("Выписка изменена в редакторе: " + Path.GetFileName(path));
                        ImportPaths(new[] { path }, true);   // импортируем сразу, чтобы видеть новые статусы
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Редактор выписки",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>Открыть в редакторе выписку, из которой пришёл выбранный платёж.</summary>
        private void EditStatementOfCurrentPayment()
        {
            if (_gridPayments.CurrentRow == null || _gridPayments.CurrentRow.Tag == null) return;
            var p = (Payment)_gridPayments.CurrentRow.Tag;

            string name = p.FirstSeenFile;
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show(this, "Для этого платежа не сохранено имя файла выписки.",
                    "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string path = null;
            if (File.Exists(name)) path = name;
            else
            {
                string folder = _txtFolder.Text.Trim();
                if (folder.Length > 0)
                {
                    string candidate = Path.Combine(folder, name);
                    if (File.Exists(candidate)) path = candidate;
                }
            }

            if (path == null)
            {
                MessageBox.Show(this,
                    "Файл выписки не найден: " + name +
                    "\r\n\r\nУкажите папку мониторинга с этим файлом или откройте его кнопкой «Редактор выписки…».",
                    "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            OpenEditorFor(path);
        }

        private async void WatchTick(object sender, EventArgs e)
        {
            if (_busy) return;
            string folder = _txtFolder.Text.Trim();
            if (folder.Length == 0 || !Directory.Exists(folder)) return;

            _busy = true;
            try
            {
                List<ImportResult> results = await Task.Run(() => _importer.ScanFolder(folder));
                bool any = false;
                foreach (ImportResult r in results)
                {
                    if (r.Error != null) { LogImport(r); continue; }
                    if (r.Skipped) continue;              // уже импортирован — молчим
                    LogImport(r);
                    any = true;
                }
                if (any) Reload();
            }
            finally
            {
                _busy = false;
            }
        }

        private void ToggleWatch()
        {
            if (_watchTimer.Enabled)
            {
                _watchTimer.Stop();
                _btnWatch.Text = "▶  Мониторинг: выкл";
                AppendLog("Мониторинг остановлен.");
                SaveSettings();
                return;
            }

            if (_txtFolder.Text.Trim().Length == 0) BrowseFolder();
            string folder = _txtFolder.Text.Trim();
            if (folder.Length == 0) return;
            if (!Directory.Exists(folder))
            {
                MessageBox.Show(this, "Папка не найдена:\r\n" + folder, "PaymentMonitor",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _watchTimer.Interval = Math.Max(1, (int)_nudInterval.Value) * 1000;
            _watchTimer.Start();
            _btnWatch.Text = "⏹  Мониторинг: вкл";
            AppendLog(string.Format(CultureInfo.InvariantCulture,
                "Мониторинг: {0} (проверка каждые {1} сек)", folder, (int)_nudInterval.Value));
            SaveSettings();
        }

        private void BrowseFolder()
        {
            using (var dlg = new FolderBrowserDialog
            {
                Description = "Папка с файлами выписок (xlsx)",
                ShowNewFolderButton = false
            })
            {
                if (!string.IsNullOrEmpty(_txtFolder.Text) && Directory.Exists(_txtFolder.Text))
                    dlg.SelectedPath = _txtFolder.Text;

                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    _txtFolder.Text = dlg.SelectedPath;
                    SaveSettings();
                }
            }
        }

        private void SetBusy(bool busy)
        {
            _btnImport.Enabled = !busy;
            _btnBrowse.Enabled = !busy;
            _btnRefresh.Enabled = !busy;
            _btnWatch.Enabled = !busy;
            _btnEditor.Enabled = !busy;
            UseWaitCursor = busy;
        }

        private void LogImport(ImportResult res)
        {
            if (res.Error != null)
            {
                AppendLog("  [ошибка] " + res.FileName + ": " + res.Error);
                return;
            }
            if (res.Skipped)
            {
                AppendLog("  [пропущен] " + res.FileName + " — уже импортирован");
                return;
            }
            ImportRun s = res.Run;
            AppendLog(string.Format(CultureInfo.InvariantCulture,
                "  [ok] {0}: создано {1}, обновлено {2}, без изменений {3} (строк {4})",
                s.FileName, s.Created, s.Updated, s.Unchanged, s.Rows));
        }

        // ── отображение ──────────────────────────────────────────────────

        private void Reload()
        {
            PaymentStatus? filter = null;
            if (_cmbStatus.SelectedIndex > 0) filter = _allStatuses[_cmbStatus.SelectedIndex - 1];

            List<Payment> items = _db.QueryPayments(filter, _txtSearch.Text, 500);
            FillPaymentsGrid(items);
            UpdateCounts();
            _lblMsg.Text = items.Count == 500
                ? "Показаны первые 500 записей (укажите фильтр или поиск)"
                : "Записей: " + items.Count;
        }

        private void FillPaymentsGrid(List<Payment> items)
        {
            _gridPayments.Rows.Clear();
            foreach (Payment p in items)
            {
                int i = _gridPayments.Rows.Add(
                    p.OpDate.HasValue ? p.OpDate.Value.ToString("dd.MM.yyyy") : "",
                    FormatMoney(p.AmountMinor, p.Currency),
                    StatusText.Russian(p.Status),
                    p.ExternalId ?? "",
                    p.Counterparty ?? "",
                    p.Purpose ?? "",
                    p.UpdatedAt == DateTime.MinValue ? "" : p.UpdatedAt.ToString("dd.MM.yy HH:mm"));

                DataGridViewRow row = _gridPayments.Rows[i];
                row.Tag = p;
                PaintStatusCell(row.Cells[2], p.Status);
            }

            if (_gridPayments.Rows.Count > 0)
                _gridPayments.CurrentCell = _gridPayments.Rows[0].Cells[0];
            else
                ShowHistory(null);
        }

        private void ShowHistoryForCurrent()
        {
            if (_gridPayments.CurrentRow == null || _gridPayments.CurrentRow.Tag == null)
            {
                ShowHistory(null);
                return;
            }
            var p = (Payment)_gridPayments.CurrentRow.Tag;
            ShowHistory(_db.QueryEvents(p.Id, 200));
        }

        private void ShowHistory(List<PaymentEvent> events)
        {
            _gridHistory.Rows.Clear();
            if (events == null) return;
            foreach (PaymentEvent ev in events)
            {
                int i = _gridHistory.Rows.Add(
                    ev.CreatedAt == DateTime.MinValue ? "" : ev.CreatedAt.ToString("dd.MM.yy HH:mm:ss"),
                    RuStatus(ev.FromStatus),
                    RuStatus(ev.ToStatus),
                    ev.File ?? "");
                _gridHistory.Rows[i].Tag = ev.Payload;
            }
        }

        private static string RuStatus(string key)
        {
            if (string.IsNullOrEmpty(key)) return "—";
            return StatusText.Russian(StatusText.ParseKey(key));
        }

        private void UpdateCounts()
        {
            Dictionary<PaymentStatus, int> counts = _db.CountByStatus();
            int total = 0;
            foreach (int n in counts.Values) total += n;

            var sb = new StringBuilder();
            sb.Append("Всего: ").Append(total);
            foreach (PaymentStatus s in _allStatuses)
            {
                int n;
                counts.TryGetValue(s, out n);
                if (n > 0)
                    sb.Append("   ·   ").Append(StatusText.Russian(s)).Append(": ").Append(n);
            }
            _lblCounts.Text = sb.ToString();
        }

        private static void PaintStatusCell(DataGridViewCell cell, PaymentStatus st)
        {
            Color back, fore;
            switch (st)
            {
                case PaymentStatus.Succeeded: back = Color.FromArgb(230, 244, 234); fore = Color.FromArgb(30, 126, 52); break;
                case PaymentStatus.Failed: back = Color.FromArgb(253, 236, 234); fore = Color.FromArgb(176, 42, 55); break;
                case PaymentStatus.Pending: back = Color.FromArgb(255, 244, 229); fore = Color.FromArgb(178, 106, 0); break;
                case PaymentStatus.Canceled: back = Color.FromArgb(238, 238, 238); fore = Color.FromArgb(85, 85, 85); break;
                case PaymentStatus.Refunded: back = Color.FromArgb(237, 231, 246); fore = Color.FromArgb(94, 53, 177); break;
                case PaymentStatus.New: back = Color.FromArgb(227, 242, 253); fore = Color.FromArgb(13, 71, 161); break;
                default: back = Color.FromArgb(255, 249, 196); fore = Color.FromArgb(130, 119, 23); break;
            }
            cell.Style.BackColor = back;
            cell.Style.ForeColor = fore;
            cell.Style.SelectionBackColor = back;
            cell.Style.SelectionForeColor = fore;
            cell.Style.Font = StatusFont;
        }

        private static string FormatMoney(long? minor, string currency)
        {
            if (!minor.HasValue) return "";
            string s = (minor.Value / 100m).ToString("N2", CultureInfo.GetCultureInfo("ru-RU"));
            return string.IsNullOrEmpty(currency) ? s : s + " " + currency;
        }

        private void AppendLog(string message)
        {
            if (_txtLog == null || _txtLog.IsDisposed) return;
            if (_txtLog.TextLength > 60000) _txtLog.Clear();
            _txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }
    }
}

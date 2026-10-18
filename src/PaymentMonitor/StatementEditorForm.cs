using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using PaymentMonitor.Core;

namespace PaymentMonitor
{
    /// <summary>
    /// Окно «Редактор выписки»: показывает таблицу выбранного xlsx и позволяет
    /// менять данные (статус, сумму, дату, назначение и т.д.), сохраняя их в тот же файл.
    /// </summary>
    public sealed class StatementEditorForm : Form
    {
        private readonly StatementDocument _doc;
        private bool _loading;

        private ComboBox _cmbSheet;
        private DataGridView _grid;
        private Button _btnAdd;
        private Button _btnDelete;
        private Button _btnSave;
        private Button _btnClose;
        private Label _lblTop;
        private Label _lblWarn;
        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _lblInfo;

        /// <summary>Путь сохранённого файла (для импорта в главном окне).</summary>
        public string SavedPath { get { return _doc.FilePath; } }

        public StatementEditorForm(string path)
        {
            _doc = new StatementDocument(path);   // исключение пробрасывается вызывающему

            InitializeComponent();
            LoadSheets();
            LoadCurrentSheet();
        }

        private void InitializeComponent()
        {
            Text = "Редактор выписки — " + Path.GetFileName(_doc.FilePath);
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(1040, 640);
            MinimumSize = new Size(760, 460);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));

            // ── строка 1: лист и файл
            var row1 = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Margin = new Padding(6, 6, 6, 0),
                Padding = new Padding(0)
            };
            row1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 42F));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190F));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50F));
            row1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            _cmbSheet = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 3, 0, 3)
            };
            _lblTop = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                ForeColor = Color.FromArgb(70, 70, 70)
            };

            row1.Controls.Add(MakeLabel("Лист:"), 0, 0);
            row1.Controls.Add(_cmbSheet, 1, 0);
            row1.Controls.Add(MakeLabel("Файл:"), 2, 0);
            row1.Controls.Add(_lblTop, 3, 0);

            // ── строка 2: кнопки
            var row2 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(6, 2, 6, 2),
                Padding = new Padding(0)
            };
            _btnAdd = new Button { Text = "➕ Добавить строку", AutoSize = true, Margin = new Padding(0, 2, 6, 0) };
            _btnDelete = new Button { Text = "➖ Удалить строку", AutoSize = true, Margin = new Padding(0, 2, 6, 0) };
            _btnSave = new Button { Text = "💾 Сохранить в файл", AutoSize = true, Margin = new Padding(16, 2, 6, 0) };
            _btnClose = new Button { Text = "Закрыть", AutoSize = true, Margin = new Padding(0, 2, 6, 0) };
            _lblWarn = new Label
            {
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(178, 106, 0),
                Margin = new Padding(12, 8, 0, 0)
            };
            row2.Controls.Add(_btnAdd);
            row2.Controls.Add(_btnDelete);
            row2.Controls.Add(_btnSave);
            row2.Controls.Add(_btnClose);
            row2.Controls.Add(_lblWarn);

            // ── таблица
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 30,
                EnableHeadersVisualStyles = false,
                MultiSelect = false,
                RowHeadersVisible = true,
                RowHeadersWidth = 52,
                RowTemplate = { Height = 22 },
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                Margin = new Padding(6, 0, 6, 0)
            };
            _grid.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            var tip = new ToolTip();
            tip.SetToolTip(_grid, "Двойной клик или F2 — правка ячейки. Серые колонки изменить нельзя.");

            // ── статус-строка
            _statusStrip = new StatusStrip { SizingGrip = false };
            _lblInfo = new ToolStripStatusLabel { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            _statusStrip.Items.Add(_lblInfo);

            root.Controls.Add(row1, 0, 0);
            root.Controls.Add(row2, 0, 1);
            root.Controls.Add(_grid, 0, 2);
            root.Controls.Add(_statusStrip, 0, 3);
            Controls.Add(root);

            AttachHandlers();
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

        private void AttachHandlers()
        {
            _cmbSheet.SelectedIndexChanged += delegate
            {
                if (_loading) return;
                if (_cmbSheet.SelectedItem == null) return;
                if (!ConfirmDiscard()) { RestoreSheetSelection(); return; }
                try
                {
                    _doc.SelectSheet(_cmbSheet.SelectedItem.ToString());
                    LoadCurrentSheet();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            _btnAdd.Click += delegate { AddRow(); };
            _btnDelete.Click += delegate { DeleteRow(); };
            _btnSave.Click += delegate { SaveAndClose(); };
            _btnClose.Click += delegate { Close(); };

            _grid.CellEndEdit += delegate(object s, DataGridViewCellEventArgs e)
            {
                if (_loading) return;
                ApplyEdit(e.RowIndex, e.ColumnIndex);
            };
            _grid.EditingControlShowing += GridEditingControlShowing;
            _grid.DataError += delegate(object s, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            _grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && _grid.Columns[e.ColumnIndex].ReadOnly == false)
                    _grid.BeginEdit(true);
            };

            FormClosing += EditorFormClosing;
        }

        // ── загрузка ────────────────────────────────────────────────────

        private void LoadSheets()
        {
            _loading = true;
            _cmbSheet.Items.Clear();
            foreach (string name in _doc.SheetNames) _cmbSheet.Items.Add(name);
            _cmbSheet.SelectedItem = _doc.SheetName;
            _loading = false;
        }

        private void RestoreSheetSelection()
        {
            _loading = true;
            _cmbSheet.SelectedItem = _doc.SheetName;
            _loading = false;
        }

        private void LoadCurrentSheet()
        {
            _loading = true;
            try
            {
                _grid.Rows.Clear();
                _grid.Columns.Clear();

                foreach (StatementColumn col in _doc.Columns)
                {
                    var gc = new DataGridViewTextBoxColumn
                    {
                        HeaderText = col.Title,
                        Name = "c" + col.Index,
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                        ReadOnly = !col.Editable,
                        Width = Math.Max(70, Math.Min(240, 24 + col.Title.Length * 9)),
                        Tag = col
                    };
                    if (!col.Editable)
                        gc.DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
                    if (col.Field == "amount")
                        gc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    _grid.Columns.Add(gc);
                }

                foreach (StatementRow row in _doc.Rows)
                {
                    int i = _grid.Rows.Add(row.Values.ToArray());
                    DataGridViewRow gr = _grid.Rows[i];
                    gr.Tag = row.Index;
                    gr.HeaderCell.Value = row.Index.ToString(CultureInfo.InvariantCulture);
                    PaintStatus(gr);
                }
            }
            finally
            {
                _loading = false;
            }

            _lblTop.Text = _doc.FilePath;
            int statusCol = StatusColumnIndex();
            _lblWarn.Text = _doc.HeaderDetected
                ? "Можно менять: " + _doc.EditableFieldsSummary()
                : "Шапка таблицы не распознана — правка всех колонок разрешена. Статус при импорте может не определиться.";
            UpdateInfo();
        }

        private int StatusColumnIndex()
        {
            foreach (StatementColumn col in _doc.Columns)
                if (col.Field == "status") return col.Index;
            return -1;
        }

        private void UpdateInfo()
        {
            string s = "Строк: " + _grid.Rows.Count;
            if (_doc.HeaderDetected) s += "   ·   строка заголовков: " + _doc.HeaderRow;
            if (_doc.Dirty) s += "   ·   ЕСТЬ НЕСОХРАНЁННЫЕ ИЗМЕНЕНИЯ";
            _lblInfo.Text = s;
        }

        // ── правка ──────────────────────────────────────────────────────

        private void ApplyEdit(int rowPos, int colPos)
        {
            if (rowPos < 0 || colPos < 0) return;
            DataGridViewRow gr = _grid.Rows[rowPos];
            if (gr.Tag == null) return;

            int sheetRow = (int)gr.Tag;
            var col = (StatementColumn)_grid.Columns[colPos].Tag;
            string text = Convert.ToString(gr.Cells[colPos].Value, CultureInfo.InvariantCulture);

            try
            {
                _doc.SetValue(sheetRow, col.Index, text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _loading = true;
            try
            {
                gr.Cells[colPos].Value = _doc.Display(sheetRow, col.Index);
                PaintStatus(gr);
            }
            finally
            {
                _loading = false;
            }
            UpdateInfo();
        }

        private void PaintStatus(DataGridViewRow gr)
        {
            int col = StatusColumnIndex();
            if (col < 0 || col > gr.Cells.Count) return;
            int pos = col - 1;
            if (pos < 0 || pos >= gr.Cells.Count) return;

            string raw = Convert.ToString(gr.Cells[pos].Value, CultureInfo.InvariantCulture);
            PaymentStatus st = XlsxReader.MapStatus(raw);
            Color back, fore;
            switch (st)
            {
                case PaymentStatus.Succeeded: back = Color.FromArgb(230, 244, 234); fore = Color.FromArgb(30, 126, 52); break;
                case PaymentStatus.Failed: back = Color.FromArgb(253, 236, 234); fore = Color.FromArgb(176, 42, 55); break;
                case PaymentStatus.Pending: back = Color.FromArgb(255, 244, 229); fore = Color.FromArgb(178, 106, 0); break;
                case PaymentStatus.Canceled: back = Color.FromArgb(238, 238, 238); fore = Color.FromArgb(85, 85, 85); break;
                case PaymentStatus.Refunded: back = Color.FromArgb(237, 231, 246); fore = Color.FromArgb(94, 53, 177); break;
                default: gr.Cells[pos].Style.BackColor = Color.Empty; gr.Cells[pos].Style.ForeColor = Color.Empty; return;
            }
            gr.Cells[pos].Style.BackColor = back;
            gr.Cells[pos].Style.ForeColor = fore;
            gr.Cells[pos].Style.SelectionBackColor = back;
            gr.Cells[pos].Style.SelectionForeColor = fore;
        }

        private void GridEditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            var tb = e.Control as TextBox;
            if (tb == null) return;

            int col = _grid.CurrentCell == null ? -1 : _grid.CurrentCell.ColumnIndex;
            var sc = col >= 0 ? (StatementColumn)_grid.Columns[col].Tag : null;

            if (sc != null && sc.Field == "status")
            {
                var src = new AutoCompleteStringCollection();
                src.AddRange(StatusSuggestions().ToArray());
                tb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                tb.AutoCompleteSource = AutoCompleteSource.CustomSource;
                tb.AutoCompleteCustomSource = src;
            }
            else
            {
                tb.AutoCompleteMode = AutoCompleteMode.None;
                tb.AutoCompleteCustomSource = null;
            }
        }

        /// <summary>Подсказки для статуса: значения из файла + типовые формулировки банков.</summary>
        private List<string> StatusSuggestions()
        {
            var list = new List<string>();
            int col = StatusColumnIndex();
            if (col > 0)
            {
                foreach (DataGridViewRow r in _grid.Rows)
                {
                    string v = Convert.ToString(r.Cells[col - 1].Value, CultureInfo.InvariantCulture);
                    if (!string.IsNullOrWhiteSpace(v) && !list.Contains(v)) list.Add(v);
                }
            }
            string[] typical =
            {
                "Исполнено", "Проведено", "Оплачено", "Зачислено", "Завершено", "Успешно",
                "В обработке", "Ожидает", "На проверке",
                "Отклонено", "Отказ", "Ошибка", "Не проведено",
                "Отменено", "Отозвано",
                "Возврат", "Сторно",
                "Новое", "Создано"
            };
            foreach (string t in typical)
                if (!list.Contains(t)) list.Add(t);
            return list;
        }

        private void AddRow()
        {
            try
            {
                int newRow = _doc.AddRow();
                LoadCurrentSheet();

                for (int i = 0; i < _grid.Rows.Count; i++)
                {
                    if ((int)_grid.Rows[i].Tag == newRow)
                    {
                        _grid.CurrentCell = _grid.Rows[i].Cells[0];
                        _grid.BeginEdit(true);
                        break;
                    }
                }
                UpdateInfo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DeleteRow()
        {
            if (_grid.CurrentRow == null)
            {
                MessageBox.Show(this, "Выберите строку для удаления.", "Редактор выписки",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int sheetRow = (int)_grid.CurrentRow.Tag;
            DialogResult ok = MessageBox.Show(this,
                "Удалить строку " + sheetRow + " из файла?", "Редактор выписки",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (ok != DialogResult.Yes) return;

            try
            {
                _doc.DeleteRow(sheetRow);
                LoadCurrentSheet();
                UpdateInfo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ── сохранение ──────────────────────────────────────────────────

        private void SaveAndClose()
        {
            if (!_doc.Dirty)
            {
                DialogResult = DialogResult.OK;   // изменений нет — просто закрываем
                Close();
                return;
            }
            try
            {
                _doc.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            MessageBox.Show(this,
                "Файл сохранён:\r\n" + _doc.FilePath + "\r\n\r\nСейчас изменения будут импортированы в базу.",
                "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Information);

            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>Есть ли несохранённые правки: спросить пользователя.</summary>
        private bool ConfirmDiscard()
        {
            if (!_doc.Dirty) return true;
            DialogResult r = MessageBox.Show(this,
                "Есть несохранённые изменения. Сохранить их?",
                "Редактор выписки", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return false;
            if (r == DialogResult.No) return true;   // выбрасываем правки (файл не меняем)
            try
            {
                _doc.Save();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Редактор выписки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private void EditorFormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing && DialogResult != DialogResult.OK)
            {
                if (!ConfirmDiscard()) { e.Cancel = true; return; }
            }
            _doc.Dispose();
        }
    }
}

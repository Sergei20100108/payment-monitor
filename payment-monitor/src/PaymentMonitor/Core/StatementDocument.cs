using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OfficeOpenXml;

namespace PaymentMonitor.Core
{
    /// <summary>Колонка выписки в редакторе.</summary>
    public sealed class StatementColumn
    {
        public int Index;          // номер столбца на листе (1-based)
        public string Title = "";  // заголовок из файла (или буква столбца)
        public string Field;       // date/amount/status/... либо null, если колонка не распознана
        public bool Editable;      // можно ли править
    }

    /// <summary>Строка выписки в редакторе.</summary>
    public sealed class StatementRow
    {
        public int Index;          // номер строки на листе (1-based)
        public List<string> Values = new List<string>();
    }

    /// <summary>
    /// Выписка, открытая для правки. Показывает таблицу листа, позволяет менять значения
    /// и сохраняет их обратно в тот же xlsx-файл. Правки держатся в памяти до вызова Save().
    /// </summary>
    public sealed class StatementDocument : IDisposable
    {
        private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

        private readonly ExcelPackage _package;
        private ExcelWorksheet _ws;

        public string FilePath { get; private set; }
        public List<string> SheetNames { get; private set; }
        public string SheetName { get { return _ws == null ? null : _ws.Name; } }
        public int HeaderRow { get; private set; }       // 0, если шапка не распознана
        public bool HeaderDetected { get; private set; }
        public List<StatementColumn> Columns { get; private set; }
        public List<StatementRow> Rows { get; private set; }
        public bool Dirty { get; private set; }

        public StatementDocument(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new ParseException("файл не найден: " + path);

            FilePath = Path.GetFullPath(path);
            try
            {
                _package = new ExcelPackage(new FileInfo(FilePath));
            }
            catch (Exception ex)
            {
                throw new ParseException(
                    "не удалось открыть Excel-файл (возможно, он занят другой программой или это старый .xls): " + ex.Message);
            }

            SheetNames = new List<string>();
            foreach (ExcelWorksheet w in _package.Workbook.Worksheets) SheetNames.Add(w.Name);
            if (SheetNames.Count == 0) throw new ParseException("в файле нет листов");

            // по умолчанию — первый лист, на котором нашлась шапка таблицы
            string pick = SheetNames[0];
            foreach (string name in SheetNames)
            {
                ExcelWorksheet w = _package.Workbook.Worksheets[name];
                int hr;
                Dictionary<string, int> m;
                if (w.Dimension != null && XlsxReader.TryDetectHeader(w, out hr, out m))
                {
                    pick = name;
                    break;
                }
            }
            SelectSheet(pick);
        }

        /// <summary>Переключиться на другой лист файла.</summary>
        public void SelectSheet(string name)
        {
            ExcelWorksheet w = _package.Workbook.Worksheets[name];
            if (w == null) throw new ParseException("лист не найден: " + name);
            _ws = w;
            Dirty = false;
            Build();
        }

        // ── построение модели листа ─────────────────────────────────────

        private void Build()
        {
            Columns = new List<StatementColumn>();
            Rows = new List<StatementRow>();

            int lastRow = 0, lastCol = 0;
            if (_ws.Dimension != null)
            {
                lastRow = _ws.Dimension.End.Row;
                lastCol = _ws.Dimension.End.Column;
            }
            if (lastCol > 100) lastCol = 100;   // защита от «широких» листов

            int headerRow = 0;
            Dictionary<string, int> map = null;
            HeaderDetected = lastRow > 0 && XlsxReader.TryDetectHeader(_ws, out headerRow, out map);
            if (!HeaderDetected) { headerRow = 0; map = new Dictionary<string, int>(); }
            HeaderRow = HeaderDetected ? headerRow : 0;

            var fieldOfCol = new Dictionary<int, string>();
            if (HeaderDetected)
                foreach (KeyValuePair<string, int> kv in map) fieldOfCol[kv.Value] = kv.Key;

            for (int c = 1; c <= lastCol; c++)
            {
                string title = HeaderDetected
                    ? Convert.ToString(_ws.Cells[HeaderRow, c].Value, CultureInfo.InvariantCulture)
                    : "";
                string field;
                fieldOfCol.TryGetValue(c, out field);

                Columns.Add(new StatementColumn
                {
                    Index = c,
                    Title = string.IsNullOrWhiteSpace(title) ? ColumnLetter(c) : title.Trim(),
                    Field = field,
                    // если шапка не распознана — разрешаем править всё, иначе только понятные колонки
                    Editable = !HeaderDetected || field != null
                });
            }

            int firstDataRow = HeaderDetected ? HeaderRow + 1 : 1;
            for (int r = firstDataRow; r <= lastRow; r++)
            {
                var row = new StatementRow { Index = r };
                for (int i = 0; i < Columns.Count; i++)
                    row.Values.Add(CellText(_ws.Cells[r, Columns[i].Index].Value, Columns[i].Field));
                Rows.Add(row);
            }

            while (Rows.Count > 0 && RowIsBlank(Rows[Rows.Count - 1]))
                Rows.RemoveAt(Rows.Count - 1);
        }

        // ── правка ──────────────────────────────────────────────────────

        /// <summary>Записать новое значение в ячейку (в памяти; файл меняется только при Save).</summary>
        public void SetValue(int rowIndex, int colIndex, string text)
        {
            if (rowIndex < 1 || colIndex < 1) throw new ParseException("некорректная ячейка");
            if (HeaderDetected && rowIndex == HeaderRow)
                throw new ParseException("строка заголовков не редактируется");

            ExcelRange cell = _ws.Cells[rowIndex, colIndex];
            object existing = cell.Value;
            string field = FieldOfColumn(colIndex);
            text = text == null ? null : text.Trim();

            if (field == "date")
            {
                DateTime? d = XlsxReader.ParseDate(text);
                cell.Value = d.HasValue ? (object)d.Value : (text == "" ? null : (object)text);
            }
            else if (field == "amount")
            {
                long? minor = XlsxReader.ParseAmount(text);
                cell.Value = minor.HasValue ? (object)(minor.Value / 100m) : (text == "" ? null : (object)text);
            }
            else if (existing is DateTime)
            {
                DateTime? d = XlsxReader.ParseDate(text);
                cell.Value = d.HasValue ? (object)d.Value : (object)text;
            }
            else if (IsNumber(existing))
            {
                // числовая ячейка: пытаемся сохранить числовой тип, иначе пишем текстом
                double num;
                if (TryParseNumber(text, out num)) cell.Value = num;
                else cell.Value = text == "" ? null : (object)text;
            }
            else
            {
                cell.Value = text == "" ? null : (object)text;
            }

            // обновляем «снимок» строки, иначе повторный вывод покажет старое значение
            int at = IndexOfRow(rowIndex);
            if (at >= 0)
            {
                StatementRow row = Rows[at];
                int ci = colIndex - 1;
                if (ci >= 0 && ci < row.Values.Count)
                    row.Values[ci] = Display(rowIndex, colIndex);
            }

            Dirty = true;
        }

        /// <summary>Добавить пустую строку сразу после последнего платежа. Возвращает номер новой строки.</summary>
        public int AddRow()
        {
            int lastData = -1;
            for (int i = Rows.Count - 1; i >= 0; i--)
            {
                if (RowLooksLikeData(Rows[i])) { lastData = Rows[i].Index; break; }
            }

            int baseRow = lastData >= 0 ? lastData : (HeaderDetected ? HeaderRow : 0);
            int newRow = baseRow + 1;

            _ws.InsertRow(newRow, 1);
            if (baseRow >= 1)
            {
                for (int c = 1; c <= Columns.Count; c++)
                {
                    try { _ws.Cells[newRow, c].StyleID = _ws.Cells[baseRow, c].StyleID; }
                    catch { /* оформление не критично */ }
                }
            }

            foreach (StatementRow r in Rows)
                if (r.Index >= newRow) r.Index++;

            var row = new StatementRow { Index = newRow };
            for (int i = 0; i < Columns.Count; i++) row.Values.Add("");

            int at = IndexOfRow(lastData);
            if (at < 0) Rows.Add(row);
            else Rows.Insert(at + 1, row);

            Dirty = true;
            return newRow;
        }

        /// <summary>Удалить строку листа целиком.</summary>
        public void DeleteRow(int rowIndex)
        {
            if (HeaderDetected && rowIndex <= HeaderRow)
                throw new ParseException("нельзя удалить строку заголовков (строка " + HeaderRow + ")");

            int at = IndexOfRow(rowIndex);
            if (at < 0) throw new ParseException("строка " + rowIndex + " не найдена");

            _ws.DeleteRow(rowIndex, 1);
            Rows.RemoveAt(at);
            foreach (StatementRow r in Rows)
                if (r.Index > rowIndex) r.Index--;

            Dirty = true;
        }

        /// <summary>Записать изменения в файл (перезаписывает исходный xlsx).</summary>
        public void Save()
        {
            try
            {
                _package.Save();
            }
            catch (Exception ex)
            {
                throw new ParseException(
                    "не удалось сохранить файл: " + ex.Message +
                    ". Если выписка открыта в Excel — закройте её и повторите сохранение.");
            }
            Dirty = false;
        }

        public void Dispose()
        {
            try { _package.Dispose(); } catch { }
        }

        // ── вспомогательное ─────────────────────────────────────────────

        /// <summary>Колонка по имени поля (date, amount, status, purpose, currency, counterparty, external_id) или по номеру.</summary>
        public int ColumnByFieldOrIndex(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new ParseException("не указана колонка (например, status или 7)");

            int n;
            if (int.TryParse(key.Trim(), out n))
            {
                if (n >= 1 && n <= Columns.Count) return n;
                throw new ParseException("нет такого столбца: " + key);
            }

            string k = key.Trim().ToLowerInvariant();
            foreach (StatementColumn c in Columns)
                if (c.Field != null && c.Field == k) return c.Index;

            switch (k)
            {
                case "дата": return ColumnByFieldOrIndex("date");
                case "сумма": return ColumnByFieldOrIndex("amount");
                case "статус": case "состояние": return ColumnByFieldOrIndex("status");
                case "назначение": case "описание": return ColumnByFieldOrIndex("purpose");
                case "валюта": return ColumnByFieldOrIndex("currency");
                case "контрагент": case "плательщик": case "получатель": return ColumnByFieldOrIndex("counterparty");
                case "номер": case "№": case "id": return ColumnByFieldOrIndex("external_id");
            }
            throw new ParseException("колонка не распознана: " + key + " (доступно: " + EditableFieldsSummary() + ")");
        }

        /// <summary>Отображаемое значение ячейки (после правки).</summary>
        public string Display(int rowIndex, int colIndex)
        {
            return CellText(_ws.Cells[rowIndex, colIndex].Value, FieldOfColumn(colIndex));
        }

        /// <summary>Перечень редактируемых полей по-русски (для подсказки).</summary>
        public string EditableFieldsSummary()
        {
            var parts = new List<string>();
            foreach (StatementColumn c in Columns)
            {
                if (!c.Editable || c.Field == null) continue;
                string ru = FieldRussian(c.Field);
                if (!parts.Contains(ru)) parts.Add(ru);
            }
            return parts.Count == 0 ? "нет распознанных колонок" : string.Join(", ", parts.ToArray());
        }

        private string FieldOfColumn(int colIndex)
        {
            foreach (StatementColumn c in Columns)
                if (c.Index == colIndex) return c.Field;
            return null;
        }

        private int IndexOfRow(int rowIndex)
        {
            if (rowIndex < 0) return -1;
            for (int i = 0; i < Rows.Count; i++)
                if (Rows[i].Index == rowIndex) return i;
            return -1;
        }

        private bool RowLooksLikeData(StatementRow row)
        {
            for (int i = 0; i < Columns.Count; i++)
            {
                string f = Columns[i].Field;
                if ((f == "status" || f == "external_id") && i < row.Values.Count &&
                    !string.IsNullOrWhiteSpace(row.Values[i]))
                    return true;
            }
            return false;
        }

        private bool RowIsBlank(StatementRow row)
        {
            foreach (string v in row.Values)
                if (!string.IsNullOrWhiteSpace(v)) return false;
            return true;
        }

        private static bool IsNumber(object v)
        {
            return v is double || v is int || v is long || v is decimal || v is short || v is float;
        }

        private static bool TryParseNumber(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string s = text.Replace(" ", "").Replace("\u00a0", "").Replace("\u202f", "");
            return double.TryParse(s, NumberStyles.Any, Ru, out value)
                || double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private static string CellText(object v, string field)
        {
            if (v == null) return "";
            if (v is DateTime) return ((DateTime)v).ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
            if (IsNumber(v))
            {
                double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                if (field == "amount") return d.ToString("N2", Ru);
                return d == Math.Floor(d) ? ((long)d).ToString(Ru) : d.ToString(Ru);
            }
            if (v is bool) return (bool)v ? "да" : "нет";
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        private static string ColumnLetter(int index)
        {
            string s = "";
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                s = (char)('A' + rem) + s;
                index = (index - 1) / 26;
            }
            return s;
        }

        private static string FieldRussian(string field)
        {
            switch (field)
            {
                case "date": return "Дата";
                case "amount": return "Сумма";
                case "status": return "Статус";
                case "purpose": return "Назначение";
                case "currency": return "Валюта";
                case "counterparty": return "Контрагент";
                case "external_id": return "Номер";
                default: return field;
            }
        }
    }
}

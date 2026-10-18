using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace PaymentMonitor.Core
{
    /// <summary>Не удалось разобрать файл выписки.</summary>
    public class ParseException : Exception
    {
        public ParseException(string message) : base(message) { }
    }

    /// <summary>
    /// Чтение выписки Excel (xlsx) через EPPlus.
    /// Шапка и колонки определяются автоматически по ключевым словам,
    /// статусы банка нормализуются в PaymentStatus.
    /// </summary>
    public static class XlsxReader
    {
        // поле -> ключевые слова заголовка (поиск по подстроке, регистр не важен)
        private static readonly Dictionary<string, string[]> FieldKeywords = new Dictionary<string, string[]>
        {
            { "date",         new[] { "дата" } },
            { "amount",       new[] { "сумма" } },
            { "status",       new[] { "статус", "состояние", "state" } },
            { "purpose",      new[] { "назнач", "описание", "основани", "комментар" } },
            { "currency",     new[] { "валют" } },
            { "counterparty", new[] { "контрагент", "контраг", "получател", "плательщик", "адресат", "наименование" } },
            { "external_id",  new[] { "номер платежа", "номер документа", "номер операции", "номер", "№", "id", "идентификатор" } }
        };

        // порядок проверки важен: «Дата документа» должна попасть в date, а не в external_id
        private static readonly string[] FieldOrder =
            { "date", "amount", "status", "purpose", "currency", "counterparty", "external_id" };

        // правила перевода статусов банка в наши значения (проверяются по порядку)
        private static readonly Tuple<PaymentStatus, string[]>[] StatusRules =
        {
            Tuple.Create(PaymentStatus.Refunded, new[] { "возврат", "возвращ", "сторно", "refund", "revers" }),
            Tuple.Create(PaymentStatus.Canceled, new[] { "отмен", "отозван", "cancel" }),
            Tuple.Create(PaymentStatus.Failed, new[] { "отклон", "отказ", "ошибка", "неуспешн", "не пров", "непровед", "не исполн", "неисполн", "reject", "fail", "error", "declin", "not booked" }),
            Tuple.Create(PaymentStatus.Pending, new[] { "в обработк", "ожидает", "на проверк", "на рассмотр", "pending", "processing", "hold" }),
            Tuple.Create(PaymentStatus.Succeeded, new[] { "исполн", "проведен", "проведён", "оплачен", "выполнен", "завершен", "завершён", "успешн", "зачислен", "executed", "success", "complete", "booked" }),
            Tuple.Create(PaymentStatus.New, new[] { "нов", "создан", "registered", "new" })
        };

        private static readonly string[] DateFormats =
        {
            "d.M.yyyy", "d.M.yyyy H:mm", "dd.MM.yyyy", "yyyy-MM-dd", "d/M/yyyy", "yyyy/MM/dd"
        };

        public static ParseResult Parse(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new ParseException("файл не найден: " + path);

            ExcelPackage package;
            try
            {
                package = new ExcelPackage(new FileInfo(path));
            }
            catch (Exception ex)
            {
                throw new ParseException("не удалось открыть Excel-файл (возможно, это старый .xls или файл ещё дописывается): " + ex.Message);
            }

            using (package)
            {
                foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                {
                    if (ws.Dimension == null) continue;

                    int headerRow;
                    Dictionary<string, int> map;
                    if (!TryDetectHeader(ws, out headerRow, out map)) continue;

                    var result = new ParseResult();
                    Extract(ws, headerRow, map, result);
                    return result;
                }
            }

            throw new ParseException(
                "не найдена шапка таблицы: нужны колонки «Сумма» и как минимум одна " +
                "(«Дата», «Номер…», «Статус», «Контрагент» и т.п.). " +
                "Если файл в старом формате .xls — сохраните его как .xlsx");
        }

        /// <summary>Маппинг статуса банка в нормализованный.</summary>
        public static PaymentStatus MapStatus(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return PaymentStatus.New;
            string s = raw.ToLowerInvariant();
            foreach (var rule in StatusRules)
            {
                foreach (var kw in rule.Item2)
                    if (s.Contains(kw)) return rule.Item1;
            }
            return PaymentStatus.Unknown;
        }

        // ── внутренняя кухня ────────────────────────────────────────────

        /// <summary>
        /// Поиск строки заголовков и колонок таблицы в листе (первые 15 строк).
        /// Используется и при импорте, и в редакторе выписок.
        /// </summary>
        public static bool TryDetectHeader(ExcelWorksheet ws, out int headerRow, out Dictionary<string, int> map)
        {
            headerRow = -1;
            map = null;

            int firstRow = ws.Dimension.Start.Row;
            int lastRow = Math.Min(ws.Dimension.End.Row, firstRow + 14);
            int lastCol = Math.Min(ws.Dimension.End.Column, 60);

            for (int r = firstRow; r <= lastRow; r++)
            {
                // Поля распределяются по столбцам в порядке FieldOrder: каждое поле
                // берёт самый левый подходящий столбец, который ещё не занят другим
                // полем. Так «Сумма в валюте счёта» уходит в сумму, а валюта —
                // в отдельную колонку «Код валюты», а не в тот же столбец.
                var candidate = new Dictionary<string, int>();
                var used = new HashSet<int>();

                foreach (string field in FieldOrder)
                {
                    string[] keywords;
                    if (!FieldKeywords.TryGetValue(field, out keywords)) continue;

                    for (int c = 1; c <= lastCol; c++)
                    {
                        if (used.Contains(c)) continue;

                        string header = NormHeader(ws.Cells[r, c].Value);
                        if (header.Length == 0) continue;

                        bool hit = false;
                        foreach (string kw in keywords)
                        {
                            if (header.Contains(kw)) { hit = true; break; }
                        }
                        if (hit)
                        {
                            candidate[field] = c;
                            used.Add(c);
                            break;
                        }
                    }
                }

                if (candidate.ContainsKey("amount") && candidate.Count >= 2)
                {
                    headerRow = r;
                    map = candidate;
                    return true;
                }
            }
            return false;
        }

        private static void Extract(ExcelWorksheet ws, int headerRow, Dictionary<string, int> map, ParseResult result)
        {
            int lastRow = ws.Dimension.End.Row;

            for (int r = headerRow + 1; r <= lastRow; r++)
            {
                object cellStatus = CellVal(ws, r, map, "status");
                object cellExt = CellVal(ws, r, map, "external_id");

                string rawStatus = AsText(cellStatus);
                string externalId = AsText(cellExt);

                // полностью пустая строка — пропускаем без счёта
                if (RowIsEmpty(ws, r, map)) { result.SkippedRows++; continue; }

                // строка без статуса и без номера (итоги, подписи) — не платёж
                if (string.IsNullOrEmpty(rawStatus) && string.IsNullOrEmpty(externalId))
                {
                    result.SkippedRows++;
                    continue;
                }

                DateTime? opDate = ParseDate(CellVal(ws, r, map, "date"));
                long? amount = ParseAmount(CellVal(ws, r, map, "amount"));
                string currency = AsText(CellVal(ws, r, map, "currency"));
                string counterparty = AsText(CellVal(ws, r, map, "counterparty"));
                string purpose = AsText(CellVal(ws, r, map, "purpose"));

                var draft = new PaymentDraft
                {
                    ExternalId = externalId,
                    OpDate = opDate,
                    AmountMinor = amount,
                    Currency = currency,
                    Counterparty = counterparty,
                    Purpose = purpose,
                    RawStatus = rawStatus,
                    Status = MapStatus(rawStatus),
                    RowKey = MakeRowKey(externalId, opDate, amount, counterparty, purpose),
                    Payload = BuildPayload(r, map, ws)
                };
                result.Drafts.Add(draft);
            }
        }

        private static object CellVal(ExcelWorksheet ws, int row, Dictionary<string, int> map, string field)
        {
            int col;
            if (!map.TryGetValue(field, out col)) return null;
            return ws.Cells[row, col].Value;
        }

        private static bool RowIsEmpty(ExcelWorksheet ws, int row, Dictionary<string, int> map)
        {
            foreach (int col in map.Values)
            {
                object v = ws.Cells[row, col].Value;
                if (v != null && Convert.ToString(v, CultureInfo.InvariantCulture).Trim().Length > 0)
                    return false;
            }
            return true;
        }

        private static string BuildPayload(int row, Dictionary<string, int> map, ExcelWorksheet ws)
        {
            var parts = new List<string>();
            foreach (var kv in map)
            {
                object v = ws.Cells[row, kv.Value].Value;
                string s = v is DateTime
                    ? ((DateTime)v).ToString("yyyy-MM-dd")
                    : Convert.ToString(v, CultureInfo.InvariantCulture);
                if (!string.IsNullOrEmpty(s)) parts.Add(kv.Key + "=" + s.Trim());
            }
            return string.Join(" | ", parts.ToArray());
        }

        private static string NormHeader(object value)
        {
            string s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            s = s.ToLowerInvariant().Trim();
            s = Regex.Replace(s, @"\s+", " ");
            return s;
        }

        private static string AsText(object v)
        {
            if (v == null) return null;
            if (v is DateTime) return ((DateTime)v).ToString("yyyy-MM-dd");
            string s = Convert.ToString(v, CultureInfo.InvariantCulture);
            if (s == null) return null;
            s = s.Trim();
            return s.Length == 0 ? null : s;
        }

        internal static DateTime? ParseDate(object v)
        {
            if (v == null) return null;

            if (v is DateTime) return ((DateTime)v).Date;

            if (v is double)
            {
                double d = (double)v;
                if (d > 20000 && d < 60000) // 1955..2064 — «разумный» диапазон дат Excel
                    return DateTime.FromOADate(d).Date;
                return null;
            }
            if (v is int) return ParseDate((double)(int)v);
            if (v is decimal) return ParseDate((double)(decimal)v);

            string s = (Convert.ToString(v, CultureInfo.InvariantCulture) ?? "").Trim();
            if (s.Length == 0) return null;

            DateTime dt;
            if (DateTime.TryParseExact(s, DateFormats, CultureInfo.GetCultureInfo("ru-RU"),
                DateTimeStyles.None, out dt))
                return dt.Date;

            if (DateTime.TryParse(s, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.None, out dt))
                return dt.Date;

            var m = Regex.Match(s, @"(\d{1,2})[./](\d{1,2})[./](\d{4})");
            if (m.Success)
            {
                int day, mon, year;
                if (int.TryParse(m.Groups[1].Value, out day) &&
                    int.TryParse(m.Groups[2].Value, out mon) &&
                    int.TryParse(m.Groups[3].Value, out year) &&
                    day >= 1 && day <= 31 && mon >= 1 && mon <= 12 && year > 1990 && year < 2100)
                {
                    return new DateTime(year, mon, day);
                }
            }
            return null;
        }

        /// <summary>Сумма в копейках: «15 000,00», «15000.00», «(5 000,00)», числа Excel.</summary>
        internal static long? ParseAmount(object v)
        {
            if (v == null) return null;
            if (v is bool) return null;

            if (v is double) return (long)Math.Round((double)v * 100.0);
            if (v is int) return (long)(int)v * 100L;
            if (v is long) return (long)((long)v * 100L);
            if (v is decimal) return (long)Math.Round((double)(decimal)v * 100.0);

            string s = (Convert.ToString(v, CultureInfo.InvariantCulture) ?? "").Trim();
            if (s.Length == 0) return null;

            bool neg = s.StartsWith("(") || s.StartsWith("-") || s.StartsWith("\u2212");
            s = Regex.Replace(s, @"[\s\u00a0\u202f]", "");
            s = s.Trim("()-\u2212".ToCharArray());
            s = Regex.Replace(s, @"[^\d.,]", ""); // убираем «руб.», «₽» и прочее

            if (s.Length == 0) return null;

            if (s.Contains(",") && s.Contains("."))
            {
                // 1.234,56 или 1,234.56 — десятичный разделитель последний
                s = s.LastIndexOf(',') > s.LastIndexOf('.')
                    ? s.Replace(".", "").Replace(",", ".")
                    : s.Replace(",", "");
            }
            else if (s.Contains(","))
            {
                s = s.Replace(",", ".");
            }

            double d;
            if (!double.TryParse(s, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out d))
                return null;

            long minor = (long)Math.Round(d * 100.0);
            return neg ? -minor : minor;
        }

        /// <summary>
        /// Ключ строки: по номеру документа, а если его нет —
        /// по отпечатку (дата+сумма+контрагент+назначение).
        /// Именно он отвечает за «тот же платёж» в разных выписках.
        /// </summary>
        internal static string MakeRowKey(string externalId, DateTime? date, long? amount,
            string counterparty, string purpose)
        {
            if (!string.IsNullOrWhiteSpace(externalId))
                return "ext:" + externalId.Trim();

            var sb = new StringBuilder();
            sb.Append(date.HasValue ? date.Value.ToString("yyyy-MM-dd") : "").Append('|');
            sb.Append(amount.HasValue ? amount.Value.ToString(CultureInfo.InvariantCulture) : "").Append('|');
            sb.Append((counterparty ?? "").Trim().ToLowerInvariant()).Append('|');
            sb.Append((purpose ?? "").Trim().ToLowerInvariant());

            byte[] hash;
            using (var sha = SHA256.Create())
                hash = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));

            var hex = new StringBuilder(32);
            for (int i = 0; i < 16; i++) hex.Append(hash[i].ToString("x2"));
            return "fp:" + hex.ToString();
        }
    }
}

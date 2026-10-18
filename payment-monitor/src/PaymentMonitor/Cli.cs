using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using PaymentMonitor.Core;

namespace PaymentMonitor
{
    /// <summary>
    /// Консольный режим (запуск из cmd/PowerShell):
    ///   PaymentMonitor.exe --import "выписка.xlsx" [--bank "Сбер"] [--force]
    ///   PaymentMonitor.exe --scan "C:\Выписки"
    ///   PaymentMonitor.exe --list [--status succeeded] [--search рога] [--limit 50]
    ///   PaymentMonitor.exe --edit "выписка.xlsx" [--row 3 --set status=Отменено --save]
    ///   PaymentMonitor.exe --runs
    ///   PaymentMonitor.exe "выписка.xlsx"    — открыть выписку в редакторе
    /// </summary>
    internal static class Cli
    {
        public static int Run(string[] args)
        {
            var log = new StringBuilder();
            int code;
            try
            {
                code = Execute(args, log);
            }
            catch (Exception ex)
            {
                log.AppendLine("Ошибка: " + ex.Message);
                code = 1;
            }

            string text = log.ToString();
            Console.WriteLine(text.TrimEnd());

            // дублируем отчёт в файл — удобно, когда консоль недоступна
            try
            {
                string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "last-run.txt"), text, Encoding.UTF8);
            }
            catch { }

            return code;
        }

        private static int Execute(string[] args, StringBuilder log)
        {
            string cmd = args[0].ToLowerInvariant();
            if (cmd == "--help" || cmd == "-h" || cmd == "/?" || cmd == "-?")
            {
                Help(log);
                return 0;
            }

            // Открыть выписку в графическом редакторе:
            //   PaymentMonitor.exe "выписка.xlsx"          (можно перетащить файл на программу)
            //   PaymentMonitor.exe --edit-gui "выписка.xlsx"
            string guiFile = null;
            if (cmd == "--edit-gui")
            {
                guiFile = PositionalAfter(args, 0);
            }
            else if ((cmd.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                      cmd.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)) && File.Exists(args[0]))
            {
                guiFile = args[0];
            }
            if (guiFile != null) return RunEditor(guiFile, log);

            AppSettings settings = SettingsStore.Load();
            string dbPath = SettingsStore.ResolveDbPath(settings);
            string bank = Opt(args, "--bank") ?? settings.Bank;

            if (cmd == "--import")
            {
                string file = PositionalAfter(args, 0);
                if (file == null || !File.Exists(file))
                {
                    log.AppendLine("Укажите файл: --import \"путь\\к\\выписке.xlsx\"");
                    return 2;
                }
                var importer = new Importer(dbPath, bank);
                ImportResult res = importer.ImportFile(file, HasFlag(args, "--force"));
                LogResult(log, res);
                return res.Error == null ? 0 : 1;
            }

            if (cmd == "--scan")
            {
                string folder = PositionalAfter(args, 0);
                if (folder == null || !Directory.Exists(folder))
                {
                    log.AppendLine("Укажите папку: --scan \"путь\\к\\папке\"");
                    return 2;
                }
                var importer = new Importer(dbPath, bank);
                var results = importer.ScanFolder(folder);
                if (results.Count == 0) log.AppendLine("В папке нет файлов xlsx/xlsm.");
                foreach (ImportResult res in results) LogResult(log, res);
                return 0;
            }

            if (cmd == "--list")
            {
                PaymentStatus? filter = null;
                string st = Opt(args, "--status");
                if (st != null)
                {
                    PaymentStatus parsed;
                    if (!Enum.TryParse(st, true, out parsed))
                    {
                        log.AppendLine("Неизвестный статус: " + st +
                            " (new, pending, succeeded, failed, canceled, refunded, unknown)");
                        return 2;
                    }
                    filter = parsed;
                }
                int limit = 30;
                string lim = Opt(args, "--limit");
                if (lim != null) int.TryParse(lim, out limit);

                using (var db = new Database(dbPath))
                {
                    var items = db.QueryPayments(filter, Opt(args, "--search"), limit);
                    if (items.Count == 0)
                    {
                        log.AppendLine("Платежей не найдено. Импортируйте выписку: --import \"файл.xlsx\"");
                        return 0;
                    }
                    log.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "{0,-10} {1,-14} {2,-11} {3,-10} {4,-24} {5}",
                        "Дата", "Сумма", "Статус", "Номер", "Контрагент", "Назначение"));
                    log.AppendLine(new string('-', 100));
                    foreach (Payment p in items)
                    {
                        log.AppendLine(string.Format(CultureInfo.InvariantCulture,
                            "{0,-10} {1,-14} {2,-11} {3,-10} {4,-24} {5}",
                            p.OpDate.HasValue ? p.OpDate.Value.ToString("dd.MM.yyyy") : "",
                            FormatMoney(p),
                            StatusText.Russian(p.Status),
                            Trunc(p.ExternalId, 10),
                            Trunc(p.Counterparty, 24),
                            Trunc(p.Purpose, 45)));
                    }
                    log.AppendLine();
                    log.AppendLine("Всего показано: " + items.Count);
                }
                return 0;
            }

            if (cmd == "--edit")
            {
                string file = PositionalAfter(args, 0);
                if (file == null || !File.Exists(file))
                {
                    log.AppendLine("Укажите файл: --edit \"путь\\к\\выписке.xlsx\" [--sheet \"Имя\"]");
                    log.AppendLine("Правка:        --edit \"файл\" --row 3 --set status=Отменено [--set amount=15000,00] [--save]");
                    return 2;
                }

                using (var doc = new StatementDocument(file))
                {
                    string sheet = Opt(args, "--sheet");
                    if (!string.IsNullOrEmpty(sheet)) doc.SelectSheet(sheet);

                    bool save = HasFlag(args, "--save");
                    bool changed = false;

                    int targetRow = -1;
                    string rowText = Opt(args, "--row");
                    if (rowText != null && !int.TryParse(rowText, out targetRow))
                    {
                        log.AppendLine("Некорректный номер строки: " + rowText);
                        return 2;
                    }

                    if (HasFlag(args, "--add-row"))
                    {
                        targetRow = doc.AddRow();
                        changed = true;
                        log.AppendLine("Добавлена строка " + targetRow + ".");
                    }

                    foreach (string kv in MultiOpt(args, "--set"))
                    {
                        if (targetRow < 1)
                        {
                            log.AppendLine("Для --set укажите --row <номер строки> (или --add-row).");
                            return 2;
                        }
                        int eq = kv.IndexOf('=');
                        if (eq <= 0)
                        {
                            log.AppendLine("Некорректное --set: " + kv + " (ожидается поле=значение)");
                            return 2;
                        }
                        string field = kv.Substring(0, eq);
                        string value = kv.Substring(eq + 1);
                        int col = doc.ColumnByFieldOrIndex(field);
                        doc.SetValue(targetRow, col, value);
                        changed = true;
                        log.AppendLine("Строка " + targetRow + ", колонка «" + field + "» = " + value);
                    }

                    string delText = Opt(args, "--delete-row");
                    if (delText != null)
                    {
                        int n;
                        if (!int.TryParse(delText, out n))
                        {
                            log.AppendLine("Некорректный номер строки для удаления: " + delText);
                            return 2;
                        }
                        doc.DeleteRow(n);
                        changed = true;
                        log.AppendLine("Удалена строка " + n + ".");
                    }

                    PrintStatement(log, doc);

                    if (changed)
                    {
                        if (save)
                        {
                            doc.Save();
                            log.AppendLine();
                            log.AppendLine("Файл сохранён: " + doc.FilePath);
                        }
                        else
                        {
                            log.AppendLine();
                            log.AppendLine("Изменения НЕ записаны (добавьте --save, чтобы сохранить в файл).");
                        }
                    }
                }
                return 0;
            }

            if (cmd == "--runs")
            {
                using (var db = new Database(dbPath))
                {
                    var runs = db.QueryRuns(15);
                    if (runs.Count == 0)
                    {
                        log.AppendLine("Журнал импортов пуст.");
                        return 0;
                    }
                    foreach (ImportRun r in runs)
                    {
                        log.AppendLine(string.Format(CultureInfo.InvariantCulture,
                            "[{0:dd.MM.yyyy HH:mm}] {1} — {2}; строк {3}, создано {4}, обновлено {5}{6}",
                            r.StartedAt, r.FileName,
                            r.Ok ? "ok" : "ОШИБКА: " + r.Error,
                            r.Rows, r.Created, r.Updated,
                            r.SkippedRows > 0 ? ", пропущено строк " + r.SkippedRows : ""));
                    }
                }
                return 0;
            }

            Help(log);
            return 2;
        }

        /// <summary>
        /// Открыть выписку в графическом редакторе. Если пользователь сохранил изменения,
        /// файл сразу импортируется в базу (как и при правке из главного окна).
        /// </summary>
        private static int RunEditor(string file, StringBuilder log)
        {
            if (!File.Exists(file))
            {
                MessageBox.Show("Файл не найден:\r\n" + file, "Редактор выписки",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 2;
            }

            AppSettings settings = SettingsStore.Load();
            string dbPath = SettingsStore.ResolveDbPath(settings);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            StatementEditorForm form;
            try
            {
                form = new StatementEditorForm(file);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Редактор выписки",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }

            using (form)
            {
                Application.Run(form);
                if (form.DialogResult == DialogResult.OK)
                {
                    var importer = new Importer(dbPath, settings.Bank);
                    LogResult(log, importer.ImportFile(file, true));
                }
            }
            return 0;
        }

        private static void LogResult(StringBuilder log, ImportResult res)
        {
            if (res.Error != null)
            {
                log.AppendLine("[ОШИБКА] " + res.FileName + ": " + res.Error);
                return;
            }
            if (res.Skipped)
            {
                log.AppendLine("[пропущен] " + res.FileName + " — уже импортирован");
                return;
            }
            ImportRun s = res.Run;
            log.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "[импорт] {0}: строк {1}, создано {2}, обновлено {3}, без изменений {4}, пропущено строк {5}",
                s.FileName, s.Rows, s.Created, s.Updated, s.Unchanged, s.SkippedRows));
        }

        private static string FormatMoney(Payment p)
        {
            if (!p.AmountMinor.HasValue) return "";
            string s = (p.AmountMinor.Value / 100m).ToString("N2", CultureInfo.GetCultureInfo("ru-RU"));
            return string.IsNullOrEmpty(p.Currency) ? s : s + " " + p.Currency;
        }

        private static string Trunc(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Length <= max ? s : s.Substring(0, max - 1) + "…";
        }

        private static string Opt(string[] args, string name)
        {
            for (int i = 0; i < args.Length - 1; i++)
                if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
                    return args[i + 1];
            return null;
        }

        private static bool HasFlag(string[] args, string name)
        {
            foreach (string a in args)
                if (string.Equals(a, name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        /// <summary>Все значения флага, который может повторяться (например, --set).</summary>
        private static List<string> MultiOpt(string[] args, string name)
        {
            var list = new List<string>();
            for (int i = 0; i < args.Length - 1; i++)
                if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
                    list.Add(args[i + 1]);
            return list;
        }

        /// <summary>Вывод таблицы выписки в консоль (та же модель, что и в редакторе).</summary>
        private static void PrintStatement(StringBuilder log, StatementDocument doc)
        {
            log.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "Лист «{0}». {1} {2}{3}",
                doc.SheetName,
                doc.HeaderDetected ? "Строка заголовков:" : "Шапка таблицы не распознана. Первая строка:",
                doc.HeaderDetected ? doc.HeaderRow : 1,
                doc.HeaderDetected ? ", колонок: " + doc.Columns.Count : ""));

            if (doc.Columns.Count == 0)
            {
                log.AppendLine("(таблица пуста)");
                return;
            }

            const int nw = 6;      // ширина колонки с номером строки
            var widths = new int[doc.Columns.Count];
            for (int i = 0; i < doc.Columns.Count; i++)
            {
                int w = doc.Columns[i].Title.Length + (doc.Columns[i].Editable ? 0 : 2);
                foreach (StatementRow r in doc.Rows)
                {
                    if (i < r.Values.Count) w = Math.Max(w, r.Values[i].Length);
                }
                widths[i] = Math.Min(Math.Max(w, 8), 26);
            }

            var head = new StringBuilder();
            head.Append("Стр.".PadRight(nw)).Append("| ");
            for (int i = 0; i < doc.Columns.Count; i++)
            {
                string t = doc.Columns[i].Title + (doc.Columns[i].Editable ? "" : " *");
                head.Append(Pad(t, widths[i])).Append(" | ");
            }
            log.AppendLine(head.ToString().TrimEnd());
            log.AppendLine(new string('-', Math.Min(200, head.Length)));

            foreach (StatementRow r in doc.Rows)
            {
                var line = new StringBuilder();
                line.Append(r.Index.ToString(CultureInfo.InvariantCulture).PadLeft(nw - 1)).Append(" | ");
                for (int i = 0; i < doc.Columns.Count; i++)
                {
                    string v = i < r.Values.Count ? r.Values[i] : "";
                    line.Append(Pad(v, widths[i])).Append(" | ");
                }
                log.AppendLine(line.ToString().TrimEnd());
            }
            log.AppendLine("(*) — колонка не распознана и не редактируется");
        }

        private static string Pad(string text, int width)
        {
            if (text == null) text = "";
            if (text.Length > width) return text.Substring(0, Math.Max(1, width - 1)) + "…";
            return text.PadRight(width);
        }

        /// <summary>Позиционный аргумент (следующий за флагами вида --xxx значение).</summary>
        private static string PositionalAfter(string[] args, int cmdIndex)
        {
            for (int i = cmdIndex + 1; i < args.Length; i++)
            {
                if (args[i].StartsWith("--", StringComparison.Ordinal))
                {
                    i++; // пропускаем значение флага
                    continue;
                }
                return args[i];
            }
            return null;
        }

        private static void Help(StringBuilder log)
        {
            log.AppendLine("PaymentMonitor — импорт выписок Excel (xlsx) и мониторинг статусов платежей");
            log.AppendLine();
            log.AppendLine("Консольные команды:");
            log.AppendLine("  --import \"выписка.xlsx\" [--bank Имя] [--force]  импорт одного файла");
            log.AppendLine("  --scan \"C:\\Выписки\"      [--bank Имя]  импорт всех новых xlsx из папки");
            log.AppendLine("  --list  [--status succeeded] [--search текст] [--limit 50]");
            log.AppendLine("  --edit \"выписка.xlsx\" [--sheet Имя]              показать таблицу выписки");
            log.AppendLine("  --edit \"выписка.xlsx\" --row 3 --set status=Отменено [--save]   изменить данные");
            log.AppendLine("  --runs                                    журнал последних импортов");
            log.AppendLine();
            log.AppendLine("  \"выписка.xlsx\"                            открыть выписку в редакторе");
            log.AppendLine("                                            (то же, если перетащить файл на программу)");
            log.AppendLine();
            log.AppendLine("Без аргументов программа запускается в графическом режиме.");
        }
    }
}

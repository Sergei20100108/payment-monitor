using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace PaymentMonitor.Core
{
    /// <summary>
    /// Импорт файлов выписок в базу.
    /// Повторный импорт того же файла (совпал sha256) пропускается.
    /// </summary>
    public sealed class Importer
    {
        public string DbPath { get; private set; }
        public string Bank { get; set; }

        public Importer(string dbPath, string bank)
        {
            DbPath = dbPath;
            Bank = string.IsNullOrWhiteSpace(bank) ? "Не указан" : bank.Trim();
        }

        /// <summary>Импорт одного файла. Никогда не бросает исключений — возвращает результат с ошибкой.</summary>
        public ImportResult ImportFile(string path, bool force)
        {
            var res = new ImportResult
            {
                FileName = Path.GetFileName(path),
                FullPath = path
            };

            string sha;
            try
            {
                sha = ComputeSha256(path);
            }
            catch (Exception ex)
            {
                res.Error = "не удалось прочитать файл: " + ex.Message;
                return res;
            }

            try
            {
                using (var db = new Database(DbPath))
                {
                    if (!force && db.RunExists(sha, true))
                    {
                        res.Skipped = true;
                        return res;
                    }

                    ParseResult parsed;
                    try
                    {
                        parsed = XlsxReader.Parse(path);
                    }
                    catch (ParseException ex)
                    {
                        db.RunInTransaction(delegate
                        {
                            db.DeleteRunBySha(sha);
                            db.InsertRun(new ImportRun
                            {
                                FilePath = path,
                                FileName = res.FileName,
                                FileSha256 = sha,
                                Bank = Bank,
                                StartedAt = DateTime.UtcNow,
                                FinishedAt = DateTime.UtcNow,
                                Ok = false,
                                Error = ex.Message
                            });
                        });
                        res.Error = ex.Message;
                        return res;
                    }

                    DateTime now = DateTime.UtcNow;
                    var run = new ImportRun
                    {
                        FilePath = path,
                        FileName = res.FileName,
                        FileSha256 = sha,
                        Bank = Bank,
                        StartedAt = now,
                        Ok = true
                    };

                    db.RunInTransaction(delegate
                    {
                        db.DeleteRunBySha(sha);
                        foreach (PaymentDraft d in parsed.Drafts)
                        {
                            string change = db.ApplyDraft(d, Bank, res.FileName, now);
                            if (change == "created") run.Created++;
                            else if (change == "updated") run.Updated++;
                            else run.Unchanged++;
                        }
                        run.Rows = parsed.Drafts.Count;
                        run.SkippedRows = parsed.SkippedRows;
                        run.FinishedAt = DateTime.UtcNow;
                        db.InsertRun(run);   // пишем журнал уже с итоговыми счётчиками
                    });

                    res.Run = run;
                    return res;
                }
            }
            catch (Exception ex)
            {
                res.Error = ex.Message;
                return res;
            }
        }

        /// <summary>Обход папки: импорт всех xlsx/xlsm, ещё не виденных (по sha256).</summary>
        public List<ImportResult> ScanFolder(string folder)
        {
            var results = new List<ImportResult>();
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                results.Add(new ImportResult { FileName = folder, Error = "папка не найдена" });
                return results;
            }

            var files = new List<string>();
            files.AddRange(Directory.GetFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly));
            files.AddRange(Directory.GetFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly));
            files.Sort(StringComparer.OrdinalIgnoreCase);

            foreach (string f in files)
            {
                string name = Path.GetFileName(f);
                if (name.StartsWith("~$", StringComparison.Ordinal)) continue; // временные файлы Excel
                results.Add(ImportFile(f, false));
            }
            return results;
        }

        private static string ComputeSha256(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(fs);
                var hex = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash) hex.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                return hex.ToString();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO;

namespace PaymentMonitor.Core
{
    /// <summary>
    /// Локальная база SQLite: платежи, история изменений, журнал импортов.
    /// Один экземпляр = одно соединение; потокобезопасности нет,
    /// поэтому фоновые импорты создают собственные экземпляры.
    /// </summary>
    public sealed class Database : IDisposable
    {
        private readonly SQLiteConnection _conn;
        private SQLiteTransaction _tx;

        public string DbPath { get; private set; }

        public Database(string dbPath)
        {
            DbPath = dbPath;
            string dir = Path.GetDirectoryName(dbPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            string cs = string.Format(CultureInfo.InvariantCulture,
                "Data Source={0};Version=3;Pooling=True;Default Timeout=30;", dbPath);
            _conn = new SQLiteConnection(cs);
            _conn.Open();

            using (var cmd = Cmd("PRAGMA journal_mode=WAL; PRAGMA busy_timeout=5000; PRAGMA foreign_keys=ON;"))
                cmd.ExecuteNonQuery();

            EnsureSchema();
        }

        private void EnsureSchema()
        {
            const string sql = @"
CREATE TABLE IF NOT EXISTS payments (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    bank            TEXT NOT NULL,
    row_key         TEXT NOT NULL,
    external_id     TEXT,
    op_date         TEXT,
    amount_minor    INTEGER,
    currency        TEXT,
    counterparty    TEXT,
    purpose         TEXT,
    status          TEXT NOT NULL,
    raw_status      TEXT,
    source          TEXT,
    first_seen_file TEXT,
    created_at      TEXT NOT NULL,
    updated_at      TEXT NOT NULL,
    UNIQUE (bank, row_key)
);
CREATE INDEX IF NOT EXISTS ix_payments_status ON payments(status);
CREATE INDEX IF NOT EXISTS ix_payments_ext    ON payments(external_id);

CREATE TABLE IF NOT EXISTS payment_events (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    payment_id  INTEGER NOT NULL,
    from_status TEXT,
    to_status   TEXT NOT NULL,
    source      TEXT,
    file        TEXT,
    payload     TEXT,
    created_at  TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_events_payment ON payment_events(payment_id);

CREATE TABLE IF NOT EXISTS import_runs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path       TEXT NOT NULL,
    file_name       TEXT NOT NULL,
    file_sha256     TEXT NOT NULL UNIQUE,
    bank            TEXT NOT NULL,
    started_at      TEXT NOT NULL,
    finished_at     TEXT,
    ok              INTEGER NOT NULL DEFAULT 0,
    rows_count      INTEGER NOT NULL DEFAULT 0,
    created_count   INTEGER NOT NULL DEFAULT 0,
    updated_count   INTEGER NOT NULL DEFAULT 0,
    unchanged_count INTEGER NOT NULL DEFAULT 0,
    skipped_count   INTEGER NOT NULL DEFAULT 0,
    error           TEXT
);";
            using (var cmd = Cmd(sql)) cmd.ExecuteNonQuery();
        }

        private SQLiteCommand Cmd(string sql)
        {
            var c = _conn.CreateCommand();
            c.CommandText = sql;
            if (_tx != null) c.Transaction = _tx;
            return c;
        }

        private static void P(SQLiteCommand c, string name, object value)
        {
            c.Parameters.AddWithValue(name, value ?? DBNull.Value);
        }

        private static string Iso(DateTime t)
        {
            return t.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
        }

        private static DateTime ParseIso(string s)
        {
            DateTime t;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out t))
                return t;
            return DateTime.MinValue;
        }

        // ── журнал импортов ─────────────────────────────────────────────

        public void RunInTransaction(Action action)
        {
            using (var tx = _conn.BeginTransaction())
            {
                _tx = tx;
                try
                {
                    action();
                    tx.Commit();
                }
                finally
                {
                    _tx = null;
                }
            }
        }

        public bool RunExists(string sha, bool onlyOk)
        {
            using (var c = Cmd(onlyOk
                ? "SELECT COUNT(*) FROM import_runs WHERE file_sha256=@s AND ok=1"
                : "SELECT COUNT(*) FROM import_runs WHERE file_sha256=@s"))
            {
                P(c, "@s", sha);
                return Convert.ToInt32(c.ExecuteScalar(), CultureInfo.InvariantCulture) > 0;
            }
        }

        public void DeleteRunBySha(string sha)
        {
            using (var c = Cmd("DELETE FROM import_runs WHERE file_sha256=@s"))
            {
                P(c, "@s", sha);
                c.ExecuteNonQuery();
            }
        }

        public void InsertRun(ImportRun run)
        {
            using (var c = Cmd(@"INSERT INTO import_runs
    (file_path, file_name, file_sha256, bank, started_at, finished_at, ok,
     rows_count, created_count, updated_count, unchanged_count, skipped_count, error)
VALUES
    (@path, @name, @sha, @bank, @started, @finished, @ok,
     @rows, @created, @updated, @unchanged, @skipped, @error)"))
            {
                P(c, "@path", run.FilePath); P(c, "@name", run.FileName); P(c, "@sha", run.FileSha256);
                P(c, "@bank", run.Bank); P(c, "@started", Iso(run.StartedAt));
                P(c, "@finished", run.FinishedAt.HasValue ? Iso(run.FinishedAt.Value) : null);
                P(c, "@ok", run.Ok ? 1 : 0); P(c, "@rows", run.Rows); P(c, "@created", run.Created);
                P(c, "@updated", run.Updated); P(c, "@unchanged", run.Unchanged);
                P(c, "@skipped", run.SkippedRows); P(c, "@error", run.Error);
                c.ExecuteNonQuery();
            }
        }

        public List<ImportRun> QueryRuns(int limit)
        {
            var list = new List<ImportRun>();
            using (var c = Cmd("SELECT * FROM import_runs ORDER BY id DESC LIMIT @lim"))
            {
                P(c, "@lim", limit);
                using (var rd = c.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        list.Add(new ImportRun
                        {
                            Id = GetLong(rd, "id"),
                            FilePath = GetStr(rd, "file_path"),
                            FileName = GetStr(rd, "file_name"),
                            FileSha256 = GetStr(rd, "file_sha256"),
                            Bank = GetStr(rd, "bank"),
                            StartedAt = ParseIso(GetStr(rd, "started_at")),
                            FinishedAt = GetStr(rd, "finished_at") == null ? (DateTime?)null : ParseIso(GetStr(rd, "finished_at")),
                            Ok = GetLong(rd, "ok") != 0,
                            Rows = (int)GetLong(rd, "rows_count"),
                            Created = (int)GetLong(rd, "created_count"),
                            Updated = (int)GetLong(rd, "updated_count"),
                            Unchanged = (int)GetLong(rd, "unchanged_count"),
                            SkippedRows = (int)GetLong(rd, "skipped_count"),
                            Error = GetStr(rd, "error")
                        });
                    }
                }
            }
            return list;
        }

        // ── платежи ─────────────────────────────────────────────────────

        /// <summary>
        /// Вставка/обновление платежа по черновику из выписки.
        /// Идемпотентно: повторный импорт того же файла ничего не меняет.
        /// Возвращает created / updated / unchanged.
        /// </summary>
        public string ApplyDraft(PaymentDraft d, string bank, string file, DateTime nowUtc)
        {
            Payment existing = null;
            using (var c = Cmd(@"SELECT id, status, raw_status, external_id, op_date, amount_minor,
                                        currency, counterparty, purpose
                                 FROM payments WHERE bank=@bank AND row_key=@key"))
            {
                P(c, "@bank", bank); P(c, "@key", d.RowKey);
                using (var rd = c.ExecuteReader())
                {
                    if (rd.Read()) existing = ReadPayment(rd);
                }
            }

            if (existing == null)
            {
                using (var c = Cmd(@"INSERT INTO payments
    (bank, row_key, external_id, op_date, amount_minor, currency, counterparty, purpose,
     status, raw_status, source, first_seen_file, created_at, updated_at)
VALUES
    (@bank, @key, @ext, @date, @amount, @cur, @cpty, @purpose,
     @status, @raw, 'parsing', @file, @now, @now)"))
                {
                    P(c, "@bank", bank); P(c, "@key", d.RowKey); P(c, "@ext", d.ExternalId);
                    P(c, "@date", d.OpDate.HasValue ? d.OpDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : null);
                    P(c, "@amount", d.AmountMinor); P(c, "@cur", d.Currency);
                    P(c, "@cpty", d.Counterparty); P(c, "@purpose", d.Purpose);
                    P(c, "@status", StatusText.Key(d.Status)); P(c, "@raw", d.RawStatus);
                    P(c, "@file", file); P(c, "@now", Iso(nowUtc));
                    c.ExecuteNonQuery();
                }
                long newId = LastInsertId();
                InsertEvent(newId, null, StatusText.Key(d.Status), file, d.Payload, nowUtc);
                return "created";
            }

            // ── уже есть: обновляем известные поля, при необходимости пишем событие
            PaymentStatus finalStatus = existing.Status;
            string finalRaw = existing.RawStatus;
            bool statusChanged = false;

            if (existing.Status != d.Status)
            {
                bool accept = !StatusText.IsWeak(d.Status) || StatusText.IsWeak(existing.Status);
                if (accept)
                {
                    finalStatus = d.Status;
                    finalRaw = d.RawStatus;
                    statusChanged = true;
                }
                else
                {
                    // известный статус не деградируем из-за пустой/неизвестной строки
                    finalRaw = d.RawStatus;
                }
            }
            else if ((existing.RawStatus ?? "") != (d.RawStatus ?? ""))
            {
                finalRaw = d.RawStatus;
            }

            string extId = d.ExternalId ?? existing.ExternalId;
            DateTime? opDate = d.OpDate ?? existing.OpDate;
            long? amount = d.AmountMinor ?? existing.AmountMinor;
            string cur = d.Currency ?? existing.Currency;
            string cpty = d.Counterparty ?? existing.Counterparty;
            string purpose = d.Purpose ?? existing.Purpose;

            bool fieldsChanged =
                (extId ?? "") != (existing.ExternalId ?? "") ||
                opDate != existing.OpDate ||
                amount != existing.AmountMinor ||
                (cur ?? "") != (existing.Currency ?? "") ||
                (cpty ?? "") != (existing.Counterparty ?? "") ||
                (purpose ?? "") != (existing.Purpose ?? "");

            bool rawChanged = (finalRaw ?? "") != (existing.RawStatus ?? "");

            if (statusChanged || fieldsChanged || rawChanged)
            {
                using (var c = Cmd(@"UPDATE payments SET
    external_id=@ext, op_date=@date, amount_minor=@amount, currency=@cur,
    counterparty=@cpty, purpose=@purpose, status=@status, raw_status=@raw, updated_at=@now
WHERE id=@id"))
                {
                    P(c, "@ext", extId);
                    P(c, "@date", opDate.HasValue ? opDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : null);
                    P(c, "@amount", amount); P(c, "@cur", cur); P(c, "@cpty", cpty);
                    P(c, "@purpose", purpose); P(c, "@status", StatusText.Key(finalStatus));
                    P(c, "@raw", finalRaw); P(c, "@now", Iso(nowUtc)); P(c, "@id", existing.Id);
                    c.ExecuteNonQuery();
                }
                if (statusChanged)
                    InsertEvent(existing.Id, StatusText.Key(existing.Status), StatusText.Key(finalStatus),
                        file, d.Payload, nowUtc);
                return "updated";
            }

            return "unchanged";
        }

        private void InsertEvent(long paymentId, string from, string to, string file, string payload, DateTime nowUtc)
        {
            using (var c = Cmd(@"INSERT INTO payment_events
    (payment_id, from_status, to_status, source, file, payload, created_at)
VALUES (@pid, @from, @to, 'parsing', @file, @payload, @now)"))
            {
                P(c, "@pid", paymentId); P(c, "@from", from); P(c, "@to", to);
                P(c, "@file", file); P(c, "@payload", payload); P(c, "@now", Iso(nowUtc));
                c.ExecuteNonQuery();
            }
        }

        private long LastInsertId()
        {
            using (var c = Cmd("SELECT last_insert_rowid()"))
                return Convert.ToInt64(c.ExecuteScalar(), CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Список платежей. Статус фильтруется в SQL; поиск — в памяти
        /// (SQLite не умеет регистронезависимый поиск по кириллице).
        /// </summary>
        public List<Payment> QueryPayments(PaymentStatus? status, string search, int limit)
        {
            var all = new List<Payment>();
            using (var c = Cmd(@"SELECT id, bank, row_key, external_id, op_date, amount_minor, currency,
                                        counterparty, purpose, status, raw_status, source, first_seen_file,
                                        created_at, updated_at
                                 FROM payments
                                 ORDER BY op_date IS NULL, op_date DESC, id DESC
                                 LIMIT 5000"))
            {
                using (var rd = c.ExecuteReader())
                {
                    while (rd.Read()) all.Add(ReadPayment(rd));
                }
            }

            string needle = string.IsNullOrWhiteSpace(search) ? null : search.Trim();
            var result = new List<Payment>();
            foreach (Payment p in all)
            {
                if (status.HasValue && p.Status != status.Value) continue;
                if (needle != null &&
                    !Contains(p.ExternalId, needle) && !Contains(p.Counterparty, needle) &&
                    !Contains(p.Purpose, needle) && !Contains(p.RawStatus, needle))
                    continue;

                result.Add(p);
                if (result.Count >= limit) break;
            }
            return result;
        }

        private static bool Contains(string hay, string needle)
        {
            return hay != null &&
                   hay.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public List<PaymentEvent> QueryEvents(long paymentId, int limit)
        {
            var list = new List<PaymentEvent>();
            using (var c = Cmd("SELECT * FROM payment_events WHERE payment_id=@pid ORDER BY id DESC LIMIT @lim"))
            {
                P(c, "@pid", paymentId); P(c, "@lim", limit);
                using (var rd = c.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        list.Add(new PaymentEvent
                        {
                            Id = GetLong(rd, "id"),
                            PaymentId = GetLong(rd, "payment_id"),
                            FromStatus = GetStr(rd, "from_status"),
                            ToStatus = GetStr(rd, "to_status"),
                            Source = GetStr(rd, "source"),
                            File = GetStr(rd, "file"),
                            Payload = GetStr(rd, "payload"),
                            CreatedAt = ParseIso(GetStr(rd, "created_at"))
                        });
                    }
                }
            }
            return list;
        }

        public Dictionary<PaymentStatus, int> CountByStatus()
        {
            var counts = new Dictionary<PaymentStatus, int>();
            using (var c = Cmd("SELECT status, COUNT(*) AS n FROM payments GROUP BY status"))
            {
                using (var rd = c.ExecuteReader())
                {
                    while (rd.Read())
                        counts[StatusText.ParseKey(GetStr(rd, "status"))] = (int)GetLong(rd, "n");
                }
            }
            return counts;
        }

        // ── чтение строк ────────────────────────────────────────────────

        private static Payment ReadPayment(IDataRecord rd)
        {
            string opDate = GetStr(rd, "op_date");
            DateTime? date = null;
            if (!string.IsNullOrEmpty(opDate))
            {
                DateTime t;
                if (DateTime.TryParse(opDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out t))
                    date = t.Date;
            }
            long? amount = null;
            if (GetStr(rd, "amount_minor") != null) amount = GetLong(rd, "amount_minor");

            return new Payment
            {
                Id = GetLong(rd, "id"),
                Bank = GetStr(rd, "bank"),
                RowKey = GetStr(rd, "row_key"),
                ExternalId = GetStr(rd, "external_id"),
                OpDate = date,
                AmountMinor = amount,
                Currency = GetStr(rd, "currency"),
                Counterparty = GetStr(rd, "counterparty"),
                Purpose = GetStr(rd, "purpose"),
                Status = StatusText.ParseKey(GetStr(rd, "status")),
                RawStatus = GetStr(rd, "raw_status"),
                Source = GetStr(rd, "source"),
                FirstSeenFile = GetStr(rd, "first_seen_file"),
                CreatedAt = ParseIso(GetStr(rd, "created_at")),
                UpdatedAt = ParseIso(GetStr(rd, "updated_at"))
            };
        }

        private static string GetStr(IDataRecord rd, string name)
        {
            int i = rd.GetOrdinal(name);
            return rd.IsDBNull(i) ? null : Convert.ToString(rd.GetValue(i), CultureInfo.InvariantCulture);
        }

        private static long GetLong(IDataRecord rd, string name)
        {
            int i = rd.GetOrdinal(name);
            return rd.IsDBNull(i) ? 0L : Convert.ToInt64(rd.GetValue(i), CultureInfo.InvariantCulture);
        }

        public void Dispose()
        {
            _conn.Dispose();
        }
    }
}

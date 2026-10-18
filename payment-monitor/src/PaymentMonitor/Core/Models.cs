using System;
using System.Collections.Generic;

namespace PaymentMonitor.Core
{
    /// <summary>Нормализованный статус платежа (общий для всех банков).</summary>
    public enum PaymentStatus
    {
        New,        // создан, статуса ещё нет
        Pending,    // в обработке
        Succeeded,  // исполнен
        Failed,     // отклонён
        Canceled,   // отменён
        Refunded,   // возврат/сторно
        Unknown     // банк прислал что-то непонятное
    }

    /// <summary>Тексты и преобразования статусов.</summary>
    public static class StatusText
    {
        public static PaymentStatus[] All()
        {
            return (PaymentStatus[])Enum.GetValues(typeof(PaymentStatus));
        }

        public static string Russian(PaymentStatus s)
        {
            switch (s)
            {
                case PaymentStatus.New: return "Новый";
                case PaymentStatus.Pending: return "В обработке";
                case PaymentStatus.Succeeded: return "Исполнен";
                case PaymentStatus.Failed: return "Отклонён";
                case PaymentStatus.Canceled: return "Отменён";
                case PaymentStatus.Refunded: return "Возврат";
                default: return "Неизвестен";
            }
        }

        /// <summary>Ключ для хранения в БД: new, pending, succeeded, failed, canceled, refunded, unknown.</summary>
        public static string Key(PaymentStatus s)
        {
            return s.ToString().ToLowerInvariant();
        }

        public static PaymentStatus ParseKey(string key)
        {
            PaymentStatus s;
            if (!string.IsNullOrEmpty(key) && Enum.TryParse(key, true, out s))
                return s;
            return PaymentStatus.Unknown;
        }

        public static bool IsWeak(PaymentStatus s)
        {
            // «Слабые» статусы не должны затирать уже известный сильный статус
            return s == PaymentStatus.New || s == PaymentStatus.Unknown;
        }
    }

    /// <summary>Платёж в локальной базе.</summary>
    public class Payment
    {
        public long Id { get; set; }
        public string Bank { get; set; }
        public string RowKey { get; set; }
        public string ExternalId { get; set; }
        public DateTime? OpDate { get; set; }
        public long? AmountMinor { get; set; }
        public string Currency { get; set; }
        public string Counterparty { get; set; }
        public string Purpose { get; set; }
        public PaymentStatus Status { get; set; }
        public string RawStatus { get; set; }
        public string Source { get; set; }
        public string FirstSeenFile { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    /// <summary>Событие об изменении статуса (история).</summary>
    public class PaymentEvent
    {
        public long Id { get; set; }
        public long PaymentId { get; set; }
        public string FromStatus { get; set; }
        public string ToStatus { get; set; }
        public string Source { get; set; }
        public string File { get; set; }
        public string Payload { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    /// <summary>Журнал импорта одного файла.</summary>
    public class ImportRun
    {
        public long Id { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FileSha256 { get; set; }
        public string Bank { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime? FinishedAt { get; set; }
        public bool Ok { get; set; }
        public int Rows { get; set; }
        public int Created { get; set; }
        public int Updated { get; set; }
        public int Unchanged { get; set; }
        public int SkippedRows { get; set; }
        public string Error { get; set; }
    }

    /// <summary>Строка выписки, приведённая к единому виду.</summary>
    public class PaymentDraft
    {
        public string ExternalId { get; set; }
        public DateTime? OpDate { get; set; }
        public long? AmountMinor { get; set; }
        public string Currency { get; set; }
        public string Counterparty { get; set; }
        public string Purpose { get; set; }
        public string RawStatus { get; set; }
        public PaymentStatus Status { get; set; }
        public string RowKey { get; set; }
        public string Payload { get; set; }
    }

    /// <summary>Результат разбора одного файла.</summary>
    public class ParseResult
    {
        public List<PaymentDraft> Drafts = new List<PaymentDraft>();
        public int SkippedRows;
    }

    /// <summary>Результат импорта одного файла.</summary>
    public class ImportResult
    {
        public string FileName;
        public string FullPath;
        public bool Skipped;      // файл уже импортирован (совпал sha256)
        public ImportRun Run;     // заполнен при успешном импорте
        public string Error;      // текст ошибки, если есть
    }
}

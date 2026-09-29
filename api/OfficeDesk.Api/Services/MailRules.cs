using OfficeDesk.Api.Data;
using OfficeDesk.Api.Models;

namespace OfficeDesk.Api.Services;

public static class MailRules
{
    public static readonly TimeZoneInfo Riyadh = TimeZoneInfo.FindSystemTimeZoneById("Asia/Riyadh");

    public static DateTimeOffset Now() => TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, Riyadh);

    public static bool IsWorkTime(DateTimeOffset local, Settings settings)
    {
        var days = settings.WorkDays.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (!days.Contains(local.DayOfWeek.ToString(), StringComparer.OrdinalIgnoreCase))
            return false;

        var start = TimeOnly.TryParse(settings.WorkStart, out var s) ? s : new TimeOnly(8, 0);
        var end = TimeOnly.TryParse(settings.WorkEnd, out var e) ? e : new TimeOnly(17, 0);
        var time = TimeOnly.FromDateTime(local.DateTime);
        return time >= start && time < end;
    }

    public static bool IsUrgent(string subject, string preview, string words)
    {
        var hay = $"{subject}\n{preview}";
        foreach (var word in words.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (hay.Contains(word, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    public static bool IsMarketing(SyncItem item)
    {
        var text = $"{item.Subject} {item.Preview} {item.Sender}";
        string[] keys = ["unsubscribe", "newsletter", "نشرة", "تسويق", "عرض خاص", "خصم حصري"];
        if (keys.Any(key => text.Contains(key, StringComparison.OrdinalIgnoreCase)))
            return true;

        var greeting = text.Contains("تهنئة", StringComparison.Ordinal)
            || text.Contains("مبروك", StringComparison.Ordinal);
        var machine = (item.SenderAddress ?? "").Contains("noreply", StringComparison.OrdinalIgnoreCase)
            || (item.SenderAddress ?? "").Contains("no-reply", StringComparison.OrdinalIgnoreCase);
        return greeting && machine && (item.Preview?.Length ?? 0) < 40;
    }

    public static bool IsTeams(SyncItem item)
    {
        if (item.IsTeams)
            return true;
        var blob = $"{item.MessageClass} {item.Sender} {item.SenderAddress} {item.Subject}";
        return blob.Contains("Teams", StringComparison.OrdinalIgnoreCase)
            || blob.Contains("teams.microsoft", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsHr(SyncItem item)
    {
        var text = $"{item.Subject} {item.Preview}";
        return text.Contains("تعميم", StringComparison.Ordinal)
            || text.Contains("موارد بشرية", StringComparison.Ordinal)
            || text.Contains("human resources", StringComparison.OrdinalIgnoreCase);
    }

    public static string Classify(SyncItem item, Settings settings)
    {
        if (item.Folder == "sent")
            return "sent";
        if (item.Folder == "drafts")
            return "draft";
        if (IsMarketing(item))
            return "marketing";
        if (IsUrgent(item.Subject, item.Preview, settings.UrgentWords))
            return "urgent";
        if (IsTeams(item))
            return "teams";
        if (IsHr(item))
            return "hr";
        return "action";
    }

    public static string ConversationId(SyncItem item) =>
        string.IsNullOrWhiteSpace(item.ConversationId) ? item.EntryId : item.ConversationId;
}

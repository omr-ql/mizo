namespace OfficeDesk.Api.Data;

public sealed class Settings
{
    public int Id { get; set; }
    public string DisplayName { get; set; } = "mizo";
    public string AccountSmtp { get; set; } = "o.bamusa@ejadah-sa.com";
    public string WorkDays { get; set; } = "Sunday,Monday,Tuesday,Wednesday,Thursday";
    public string WorkStart { get; set; } = "08:00";
    public string WorkEnd { get; set; } = "17:00";
    public string UrgentWords { get; set; } = "عاجل,urgent,asap,فوري,immediate,deadline";
    public int PollIntervalMinutes { get; set; } = 3;
    public bool StartWithWindows { get; set; }
}

public sealed class MailMessage
{
    public int Id { get; set; }
    public string EntryId { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string ConversationId { get; set; } = "";
    public string Folder { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string SenderAddress { get; set; } = "";
    public DateTimeOffset ReceivedAt { get; set; }
    public string Preview { get; set; } = "";
    public bool Unread { get; set; }
    public string Classification { get; set; } = "";
    public DateTimeOffset SyncedAt { get; set; }
}

public sealed class DeskThread
{
    public int Id { get; set; }
    public string ConversationId { get; set; } = "";
    public string Subject { get; set; } = "";
    public string State { get; set; } = "awaiting_me";
    public string RecommendedStep { get; set; } = "";
    public string EntryId { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string CounterpartName { get; set; } = "";
    public string CounterpartAddress { get; set; } = "";
    public DateTimeOffset LastActivityAt { get; set; }
}

public sealed class DeskAlert
{
    public int Id { get; set; }
    public string Title { get; set; } = "";
    public string Message { get; set; } = "";
    public string Source { get; set; } = "mail";
    public string Kind { get; set; } = "action";
    public string? EntryId { get; set; }
    public string? StoreId { get; set; }
    public DateTimeOffset CreatedAt { get; set; }
    public DateTimeOffset? DismissedAt { get; set; }
}

public sealed class BriefingRecord
{
    public int Id { get; set; }
    public string Slot { get; set; } = "";
    public string Body { get; set; } = "";
    public DateTimeOffset CreatedAt { get; set; }
}

public sealed class Draft
{
    public int Id { get; set; }
    public int? ThreadId { get; set; }
    public string? EntryId { get; set; }
    public string? StoreId { get; set; }
    public string To { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Body { get; set; } = "";
    public string Status { get; set; } = "editing";
    public DateTimeOffset CreatedAt { get; set; }
    public DateTimeOffset UpdatedAt { get; set; }
}

public sealed class OutlookJob
{
    public int Id { get; set; }
    public string Kind { get; set; } = "";
    public string PayloadJson { get; set; } = "{}";
    public string Status { get; set; } = "pending";
    public string? Result { get; set; }
    public int? DraftId { get; set; }
    public DateTimeOffset CreatedAt { get; set; }
    public DateTimeOffset? CompletedAt { get; set; }
}

public sealed class PendingAction
{
    public int Id { get; set; }
    public string Title { get; set; } = "";
    public string Notes { get; set; } = "";
    public bool Done { get; set; }
    public DateTimeOffset CreatedAt { get; set; }
}

public sealed class DecisionLog
{
    public int Id { get; set; }
    public string Kind { get; set; } = "";
    public string Detail { get; set; } = "";
    public DateTimeOffset CreatedAt { get; set; }
}

public sealed class HostStatus
{
    public int Id { get; set; }
    public bool Ok { get; set; }
    public string Message { get; set; } = "";
    public DateTimeOffset UpdatedAt { get; set; }
}

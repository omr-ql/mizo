namespace OfficeDesk.Api.Models;

public sealed class SyncItem
{
    public string EntryId { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string ConversationId { get; set; } = "";
    public string Folder { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string SenderAddress { get; set; } = "";
    public DateTimeOffset ReceivedAt { get; set; }
    public string Preview { get; set; } = "";
    public string Body { get; set; } = "";
    public bool Unread { get; set; }
    public bool IsTeams { get; set; }
    public string MessageClass { get; set; } = "";
}

public sealed class SyncRequest
{
    public List<SyncItem> Items { get; set; } = [];
}

public sealed class SettingsWrite
{
    public string DisplayName { get; set; } = "";
    public string AccountSmtp { get; set; } = "";
    public string WorkDays { get; set; } = "";
    public string WorkStart { get; set; } = "";
    public string WorkEnd { get; set; } = "";
    public string UrgentWords { get; set; } = "";
    public int PollIntervalMinutes { get; set; } = 3;
    public bool StartWithWindows { get; set; }
}

public sealed class PatchThreadRequest
{
    public string? State { get; set; }
}

public sealed class DraftWrite
{
    public int? ThreadId { get; set; }
    public string? EntryId { get; set; }
    public string? StoreId { get; set; }
    public string To { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Body { get; set; } = "";
}

public sealed class OpenMailRequest
{
    public string EntryId { get; set; } = "";
    public string? StoreId { get; set; }
}

public sealed class PendingWrite
{
    public string Title { get; set; } = "";
    public string Notes { get; set; } = "";
    public bool? Done { get; set; }
}

public sealed class CompleteJobRequest
{
    public bool Success { get; set; }
    public string? Error { get; set; }
}

public sealed class HostReport
{
    public bool Ok { get; set; }
    public string Message { get; set; } = "";
}

public sealed class GenerateBriefingRequest
{
    public bool QuietIfEmpty { get; set; } = true;
}

public sealed class MailDigestView
{
    public string Lead { get; set; } = "";
    public List<MailDigestItem> Items { get; set; } = [];
}

public sealed class MailDigestItem
{
    public int Id { get; set; }
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string Summary { get; set; } = "";
    public DateTimeOffset ReceivedAt { get; set; }
    public string EntryId { get; set; } = "";
    public string StoreId { get; set; } = "";
    public bool Unread { get; set; }
    public int? ThreadId { get; set; }
    public string State { get; set; } = "";
    public string CounterpartAddress { get; set; } = "";
}

public sealed class BriefingItem
{
    public int ThreadId { get; set; }
    public string Subject { get; set; } = "";
    public string State { get; set; } = "";
    public string RecommendedStep { get; set; } = "";
    public string EntryId { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string CounterpartName { get; set; } = "";
    public string CounterpartAddress { get; set; } = "";
    public DateTimeOffset LastActivityAt { get; set; }
}

public sealed class BriefingView
{
    public List<BriefingItem> NeedsYou { get; set; } = [];
    public List<BriefingItem> Waiting { get; set; } = [];
    public List<BriefingItem> Info { get; set; } = [];
    public DateTimeOffset GeneratedAt { get; set; }
}

using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using OfficeDesk.Api.Data;
using OfficeDesk.Api.Models;

namespace OfficeDesk.Api.Services;

public sealed class OfficeService(OfficeDb db)
{
    static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    static readonly HashSet<string> States = ["awaiting_me", "awaiting_them", "info", "closed"];

    public Settings EnsureSettings()
    {
        var settings = db.Settings.FirstOrDefault();
        if (settings is not null)
            return settings;
        settings = new Settings();
        db.Settings.Add(settings);
        db.SaveChanges();
        return settings;
    }

    public Settings UpdateSettings(SettingsWrite write)
    {
        if (string.IsNullOrWhiteSpace(write.AccountSmtp))
            throw new InvalidOperationException("حساب البريد مطلوب.");
        if (write.PollIntervalMinutes is < 1 or > 60)
            throw new InvalidOperationException("فترة المراقبة بين 1 و 60 دقيقة.");

        var settings = EnsureSettings();
        settings.DisplayName = string.IsNullOrWhiteSpace(write.DisplayName) ? "mizo" : write.DisplayName.Trim();
        settings.AccountSmtp = write.AccountSmtp.Trim();
        settings.WorkDays = string.IsNullOrWhiteSpace(write.WorkDays)
            ? "Sunday,Monday,Tuesday,Wednesday,Thursday"
            : write.WorkDays.Trim();
        settings.WorkStart = string.IsNullOrWhiteSpace(write.WorkStart) ? "08:00" : write.WorkStart.Trim();
        settings.WorkEnd = string.IsNullOrWhiteSpace(write.WorkEnd) ? "17:00" : write.WorkEnd.Trim();
        settings.UrgentWords = string.IsNullOrWhiteSpace(write.UrgentWords)
            ? "عاجل,urgent,asap,فوري,immediate,deadline"
            : write.UrgentWords.Trim();
        settings.PollIntervalMinutes = write.PollIntervalMinutes;
        settings.StartWithWindows = write.StartWithWindows;
        db.SaveChanges();
        return settings;
    }

    public List<DeskAlert> Sync(SyncRequest request)
    {
        var settings = EnsureSettings();
        var now = DateTimeOffset.UtcNow;
        var work = MailRules.IsWorkTime(MailRules.Now(), settings);
        var created = new List<DeskAlert>();

        foreach (var item in request.Items)
        {
            if (string.IsNullOrWhiteSpace(item.EntryId))
                continue;

            var row = db.MailMessages.FirstOrDefault(m => m.EntryId == item.EntryId);
            var isNew = row is null;
            row ??= new MailMessage();
            if (isNew)
                db.MailMessages.Add(row);

            row.EntryId = item.EntryId;
            row.StoreId = item.StoreId ?? "";
            row.ConversationId = MailRules.ConversationId(item);
            row.Folder = item.Folder ?? "";
            row.Subject = item.Subject ?? "";
            row.Sender = item.Sender ?? "";
            row.SenderAddress = item.SenderAddress ?? "";
            row.ReceivedAt = item.ReceivedAt;
            row.Preview = item.Preview ?? "";
            row.Unread = item.Unread;
            row.Classification = MailRules.Classify(item, settings);
            row.SyncedAt = now;

            if (!isNew || row.Folder != "inbox" || row.Classification == "marketing")
                continue;

            var urgent = row.Classification == "urgent";
            if (!work && !urgent)
                continue;
            if (AlreadyAnswered(request, row))
                continue;
            if (db.Alerts.Any(a => a.EntryId == row.EntryId && a.DismissedAt == null))
                continue;

            var alert = MakeAlert(row);
            db.Alerts.Add(alert);
            created.Add(alert);
        }

        db.SaveChanges();
        UpdateThreads(request);
        db.DecisionLogs.Add(new DecisionLog
        {
            Kind = "sync",
            Detail = $"عناصر={request.Items.Count} تنبيهات={created.Count}",
            CreatedAt = now
        });
        db.SaveChanges();
        return created;
    }

    public List<MailMessage> Mail(string? folder)
    {
        var query = db.MailMessages.AsQueryable();
        if (!string.IsNullOrWhiteSpace(folder))
            query = query.Where(m => m.Folder == folder);
        return query.AsEnumerable().OrderByDescending(m => m.ReceivedAt).Take(80).ToList();
    }

    public List<DeskThread> Threads() =>
        db.Threads.AsEnumerable().OrderByDescending(t => t.LastActivityAt).Take(100).ToList();

    public DeskThread PatchThread(int id, PatchThreadRequest request)
    {
        var thread = db.Threads.FirstOrDefault(t => t.Id == id) ?? throw new KeyNotFoundException("الخيط غير موجود.");
        if (string.IsNullOrWhiteSpace(request.State) || !States.Contains(request.State))
            throw new InvalidOperationException("حالة الخيط غير معروفة.");
        thread.State = request.State;
        if (request.State == "closed")
            thread.RecommendedStep = "مغلق";
        db.SaveChanges();
        return thread;
    }

    public BriefingView Briefing()
    {
        var threads = db.Threads.AsNoTracking().Where(t => t.State != "closed").ToList();
        return new BriefingView
        {
            NeedsYou = Map(threads.Where(t => t.State == "awaiting_me")),
            Waiting = Map(threads.Where(t => t.State == "awaiting_them")),
            Info = Map(threads.Where(t => t.State == "info")),
            GeneratedAt = DateTimeOffset.UtcNow
        };
    }

    public DeskAlert? GenerateBriefing(bool quietIfEmpty)
    {
        var local = MailRules.Now();
        var settings = EnsureSettings();
        var days = settings.WorkDays.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (!days.Contains(local.DayOfWeek.ToString(), StringComparer.OrdinalIgnoreCase))
            return null;

        var view = Briefing();
        var last = db.Briefings.AsEnumerable().OrderByDescending(b => b.CreatedAt).FirstOrDefault();
        var firstToday = last is null || TimeZoneInfo.ConvertTime(last.CreatedAt, MailRules.Riyadh).Date != local.Date;
        var changed = last is not null && db.Alerts.Any(a => a.CreatedAt > last.CreatedAt);
        if (quietIfEmpty && !firstToday && !changed)
            return null;

        var slot = local.Hour < 12 ? "morning" : "afternoon";
        var lines = new List<string>();
        if (view.NeedsYou.Count == 0 && view.Waiting.Count == 0 && view.Info.Count == 0)
            lines.Add("لا جديد.");
        foreach (var item in view.NeedsYou.Take(5))
            lines.Add($"يحتاجك: {item.Subject} — {item.RecommendedStep}");
        foreach (var item in view.Waiting.Take(3))
            lines.Add($"بانتظارهم: {item.Subject}");
        foreach (var item in view.Info.Take(3))
            lines.Add($"للعلم: {item.Subject}");

        var body = string.Join("\n", lines);
        db.Briefings.Add(new BriefingRecord { Slot = slot, Body = body, CreatedAt = DateTimeOffset.UtcNow });
        var alert = new DeskAlert
        {
            Title = slot == "morning" ? "إحاطة الصباح" : "إحاطة العصر",
            Message = body,
            Source = "briefing",
            Kind = "briefing",
            CreatedAt = DateTimeOffset.UtcNow
        };
        db.Alerts.Add(alert);
        db.SaveChanges();
        return alert;
    }

    public List<DeskAlert> PendingAlerts() =>
        db.Alerts.Where(a => a.DismissedAt == null).AsEnumerable().OrderBy(a => a.CreatedAt).ToList();

    public void DismissAlert(int id)
    {
        var alert = db.Alerts.FirstOrDefault(a => a.Id == id) ?? throw new KeyNotFoundException("التنبيه غير موجود.");
        alert.DismissedAt = DateTimeOffset.UtcNow;
        db.DecisionLogs.Add(new DecisionLog { Kind = "dismiss", Detail = $"alert={id}", CreatedAt = DateTimeOffset.UtcNow });
        db.SaveChanges();
    }

    public OutlookJob? OpenAlert(int id)
    {
        var alert = db.Alerts.FirstOrDefault(a => a.Id == id) ?? throw new KeyNotFoundException("التنبيه غير موجود.");
        if (alert.Kind == "briefing")
            return null;
        if (alert.Kind == "teams" || alert.Source == "teams")
            return Queue("open_teams", "{}", null);
        if (string.IsNullOrWhiteSpace(alert.EntryId))
            throw new InvalidOperationException("لا توجد رسالة لفتحها.");
        return Queue("open", JsonSerializer.Serialize(new { entryId = alert.EntryId, storeId = alert.StoreId }, Json), null);
    }

    public OutlookJob OpenMail(OpenMailRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.EntryId))
            throw new InvalidOperationException("معرّف الرسالة مطلوب.");
        return Queue("open", JsonSerializer.Serialize(new { entryId = request.EntryId, storeId = request.StoreId }, Json), null);
    }

    public List<Draft> Drafts() => db.Drafts.AsEnumerable().OrderByDescending(d => d.UpdatedAt).Take(40).ToList();

    public Draft CreateDraft(DraftWrite write)
    {
        var now = DateTimeOffset.UtcNow;
        var draft = new Draft
        {
            ThreadId = write.ThreadId,
            EntryId = write.EntryId,
            StoreId = write.StoreId,
            To = write.To?.Trim() ?? "",
            Subject = write.Subject?.Trim() ?? "",
            Body = write.Body ?? "",
            Status = "editing",
            CreatedAt = now,
            UpdatedAt = now
        };
        db.Drafts.Add(draft);
        db.SaveChanges();
        return draft;
    }

    public Draft UpdateDraft(int id, DraftWrite write)
    {
        var draft = db.Drafts.FirstOrDefault(d => d.Id == id) ?? throw new KeyNotFoundException("المسودة غير موجودة.");
        if (draft.Status is "sent" or "send_requested")
            throw new InvalidOperationException("هذه المسودة أُرسلت أو بانتظار الإرسال.");
        draft.To = write.To?.Trim() ?? "";
        draft.Subject = write.Subject?.Trim() ?? "";
        draft.Body = write.Body ?? "";
        draft.EntryId = write.EntryId ?? draft.EntryId;
        draft.StoreId = write.StoreId ?? draft.StoreId;
        draft.ThreadId = write.ThreadId ?? draft.ThreadId;
        draft.Status = "editing";
        draft.UpdatedAt = DateTimeOffset.UtcNow;
        db.SaveChanges();
        return draft;
    }

    public OutlookJob RequestSend(int id)
    {
        var draft = db.Drafts.FirstOrDefault(d => d.Id == id) ?? throw new KeyNotFoundException("المسودة غير موجودة.");
        if (draft.Status == "sent")
            throw new InvalidOperationException("أُرسلت هذه المسودة.");
        if (draft.Status == "send_requested")
            throw new InvalidOperationException("الإرسال قيد التنفيذ.");
        if (string.IsNullOrWhiteSpace(draft.Body))
            throw new InvalidOperationException("المسودة فارغة.");
        if (string.IsNullOrWhiteSpace(draft.To) && string.IsNullOrWhiteSpace(draft.EntryId))
            throw new InvalidOperationException("حدد المستلم أو رسالة للرد عليها.");
        draft.Status = "send_requested";
        draft.UpdatedAt = DateTimeOffset.UtcNow;
        var job = Queue("send", JsonSerializer.Serialize(new
        {
            draftId = draft.Id,
            to = draft.To,
            subject = draft.Subject,
            body = draft.Body,
            entryId = draft.EntryId,
            storeId = draft.StoreId
        }, Json), draft.Id);
        db.DecisionLogs.Add(new DecisionLog { Kind = "send_requested", Detail = $"draft={draft.Id}", CreatedAt = DateTimeOffset.UtcNow });
        db.SaveChanges();
        return job;
    }

    public OutlookJob RequestSave(int id)
    {
        var draft = db.Drafts.FirstOrDefault(d => d.Id == id) ?? throw new KeyNotFoundException("المسودة غير موجودة.");
        draft.Status = "save_requested";
        draft.UpdatedAt = DateTimeOffset.UtcNow;
        return Queue("save_draft", JsonSerializer.Serialize(new
        {
            draftId = draft.Id,
            to = draft.To,
            subject = draft.Subject,
            body = draft.Body,
            entryId = draft.EntryId,
            storeId = draft.StoreId
        }, Json), draft.Id);
    }

    public List<OutlookJob> TakeJobs()
    {
        var jobs = db.OutlookJobs.Where(j => j.Status == "pending").OrderBy(j => j.Id).Take(5).ToList();
        foreach (var job in jobs)
            job.Status = "running";
        if (jobs.Count > 0)
            db.SaveChanges();
        return jobs;
    }

    public void CompleteJob(int id, CompleteJobRequest request)
    {
        var job = db.OutlookJobs.FirstOrDefault(j => j.Id == id) ?? throw new KeyNotFoundException("المهمة غير موجودة.");
        job.Status = request.Success ? "done" : "failed";
        job.Result = request.Error;
        job.CompletedAt = DateTimeOffset.UtcNow;
        if (job.DraftId is int draftId)
        {
            var draft = db.Drafts.FirstOrDefault(d => d.Id == draftId);
            if (draft is not null)
            {
                if (request.Success && job.Kind == "send")
                    draft.Status = "sent";
                else if (request.Success && job.Kind == "save_draft")
                    draft.Status = "saved";
                else if (!request.Success && draft.Status is "send_requested" or "save_requested")
                    draft.Status = "editing";
                draft.UpdatedAt = DateTimeOffset.UtcNow;
            }
        }
        if (job.Kind == "send")
        {
            db.DecisionLogs.Add(new DecisionLog
            {
                Kind = request.Success ? "sent" : "send_failed",
                Detail = request.Error ?? $"draft={job.DraftId}",
                CreatedAt = DateTimeOffset.UtcNow
            });
        }
        db.SaveChanges();
    }

    public List<PendingAction> Pending() =>
        db.PendingActions.AsEnumerable().OrderBy(p => p.Done).ThenByDescending(p => p.CreatedAt).Take(50).ToList();

    public PendingAction AddPending(PendingWrite write)
    {
        if (string.IsNullOrWhiteSpace(write.Title))
            throw new InvalidOperationException("عنوان المهمة مطلوب.");
        var row = new PendingAction
        {
            Title = write.Title.Trim(),
            Notes = write.Notes?.Trim() ?? "",
            Done = write.Done ?? false,
            CreatedAt = DateTimeOffset.UtcNow
        };
        db.PendingActions.Add(row);
        db.SaveChanges();
        return row;
    }

    public PendingAction PatchPending(int id, PendingWrite write)
    {
        var row = db.PendingActions.FirstOrDefault(p => p.Id == id) ?? throw new KeyNotFoundException("المهمة غير موجودة.");
        if (!string.IsNullOrWhiteSpace(write.Title))
            row.Title = write.Title.Trim();
        if (write.Notes is not null)
            row.Notes = write.Notes.Trim();
        if (write.Done is bool done)
            row.Done = done;
        db.SaveChanges();
        return row;
    }

    public HostStatus ReportHost(HostReport report)
    {
        var row = db.HostStatuses.FirstOrDefault();
        if (row is null)
        {
            row = new HostStatus();
            db.HostStatuses.Add(row);
        }
        row.Ok = report.Ok;
        row.Message = report.Message ?? "";
        row.UpdatedAt = DateTimeOffset.UtcNow;
        db.SaveChanges();
        return row;
    }

    public HostStatus Host() =>
        db.HostStatuses.FirstOrDefault() ?? new HostStatus { Ok = false, Message = "بانتظار Outlook", UpdatedAt = DateTimeOffset.UtcNow };

    void UpdateThreads(SyncRequest request)
    {
        var ids = request.Items
            .Where(i => !string.IsNullOrWhiteSpace(i.EntryId))
            .Select(MailRules.ConversationId)
            .Distinct()
            .ToList();

        foreach (var conversationId in ids)
        {
            var messages = db.MailMessages
                .Where(m => m.ConversationId == conversationId && m.Classification != "marketing")
                .ToList();
            if (messages.Count == 0)
                continue;

            var newest = messages.OrderBy(m => m.ReceivedAt).Last();
            var thread = db.Threads.FirstOrDefault(t => t.ConversationId == conversationId);
            if (thread?.State == "closed" && newest.ReceivedAt <= thread.LastActivityAt)
                continue;

            thread ??= new DeskThread { ConversationId = conversationId };
            if (thread.Id == 0)
                db.Threads.Add(thread);

            var real = messages.Where(m => m.Folder != "drafts").OrderBy(m => m.ReceivedAt).LastOrDefault() ?? newest;
            var inbox = messages.Where(m => m.Folder == "inbox").OrderBy(m => m.ReceivedAt).LastOrDefault();
            thread.Subject = string.IsNullOrWhiteSpace(real.Subject) ? "(بدون عنوان)" : real.Subject;
            thread.EntryId = real.EntryId;
            thread.StoreId = real.StoreId;
            thread.LastActivityAt = newest.ReceivedAt;
            thread.CounterpartName = inbox?.Sender ?? thread.CounterpartName;
            thread.CounterpartAddress = inbox?.SenderAddress ?? thread.CounterpartAddress;
            thread.State = real.Folder == "sent"
                ? "awaiting_them"
                : real.Classification is "hr" or "teams" ? "info" : "awaiting_me";
            thread.RecommendedStep = thread.State switch
            {
                "awaiting_them" => "بانتظارهم — لا رد الآن",
                "info" when real.Classification == "teams" => "افتح Teams واقرأ هناك",
                "info" => "للقراءة فقط",
                _ when real.Classification == "urgent" => "رد اليوم — جهّز المسودة",
                _ => "راجع الرسالة وجهّز الرد"
            };
        }
    }

    OutlookJob Queue(string kind, string payload, int? draftId)
    {
        var job = new OutlookJob
        {
            Kind = kind,
            PayloadJson = payload,
            Status = "pending",
            DraftId = draftId,
            CreatedAt = DateTimeOffset.UtcNow
        };
        db.OutlookJobs.Add(job);
        db.SaveChanges();
        return job;
    }

    static bool AlreadyAnswered(SyncRequest request, MailMessage row) =>
        request.Items.Any(item =>
            item.Folder == "sent"
            && MailRules.ConversationId(item) == row.ConversationId
            && item.ReceivedAt >= row.ReceivedAt);

    static DeskAlert MakeAlert(MailMessage row)
    {
        var (title, message, kind, source) = row.Classification switch
        {
            "urgent" => ("عاجل", $"{row.Sender}: {row.Subject}", "urgent", "mail"),
            "teams" => ("Teams", $"إشعار من {row.Sender}. افتح Teams واقرأ هناك.", "teams", "teams"),
            "hr" => ("تعميم", string.IsNullOrWhiteSpace(row.Subject) ? "تعميم للقراءة" : row.Subject, "read", "mail"),
            _ => ("يحتاج إجراء", $"{row.Sender}: {row.Subject}", "action", "mail")
        };
        return new DeskAlert
        {
            Title = title,
            Message = message,
            Kind = kind,
            Source = source,
            EntryId = row.EntryId,
            StoreId = row.StoreId,
            CreatedAt = DateTimeOffset.UtcNow
        };
    }

    static List<BriefingItem> Map(IEnumerable<DeskThread> threads) =>
        threads.OrderByDescending(t => t.LastActivityAt).Take(12).Select(t => new BriefingItem
        {
            ThreadId = t.Id,
            Subject = t.Subject,
            State = t.State,
            RecommendedStep = t.RecommendedStep,
            EntryId = t.EntryId,
            StoreId = t.StoreId,
            CounterpartName = t.CounterpartName,
            CounterpartAddress = t.CounterpartAddress,
            LastActivityAt = t.LastActivityAt
        }).ToList();
}

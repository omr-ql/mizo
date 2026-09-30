using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OfficeDesk.OutlookHost;

public sealed class OutlookClient : IDisposable
{
    object? _app;
    object? _namespace;
    bool _ready;
    bool _background;
    string _account = "";

    public string StatusText => _background ? "Outlook يعمل في الخلفية" : "متصل بـ Outlook";

    public void Connect(string accountSmtp)
    {
        _account = accountSmtp.Trim();
        if (_ready && IsAlive())
            return;

        Drop();
        _background = EnsureOutlookProcess();
        Exception? last = null;
        for (var attempt = 0; attempt < 12; attempt++)
        {
            try
            {
                _app = Com.Create("Outlook.Application");
                _namespace = Com.Call(_app, "GetNamespace", "MAPI");
                _ready = true;
                return;
            }
            catch (Exception ex)
            {
                last = ex;
                Drop();
                Thread.Sleep(1000);
            }
        }

        throw new InvalidOperationException("تعذر الاتصال بـ Outlook.", last);
    }

    public void Drop()
    {
        try
        {
            if (_namespace is not null)
                Marshal.ReleaseComObject(_namespace);
        }
        catch
        {
            // The Outlook process may already be gone.
        }

        try
        {
            if (_app is not null)
                Marshal.ReleaseComObject(_app);
        }
        catch
        {
            // The Outlook process may already be gone.
        }

        _namespace = null;
        _app = null;
        _ready = false;
    }

    public string InboxStamp()
    {
        if (_namespace is null)
            throw new InvalidOperationException("Outlook غير متصل.");

        var account = FindAccount(_namespace, _account);
        var store = Com.Get(account, "DeliveryStore");
        object folder;
        try
        {
            folder = Com.Call(store, "GetDefaultFolder", 6);
        }
        catch
        {
            return "";
        }

        object items;
        try
        {
            items = Com.Call(folder, "Items");
            try
            {
                Com.Call(items, "Sort", "[ReceivedTime]", true);
            }
            catch
            {
                // Keep the folder's own order when ReceivedTime is missing.
            }
        }
        catch
        {
            return "";
        }

        var count = Convert.ToInt32(Com.Get(items, "Count"));
        var take = Math.Min(count, 8);
        var parts = new string[take];
        for (var i = 1; i <= take; i++)
        {
            try
            {
                var item = Com.Index(items, i);
                parts[i - 1] = TryStr(item, "EntryID") + ":" + TryBool(item, "UnRead");
            }
            catch
            {
                parts[i - 1] = "";
            }
        }

        return string.Join("|", parts);
    }

    public List<MailItemDto> ReadRecent()
    {
        if (_namespace is null)
            throw new InvalidOperationException("Outlook غير متصل.");

        var account = FindAccount(_namespace, _account);
        var store = Com.Get(account, "DeliveryStore");
        var storeId = Com.Str(store, "StoreID");
        var items = new List<MailItemDto>();
        ReadFolder(store, storeId, 6, "inbox", items);
        ReadFolder(store, storeId, 5, "sent", items);
        ReadFolder(store, storeId, 16, "drafts", items);
        return items;
    }

    public void Open(string entryId, string? storeId)
    {
        var item = GetItem(entryId, storeId);
        Com.Call(item, "Display");
    }

    public void OpenTeams()
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = "msteams:",
            UseShellExecute = true
        });
    }

    public void Send(string? entryId, string? storeId, string to, string subject, string body)
    {
        if (!string.IsNullOrWhiteSpace(entryId))
        {
            var source = GetItem(entryId, storeId);
            var reply = Com.Call(source, "Reply");
            Com.Set(reply, "Body", body);
            Com.Call(reply, "Send");
            return;
        }

        var mail = Com.Call(App, "CreateItem", 0);
        Com.Set(mail, "To", to);
        Com.Set(mail, "Subject", subject);
        Com.Set(mail, "Body", body);
        Com.Call(mail, "Send");
    }

    public void SaveDraft(string to, string subject, string body)
    {
        var mail = Com.Call(App, "CreateItem", 0);
        Com.Set(mail, "To", to);
        Com.Set(mail, "Subject", subject);
        Com.Set(mail, "Body", body);
        Com.Call(mail, "Save");
    }

    public void Dispose() => Drop();

    bool IsAlive()
    {
        if (_app is null || Process.GetProcessesByName("OUTLOOK").Length == 0)
            return false;
        try
        {
            _ = Com.Get(_app, "Version");
            return true;
        }
        catch
        {
            return false;
        }
    }

    static bool EnsureOutlookProcess()
    {
        if (Process.GetProcessesByName("OUTLOOK").Length > 0)
            return false;

        var exe = FindOutlookExe() ?? throw new InvalidOperationException("Outlook غير مثبت على هذا الجهاز.");
        Process.Start(new ProcessStartInfo
        {
            FileName = exe,
            Arguments = "/embedding",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden
        });

        for (var attempt = 0; attempt < 20 && Process.GetProcessesByName("OUTLOOK").Length == 0; attempt++)
            Thread.Sleep(500);

        return true;
    }

    static string? FindOutlookExe()
    {
        foreach (var hive in new[] { Registry.LocalMachine, Registry.CurrentUser })
        {
            using var key = hive.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE");
            if (key?.GetValue(null) is string registered && File.Exists(registered))
                return registered;
        }

        string[] candidates =
        [
            @"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
        ];
        return candidates.FirstOrDefault(File.Exists);
    }

    object App => _app ?? throw new InvalidOperationException("Outlook غير متصل.");

    object GetItem(string entryId, string? storeId)
    {
        if (_namespace is null)
            throw new InvalidOperationException("Outlook غير متصل.");
        return string.IsNullOrWhiteSpace(storeId)
            ? Com.Call(_namespace, "GetItemFromID", entryId)
            : Com.Call(_namespace, "GetItemFromID", entryId, storeId);
    }

    static object FindAccount(object ns, string smtp)
    {
        var accounts = Com.Get(ns, "Accounts");
        var count = Convert.ToInt32(Com.Get(accounts, "Count"));
        for (var i = 1; i <= count; i++)
        {
            var account = Com.Index(accounts, i);
            var address = Com.Str(account, "SmtpAddress");
            if (address.Equals(smtp, StringComparison.OrdinalIgnoreCase))
                return account;
        }
        throw new InvalidOperationException($"حساب Outlook غير موجود: {smtp}");
    }

    static void ReadFolder(object store, string storeId, int folderId, string folderName, List<MailItemDto> sink)
    {
        object folder;
        try
        {
            folder = Com.Call(store, "GetDefaultFolder", folderId);
        }
        catch
        {
            return;
        }

        object items;
        try
        {
            items = Com.Call(folder, "Items");
            try
            {
                Com.Call(items, "Sort", "[ReceivedTime]", true);
            }
            catch
            {
                // Some folders have no ReceivedTime. Keep the default order.
            }
        }
        catch
        {
            return;
        }

        var count = Convert.ToInt32(Com.Get(items, "Count"));
        var take = Math.Min(count, 30);
        for (var i = 1; i <= take; i++)
        {
            try
            {
                var item = Com.Index(items, i);
                var messageClass = TryStr(item, "MessageClass");
                if (!messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)
                    && !messageClass.Contains("Teams", StringComparison.OrdinalIgnoreCase))
                    continue;

                var preview = TryStr(item, "Body").Replace('\r', ' ').Replace('\n', ' ').Trim();
                if (preview.Length > 240)
                    preview = preview[..240];

                sink.Add(new MailItemDto
                {
                    EntryId = TryStr(item, "EntryID"),
                    StoreId = storeId,
                    ConversationId = TryStr(item, "ConversationID"),
                    Folder = folderName,
                    Subject = TryStr(item, "Subject"),
                    Sender = FirstNonEmpty(TryStr(item, "SenderName"), TryStr(item, "SenderEmailAddress")),
                    SenderAddress = Smtp(item),
                    ReceivedAt = When(item),
                    Preview = preview,
                    Unread = TryBool(item, "UnRead"),
                    IsTeams = messageClass.Contains("Teams", StringComparison.OrdinalIgnoreCase),
                    MessageClass = messageClass
                });
            }
            catch
            {
                // Skip items Outlook cannot expose.
            }
        }
    }

    static string Smtp(object item)
    {
        try
        {
            var accessor = Com.Call(item, "PropertyAccessor");
            var value = Com.Call(accessor, "GetProperty", "http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
            var text = value?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(text))
                return text;
        }
        catch
        {
            // Exchange items sometimes hide the SMTP address.
        }
        return TryStr(item, "SenderEmailAddress");
    }

    static DateTimeOffset When(object item)
    {
        foreach (var name in new[] { "ReceivedTime", "SentOn", "CreationTime" })
        {
            try
            {
                var value = Com.Get(item, name);
                if (value is DateTime dt && dt.Year > 1990)
                    return new DateTimeOffset(DateTime.SpecifyKind(dt, DateTimeKind.Local));
            }
            catch
            {
                // Try the next timestamp property.
            }
        }
        return DateTimeOffset.Now;
    }

    static string TryStr(object target, string name)
    {
        try
        {
            return Com.Str(target, name);
        }
        catch
        {
            return "";
        }
    }

    static bool TryBool(object target, string name)
    {
        try
        {
            return Convert.ToBoolean(Com.Get(target, name));
        }
        catch
        {
            return false;
        }
    }

    static string FirstNonEmpty(params string[] values) =>
        values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
}

public sealed class MailItemDto
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
    public bool Unread { get; set; }
    public bool IsTeams { get; set; }
    public string MessageClass { get; set; } = "";
}

static class Com
{
    const BindingFlags Method = BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance;
    const BindingFlags GetProp = BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance;
    const BindingFlags SetProp = BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance;

    public static object Create(string progId)
    {
        var type = Type.GetTypeFromProgID(progId, throwOnError: false)
            ?? throw new InvalidOperationException("Outlook غير مثبت على هذا الجهاز.");
        return Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("تعذر تشغيل Outlook.");
    }

    public static object Call(object target, string name, params object?[] args) =>
        target.GetType().InvokeMember(name, Method, null, target, args)!;

    public static object Get(object target, string name) =>
        target.GetType().InvokeMember(name, GetProp, null, target, null)!;

    public static void Set(object target, string name, object? value) =>
        target.GetType().InvokeMember(name, SetProp, null, target, [value]);

    public static string Str(object target, string name) => Get(target, name)?.ToString() ?? "";

    public static object Index(object target, int index)
    {
        try
        {
            return target.GetType().InvokeMember("Item", GetProp, null, target, [index])!;
        }
        catch (MissingMemberException)
        {
            return target.GetType().InvokeMember("Item", Method, null, target, [index])!;
        }
        catch (TargetInvocationException ex) when (ex.InnerException is MissingMemberException or COMException)
        {
            return target.GetType().InvokeMember("Item", Method, null, target, [index])!;
        }
    }
}

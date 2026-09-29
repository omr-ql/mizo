using System.Net.Http.Json;
using System.Text.Json;

namespace OfficeDesk.OutlookHost;

public static class Program
{
    static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };

    [STAThread]
    public static void Main()
    {
        var api = Environment.GetEnvironmentVariable("OFFICEDESK_API") ?? "http://127.0.0.1:5080";
        using var http = new HttpClient { BaseAddress = new Uri(api.TrimEnd('/') + "/"), Timeout = TimeSpan.FromSeconds(60) };
        using var outlook = new OutlookClient();
        var nextSync = DateTimeOffset.MinValue;

        while (true)
        {
            try
            {
                var settings = GetSettings(http);
                outlook.Connect(settings.AccountSmtp);

                Drain(http, outlook);

                if (DateTimeOffset.UtcNow >= nextSync)
                {
                    var items = outlook.ReadRecent();
                    Post(http, "/api/mail/sync", new { items });
                    Post(http, "/api/host/status", new { ok = true, message = outlook.StatusText });
                    nextSync = DateTimeOffset.UtcNow.AddMinutes(Math.Clamp(settings.PollIntervalMinutes, 1, 60));
                }
            }
            catch (Exception ex)
            {
                outlook.Drop();
                Console.Error.WriteLine(ex.Message);
                try
                {
                    Post(http, "/api/host/status", new { ok = false, message = ex.Message });
                }
                catch
                {
                    // The API may still be starting.
                }
                Thread.Sleep(TimeSpan.FromSeconds(8));
                continue;
            }

            Thread.Sleep(TimeSpan.FromSeconds(3));
        }
    }

    static SettingsDto GetSettings(HttpClient http)
    {
        using var response = http.Send(new HttpRequestMessage(HttpMethod.Get, "api/settings"));
        response.EnsureSuccessStatusCode();
        using var stream = response.Content.ReadAsStream();
        return JsonSerializer.Deserialize<SettingsDto>(stream, Json)
            ?? throw new InvalidOperationException("الإعدادات فارغة.");
    }

    static void Drain(HttpClient http, OutlookClient outlook)
    {
        using var response = http.Send(new HttpRequestMessage(HttpMethod.Get, "api/outlook/jobs"));
        response.EnsureSuccessStatusCode();
        using var stream = response.Content.ReadAsStream();
        var jobs = JsonSerializer.Deserialize<List<JobDto>>(stream, Json) ?? [];
        foreach (var job in jobs)
        {
            try
            {
                Run(outlook, job);
                Post(http, $"/api/outlook/jobs/{job.Id}/complete", new { success = true });
            }
            catch (Exception ex)
            {
                Post(http, $"/api/outlook/jobs/{job.Id}/complete", new { success = false, error = ex.Message });
            }
        }
    }

    static void Run(OutlookClient outlook, JobDto job)
    {
        var payload = JsonSerializer.Deserialize<JobPayload>(job.PayloadJson, Json) ?? new JobPayload();
        switch (job.Kind)
        {
            case "open":
                outlook.Open(payload.EntryId ?? "", payload.StoreId);
                break;
            case "open_teams":
                outlook.OpenTeams();
                break;
            case "send":
                outlook.Send(payload.EntryId, payload.StoreId, payload.To ?? "", payload.Subject ?? "", payload.Body ?? "");
                break;
            case "save_draft":
                outlook.SaveDraft(payload.To ?? "", payload.Subject ?? "", payload.Body ?? "");
                break;
            default:
                throw new InvalidOperationException($"مهمة غير معروفة: {job.Kind}");
        }
    }

    static void Post(HttpClient http, string path, object body)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, path.TrimStart('/'))
        {
            Content = JsonContent.Create(body, options: Json)
        };
        using var response = http.Send(request);
        response.EnsureSuccessStatusCode();
    }
}

sealed class SettingsDto
{
    public string AccountSmtp { get; set; } = "o.bamusa@ejadah-sa.com";
    public int PollIntervalMinutes { get; set; } = 3;
}

sealed class JobDto
{
    public int Id { get; set; }
    public string Kind { get; set; } = "";
    public string PayloadJson { get; set; } = "{}";
}

sealed class JobPayload
{
    public string? EntryId { get; set; }
    public string? StoreId { get; set; }
    public string? To { get; set; }
    public string? Subject { get; set; }
    public string? Body { get; set; }
}

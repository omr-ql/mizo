using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using OfficeDesk.Api.Data;
using OfficeDesk.Api.Models;
using OfficeDesk.Api.Services;

var builder = WebApplication.CreateBuilder(args);

var connectionString = builder.Configuration.GetConnectionString("Office") ?? "Data Source=office.db";
var dataSource = DataSourceOf(connectionString);
if (!string.IsNullOrEmpty(dataSource))
{
    var directory = Path.GetDirectoryName(dataSource);
    if (!string.IsNullOrEmpty(directory))
        Directory.CreateDirectory(directory);
}

builder.Services.AddDbContext<OfficeDb>(options => options.UseSqlite(connectionString));
builder.Services.AddScoped<OfficeService>();
builder.Services.ConfigureHttpJsonOptions(options =>
{
    options.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
});

var app = builder.Build();

using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<OfficeDb>();
    db.Database.EnsureCreated();
    db.Database.ExecuteSqlRaw("PRAGMA journal_mode=WAL;");
    db.Database.ExecuteSqlRaw("PRAGMA busy_timeout=5000;");
    EnsureColumn(db, "MailMessages", "Body");
    EnsureColumn(db, "MailMessages", "Summary");
    var office = scope.ServiceProvider.GetRequiredService<OfficeService>();
    office.EnsureSettings();
    office.RefreshSummaries();
}

app.Use(async (context, next) =>
{
    try
    {
        await next();
    }
    catch (KeyNotFoundException ex)
    {
        context.Response.StatusCode = StatusCodes.Status404NotFound;
        await context.Response.WriteAsJsonAsync(new { error = ex.Message });
    }
    catch (InvalidOperationException ex)
    {
        context.Response.StatusCode = StatusCodes.Status400BadRequest;
        await context.Response.WriteAsJsonAsync(new { error = ex.Message });
    }
});

app.MapGet("/health", () => Results.Ok(new { ok = true }));

var api = app.MapGroup("/api");

api.MapPost("/mail/sync", (SyncRequest body, OfficeService office) =>
    Results.Ok(new { alerts = office.Sync(body) }));

api.MapGet("/mail", (string? folder, OfficeService office) => Results.Ok(office.Mail(folder)));

api.MapGet("/mail/digest", (OfficeService office) => Results.Ok(office.Digest()));

api.MapPost("/mail/open", (OpenMailRequest body, OfficeService office) => Results.Ok(office.OpenMail(body)));

api.MapGet("/briefing", (OfficeService office) => Results.Ok(office.Briefing()));

api.MapPost("/briefing/generate", (GenerateBriefingRequest body, OfficeService office) =>
    Results.Ok(new { alert = office.GenerateBriefing(body.QuietIfEmpty) }));

api.MapGet("/threads", (OfficeService office) => Results.Ok(office.Threads()));

api.MapPatch("/threads/{id:int}", (int id, PatchThreadRequest body, OfficeService office) =>
    Results.Ok(office.PatchThread(id, body)));

api.MapGet("/pending", (OfficeService office) => Results.Ok(office.Pending()));

api.MapPost("/pending", (PendingWrite body, OfficeService office) => Results.Ok(office.AddPending(body)));

api.MapPatch("/pending/{id:int}", (int id, PendingWrite body, OfficeService office) =>
    Results.Ok(office.PatchPending(id, body)));

api.MapGet("/settings", (OfficeService office) => Results.Ok(office.EnsureSettings()));

api.MapPut("/settings", (SettingsWrite body, OfficeService office) => Results.Ok(office.UpdateSettings(body)));

api.MapGet("/drafts", (OfficeService office) => Results.Ok(office.Drafts()));

api.MapPost("/drafts", (DraftWrite body, OfficeService office) => Results.Ok(office.CreateDraft(body)));

api.MapPut("/drafts/{id:int}", (int id, DraftWrite body, OfficeService office) =>
    Results.Ok(office.UpdateDraft(id, body)));

api.MapPost("/drafts/{id:int}/send", (int id, OfficeService office) => Results.Ok(office.RequestSend(id)));

api.MapPost("/drafts/{id:int}/save", (int id, OfficeService office) => Results.Ok(office.RequestSave(id)));

api.MapGet("/alerts/pending", (OfficeService office) => Results.Ok(office.PendingAlerts()));

api.MapPost("/alerts/{id:int}/dismiss", (int id, OfficeService office) =>
{
    office.DismissAlert(id);
    return Results.Ok(new { ok = true });
});

api.MapPost("/alerts/{id:int}/open", (int id, OfficeService office) =>
    Results.Ok(new { job = office.OpenAlert(id) }));

api.MapGet("/outlook/jobs", (OfficeService office) => Results.Ok(office.TakeJobs()));

api.MapPost("/outlook/jobs/{id:int}/complete", (int id, CompleteJobRequest body, OfficeService office) =>
{
    office.CompleteJob(id, body);
    return Results.Ok(new { ok = true });
});

api.MapGet("/host/status", (OfficeService office) => Results.Ok(office.Host()));

api.MapPost("/host/status", (HostReport body, OfficeService office) => Results.Ok(office.ReportHost(body)));

app.Run();

static void EnsureColumn(OfficeDb db, string table, string column)
{
    var connection = db.Database.GetDbConnection();
    if (connection.State != System.Data.ConnectionState.Open)
        connection.Open();
    using var check = connection.CreateCommand();
    check.CommandText = $"SELECT COUNT(*) FROM pragma_table_info('{table}') WHERE name = $name";
    var name = check.CreateParameter();
    name.ParameterName = "$name";
    name.Value = column;
    check.Parameters.Add(name);
    var exists = Convert.ToInt32(check.ExecuteScalar()) > 0;
    if (exists)
        return;
    using var alter = connection.CreateCommand();
    alter.CommandText = $"ALTER TABLE \"{table}\" ADD COLUMN \"{column}\" TEXT NOT NULL DEFAULT ''";
    alter.ExecuteNonQuery();
}

static string? DataSourceOf(string connectionString)
{
    foreach (var part in connectionString.Split(';', StringSplitOptions.RemoveEmptyEntries))
    {
        var split = part.IndexOf('=');
        if (split <= 0)
            continue;
        if (part[..split].Trim().Equals("Data Source", StringComparison.OrdinalIgnoreCase))
            return part[(split + 1)..].Trim();
    }
    return null;
}

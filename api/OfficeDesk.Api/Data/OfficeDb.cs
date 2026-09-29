using Microsoft.EntityFrameworkCore;

namespace OfficeDesk.Api.Data;

public sealed class OfficeDb(DbContextOptions<OfficeDb> options) : DbContext(options)
{
    public DbSet<Settings> Settings => Set<Settings>();
    public DbSet<MailMessage> MailMessages => Set<MailMessage>();
    public DbSet<DeskThread> Threads => Set<DeskThread>();
    public DbSet<DeskAlert> Alerts => Set<DeskAlert>();
    public DbSet<BriefingRecord> Briefings => Set<BriefingRecord>();
    public DbSet<Draft> Drafts => Set<Draft>();
    public DbSet<OutlookJob> OutlookJobs => Set<OutlookJob>();
    public DbSet<PendingAction> PendingActions => Set<PendingAction>();
    public DbSet<DecisionLog> DecisionLogs => Set<DecisionLog>();
    public DbSet<HostStatus> HostStatuses => Set<HostStatus>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<MailMessage>().HasIndex(m => m.EntryId).IsUnique();
        modelBuilder.Entity<MailMessage>().HasIndex(m => m.ConversationId);
        modelBuilder.Entity<DeskThread>().HasIndex(t => t.ConversationId).IsUnique();
        modelBuilder.Entity<DeskAlert>().HasIndex(a => a.DismissedAt);
        modelBuilder.Entity<OutlookJob>().HasIndex(j => j.Status);
    }
}

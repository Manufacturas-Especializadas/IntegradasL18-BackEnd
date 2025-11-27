using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace Integradas.Models;

public partial class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Integradas> Integradas { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Integradas>(entity =>
        {
            entity.HasKey(e => e.Id).HasName("PK__Integrad__3214EC0720D75283");

            entity.Property(e => e.Amount).HasColumnName("amount");
            entity.Property(e => e.Completed).HasColumnName("completed");
            entity.Property(e => e.CompletedDate)
                .HasColumnType("datetime")
                .HasColumnName("completedDate");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("createdAt");
            entity.Property(e => e.PartNumber)
                .IsUnicode(false)
                .HasColumnName("part_Number");
            entity.Property(e => e.Pipeline)
                .IsUnicode(false)
                .HasColumnName("pipeline");
            entity.Property(e => e.ScannedQuantity).HasColumnName("scannedQuantity");
            entity.Property(e => e.TotalQuantity).HasColumnName("totalQuantity");
            entity.Property(e => e.Type)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("type");
            entity.Property(e => e.WeekNumber).HasColumnName("weekNumber");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
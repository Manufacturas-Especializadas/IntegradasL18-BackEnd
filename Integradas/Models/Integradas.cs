using System;
using System.Collections.Generic;

namespace Integradas.Models;

public partial class Integradas
{
    public int Id { get; set; }

    public string Type { get; set; }

    public string PartNumber { get; set; }

    public string Pipeline { get; set; }

    public int? Amount { get; set; }

    public int? WeekNumber { get; set; }

    public DateTime? CreatedAt { get; set; }

    public bool Completed { get; set; }

    public int? ScannedQuantity { get; set; }

    public int? TotalQuantity { get; set; }

    public DateTime? CompletedDate { get; set; }
}
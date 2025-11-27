namespace Integradas.Dtos
{
    public class OrderSummaryDto
    {
        public int WeekNumber { get; set; }

        public int TotalOrders { get; set; }

        public int CompletedOrders { get; set; }

        public int PendingOrders { get; set; }

        public int TotalQuantity { get; set; }

        public int ScannedQuantity { get; set; }

        public double ProgressPercentage { get; set; }

        public DateTime? LastUpdate { get; set; }
    }
}
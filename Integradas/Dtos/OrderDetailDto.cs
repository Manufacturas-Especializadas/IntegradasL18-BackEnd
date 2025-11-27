namespace Integradas.Dtos
{
    public class OrderDetailDto
    {
        public int Id { get; set; }

        public string Type { get; set; }

        public string PartNumber { get; set; }

        public string Pipeline { get; set; }

        public int? Amount { get; set; }

        public int? ScannedQuantity { get; set; }

        public bool? Completed { get; set; }

        public DateTime? CreatedAt { get; set; }

        public DateTime? CompletedDate { get; set; }

        public int? Remaining => (Amount.HasValue && ScannedQuantity.HasValue) ? Amount - ScannedQuantity : null;

        public double? ProgressPercentage
        {
            get
            {
                if (!Amount.HasValue || Amount.Value <= 0 || !ScannedQuantity.HasValue)
                    return 0;

                return (double)ScannedQuantity.Value / Amount.Value * 100;
            }
        }
    }
}
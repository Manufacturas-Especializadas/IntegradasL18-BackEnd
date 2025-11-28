namespace Integradas.Dtos
{
    public class ScanResponseDto
    {
        public bool Success { get; set; }

        public string Message { get; set; } = string.Empty;

        public string PartNumber { get; set; } = string.Empty;

        public int? ScannedQuantity { get; set; }

        public int RequiredQuantity { get; set; }

        public int AddedQuantity { get; set; }

        public bool Completed { get; set; }

        public int Remaining { get; set; }

        public double ProgressPercentage { get; set; }

        public OrderDetailDto? Order {  get; set; }
    }
}
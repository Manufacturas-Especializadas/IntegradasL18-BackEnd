namespace Integradas.Dtos
{
    public class ScanRequestDto
    {
        public string PartNumber { get; set; } = string.Empty;

        public int WeekNumber { get; set; }

        public int Quantity { get; set; }
    }
}
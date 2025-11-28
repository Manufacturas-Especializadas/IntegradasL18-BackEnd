namespace Integradas.Dtos
{
    public class ReportRequestDto
    {
        public int WeekNumber { get; set; }

        public string ReportType { get; set; } = "detailed";
    }
}
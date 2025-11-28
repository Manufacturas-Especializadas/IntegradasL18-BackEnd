namespace Integradas.Dtos
{
    public class ReportResponseDto
    {
        public bool Success { get; set; }

        public string Message { get; set; } = string.Empty;

        public string FileName { get; set; } = string.Empty;

        public byte[] FileContent { get; set; } = Array.Empty<byte>();

        public string ContentType {  get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    }
}
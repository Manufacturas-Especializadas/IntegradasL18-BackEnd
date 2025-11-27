namespace Integradas.Dtos
{
    public class ExcelUploadRequest
    {
        public IFormFile File { get; set; }

        public int WeekNumber {  get; set; }
    }
}
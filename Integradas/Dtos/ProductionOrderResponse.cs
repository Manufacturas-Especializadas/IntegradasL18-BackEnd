namespace Integradas.Dtos
{
    public class ProductionOrderResponse
    {
        public bool Success { get; set; }

        public string Message { get; set; }

        public int RecordProcessed { get; set; }

        public List<string> Errors { get; set; } = new List<string>();
    }
}
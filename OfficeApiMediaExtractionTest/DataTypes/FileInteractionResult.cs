namespace OfficeApiMediaExtractionTest.DataTypes
{
    public class FileInteractionResult
    {
        public bool IsSuccess { get; set; }
        public string? Value { get; set; }
        public string? Message { get; set; }
        public Exception? InteractionException { get; set; }
    }
}

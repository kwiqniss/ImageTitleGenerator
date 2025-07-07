namespace OfficeApiMediaExtractionProgram
{
    public interface IWorker
    {
        public Task ExecuteProgramAsync(string docPath);
    }
}

namespace OfficeApiMediaExtractionTest.Interfaces
{
    public interface IWorker
    {
        public Task ExecuteProgramAsync(string docPath);
    }
}

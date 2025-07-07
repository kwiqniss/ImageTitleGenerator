namespace ImageAnalyzer.Interfaces
{
    public interface IWorker
    {
        public Task ExecuteProgramAsync(string docPath);
    }
}

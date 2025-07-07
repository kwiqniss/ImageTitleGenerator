namespace ImageAnalyzerProgram
{
    public interface IWorker
    {
        public Task ExecuteProgramAsync(string docPath);
    }
}

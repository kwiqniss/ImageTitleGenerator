using ImageAnalyzer.AI;
using ImageAnalyzer.DataTypes;
using ImageAnalyzer.DocumentInteractions;
using ImageAnalyzerProgram.Loggers;

namespace ImageAnalyzerProgram
{
    public class Worker : IWorker
    {
        private readonly IDocManager _docManager;
        private readonly IImageAnalyzer _imageAnalyzer;
        private readonly List<ILogger> _loggers = new(); 

        public Worker(IDocManager docManager, 
            IImageAnalyzer imageAnalyzer, 
            IEnumerable<ILogger> loggers)
        {
            _docManager = docManager ?? throw new ArgumentNullException(nameof(docManager));
            _imageAnalyzer = imageAnalyzer ?? throw new ArgumentNullException(nameof(imageAnalyzer));
            if (loggers.Any())
            {
                _loggers = loggers.ToList();
            }
        }

        public async Task ExecuteProgramAsync(string docPath)
        {
            _loggers.ForEach(logger => logger.WriteLogMessage($"Starting job with document path: {docPath}"));

            // Copy the doc
            var copyResult = _docManager.CopyDoc(docPath);
            docPath = string.Empty; // to prevent saving over the original
            
            if (!copyResult.IsSuccess)
            {
                _loggers.ForEach(logger => logger.WriteLogError(copyResult.Message, copyResult.InteractionException));
                return;
            }
            if (copyResult == null || string.IsNullOrWhiteSpace(copyResult.Value))
            {
                _loggers.ForEach(logger => logger.WriteLogError("There was no result from the copy operation.", null));
                return;
            }

            var docCopyPath = copyResult.Value;
            _loggers.ForEach(logger => logger.WriteLogMessage($"Document copied to: {docCopyPath}"));

            // Get images from the doc copy
            var images = _docManager.GetImages(docCopyPath)?.ToArray() ?? Enumerable.Empty<DocumentImage>();
            var imageCount = images?.Count() ?? 0;
            _loggers.ForEach(logger => logger.WriteLogMessage($"There were {imageCount} images in the doc."));
            if (imageCount < 1) return;

            // Add descriptions to the images
            var addedDescriptionsSuccessfully = await _imageAnalyzer.AddImageDescriptionsAsync(images!);
            _loggers.ForEach(logger => logger.WriteLogMessage($"Added image descriptions successfully?: {addedDescriptionsSuccessfully}"));
            if (!addedDescriptionsSuccessfully) return;

            // Save the new image titles to the document copy
            var saveTitlesResult = _docManager.SaveImageTitles(images!, docCopyPath);
            if (!saveTitlesResult.IsSuccess)
            {
                _loggers.ForEach(logger => logger.WriteLogError(saveTitlesResult.Message, saveTitlesResult.InteractionException));
                return;
            }

            _loggers.ForEach(logger => logger.WriteLogMessage($"Job successful? {saveTitlesResult.IsSuccess}"));
        }
    }
}

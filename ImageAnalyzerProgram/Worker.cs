using ImageAnalyzer.DataTypes;
using ImageAnalyzer.Interfaces;

namespace ImageAnalyzerProgram
{
    public class Worker : IWorker
    {
        private readonly IDocManager _docManager;
        private readonly IImageAnalyzer _imageAnalyzer;
        private readonly List<ILogger> _loggers = new[]; 

        public Worker(IDocManager docManager, 
            IImageAnalyzer imageAnalyzer, 
            IEnumerable<ILogger> loggers)
        {
            _docManager = docManager ?? throw new ArgumentNullException(nameof(docManager));
            _imageAnalyzer = imageAnalyzer ?? throw new ArgumentNullException(nameof(imageAnalyzer));
            if (_loggers != null && !loggers.Any())
            {
                _loggers = loggers.ToList();
            }
        }

        public async Task ExecuteProgramAsync(string docPath)
        {
            // Copy the doc
            var copyResult = CopyDoc(docPath);
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
            var images = GetImages(docCopyPath);
            var imageCount = images?.Count() ?? 0;
            _loggers.ForEach(logger => logger.WriteLogMessage($"There were {imageCount} images in the doc."));
            if (imageCount < 1) return;

            // Add descriptions to the images
            var addedDescriptionsSuccessfully = await AddImageDescriptionsAsync(images!);
            _loggers.ForEach(logger => logger.WriteLogMessage($"Added image descriptions successfully?: {addedDescriptionsSuccessfully}"));
            if (!addedDescriptionsSuccessfully) return;

            // Save the new image titles to the document copy
            var saveTitlesResult = SaveNewImageTitles(images!, copyResult.Value);
            if (!saveTitlesResult.IsSuccess)
            {
                _loggers.ForEach(logger => logger.WriteLogError(saveTitlesResult.Message, saveTitlesResult.InteractionException));
                return;
            }

            _loggers.ForEach(logger => logger.WriteLogMessage($"Job successful? {saveTitlesResult.IsSuccess}"));
        }

        private FileInteractionResult CopyDoc(string originalDocPath)
        {
            return _docManager.CopyDoc(originalDocPath);
        }

        private IEnumerable<DocumentImage> GetImages(string docCopyPath)
        {
            var images = _docManager.GetImages(docCopyPath)?.ToArray();
            if (images == null || images.Length < 1)
            {
                return Enumerable.Empty<DocumentImage>();
            }

            return images;
        }

        private async Task<bool> AddImageDescriptionsAsync(IEnumerable<DocumentImage> images)
        {
            return (images != null
                && images.Count() > 0
                && await _imageAnalyzer.AddImageDescriptionsAsync(images));
        }

        private FileInteractionResult SaveNewImageTitles(IEnumerable<DocumentImage> images, string docPath)
        {
            return _docManager.SaveTitles(
                images,
                docPath);
        }
    }
}

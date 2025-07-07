namespace ImageAnalyzer.DataTypes
{
    public class DocumentImage
    {
        public string RelId { get; private set; }

        public string ImageExtension { get; private set; }

        public Stream? ImageStream { get; private set; }

        public string Title { get; set; }

        public int SheetSlideIndex { get; private set; }

        internal DocumentImage(
            
            string extension
            , Stream? imageStream
            , string relId
            , string title = "DefaultTitle"
            , int sheetSlideIndex = -1)
        {
            ImageExtension = extension;
            ImageStream = imageStream;
            RelId = relId;
            Title = $"{title} - relId#: {relId}";
            SheetSlideIndex = sheetSlideIndex;
        }
    }
}

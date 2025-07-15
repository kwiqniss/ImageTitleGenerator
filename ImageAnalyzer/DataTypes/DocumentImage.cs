namespace ImageAnalyzer.DataTypes
{
    public class DocumentImage
    {
        public string RelId { get; private set; }

        public Stream? ImageStream { get; private set; }

        public string Title { get; set; }

        internal DocumentImage(
            Stream? imageStream
            , string relId
            , string title = "DefaultTitle")
        {
            ImageStream = imageStream;
            RelId = relId;
            Title = $"{title} - relId#: {relId}";
        }
    }
}

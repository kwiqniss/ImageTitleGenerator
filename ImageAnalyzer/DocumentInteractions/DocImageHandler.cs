using System.IO.Packaging;
using System.Xml;
using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.DocumentInteractions
{
    public class DocImageHandler : IImageHandler
    {
        public IEnumerable<DocumentImage> GetImages(string docPath)
        {
            using var package = Package.Open(docPath, FileMode.Open, FileAccess.Read);
            int index = 0;
            foreach (var part in GetImageParts(package))
            {
                var ms = new MemoryStream();
                part.GetStream().CopyTo(ms);
                ms.Position = 0;
                yield return new DocumentImage(
                    ms,
                    part.Uri.ToString(),
                    $"Image{++index}");
            }
        }

        public FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath)
        {
            try
            {
                using var package = Package.Open(docPath, FileMode.Open, FileAccess.ReadWrite);
                var imagesByUri = images
                    .Where(img => !string.IsNullOrEmpty(img.RelId))
                    .ToDictionary(img => img.RelId, img => img);

                foreach (var part in package.GetParts())
                {
                    if (!part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var xmlDoc = new XmlDocument();
                    using (var stream = part.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        if (stream.Length == 0) continue;
                        xmlDoc.Load(stream);
                    }

                    bool modified = false;
                    var nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                    nsmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    nsmgr.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                    nsmgr.AddNamespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
                    nsmgr.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                    nsmgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

                    var xmlNodes = xmlDoc.SelectNodes("//wp:docPr | //p:cNvPr | //xdr:cNvPr", nsmgr);
                    if (xmlNodes != null)
                    {
                        foreach (XmlElement el in xmlNodes)
                        {
                            var blip = 
                                el.ParentNode?.SelectSingleNode(".//a:blip", nsmgr) as XmlElement 
                                ?? el.ParentNode?.ParentNode?.SelectSingleNode(".//a:blip | .//xdr:blip", nsmgr) as XmlElement;
                            if (blip == null) continue;

                            var relId = blip.GetAttribute("r:embed") ?? blip.GetAttribute("r:link");
                            if (string.IsNullOrEmpty(relId)) continue;
                            
                            var rel = part.GetRelationship(relId);
                            if (rel == null) continue;
                            
                            var imageUri = PackUriHelper.ResolvePartUri(part.Uri, rel.TargetUri).ToString();
                            if (imagesByUri.TryGetValue(imageUri, out var docImage))
                            {
                                el.SetAttribute("title", docImage.Title ?? "Image");
                                el.SetAttribute("descr", string.Empty);
                                modified = true;
                            }
                        }
                    }

                    if (modified)
                    {
                        using var outStream = part.GetStream(FileMode.Create, FileAccess.Write);
                        xmlDoc.Save(outStream);
                    }
                }

                package.Flush();
                return new FileInteractionResult { IsSuccess = true, Value = docPath };
            }
            catch (Exception ex)
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Value = docPath,
                    Message = "Failed to save image titles.",
                    InteractionException = ex
                };
            }
        }

        private static IEnumerable<PackagePart> GetImageParts(Package package)
        {
            foreach (var part in package.GetParts())
            {
                if (part.ContentType.StartsWith("image/"))
                    yield return part;
            }
        }
    }
}
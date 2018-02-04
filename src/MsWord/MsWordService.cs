using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace Kodiak.Facade.OpenXml.MsWord
{
    public class MsWordService : IMsWordService
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MsWordService"/> class.
        /// </summary>
        public MsWordService()
        {
            this.MemoryStream = new MemoryStream();
        }

        /// <summary>
        /// Sets the page template.
        /// </summary>
        /// <param name="pageTemplatePath">The page template path.</param>
        public void SetPageTemplate(string pageTemplatePath)
        {
            byte[] byteArray = System.IO.File.ReadAllBytes(pageTemplatePath);
            this.MemoryStream = new MemoryStream();
            this.MemoryStream.Write(byteArray, 0, byteArray.Length);
        }

        /// <summary>
        /// Gets or sets the memory stream.
        /// </summary>
        /// <value>
        /// The memory stream.
        /// </value>
        private MemoryStream MemoryStream { get; set; }

        /// <summary>
        /// Merges the text to field in main document part.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="substituteText">The substitute text.</param>
        public void MergeTextToFieldInMainDocumentPart(string fieldName, string substituteText)
        {
            using (WordprocessingDocument template = WordprocessingDocument.Open(this.MemoryStream, true))
            {
                template.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                ReplaceTextInMainDocumentPart(template.MainDocumentPart, $"«{fieldName}»", substituteText);
            }
        }

        /// <summary>
        /// Replaces the text in main document part.
        /// </summary>
        /// <param name="docPart">The document part.</param>
        /// <param name="match">The match.</param>
        /// <param name="value">The value.</param>
        private static void ReplaceTextInMainDocumentPart(MainDocumentPart docPart, string match, string value)
        {
            var body = docPart.Document.Body;

            foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
            {
                if (text.Text.Contains(match))
                {
                    text.Text = text.Text.Replace(match, value);
                }
            }
        }

        public MemoryStream GetMemoryStream()
        {
            return this.MemoryStream;
        }
    }
}

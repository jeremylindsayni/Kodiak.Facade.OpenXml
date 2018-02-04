using System.IO;

namespace Kodiak.Facade.OpenXml.MsWord
{
    interface IMsWordService
    {
        /// <summary>
        /// Sets the page template.
        /// </summary>
        /// <param name="pageTemplatePath">The page template path.</param>
        void SetPageTemplate(string pageTemplatePath);

        /// <summary>
        /// Merges the text to field in main document part.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="substituteText">The substitute text.</param>
        void MergeTextToFieldInMainDocumentPart(string fieldName, string substituteText);

        /// <summary>
        /// Gets the memory stream.
        /// </summary>
        /// <returns></returns>
        MemoryStream GetMemoryStream();
    }
}

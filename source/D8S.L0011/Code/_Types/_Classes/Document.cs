using System;

using Wd = Microsoft.Office.Interop.Word;

using R5T.T0142;

using D8S.L0011.Extensions;


namespace D8S.L0011
{
    /// <summary>
    /// Represents a Word document.
    /// </summary>
    /// <remarks>
    /// Not disposable since "disposing" a document would mean losing work unless the document was saved.
    /// Thus documents are saved then closed.
    /// </remarks>
    [UtilityTypeMarker]
    public class Document
    {
        internal Wd.Document WdDocument { get; private set; }

        public Application Application { get; private set; }


        internal Document(Wd.Document wdDocument, Application application)
        {
            this.WdDocument = wdDocument;
            this.Application = application;
        }

        /// <summary>
        /// Closes the Excel workbook without saving changes.
        /// </summary>
        public void Close()
        {
            this.WdDocument.Close(false);
        }

        public void SaveAs(string filePath, WordFileFormat fileFormat, bool overwrite = true)
        {
            // Workaround for Document.SaveAs() not having an easy overwrite option.
            if (overwrite && Instances.FileSystemOperator.Exists_File(filePath))
            {
                Instances.FileSystemOperator.Delete_File_OkIfNotExists(filePath);
            }

            var wdSaveFormat = fileFormat.To_WdSaveFormat();

            this.WdDocument.SaveAs(filePath, wdSaveFormat);
        }

        public void SaveAs(string filePath, bool overwrite = true)
        {
            this.SaveAs(filePath, WordFileFormat.DOCX, overwrite);
        }

        public void Write(string text)
            => this.Write_AtSelection(text);

        public void Write_AtSelection(string text)
        {
            var currentSelection = this.Application.WdApplication.Selection;

            currentSelection.TypeText(text);
        }
    }
}

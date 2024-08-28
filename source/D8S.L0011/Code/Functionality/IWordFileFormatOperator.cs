using System;

using R5T.T0132;

using Wd = Microsoft.Office.Interop.Word;


namespace D8S.L0011
{
    [FunctionalityMarker]
    public partial interface IWordFileFormatOperator : IFunctionalityMarker
    {
        internal Wd.WdSaveFormat To_WdSaveFormat(WordFileFormat wordFileFormat)
        {
            var output = wordFileFormat switch
            {
                WordFileFormat.DOC => Wd.WdSaveFormat.wdFormatDocument,
                WordFileFormat.DOCM => Wd.WdSaveFormat.wdFormatXMLDocumentMacroEnabled,
                WordFileFormat.DOCX => Wd.WdSaveFormat.wdFormatXMLDocument,
                _ => Wd.WdSaveFormat.wdFormatDocumentDefault,
            };

            return output;
        }
    }
}

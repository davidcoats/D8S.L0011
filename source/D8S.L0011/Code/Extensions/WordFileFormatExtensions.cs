using System;

using Wd = Microsoft.Office.Interop.Word;


namespace D8S.L0011.Extensions
{
    public static class WordFileFormatExtensions
    {
        internal static Wd.WdSaveFormat To_WdSaveFormat(this WordFileFormat wordFileFormat)
            => Instances.WordFileFormatOperator.To_WdSaveFormat(wordFileFormat);
    }
}

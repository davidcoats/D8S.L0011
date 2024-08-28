using System;


namespace D8S.L0011
{
    public class WordFileFormatOperator : IWordFileFormatOperator
    {
        #region Infrastructure

        public static IWordFileFormatOperator Instance { get; } = new WordFileFormatOperator();


        private WordFileFormatOperator()
        {
        }

        #endregion
    }
}

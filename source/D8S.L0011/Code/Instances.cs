using System;


namespace D8S.L0011
{
    public static class Instances
    {
        public static R5T.L0066.IFileSystemOperator FileSystemOperator => R5T.L0066.FileSystemOperator.Instance;
        public static IValues Values => L0011.Values.Instance;
        public static IWordFileFormatOperator WordFileFormatOperator => L0011.WordFileFormatOperator.Instance;
    }
}
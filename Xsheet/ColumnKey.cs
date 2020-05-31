namespace Xsheet
{
    public struct ColumnKey
    {
        public readonly string Key;
        public readonly int Index;
        public readonly MatrixKey MatrixKey;

        public ColumnKey(MatrixKey matKey, int index, string colKey)
        {
            MatrixKey = matKey;
            Index = index;
            Key = colKey;
        }
    }
}

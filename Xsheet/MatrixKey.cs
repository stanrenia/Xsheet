namespace Xsheet
{
    public struct MatrixKey
    {
        public readonly string Key;
        public readonly int Index;

        public MatrixKey(string key, int index)
        {
            Key = key;
            Index = index;
        }
    }
}

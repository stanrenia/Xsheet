namespace Xsheet
{
    public class ColumnDefinition
    {
        public const int UNDEFINED_INDEX_VALUE = -1;

        public ColumnKey Key { get; internal set; }

        private string _label;
        public string Label {
            get => _label;
            set {
                _label = value;
                if (Name is null)
                {
                    Name = _label;
                }
            }
        }
        public DataTypes DataType { get; set; } = DataTypes.Text;
        public string Name { get; set; }
        public int Index { get; internal set; } = UNDEFINED_INDEX_VALUE;
        public IFormat HeaderCellFormat { get; set; }
    }
}
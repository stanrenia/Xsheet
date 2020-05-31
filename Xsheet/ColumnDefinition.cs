namespace Xsheet
{
    public class ColumnDefinition
    {
        public const int UNDEFINED_INDEX_VALUE = -1;

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
        public Format HeaderCellFormat { get; set; }
    }
}
using System;
using System.Collections.Generic;

namespace Xsheet
{
    public interface IRowValuesBuilder
    {
        IMatrixBuilder RowValues(IEnumerable<RowValue> rowValues);
        Matrix Build();
    }

    public interface IMatrixBuilder : IRowValuesBuilder
    {
        IMatrixBuilder Key(string key = "", int index = 0);
        IMatrixBuilder Dimensions(int rowsCount, int colsCount);
        IMatrixBuilder RowsCount(int rowsCount);
        IMatrixBuilder ColsCount(int colsCount);

        IColsBuilder Cols();
        IMatrixBuilder Cols(List<ColumnDefinition> cols);
        IRowsBuilder Rows();
        IMatrixBuilder Rows(List<RowDefinition> rows);

        IMatrixBuilder WithoutHeadersRow();
    }

    public interface IColsBuilder : IRowValuesBuilder
    {
        IColsBuilder Col(string name = null, string label = null, DataTypes dataType = DataTypes.Text, IFormat headerCellFormat = null);
        IRowsBuilder Rows();
        IRowsBuilder Rows(List<RowDefinition> rows);
    }

    public interface IRowsBuilder : IRowValuesBuilder
    {
        IRowBuilder Row(string key = null, IFormat defaultCellFormat = null);
    }

    public interface IRowBuilder : IRowsBuilder
    {
        IRowBuilder ValueMap(string colName, Func<Matrix, MatrixCellValue, object> func);
        IRowBuilder Format(string colName, IFormat format);
    }
}

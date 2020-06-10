using System;
using System.Collections.Generic;
using System.Text;

namespace Xsheet
{
    public class Builders : IMatrixBuilder, IColsBuilder, IRowsBuilder, IRowBuilder
    {
        public static Builders New => new Builders();

        public IColsBuilder Add(string name = null, string label = null, DataTypes dataType = DataTypes.Text, IFormat headerCellFormat = null)
        {
            throw new NotImplementedException();
        }

        public IRowsBuilder Add(string key = null, IFormat defaultCellFormat = null, Dictionary<string, IFormat> formatsByColName = null, Dictionary<string, Func<Matrix, MatrixCellValue, object>> valuesMapping = null)
        {
            throw new NotImplementedException();
        }

        public IRowBuilder Add(string key = null, IFormat defaultCellFormat = null)
        {
            throw new NotImplementedException();
        }

        public IRowBuilder AddFormat(string colName, IFormat format)
        {
            throw new NotImplementedException();
        }

        public IRowBuilder AddValueMap(string colName, Func<Matrix, MatrixCellValue, object> func)
        {
            throw new NotImplementedException();
        }

        public Matrix Build()
        {
            throw new NotImplementedException();
        }

        public IColsBuilder Cols()
        {
            throw new NotImplementedException();
        }

        public IMatrixBuilder Continue()
        {
            throw new NotImplementedException();
        }

        public IRowsBuilder EndRow()
        {
            throw new NotImplementedException();
        }

        public IRowsBuilder Rows()
        {
            throw new NotImplementedException();
        }

        public IMatrixBuilder RowValues(IEnumerable<RowValue> rowValues)
        {
            throw new NotImplementedException();
        }
    }

    public interface IMatrixBuilder
    {
        IColsBuilder Cols();
        IRowsBuilder Rows();
        IMatrixBuilder RowValues(IEnumerable<RowValue> rowValues);
        Matrix Build();
    }


    public interface IColsBuilder
    {
        IColsBuilder Add(string name = null, string label = null, DataTypes dataType = DataTypes.Text, IFormat headerCellFormat = null);
        IColsBuilder Add(params ColumnDefinition[] columnDefinitions);
        IMatrixBuilder Continue();
    }

    public interface IRowsBuilder
    {
        //IRowsBuilder Add(string key = null, IFormat defaultCellFormat = null, Dictionary<string, IFormat> formatsByColName = null,
        //    Dictionary<string, Func<Matrix, MatrixCellValue, object>> valuesMapping = null);
        IRowsBuilder Add(string key = null, IFormat defaultCellFormat = null);
        IRowsBuilder Add(params RowDefinition[] rowDefinitions);
        IRowsBuilder AddValueMap(string colName, Func<Matrix, MatrixCellValue, object> func);
        IRowsBuilder AddFormat(string colName, IFormat format);
        IMatrixBuilder Continue();
    }

    public interface IRowBuilder
    {
        IRowBuilder AddValueMap(string colName, Func<Matrix, MatrixCellValue, object> func);
        IRowBuilder AddFormat(string colName, IFormat format);
        IRowsBuilder EndRow();
    }
}

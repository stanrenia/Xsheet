using ClosedXML.Excel;
using System;
using Xsheet;

namespace XSheet.Renderers.ClosedXml
{
    public class ClosedXmlFormatApplier : IFormatApplier
    {
        public void ApplyFormatToCell(IXLWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, int columnIndex, IXLCell cell)
        {
            throw new NotImplementedException();
        }

        public void ApplyFormatToCell(IXLWorkbook wb, IXLCell cell, IFormat format)
        {
            throw new NotImplementedException();
        }
    }
}

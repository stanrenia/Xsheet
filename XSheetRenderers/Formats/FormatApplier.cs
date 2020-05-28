﻿using NPOI.SS.UserModel;
using Xsheet;

namespace XSheet.Renderers.Formats
{
    public interface FormatApplier
    {
        void ApplyFormatToCell(IWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, string columnKey, ICell cell);
        void ApplyFormatToCell(IWorkbook wb, ICell cell, Format format);
    }
}
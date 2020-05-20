namespace Xsheet
{
    public interface IMatrixRenderer
    {
        T GenerateExcelWorksheet<T>(Matrix mat);
    }
}

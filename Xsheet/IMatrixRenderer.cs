using System.IO;

namespace Xsheet
{
    public interface IMatrixRenderer
    {
        void GenerateExcelFile(Matrix mat, Stream stream);
    }
}

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using System.Linq;

namespace Xsheet.Tests
{
    public static class TestExtensions
    {
        public static byte[] ToARGB(this Color color)
        {
            return new byte[] { color.A, color.R, color.G, color.B };
        }

        public static byte[] ToARGB(this IndexedColors color, byte alpha = 255)
        {
            return (new byte[] { alpha }).ToList().Concat(color.RGB).ToArray();
        }

        public static byte[] ToARGB(this IColor color)
        {
            return ((XSSFColor)color).ARGB;
        }
    }
}

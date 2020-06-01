using System;

namespace Xsheet
{
    public static class CellAddressCalculator
    {
        const char CHAR_A = 'A';
        const byte NUMBER_OF_LETTERS = 'Z' - 'A' + 1;

        public static string GetColumnLetters(int colIndex)
        {
            string letters = String.Empty;
            do
            {
                letters = Convert.ToChar(CHAR_A + colIndex % NUMBER_OF_LETTERS) + letters;
                colIndex = colIndex / NUMBER_OF_LETTERS - 1;
            }
            while (colIndex >= 0);
            return letters;
        }
    }
}

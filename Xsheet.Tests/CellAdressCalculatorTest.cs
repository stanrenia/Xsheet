using NFluent;
using Xunit;

namespace Xsheet.Tests
{
    public class CellAdressCalculatorTest
    {
        public CellAdressCalculatorTest()
        {

        }

        [Theory]
        [InlineData(0, "A")]
        [InlineData(1, "B")]
        [InlineData(25, "Z")]
        [InlineData(26, "AA")]
        [InlineData(27, "AB")]
        [InlineData(52, "BA")]
        [InlineData(701, "ZZ")]
        [InlineData(702, "AAA")]
        [InlineData(728, "ABA")]
        [InlineData(17602, "ZAA")]
        [InlineData(18277, "ZZZ")]
        public void Should_Return_Correct_Column_Letters(int index, string letters)
        {
            Check.That(CellAddressCalculator.GetColumnLetters(index)).IsEqualTo(letters);
        }
    }
}

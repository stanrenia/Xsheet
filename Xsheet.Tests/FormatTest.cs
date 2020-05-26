using NFluent;
using System.Collections.Generic;
using Xunit;

namespace Xsheet.Tests
{
    public class FormatTest
    {
        public FormatTest()
        {

        }

        [Fact]
        public void Should_Merge_Formats()
        {
            // GIVEN
            var f1 = new Format { BackgroundColor = "123", FontSize = 10, FontStyle = FontStyle.Bold };
            var f2 = new Format { BackgroundColor = "234", FontSize = 12 };
            var f3 = new Format { FontSize = 14 };

            // WHEN
            var mergedFormat = Format.MergeFormats(new List<Format> { f1, f2, f3 });

            // THEN
            Check.That(mergedFormat.BackgroundColor).IsEqualTo("234");
            Check.That(mergedFormat.FontSize).IsEqualTo(14);
            Check.That(mergedFormat.FontStyle).IsEqualTo(FontStyle.Bold);
        }
    }
}

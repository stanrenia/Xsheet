using NFluent;
using System.Collections.Generic;
using XSheet.Renderers.Formats;
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
            var f1 = new BasicFormat { BackgroundColor = "123", FontSize = 10, FontStyle = FontStyle.Bold };
            var f2 = new BasicFormat { BackgroundColor = "234", FontSize = 12 };
            var f3 = new BasicFormat { FontSize = 14 };

            // WHEN
            var mergedFormat = BasicFormat.MergeFormats(new List<BasicFormat> { f1, f2, f3 });

            // THEN
            Check.That(mergedFormat.BackgroundColor).IsEqualTo("234");
            Check.That(mergedFormat.FontSize).IsEqualTo(14);
            Check.That(mergedFormat.FontStyle).IsEqualTo(FontStyle.Bold);
        }
    }
}

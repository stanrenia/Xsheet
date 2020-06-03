using NFluent;
using System;
using System.Collections.Generic;
using System.IO;
using XSheet.Renderers;
using Xunit;

namespace Xsheet.Tests
{
    public class OpenXmlRendererTest : IDisposable
    {
        private readonly OpenXmlRenderer _renderer;
        private Stream _fileStream;

        public OpenXmlRendererTest()
        {
            _renderer = new OpenXmlRenderer();
        }

        public void Dispose()
        {
            
        }

        [Fact]
        public void Should_Render_Basic_Matrix()
        {
            // GIVEN
            var mat = Matrix.With()
                .RowValues(new List<RowValue>
                {
                    new RowValue { ValuesByColName = new Dictionary<string, object>
                        {
                            { "Lastname", "Bros" },
                            { "Firstname", "Mario" },
                        }
                    }
                })
                .Build();

            _fileStream = File.Create($"openxml_basic_{DateTime.Now:HHmmss}.xlsx");

            // WHEN
            _renderer.GenerateExcelFile(mat, _fileStream);

            // THEN
            Check.That(true).IsTrue();
        }
    }
}

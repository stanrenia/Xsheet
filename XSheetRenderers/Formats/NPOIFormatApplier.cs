using AutoMapper;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace XSheet.Renderers.Formats
{
    public class NPOIFormatApplier : BaseFormatApplier<NPOIFormat>
    {
        private readonly Mapper _mapper = new Mapper(new MapperConfiguration(conf =>
        {
            conf.CreateMap<NPOIFormat, NPOIFormat>()
            .ForAllMembers(source =>
            {
                source.Condition((sourceObject, destObject, sourceProperty, destProperty) =>
                {
                    var initial = default(NPOIFormat);
                    if (sourceProperty == null)
                        return !(destProperty == null);

                    return !sourceProperty.Equals(destProperty);
                });
            });
        }));
        
        public override void ApplyFormatToCell(IWorkbook wb, ICell cell, NPOIFormat format)
        {
            cell.CellStyle = format.CellStyle;
        }

        //protected override NPOIFormat MergeFormats(List<NPOIFormat> formats)
        //{
        //    var firstFormat = formats?.FirstOrDefault() ?? new NPOIFormat { CellStyle = null };
        //    return formats.Aggregate(firstFormat, (mergedFormat, nextFormat) =>
        //    {
        //        if (nextFormat.CellStyle != null)
        //        {
        //            _mapper.Map(nextFormat, mergedFormat);
        //        }
        //        return mergedFormat;
        //    });
        //}
    }
}

using BTS.Data.Core.Excel;
using ExcelUtility.Model.Base;

namespace ExcelUtility.Model
{
    public class DefaultSheetStyle : ISheetStyle
    {
        private static DefaultSheetStyle _DefaultSheetStyle;
        public static DefaultSheetStyle Default => _DefaultSheetStyle ??= CreatDefaultSheetStyle();

        public int FirstDataTableRow { get; set; }
        public bool IsAutoSizeColumn { get; set; }
        public bool IsAutoFilter { get; set; }
        public bool IsCreateFreezePane { get; set; }
        public ISheetFont Font { get; set; }
        public ChartInfo ChartInfo { get; set; }
        public bool ShowRowNum { get; set; }

        public static DefaultSheetStyle CreatDefaultSheetStyle() => new DefaultSheetStyle()
        {
            FirstDataTableRow = 0,
            IsAutoSizeColumn = true,
            IsAutoFilter = true,
            IsCreateFreezePane = true,
            Font = new SheetFont() { FontName = "Microsoft YaHei UI", FontSize = 10 }
        };
    }
}
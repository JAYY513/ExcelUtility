using ExcelUtility.Model.Base;

namespace BTS.Data.Core.Excel
{
    public class SheetFont : ISheetFont
    {
        public string FontName { get; set; }
        public double FontSize { get; set; }
    }
}
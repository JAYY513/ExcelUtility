using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelUtility.Model.Base
{
    public class BaseExcelSheet<T> : IExcelSheet where T : class
    {
        public List<T> Datas { get; set; }
        public virtual ISheetStyle SheetStyle { get; set; } = DefaultSheetStyle.CreatDefaultSheetStyle();

        public async ValueTask WriteToExcelAsync(string filePath, IWorkbook workBook)
        {
            await ExcelHelper.WriteToExcelAsync(workBook, Datas, SheetStyle);
        }
    }
}
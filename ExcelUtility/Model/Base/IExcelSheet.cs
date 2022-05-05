using NPOI.SS.UserModel;
using System.Threading.Tasks;

namespace ExcelUtility.Model.Base
{
    public interface IExcelSheet
    {
        public ISheetStyle SheetStyle { get; set; }

        public ValueTask WriteToExcelAsync(string filePath, IWorkbook workBook);
    }
}
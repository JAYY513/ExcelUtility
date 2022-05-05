namespace ExcelUtility.Model.Base
{
    public enum ChartType
    {
        Line,
        Bar
    }

    public interface ISheetStyle
    {
        public int FirstDataTableRow { get; set; }
        public bool IsAutoSizeColumn { get; set; }
        public bool IsAutoFilter { get; set; }
        public bool IsCreateFreezePane { get; set; }
        public ISheetFont Font { get; set; }
        public ChartInfo ChartInfo { get; set; }
        public bool ShowRowNum { get; set; }
    }

    public interface ISheetFont
    {
        public string FontName { get; set; }
        public double FontSize { get; set; }
    }

    public class ChartInfo
    {
        public string SerieTitle { get; set; }

        public ChartInfo(int col1, int row1, int col2, int row2)
        {
            Col1 = col1;
            Col2 = col2;
            Row1 = row1;
            Row2 = row2;
        }

        public ChartType ChartType { get; set; }
        public int Col1 { get; set; }
        public int Row1 { get; set; }
        public int Col2 { get; set; }
        public int Row2 { get; set; }
        public int ValueColumnIndex { get; set; }
        public int CategoryColumnIndex { get; set; }
        public string ChartTitle { get; set; }
        public string ValueAxisTitle { get; set; }
        public string CatAxisTitle { get; set; }

        //public static ChartInfo CreatChartInfo()
        //{
        //    return new ChartInfo(1, 1, 9, 18)
        //    {
        //        CategoryColumnIndex = 8,
        //        ValueColumnIndex = 5,
        //        SerieTitle = "采集电压",
        //        CatAxisTitle = "测试时间",
        //        ValueAxisTitle = "采集电压",
        //        ChartTitle = "采集电压 \\ 测试时间(记录)"
        //    };
        //}
    }
}
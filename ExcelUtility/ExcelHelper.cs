using ExcelUtility.Attributes;
using ExcelUtility.Model.Base;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelUtility
{
    public static class ExcelHelper
    {
        public static async ValueTask WriteToExcelAsync(string path, List<IExcelSheet> sheets)
        {
            IWorkbook workbook = new XSSFWorkbook();

            if (File.Exists(path))
                File.Delete(path);

            foreach (var sheet in sheets)
            {
                await sheet.WriteToExcelAsync(path, workbook);
            }

            using (FileStream fs = File.Open(path, FileMode.OpenOrCreate))
            {
                workbook.Write(fs);
                workbook.Close();
            }
        }

        public static async ValueTask WriteToExcelAsync<T>(IWorkbook workbook, List<T> list, ISheetStyle sheetStyle) where T : class
        {
            await Task.Run(() =>
            {
                string name = PropertyInfoHelper.GetAlias(typeof(T));
                string tempName = name;
                for (int i = 2; workbook.GetSheet(tempName) != null; i++)
                {
                    tempName = name + i;
                }
                ISheet sheet = workbook.CreateSheet(tempName);
                int rowIndex = sheetStyle?.FirstDataTableRow ?? 0;

                #region 设置Excel格式

                var allProp = PropertyInfoHelper.GetPropertyInfosWithListType(typeof(T));
                if (sheetStyle != null)
                {
                    if (sheetStyle.IsAutoSizeColumn)//自适应宽度
                    {
                        for (int columnIndex = 0; columnIndex < allProp.Count; columnIndex++)
                        {
                            sheet.AutoSizeColumn(columnIndex);
                        }
                    }
                    if (sheetStyle.IsAutoFilter)
                        sheet.SetAutoFilter(new CellRangeAddress(rowIndex, rowIndex, 0, allProp.Count - 1 + (sheetStyle.ShowRowNum == true ? 1 : 0))); //首行筛选
                    if (sheetStyle.IsCreateFreezePane)
                        sheet.CreateFreezePane(0, rowIndex + 1); //首行冻结
                }

                #endregion 设置Excel格式

                for (int i = 0; i < rowIndex; i++)
                {
                    sheet.CreateRow(i);
                }

                IRow row = sheet.CreateRow(rowIndex++);
                int indexColumn = 0;
                if (sheetStyle?.ShowRowNum == true)//设置编号
                {
                    var cell = row.CreateCell(indexColumn);
                    var value = "NO";
                    int exWidth = 0;

                    #region 编号列头格式设置

                    if (sheetStyle != null)
                    {
                        if (sheetStyle.Font != null)
                        {
                            ICellStyle cellStyle = workbook.CreateCellStyle();
                            IFont font = workbook.CreateFont();
                            font.FontName = sheetStyle.Font.FontName;
                            font.FontHeightInPoints = sheetStyle.Font.FontSize;
                            cellStyle.SetFont(font);
                            cell.CellStyle = cellStyle;
                        }
                        if (sheetStyle.IsAutoFilter)
                            exWidth += 2 * 256;
                    }

                    #endregion 编号列头格式设置

                    cell.SetCellValue(value);

                    //if ((value?.StrLength() > 8 || (value?.StrLength() > 4 && exWidth > 0)) && sheet.GetColumnWidth(startIndexColumn) < (value?.StrLength() + 2) * 256 + exWidth)
                    //    sheet.SetColumnWidth(startIndexColumn, (value.StrLength() + 2) * 256 + exWidth);
                    indexColumn++;
                }//设置编号

                SetColumnHead(typeof(T), ref indexColumn);
                void SetColumnHead(Type type, ref int startIndexColumn)
                {
                    var properties = PropertyInfoHelper.GetPropertyInfos(type);
                    for (int i = 0; i < properties.Count; i++)
                    {
                        var aliasType = PropertyInfoHelper.GetAliasType(properties[i]);
                        if (aliasType == AliasAttributeType.List)
                        {
                            SetColumnHead(properties[i].PropertyType.GenericTypeArguments[0], ref startIndexColumn);
                        }
                        else
                        {
                            var cell = row.CreateCell(startIndexColumn);
                            var value = PropertyInfoHelper.GetAliasName(properties[i]);
                            int exWidth = 0;

                            #region 列头格式设置

                            if (sheetStyle != null)
                            {
                                if (sheetStyle.Font != null)
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    IFont font = workbook.CreateFont();
                                    font.FontName = sheetStyle.Font.FontName;
                                    font.FontHeightInPoints = sheetStyle.Font.FontSize;
                                    cellStyle.SetFont(font);
                                    cell.CellStyle = cellStyle;
                                }
                                if (sheetStyle.IsAutoFilter)
                                    exWidth += 2 * 256;
                            }

                            #endregion 列头格式设置

                            cell.SetCellValue(value);

                            //if ((value?.StrLength() > 8 || (value?.StrLength() > 4 && exWidth > 0)) && sheet.GetColumnWidth(startIndexColumn) < (value?.StrLength() + 2) * 256 + exWidth)
                            //    sheet.SetColumnWidth(startIndexColumn, (value.StrLength() + 2) * 256 + exWidth);
                            startIndexColumn++;
                        }
                    }
                }

                list?.ForEach(r =>
                {
                    rowIndex = SetColumnValue(typeof(T), rowIndex, 0, r).Item1;
                });

                (int, int) SetColumnValue(Type type, int startRowIndex, int startColumnIndex, object obj)
                {
                    var properties = PropertyInfoHelper.GetPropertyInfos(type);
                    int outRow = startRowIndex + 1;
                    int outColumn = startColumnIndex;
                    var row = sheet.GetRow(startRowIndex);
                    if (row == null)
                    {
                        row = sheet.CreateRow(startRowIndex);
                        if(sheetStyle?.ShowRowNum == true)
                        {
                            var cell = row.CreateCell(0);
                            StrategyHelper.SetCellValue(workbook, cell, startRowIndex, null, sheetStyle);
                            if (startColumnIndex == 0)
                                startColumnIndex++;
                        }
                    }
                    for (int i = 0; i < properties.Count; i++)
                    {
                        var aliasType = PropertyInfoHelper.GetAliasType(properties[i]);
                        if (aliasType == AliasAttributeType.List)
                        {
                            IEnumerable list = obj == null ? null : properties[i].GetValue(obj) as IEnumerable;
                            var propertieItems = PropertyInfoHelper.GetPropertyInfos(properties[i].PropertyType.GenericTypeArguments[0]);

                            if (list != null)
                            {
                                int listStartRowIndex = startRowIndex;
                                int listStartColumnIndex = startColumnIndex;
                                foreach (var item in list)
                                {
                                    if (PropertyInfoHelper.ObjectIsAllNullOrEmpty(item))
                                        continue;
                                    listStartColumnIndex = startColumnIndex;
                                    var tuple = SetColumnValue(properties[i].PropertyType.GenericTypeArguments[0], listStartRowIndex, listStartColumnIndex, item);
                                    listStartRowIndex = tuple.Item1;
                                    listStartColumnIndex = tuple.Item2;
                                    outRow = Math.Max(outRow, listStartRowIndex);
                                    outColumn = Math.Max(outColumn, listStartColumnIndex);
                                }
                                startColumnIndex = outColumn;
                            }
                            else
                            {
                                int listStartRowIndex = startRowIndex;
                                int listStartColumnIndex = startColumnIndex;
                                var tuple = SetColumnValue(properties[i].PropertyType.GenericTypeArguments[0], listStartRowIndex, listStartColumnIndex, null);
                                listStartRowIndex = tuple.Item1;
                                listStartColumnIndex = tuple.Item2;
                                outRow = Math.Max(outRow, listStartRowIndex - 1);
                                outColumn = Math.Max(outColumn, listStartColumnIndex);
                                startColumnIndex = outColumn;
                            }
                        }
                        else
                        {
                            if (obj != null)
                            {
                                var cell = row.CreateCell(startColumnIndex);
                                StrategyHelper.SetCellValue(workbook, cell, properties[i].GetValue(obj), properties[i], sheetStyle);
                            }
                            startColumnIndex++;
                        }
                    }
                    return (outRow, startColumnIndex);
                }

                #region 设置Excel格式

                if (sheetStyle != null)
                {
                    sheet.Autobreaks = true;

                    if (sheetStyle.IsAutoSizeColumn)//自适应宽度
                    {
                        for (int columnIndex = 0; columnIndex < allProp.Count; columnIndex++)
                        {
                            sheet.AutoSizeColumn(columnIndex);
                            if (sheetStyle.IsAutoFilter)
                            {
                                sheet.SetColumnWidth(columnIndex, sheet.GetColumnWidth(columnIndex) + 588);
                            }
                        }
                    }
                }

                #endregion 设置Excel格式

                #region 设置Chart

                if (sheetStyle?.ChartInfo != null)
                {
                    var properties = PropertyInfoHelper.GetPropertyInfos<T>();

                    IDrawing drawing = sheet.CreateDrawingPatriarch();
                    int startChartDataRow = (sheetStyle?.FirstDataTableRow ?? 0) + 1;
                    //锚点
                    IClientAnchor anchor1 = drawing.CreateAnchor(0, 0, 0, 0, sheetStyle.ChartInfo.Col1, sheetStyle.ChartInfo.Row1, sheetStyle.ChartInfo.Col2, sheetStyle.ChartInfo.Row2);
                    if (sheetStyle.ChartInfo.ChartType == ChartType.Line)
                        CreateLineChart(sheet, drawing, anchor1, sheetStyle.ChartInfo.SerieTitle, sheetStyle.ChartInfo.ChartTitle, sheetStyle.ChartInfo.ValueAxisTitle, sheetStyle.ChartInfo.CatAxisTitle, startChartDataRow, list.Count + startChartDataRow - 1, sheetStyle.ChartInfo.ValueColumnIndex, sheetStyle.ChartInfo.CategoryColumnIndex);
                    else
                        CreateBarChart(sheet, drawing, anchor1, sheetStyle.ChartInfo.SerieTitle, sheetStyle.ChartInfo.ChartTitle, sheetStyle.ChartInfo.ValueAxisTitle, sheetStyle.ChartInfo.CatAxisTitle, startChartDataRow, list.Count + startChartDataRow - 1, sheetStyle.ChartInfo.ValueColumnIndex, sheetStyle.ChartInfo.CategoryColumnIndex);
                }

                #endregion 设置Chart
            });
        }

        private static void CreateLineChart(ISheet sheet, IDrawing drawing, IClientAnchor anchor, string serieTitle, string chartTitle, string valueAxisTitle, string catAxisTitle, int startDataRow, int endDataRow, int columnIndex, int categorycolumnIndex)
        {
            XSSFChart chart = (XSSFChart)drawing.CreateChart(anchor);

            ILineChartData<string, double> barChartData = chart.ChartDataFactory.CreateLineChartData<string, double>();

            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = LegendPosition.Right;

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            bottomAxis.MajorTickMark = AxisTickMark.None;
            bottomAxis.IsVisible = true;
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.MajorTickMark = AxisTickMark.Out;
            leftAxis.SetCrossBetween(AxisCrossBetween.Between);

            IChartDataSource<string> categoryAxis = DataSources.FromStringCellRange(sheet, new CellRangeAddress(startDataRow, endDataRow, categorycolumnIndex, categorycolumnIndex));
            IChartDataSource<double> valueAxis = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(startDataRow, endDataRow, columnIndex, columnIndex));
            var serie = barChartData.AddSeries(categoryAxis, valueAxis);
            serie.SetTitle(serieTitle);
            chart.SetTitle(chartTitle);
            chart.GetCTChartSpace().chart.title.overlay = new CT_Boolean() { val = 0 };
            var p = chart.GetCTChartSpace().chart.title.tx.rich.p[0];
            p.AddNewPPr();
            p.pPr.defRPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            setCatAxisTitle(chart, 0, catAxisTitle);
            setValueAxisTitle(chart, 0, valueAxisTitle);
            chart.Plot(barChartData, bottomAxis, leftAxis);
        }

        //set Value Axis
        private static void setValueAxisTitle(XSSFChart chart, int axisIdx, String title)
        {
            NPOI.OpenXmlFormats.Dml.Chart.CT_ValAx valAx = chart.GetCTChart().plotArea.valAx[axisIdx];
            valAx.title = new NPOI.OpenXmlFormats.Dml.Chart.CT_Title();
            NPOI.OpenXmlFormats.Dml.Chart.CT_Title ctTitle = valAx.title;
            ctTitle.layout = new NPOI.OpenXmlFormats.Dml.Chart.CT_Layout();
            ctTitle.overlay = new NPOI.OpenXmlFormats.Dml.Chart.CT_Boolean();
            ctTitle.overlay.val = 0;
            ctTitle.AddNewTx();

            NPOI.OpenXmlFormats.Dml.Chart.CT_TextBody rich = ctTitle.tx.AddNewRich();
            rich.AddNewBodyPr();
            rich.AddNewLstStyle();
            rich.AddNewP();
            NPOI.OpenXmlFormats.Dml.CT_TextParagraph p = rich.p[0];
            p.AddNewPPr();
            p.pPr.defRPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            p.AddNewR().t = title;
            p.AddNewEndParaRPr();
        }

        //set cat Axis
        private static void setCatAxisTitle(XSSFChart chart, int axisIdx, string title)
        {
            chart.GetCTChart().plotArea.catAx[axisIdx].title = new NPOI.OpenXmlFormats.Dml.Chart.CT_Title();
            NPOI.OpenXmlFormats.Dml.Chart.CT_Title ctTitle = chart.GetCTChart().plotArea.catAx[axisIdx].title;// new NPOI.OpenXmlFormats.Dml.Chart.CT_Title();
            ctTitle.layout = new NPOI.OpenXmlFormats.Dml.Chart.CT_Layout();
            ctTitle.layout.AddNewManualLayout();
            NPOI.OpenXmlFormats.Dml.Chart.CT_Boolean ctbool = new NPOI.OpenXmlFormats.Dml.Chart.CT_Boolean();
            ctbool.val = 0;
            ctTitle.overlay = ctbool;
            ctTitle.AddNewTx();
            NPOI.OpenXmlFormats.Dml.Chart.CT_TextBody rich = ctTitle.tx.AddNewRich();
            rich.AddNewBodyPr();

            rich.AddNewLstStyle();
            rich.AddNewP();
            NPOI.OpenXmlFormats.Dml.CT_TextParagraph p = rich.p[0];
            p.AddNewPPr();
            p.pPr.defRPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            p.AddNewR().t = title;
            p.AddNewEndParaRPr();
        }

        private static void CreateBarChart(ISheet sheet, IDrawing drawing, IClientAnchor anchor, string serieTitle, string chartTitle, string valueAxisTitle, string catAxisTitle, int startDataRow, int endDataRow, int columnIndex, int categorycolumnIndex)
        {
            XSSFChart chart = (XSSFChart)drawing.CreateChart(anchor);

            IBarChartData<string, double> barChartData = chart.ChartDataFactory.CreateBarChartData<string, double>();

            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = LegendPosition.Right;

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            bottomAxis.MajorTickMark = AxisTickMark.None;
            bottomAxis.IsVisible = true;
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.MajorTickMark = AxisTickMark.Out;
            leftAxis.SetCrossBetween(AxisCrossBetween.Between);

            IChartDataSource<string> categoryAxis = DataSources.FromStringCellRange(sheet, new CellRangeAddress(startDataRow, endDataRow, categorycolumnIndex, categorycolumnIndex));
            IChartDataSource<double> valueAxis = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(startDataRow, endDataRow, columnIndex, columnIndex));
            var serie = barChartData.AddSeries(categoryAxis, valueAxis);
            serie.SetTitle(serieTitle);
            chart.SetTitle(chartTitle);
            chart.GetCTChartSpace().chart.title.overlay = new CT_Boolean() { val = 0 };
            var p = chart.GetCTChartSpace().chart.title.tx.rich.p[0];
            p.AddNewPPr();
            p.pPr.defRPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            setCatAxisTitle(chart, 0, catAxisTitle);
            setValueAxisTitle(chart, 0, valueAxisTitle);
            chart.Plot(barChartData, bottomAxis, leftAxis);
        }
    }
}
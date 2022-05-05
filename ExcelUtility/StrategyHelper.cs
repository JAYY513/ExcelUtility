using ExcelUtility.Attributes;
using ExcelUtility.Model.Base;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using static ExcelUtility.Attributes.AliasAttribute;

namespace ExcelUtility
{
    public static class StrategyHelper
    {
        public static DataTable FieldsToTable<TValues>(List<TValues> Fields)
        {
            DataTable toReturn = new DataTable();
            toReturn.Columns.Add(new DataColumn("No", typeof(int)));

            var properties = PropertyInfoHelper.GetPropertyInfos<TValues>();

            var columns = PropertyInfoHelper.GetPropertyInfosWithListType(typeof(TValues)).Select(m =>
            {
                string columnName = PropertyInfoHelper.GetAliasName(m);
                var attr = m.GetCustomAttribute(typeof(BaseConvertAttribute)) as BaseConvertAttribute;
                return new DataColumn(columnName, attr?.GetType() ?? m.PropertyType);
            });

            toReturn.Columns.AddRange(columns.ToArray());
            int num = 0;
            int indexi = 1;
            Fields.ForEach(r =>
            {
                var row = SetValueValue(typeof(TValues), num, r);
                var dr = toReturn.Rows[num];
                dr["No"] = indexi;
                num = row;
                indexi++;
            });
            int SetValueValue(Type type, int startRowIndex, object obj)
            {
                DataRow dr;
                var properties = PropertyInfoHelper.GetPropertyInfos(type);
                int outRow = startRowIndex + 1;
                if (toReturn.Rows.Count <= startRowIndex)
                {
                    string oldRow = null;
                    for (int i = toReturn.Rows.Count - 1; i >= 0; i--)
                    {
                        var str = toReturn.Rows[i]["No"].ToString();
                        if (!string.IsNullOrEmpty(str))
                        {
                            oldRow = str;
                            break;
                        }
                    }
                    dr = toReturn.NewRow();
                    if (oldRow != indexi.ToString())
                        dr["No"] = indexi;
                    toReturn.Rows.Add(dr);
                }
                else
                    dr = toReturn.Rows[startRowIndex];
                for (int i = 0; i < properties.Count; i++)
                {
                    var aliasType = PropertyInfoHelper.GetAliasType(properties[i]);
                    if (aliasType == AliasAttributeType.List)
                    {
                        var list = properties[i].GetValue(obj) as IEnumerable;
                        var propertieItems = PropertyInfoHelper.GetPropertyInfos(properties[i].PropertyType.GenericTypeArguments[0]);

                        if (list != null)
                        {
                            int listStartRowIndex = startRowIndex;
                            foreach (var item in list)
                            {
                                if (PropertyInfoHelper.ObjectIsAllNullOrEmpty(item))
                                    continue;
                                var tuple = SetValueValue(properties[i].PropertyType.GenericTypeArguments[0], listStartRowIndex, item);
                                listStartRowIndex = tuple;
                                outRow = Math.Max(outRow, listStartRowIndex);
                            }
                        }
                    }
                    else
                    {
                        string columnName = PropertyInfoHelper.GetAliasName(properties[i]);
                        if (toReturn.Columns.Contains(columnName))//或者别名
                        {
                            try
                            {
                                object value = PropertyInfoHelper.GetDefaultOrValueByConvertAttribute(properties[i], properties[i].GetValue(obj));
                                if (value is TimeSpan timeSpan)
                                {
                                    var str = timeSpan.ToString(@"hh\:mm\:ss\.fff");
                                    value = str;
                                }
                                dr[columnName] = obj == null ? string.Empty : value;
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                    }
                }
                return outRow;
            }
            return toReturn;
        }

        public static void SetCellValue(IWorkbook workbook, ICell cell, object obj, PropertyInfo propertyInfo, ISheetStyle sheetStyle)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            if (sheetStyle != null && sheetStyle.Font != null)
            {
                IFont font = workbook.CreateFont();
                font.FontName = sheetStyle.Font.FontName;
                font.FontHeightInPoints = sheetStyle.Font.FontSize;
                cellStyle.SetFont(font);
                cellStyle.Alignment = HorizontalAlignment.Center;
            }
            if (obj == null)
                return;
            string str = obj.ToString();
            obj = PropertyInfoHelper.GetDefaultOrValueByConvertAttribute(propertyInfo, obj);
            if (obj is TimeSpan timeSpan)
            {
                str = timeSpan.ToString(@"hh\:mm\:ss\.fff");
                cell.CellStyle = cellStyle;
                cell.SetCellValue(str);
            }
            else if (obj is DateTime dt)
            {
                //ICellStyle cellStyle = workbook.CreateCellStyle();
                //IFont font = workbook.CreateFont();
                //font.FontName = "Microsoft Sans Serif";
                //cellStyle.SetFont(font);
                IDataFormat datastyle = workbook.CreateDataFormat();
                cellStyle.DataFormat = datastyle.GetFormat("yyyy/mm/dd hh:mm:ss.000");
                cell.CellStyle = cellStyle;
                cell.SetCellValue(dt);
            }
            else if (double.TryParse(str, out double dou))
            {
                cell.CellStyle = cellStyle;
                cell.SetCellValue(dou);
            }
            else if (obj is IRichTextString irts)
            {
                cell.CellStyle = cellStyle;
                cell.SetCellValue(irts);
            }
            else if (obj is bool b)
            {
                cell.CellStyle = cellStyle;
                cell.SetCellValue(b);
            }
            else
            {
                cell.CellStyle = cellStyle;
                cell.SetCellValue(obj.ToString());
            }
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:
                    return null;

                case CellType.Boolean: //BOOLEAN:
                    return cell.BooleanCellValue;

                case CellType.Numeric: //NUMERIC:
                    short format = cell.CellStyle.DataFormat;
                    if (format != 0) { return cell.DateCellValue; } else { return cell.NumericCellValue; }
                case CellType.String: //STRING:
                    return cell.StringCellValue;

                case CellType.Error: //ERROR:
                    return cell.ErrorCellValue;

                case CellType.Formula: //FORMULA:
                default:
                    return "=" + cell.CellFormula;
            }
        }
    }
}
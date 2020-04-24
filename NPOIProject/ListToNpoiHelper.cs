using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
/********************************************************************************
** 类名称： 泛型List导出Excel
** 描述：指定导出列或全部导出Excel
** 作者： Time
** 创建时间：2017-5-8
** 最后修改人：（无）
** 最后修改时间：（无）
** 版权所有 (C) :Time
*********************************************************************************/
namespace NPOIProject
{
    /// <summary>
    /// 
    /// </summary>
    public class ListToNpoiHelper
    {
        /// <summary>
        /// 指定导出列名称
        /// </summary>
        public enum ListToNpoiEnum
        {
            /// <summary>
            /// Name
            /// </summary>
            English = 0,
            /// <summary>
            /// 标记名称
            /// </summary>
            Chinese = 1
        }
        #region List To Excel
        /// <summary>
        /// 集合列表导出Nopi Time 2017-4-14
        /// </summary>
        /// <param name="list">列表集合</param>
        /// <param name="fileName">文件位置</param>
        /// <param name="isIndex">是否有序号</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static ResultStateDto ListToNpoi<T>(List<T> list, string fileName, bool isIndex)
        {
            return ListToNpoi(list, fileName, null, ListToNpoiEnum.English, isIndex);
        }
        /// <summary>
        /// 集合列表导出Nopi Time 2017-4-14
        /// </summary>
        /// <param name="list">列表集合</param>
        /// <param name="fileName">文件位置</param>
        /// <param name="describe">是用中文描述匹配还是属性名匹配</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static ResultStateDto ListToNpoi<T>(List<T> list, string fileName, ListToNpoiEnum describe)
        {
            return ListToNpoi(list, fileName, null, describe, false);
        }
        /// <summary>
        /// 集合列表导出Nopi Time 2017-4-14
        /// </summary>
        /// <param name="list">列表集合</param>
        /// <param name="fileName">文件位置</param>
        /// <param name="columnAll">自定列</param>
        /// <param name="describe">是用中文描述匹配还是属性名匹配</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static ResultStateDto ListToNpoi<T>(List<T> list, string fileName, List<string> columnAll, ListToNpoiEnum describe)
        {
            return ListToNpoi(list, fileName, columnAll, describe, false);
        }

        /// <summary>
        /// 集合列表导出Nopi Time 2017-4-14
        /// </summary>
        /// <param name="list">列表集合</param>
        /// <param name="fileName">文件位置</param>
        /// <param name="columnAll">自定列</param>
        /// <param name="isShowindex"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static ResultStateDto ListToNpoi<T>(List<T> list, string fileName, List<string> columnAll, bool isShowindex)
        {
            return ListToNpoi(list, fileName, columnAll, ListToNpoiEnum.English, isShowindex);
        }
        /// <summary>
        /// 集合列表导出Nopi Time 2017-4-14
        /// </summary>
        /// <param name="list">列表集合</param>
        /// <param name="fileName">文件位置</param>
        /// <param name="columnList">导出列</param>
        /// <param name="isShowindex">是否有序号</param>
        /// <param name="mergeDto">合并参数</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static ResultStateDto ListToNpoi<T>(List<T> list, string fileName, ListToNpoiColumnList columnList, bool isShowindex, ListToNpoiMergeDto mergeDto)
        {
            if (mergeDto != null)
            {
                if (mergeDto.ListToNpoiMergeRange == null)
                {
                    mergeDto.ListToNpoiMergeRange = new List<ListToNpoiMergeRange>();
                    mergeDto.ListToNpoiMergeRange.AddRange(columnList.ColumnAll.Where(s => s.Value).Select(s => new ListToNpoiMergeRange
                    {
                        StartName = s.Key,
                        EndName = s.Key,
                    }));
                }
                else
                {
                    foreach (var item in columnList.ColumnAll.Where(s => s.Value).Where(item => mergeDto.ListToNpoiMergeRange.Count(s => s.StartName.ToLower() != item.Key.ToLower()) <= 0))
                    {
                        mergeDto.ListToNpoiMergeRange.Add(new ListToNpoiMergeRange
                        {
                            StartName = item.Key,
                            EndName = item.Key,
                        });
                    }
                }
            }
            return ListToNpoi(list, fileName, columnList.ColumnAll.Select(s => s.Key.ToLower()).ToList(), columnList.Describe, isShowindex, mergeDto);
        }
        /// <summary>
        /// 集合列表导出Nopi Time 2017-4-14
        /// </summary>
        /// <param name="list">列表集合</param>
        /// <param name="fileName">文件位置</param>
        /// <param name="columnAll">需要展示列默认取 ExportAttribute</param>
        /// <param name="describe"></param>
        /// <param name="isShowindex">是否有序号</param>
        /// <param name="mergeDto">合并参数</param>
        /// <param name="customerColumnName">合并参数</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        // ReSharper disable once MethodOverloadWithOptionalParameter
        public static ResultStateDto ListToNpoi<T>(List<T> list, string fileName, List<string> columnAll = null,
            ListToNpoiEnum describe = ListToNpoiEnum.English, bool isShowindex = true,
            ListToNpoiMergeDto mergeDto = null, Dictionary<string, string> customerColumnName = null)
        {
            //Key{ Key 类属性 Value Index 位置 }  属性  Value { Key : description Value :  Name}
            var exportNames = new List<KeyValue<KeyValue<PropertyInfo, int>, KeyValue<string, string>>>();

            var colIndex = 0;

            #region 获取T所有属性 名称  描述  类型  下标等等
            //获取所有的自定义列
            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                if (propertyInfo.GetCustomAttributes(typeof(NoneExportAttribute), false).Length > 0)
                    continue;
                var descAttrs = propertyInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
                var description = string.Empty;
                if (descAttrs.Length > 0)
                {
                    var des = (DescriptionAttribute)descAttrs[0];
                    description = des.Description;
                }
                string name;
                var exportAttr = propertyInfo.GetCustomAttributes(typeof(ExportAttribute), false);
                if (exportAttr.Length <= 0)
                {
                    name = propertyInfo.Name;
                }
                else
                {
                    var export = (ExportAttribute)exportAttr[0];
                    name = export.HeaderName ?? propertyInfo.Name;
                }
                exportNames.Add(new KeyValue<KeyValue<PropertyInfo, int>, KeyValue<string, string>>(new KeyValue<PropertyInfo, int>(propertyInfo, colIndex),
                    new KeyValue<string, string>(description, name.ToLower())));
                colIndex++;
            }
            #endregion

            #region 过滤需要导出的字段
            if (columnAll != null)
            {
                switch (describe)
                {
                    case ListToNpoiEnum.English:
                        exportNames = columnAll.Select(items => exportNames.FirstOrDefault(s => String.Equals(s.Key.Key.Name, items, StringComparison.CurrentCultureIgnoreCase))).ToList();
                        break;
                    case ListToNpoiEnum.Chinese:
                        exportNames = columnAll.Select(items => exportNames.FirstOrDefault(s => s.Value.Key.ToLower() == items.ToLower())).ToList();
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(describe), describe, null);
                }
            }
            #endregion

            #region 是否支持序号
            if (isShowindex)
            {
                exportNames.Insert(0, new KeyValue<KeyValue<PropertyInfo, int>, KeyValue<string, string>>(null, new KeyValue<string, string>("序号", "序号")));
            }
            #endregion

            #region Excel组装
            var workbook = new XSSFWorkbook();

            var sheet = workbook.CreateSheet();
            //行索引
            var rowIndex = 0;

            #region 替换成自定义列名
            if (customerColumnName != null && customerColumnName.Count > 0)
            {
                foreach (var item in exportNames)
                {
                    if (item.Key != null && item.Key.Key != null)
                        if (customerColumnName.Keys.Contains(item.Key.Key.Name))
                        {
                            item.Value.Key = customerColumnName[item.Key.Key.Name];
                        }
                }
            }
            #endregion


            #region 添加标题列
            //列总数
            var row = sheet.CreateRow(rowIndex++);
            for (var i = 0; i < exportNames.Count; i++)
            {
                #region 时间类型长度加宽
                if (i != 0)
                {
                    if (new List<object> { typeof(DateTime), typeof(DateTime?) }.Contains(exportNames[i].Key.Key.PropertyType))
                    {
                        sheet.SetColumnWidth(i, 3500);
                    }
                }
                #endregion

                var name = !string.IsNullOrEmpty(exportNames[i].Value.Key) ? exportNames[i].Value.Key : exportNames[i].Value.Value;
                row.CreateCell(i, CellType.String).SetCellValue(name);
            }
            #endregion

            #region 时间样式
            //时间样式
            var styleDateTime = workbook.CreateCellStyle();
            var format = workbook.CreateDataFormat();
            styleDateTime.DataFormat = format.GetFormat("yyyy-MM-dd");
            styleDateTime.VerticalAlignment = VerticalAlignment.Center;//垂直对齐(默认应该为center，如果center无效则用justify)
            styleDateTime.Alignment = HorizontalAlignment.Center;//水平对齐
            #endregion

            #region 普通样式
            //居中样式
            var cellstyleCenter = workbook.CreateCellStyle();//设置垂直居中格式
            cellstyleCenter.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            #endregion

            //合并行数
            var listrowIndex = 0;
            var mergeRow = 0;
            foreach (var item in list)
            {
                row = sheet.CreateRow(rowIndex++); //创建内容行    
                #region 填充一行
                for (var i = 0; i < exportNames.Count; i++)
                {
                    //创建单元格
                    var cell = row.CreateCell(i);
                    //默认字符串格式
                    var ctype = CellType.String;
                    #region 是否开启索引号
                    if (isShowindex)
                    {
                        if (i == 0)
                        {
                            row.CreateCell(i, CellType.Numeric).SetCellValue(rowIndex - 1);
                            continue;
                        }
                    }
                    #endregion
                    #region 列判断类型并赋值
                    var attr = item.GetAttributesArray();
                    if (exportNames[i].Key.Key.PropertyType == typeof(double))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((double)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(float))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((float)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(short))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((short)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(int))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((int)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(int?))
                    {
                        if (attr[exportNames[i].Value.Value] != null)
                        {
                            ctype = CellType.Numeric;
                            cell.SetCellValue((int)attr[exportNames[i].Value.Value]);
                        }
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(long))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((long)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(decimal))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue(Convert.ToDouble(attr[exportNames[i].Value.Value]));
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(bool))
                    {
                        ctype = CellType.Boolean;
                        cell.SetCellValue((bool)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(DateTime))
                    {
                        if ((DateTime)attr[exportNames[i].Value.Value] != DateTime.MinValue)
                            cell.SetCellValue((DateTime)attr[exportNames[i].Value.Value]);
                        cell.CellStyle = styleDateTime;
                        continue;
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(DateTime?))
                    {
                        if (attr[exportNames[i].Value.Value] != null)
                        {
                            if ((DateTime)attr[exportNames[i].Value.Value] != DateTime.MinValue)
                                cell.SetCellValue((DateTime)attr[exportNames[i].Value.Value]);
                            cell.CellStyle = styleDateTime;
                            continue;
                        }
                    }
                    else
                    {
                        if (attr[exportNames[i].Value.Value] != null)
                        {
                            ctype = CellType.String;
                            cell.SetCellValue(attr[exportNames[i].Value.Value].ToString());
                        }
                    }
                    cell.SetCellType(ctype);
                    cell.CellStyle = cellstyleCenter;
                    #endregion
                }
                #endregion

                if (list.Count == 1)
                    break;
                #region 合并
                if (mergeDto != null && mergeDto.Contrast.Count != 0)
                {
                    //是否是最后一个合并行
                    bool isEndMerge = false;
                    //获取当前行
                    var trow = item.GetAttributesArray();
                    //获取下一行
                    Dictionary<string, object> newtrow = null;
                    #region 判断最后一行
                    if (listrowIndex < list.Count - 1)
                        newtrow = list[listrowIndex + 1].GetAttributesArray();
                    #endregion
                    #region HasSet 比较 结果集
                    var thisrow = new HashSet<object>();
                    var nextrow = new HashSet<object>();
                    #endregion

                    if (newtrow != null)
                    {
                        #region 判断标准填充结果集 需要选择中英文
                        foreach (var items in mergeDto.Contrast)
                        {
                            switch (describe)
                            {
                                case ListToNpoiEnum.English:
                                    thisrow.Add(trow[items.ToLower()]);
                                    nextrow.Add(newtrow[items.ToLower()]);
                                    break;
                                case ListToNpoiEnum.Chinese:
                                    thisrow.Add(trow[exportNames.Single(s => String.Equals(s.Value.Key, items.ToLower(), StringComparison.CurrentCultureIgnoreCase)).Value.Value]);
                                    nextrow.Add(newtrow[exportNames.Single(s => String.Equals(s.Value.Key, items.ToLower(), StringComparison.CurrentCultureIgnoreCase)).Value.Value]);
                                    break;
                                default:
                                    throw new ArgumentOutOfRangeException(nameof(describe), describe, null);
                            }
                        }
                        #endregion
                    }

                    #region 判断当前行和下一行还是可以合并的,如果可已合并并且不是最后一行就继续向下获取直到N+1行不相同则合并前面行
                    if (!(thisrow.SetEquals(nextrow) && newtrow != null))
                        isEndMerge = true;
                    else
                        mergeRow++;

                    #endregion

                    #region 进行合并
                    if (isEndMerge && mergeRow != 0)
                    {
                        //算出当前行
                        var firstRowIndex = rowIndex - mergeRow - 1;
                        var lastRowIndex = rowIndex - 1;
                        if (mergeDto.ListToNpoiMergeRange != null)
                        {
                            foreach (var range in mergeDto.ListToNpoiMergeRange)
                            {
                                int firstColIndex;
                                int lastColIndex;
                                switch (describe)
                                {
                                    case ListToNpoiEnum.English:
                                        firstColIndex = exportNames.FindIndex(s => s.Value.Value == range.StartName.ToLower());
                                        lastColIndex = exportNames.FindIndex(s => s.Value.Value == range.EndName.ToLower());
                                        break;
                                    case ListToNpoiEnum.Chinese:
                                        firstColIndex = exportNames.FindIndex(s => s.Value.Key == range.StartName.ToLower());
                                        lastColIndex = exportNames.FindIndex(s => s.Value.Key == range.EndName.ToLower());
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException(nameof(describe), describe, null);
                                }
                                SetValueRegionCell(sheet, firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
                            }
                        }
                        SetValueRegionCell(sheet, firstRowIndex, lastRowIndex, 0, 0);
                        mergeRow = 0;
                    }
                    #endregion
                }
                #endregion

                listrowIndex++;
            }
            #endregion

            #region 自适应宽度
            //for (var i = 0; i < exportNames.Count; i++)
            //    sheet.AutoSizeColumn(i, true);
            #endregion

            #region 写入流
            var ms = new MemoryStream();
            workbook.Write(ms);
            var fl = new FileInfo(fileName);
            if (fl.Directory != null && !fl.Directory.Exists)
                fl.Directory.Create();
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                var data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
            #endregion
            return new ResultStateDto(true);
        }
        #endregion

        #region DataTable To Excel
        /// <summary>
        /// Datas the table to npoi.
        /// </summary>
        /// <param name="datatable">DataTable</param>
        /// <param name="fileName">文件位置</param>
        /// <param name="isShowindex">是否有序号</param>
        /// <param name="mergeDto">合并参数</param>
        /// <param name="type">类型</param>
        /// <returns>
        /// 返回值：ResultStateDto
        /// </returns>
        /// 创建者：万浩
        /// 创建日期：2017/10/27 10:31
        /// 修改者：
        /// 修改时间：
        /// ----------------------------------------------------------------------------------------
        public static ResultStateDto DataTableToNpoi(DataTable datatable, string fileName, bool isShowindex = true, ListToNpoiMergeDto mergeDto = null, string type = null)
        {
            #region Excel 组装
            var workbook = new XSSFWorkbook();
            #region 颜色
            ICellStyle color1 = workbook.CreateCellStyle();
            color1.FillForegroundColor = HSSFColor.Red.Index;
            color1.FillPattern = FillPattern.SolidForeground;
            ICellStyle color2 = workbook.CreateCellStyle();
            color2.FillForegroundColor = HSSFColor.Yellow.Index;
            color2.FillPattern = FillPattern.SolidForeground;
            ICellStyle color3 = workbook.CreateCellStyle();
            color3.FillForegroundColor = HSSFColor.Blue.Index;
            color3.FillPattern = FillPattern.SolidForeground;
            #endregion
            var sheet = workbook.CreateSheet();
            if (isShowindex)
            {
                var Col = datatable.Columns.Add("序号", typeof(int));
                Col.SetOrdinal(0);
            }
            var columnsNameList = (from DataColumn item in datatable.Columns select item.ColumnName).ToList();
            //行索引
            var rowIndex = 0;
            //列总数
            var row = sheet.CreateRow(rowIndex++);
            //创建excel Title
            for (var i = 0; i < datatable.Columns.Count; i++)
            {
                #region 时间类型长度加宽
                if (i != 0)
                {
                    if (new List<object> { typeof(DateTime), typeof(DateTime?) }.Contains(datatable.Columns[i].GetType()))
                        sheet.SetColumnWidth(i, 3500);
                }
                #endregion
                row.CreateCell(i, CellType.String).SetCellValue(datatable.Columns[i].ColumnName);

                if (type == "SKU销量报表")
                {
                    if (datatable.Columns[i].ColumnName.Contains("销量"))
                    {
                        row.GetCell(i).CellStyle = color1;
                    }
                    else if (datatable.Columns[i].ColumnName.Contains("销售额"))
                    {
                        row.GetCell(i).CellStyle = color2;
                    }
                    else if (datatable.Columns[i].ColumnName.Contains("利润"))
                    {
                        row.GetCell(i).CellStyle = color3;
                    }
                }
            }
            #endregion

            #region 时间样式
            //时间样式
            var styleDateTime = workbook.CreateCellStyle();
            var format = workbook.CreateDataFormat();
            styleDateTime.DataFormat = format.GetFormat("yyyy-MM-dd");
            styleDateTime.VerticalAlignment = VerticalAlignment.Center;//垂直对齐(默认应该为center，如果center无效则用justify)
            styleDateTime.Alignment = HorizontalAlignment.Center;//水平对齐
            #endregion

            #region 普通样式
            //居中样式
            var cellstyleCenter = workbook.CreateCellStyle();//设置垂直居中格式
            cellstyleCenter.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            #endregion

            #region 创建cell  并且合并
            // DataTable  Row行
            var dataRowIndex = 0;
            // 合并
            var mergeRow = 0;

            for (var i = 0; i < datatable.Rows.Count; i++)
            {
                var mergeRowIndex = 1;
                //创建内容行 
                row = sheet.CreateRow(rowIndex++);

                #region 填充一行
                for (var j = 0; j < datatable.Columns.Count; j++)
                {
                    //创建单元格
                    var cell = row.CreateCell(j);
                    //默认字符串格式
                    var ctype = CellType.String;

                    if (isShowindex)
                    {
                        if (j == 0)
                        {
                            row.CreateCell(j, CellType.Numeric).SetCellValue(rowIndex - 1);
                            continue;
                        }
                    }

                    #region 列判断类型并赋值

                    var value = datatable.Rows[i][datatable.Columns[j].ColumnName];
                    var valueType = datatable.Columns[j].DataType;
                    if (valueType == typeof(double))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((double)value);
                    }
                    else if (valueType == typeof(float))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((float)value);
                    }
                    else if (valueType == typeof(short))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((short)value);
                    }
                    else if (valueType == typeof(int))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((int)value);
                    }
                    else if (valueType == typeof(int?))
                    {
                        if (value != null)
                        {
                            ctype = CellType.Numeric;
                            cell.SetCellValue((int)value);
                        }
                    }
                    else if (valueType == typeof(long))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((long)value);
                    }
                    else if (valueType == typeof(decimal))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else if (valueType == typeof(bool))
                    {
                        ctype = CellType.Boolean;
                        cell.SetCellValue((bool)value);
                    }
                    else if (valueType == typeof(DateTime))
                    {
                        if ((DateTime)value != DateTime.MinValue)
                            cell.SetCellValue((DateTime)value);
                        cell.CellStyle = styleDateTime;
                        continue;
                    }
                    else if (valueType == typeof(DateTime?))
                    {
                        if (value != null)
                        {
                            if ((DateTime)value != DateTime.MinValue)
                                cell.SetCellValue((DateTime)value);
                            cell.CellStyle = styleDateTime;
                            continue;
                        }
                    }
                    else
                    {
                        if (value != null)
                        {
                            ctype = CellType.String;
                            cell.SetCellValue(value.ToString());
                        }
                    }
                    cell.SetCellType(ctype);
                    cell.CellStyle = cellstyleCenter;
                    #endregion
                }
                #endregion

                //DataTable只有一行则不会合并
                if (datatable.Rows.Count == 1)
                {
                    break;
                }

                #region 合并
                if (mergeDto != null && mergeDto.Contrast.Count != 0)
                {
                    //是否是最后一个合并行
                    bool isEndMerge;

                    #region 判断最后一行
                    if (dataRowIndex + mergeRowIndex >= datatable.Rows.Count)
                    {
                        mergeRowIndex = -1;
                    }
                    #endregion

                    //获取当前行
                    var thisDataTableRow = datatable.Rows[i];
                    //获取下一行
                    var newtrow = datatable.Rows[i + mergeRowIndex];

                    #region HasSet 比较 结果集
                    var thisrow = new HashSet<object>();
                    var nextrow = new HashSet<object>();
                    #endregion

                    #region 合并行数据进行Hash对比
                    foreach (var items in mergeDto.Contrast)
                    {
                        thisrow.Add(thisDataTableRow[items]);
                        nextrow.Add(newtrow[items]);
                    }
                    #endregion

                    #region 判断当前行和下一行还是可以合并的,如果可已合并并且不是最后一行就继续向下获取直到N+1行不相同则合并前面行
                    if (thisrow.SetEquals(nextrow))
                    {
                        mergeRow++;
                        isEndMerge = dataRowIndex == datatable.Rows.Count - 1;
                    }
                    else
                    {
                        isEndMerge = true;
                    }
                    #endregion

                    #region 进行合并
                    if (isEndMerge && mergeRow != 0)
                    {
                        var firstRowIndex = rowIndex - mergeRow - (dataRowIndex == datatable.Rows.Count - 1 ? 0 : 1);
                        var lastRowIndex = (rowIndex - 1) - (dataRowIndex == datatable.Rows.Count - 1 ? 0 : 1);
                        if (mergeDto.ListToNpoiMergeRange != null)
                        {
                            foreach (var range in mergeDto.ListToNpoiMergeRange)
                            {
                                var firstColIndex = columnsNameList
                                    .FindIndex(s => String.Equals(s, range.StartName, StringComparison.CurrentCultureIgnoreCase));
                                var lastColIndex = columnsNameList
                                    .FindIndex(s => String.Equals(s, range.EndName, StringComparison.CurrentCultureIgnoreCase));
                                SetValueRegionCell(sheet, firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
                            }
                        }
                        SetValueRegionCell(sheet, firstRowIndex, lastRowIndex, 0, 0);
                        mergeRow = 0;
                    }
                    #endregion
                }
                #endregion
                dataRowIndex++;
            }
            #endregion

            #region 自适应宽度
            //for (var i = 0; i < columnsNameList.Count; i++)
            //    sheet.AutoSizeColumn(i, true);
            #endregion

            #region 写入流
            var ms = new MemoryStream();
            workbook.Write(ms);
            var fl = new FileInfo(fileName);
            if (fl.Directory != null && !fl.Directory.Exists)
                fl.Directory.Create();
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                var data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
            #endregion
            return new ResultStateDto(true);
        }
        #endregion


        /// <summary>
        /// 同一个数据源，分组多个Sheet，生成一个Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fun"></param>
        /// <param name="list"></param>
        /// <param name="dic"></param>
        /// <param name="fileName"></param>
        /// <param name="columnAll"></param>
        /// <param name="describe"></param>
        /// <param name="isShowindex"></param>
        /// <param name="mergeDto"></param>
        /// <returns></returns>
        public static ResultStateDto ListToSheets<T>(Func<List<T>, object, List<T>> fun, List<T> list, Dictionary<string, string> dic, string fileName, List<string> columnAll = null, ListToNpoiEnum describe = ListToNpoiEnum.English, bool isShowindex = true, ListToNpoiMergeDto mergeDto = null)
        {
            //Key{ Key 类属性 Value Index 位置 }  属性  Value { Key : description Value :  Name}
            var exportNames = new List<KeyValue<KeyValue<PropertyInfo, int>, KeyValue<string, string>>>();

            var colIndex = 0;

            #region 获取T所有属性 名称  描述  类型  下标等等
            //获取所有的自定义列
            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                if (propertyInfo.GetCustomAttributes(typeof(NoneExportAttribute), false).Length > 0)
                    continue;
                var descAttrs = propertyInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
                var description = string.Empty;
                if (descAttrs.Length > 0)
                {
                    var des = (DescriptionAttribute)descAttrs[0];
                    description = des.Description;
                }
                string name;
                var exportAttr = propertyInfo.GetCustomAttributes(typeof(ExportAttribute), false);
                if (exportAttr.Length <= 0)
                {
                    name = propertyInfo.Name;
                }
                else
                {
                    var export = (ExportAttribute)exportAttr[0];
                    name = export.HeaderName ?? propertyInfo.Name;
                }
                exportNames.Add(new KeyValue<KeyValue<PropertyInfo, int>, KeyValue<string, string>>(new KeyValue<PropertyInfo, int>(propertyInfo, colIndex),
                    new KeyValue<string, string>(description, name.ToLower())));
                colIndex++;
            }
            #endregion

            #region 过滤需要导出的字段
            if (columnAll != null)
            {
                switch (describe)
                {
                    case ListToNpoiEnum.English:
                        exportNames = columnAll.Select(items => exportNames.FirstOrDefault(s => String.Equals(s.Key.Key.Name, items, StringComparison.CurrentCultureIgnoreCase))).ToList();
                        break;
                    case ListToNpoiEnum.Chinese:
                        exportNames = columnAll.Select(items => exportNames.FirstOrDefault(s => s.Value.Key.ToLower() == items.ToLower())).ToList();
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(describe), describe, null);
                }
            }
            #endregion

            #region 是否支持序号
            if (isShowindex)
            {
                exportNames.Insert(0, new KeyValue<KeyValue<PropertyInfo, int>, KeyValue<string, string>>(null, new KeyValue<string, string>("序号", "序号")));
            }
            #endregion

            #region Excel组装
            var workbook = new XSSFWorkbook();

            foreach (var item in dic)
            {
                var sheetList = fun(list, item.Value);
                CreateSheet(item.Key, ref workbook, sheetList, exportNames, columnAll, describe, isShowindex, mergeDto);
            }


            #region 写入流
            var ms = new MemoryStream();
            workbook.Write(ms);
            var fl = new FileInfo(fileName);
            if (fl.Directory != null && !fl.Directory.Exists)
                fl.Directory.Create();
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                var data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
            #endregion
            return new ResultStateDto(true);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="workbook"></param>
        /// <param name="exportNames"></param>
        /// <param name="list"></param>
        /// <param name="columnAll"></param>
        /// <param name="describe"></param>
        /// <param name="isShowindex"></param>
        /// <param name="mergeDto"></param>
        public static void CreateSheet<T>(string sheetName, ref XSSFWorkbook workbook, List<T> list, List<KeyValue<KeyValue<PropertyInfo, int>, KeyValue<string, string>>> exportNames,
             List<string> columnAll = null, ListToNpoiEnum describe = ListToNpoiEnum.English, bool isShowindex = true, ListToNpoiMergeDto mergeDto = null)
        {
            var sheet = workbook.CreateSheet(sheetName);
            //行索引
            var rowIndex = 0;

            #region 添加标题列
            //列总数
            var row = sheet.CreateRow(rowIndex++);
            for (var i = 0; i < exportNames.Count; i++)
            {
                #region 时间类型长度加宽
                if (i != 0)
                {
                    if (new List<object> { typeof(DateTime), typeof(DateTime?) }.Contains(exportNames[i].Key.Key.PropertyType))
                    {
                        sheet.SetColumnWidth(i, 3500);
                    }
                }
                #endregion

                var name = !string.IsNullOrEmpty(exportNames[i].Value.Key) ? exportNames[i].Value.Key : exportNames[i].Value.Value;
                row.CreateCell(i, CellType.String).SetCellValue(name);
            }
            #endregion

            #region 时间样式
            //时间样式
            var styleDateTime = workbook.CreateCellStyle();
            var format = workbook.CreateDataFormat();
            styleDateTime.DataFormat = format.GetFormat("yyyy-MM-dd");
            styleDateTime.VerticalAlignment = VerticalAlignment.Center;//垂直对齐(默认应该为center，如果center无效则用justify)
            styleDateTime.Alignment = HorizontalAlignment.Center;//水平对齐
            #endregion

            #region 普通样式
            //居中样式
            var cellstyleCenter = workbook.CreateCellStyle();//设置垂直居中格式
            cellstyleCenter.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            #endregion

            //合并行数
            var listrowIndex = 0;
            var mergeRow = 0;
            foreach (var item in list)
            {
                row = sheet.CreateRow(rowIndex++); //创建内容行    
                #region 填充一行
                for (var i = 0; i < exportNames.Count; i++)
                {
                    //创建单元格
                    var cell = row.CreateCell(i);
                    //默认字符串格式
                    var ctype = CellType.String;
                    #region 是否开启索引号
                    if (isShowindex)
                    {
                        if (i == 0)
                        {
                            row.CreateCell(i, CellType.Numeric).SetCellValue(rowIndex - 1);
                            continue;
                        }
                    }
                    #endregion
                    #region 列判断类型并赋值
                    var attr = item.GetAttributesArray();
                    if (exportNames[i].Key.Key.PropertyType == typeof(double))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((double)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(float))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((float)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(short))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((short)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(int))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((int)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(int?))
                    {
                        if (attr[exportNames[i].Value.Value] != null)
                        {
                            ctype = CellType.Numeric;
                            cell.SetCellValue((int)attr[exportNames[i].Value.Value]);
                        }
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(long))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue((long)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(decimal))
                    {
                        ctype = CellType.Numeric;
                        cell.SetCellValue(Convert.ToDouble(attr[exportNames[i].Value.Value]));
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(bool))
                    {
                        ctype = CellType.Boolean;
                        cell.SetCellValue((bool)attr[exportNames[i].Value.Value]);
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(DateTime))
                    {
                        if ((DateTime)attr[exportNames[i].Value.Value] != DateTime.MinValue)
                            cell.SetCellValue((DateTime)attr[exportNames[i].Value.Value]);
                        cell.CellStyle = styleDateTime;
                        continue;
                    }
                    else if (exportNames[i].Key.Key.PropertyType == typeof(DateTime?))
                    {
                        if (attr[exportNames[i].Value.Value] != null)
                        {
                            if ((DateTime)attr[exportNames[i].Value.Value] != DateTime.MinValue)
                                cell.SetCellValue((DateTime)attr[exportNames[i].Value.Value]);
                            cell.CellStyle = styleDateTime;
                            continue;
                        }
                    }
                    else
                    {
                        if (attr[exportNames[i].Value.Value] != null)
                        {
                            ctype = CellType.String;
                            cell.SetCellValue(attr[exportNames[i].Value.Value].ToString());
                        }
                    }
                    cell.SetCellType(ctype);
                    cell.CellStyle = cellstyleCenter;
                    #endregion
                }
                #endregion

                if (list.Count == 1)
                    break;
                #region 合并
                if (mergeDto != null && mergeDto.Contrast.Count != 0)
                {
                    //是否是最后一个合并行
                    bool isEndMerge = false;
                    //获取当前行
                    var trow = item.GetAttributesArray();
                    //获取下一行
                    Dictionary<string, object> newtrow = null;
                    #region 判断最后一行
                    if (listrowIndex < list.Count - 1)
                        newtrow = list[listrowIndex + 1].GetAttributesArray();
                    #endregion
                    #region HasSet 比较 结果集
                    var thisrow = new HashSet<object>();
                    var nextrow = new HashSet<object>();
                    #endregion

                    if (newtrow != null)
                    {
                        #region 判断标准填充结果集 需要选择中英文
                        foreach (var items in mergeDto.Contrast)
                        {
                            switch (describe)
                            {
                                case ListToNpoiEnum.English:
                                    thisrow.Add(trow[items.ToLower()]);
                                    nextrow.Add(newtrow[items.ToLower()]);
                                    break;
                                case ListToNpoiEnum.Chinese:
                                    thisrow.Add(trow[exportNames.Single(s => String.Equals(s.Value.Key, items.ToLower(), StringComparison.CurrentCultureIgnoreCase)).Value.Value]);
                                    nextrow.Add(newtrow[exportNames.Single(s => String.Equals(s.Value.Key, items.ToLower(), StringComparison.CurrentCultureIgnoreCase)).Value.Value]);
                                    break;
                                default:
                                    throw new ArgumentOutOfRangeException(nameof(describe), describe, null);
                            }
                        }
                        #endregion
                    }

                    #region 判断当前行和下一行还是可以合并的,如果可已合并并且不是最后一行就继续向下获取直到N+1行不相同则合并前面行
                    if (!(thisrow.SetEquals(nextrow) && newtrow != null))
                        isEndMerge = true;
                    else
                        mergeRow++;

                    #endregion

                    #region 进行合并
                    if (isEndMerge && mergeRow != 0)
                    {
                        //算出当前行
                        var firstRowIndex = rowIndex - mergeRow - 1;
                        var lastRowIndex = rowIndex - 1;
                        if (mergeDto.ListToNpoiMergeRange != null)
                        {
                            foreach (var range in mergeDto.ListToNpoiMergeRange)
                            {
                                int firstColIndex;
                                int lastColIndex;
                                switch (describe)
                                {
                                    case ListToNpoiEnum.English:
                                        firstColIndex = exportNames.FindIndex(s => s.Value.Value == range.StartName.ToLower());
                                        lastColIndex = exportNames.FindIndex(s => s.Value.Value == range.EndName.ToLower());
                                        break;
                                    case ListToNpoiEnum.Chinese:
                                        firstColIndex = exportNames.FindIndex(s => s.Value.Key == range.StartName.ToLower());
                                        lastColIndex = exportNames.FindIndex(s => s.Value.Key == range.EndName.ToLower());
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException(nameof(describe), describe, null);
                                }
                                SetValueRegionCell(sheet, firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
                            }
                        }
                        SetValueRegionCell(sheet, firstRowIndex, lastRowIndex, 0, 0);
                        mergeRow = 0;
                    }
                    #endregion
                }
                #endregion

                listrowIndex++;
            }
            #endregion

            #region 自适应宽度
            //for (var i = 0; i < exportNames.Count; i++)
            //    sheet.AutoSizeColumn(i, true);
            #endregion
        }
        /// <summary>
        /// 设置合并单元格
        /// </summary>
        /// <param name="sheet">表格sheet对象</param>
        /// <param name="firstRow">开始行索引</param>
        /// <param name="lastRow">结束行索引</param>
        /// <param name="firstCol">开始列索引</param>
        /// <param name="lastCol">结束列索引</param>
        /// <param name="style"></param>
        public static void SetValueRegionCell(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, BorderStyle style = BorderStyle.None)
        {
            var region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            sheet.AddMergedRegion(region);
            if (style != BorderStyle.None)
            {
                ((HSSFSheet)sheet).SetBorderBottomOfRegion(region, style, HSSFColor.Black.Index);
            }
        }
    }


    public static class ListToNpoi_HelperTools
    {
        /// <summary>
        /// 获取所有排序条件 Time 2017-4-11
        /// </summary>
        /// <param name="class"></param>
        /// <returns></returns>
        public static Dictionary<string, object> GetAttributesArray(this object @class)
        {
            if (@class == null)
                return null;
            var t = @class.GetType();//获得该类的Type
            var result = new Dictionary<string, object>();
            //再用Type.GetProperties获得PropertyInfo[],然后就可以用foreach 遍历了
            var tlist = t.GetProperties();
            for (var i = 0; i < tlist.Length; i++)
            {
                if (result.ContainsKey(tlist[i].Name.ToLower()))
                    continue;
                var value = tlist[i].GetValue(@class, null);//用pi.GetValue获得值
                result.Add(tlist[i].Name.ToLower(), value);
            }
            return result;
        }
    }

    /// <summary>
    /// NPOI 合并实体
    /// </summary>
    public class ListToNpoiMergeDto
    {
        /// <summary>
        /// 数据对比依据
        /// </summary>
        public List<string> Contrast
        {
            get; set;
        }
        /// <summary>
        /// 是否开启合并列
        /// </summary>
        public bool IsMergeCol
        {
            get; set;
        }
        /// <summary>
        /// 合并列范围
        /// </summary>
        public List<ListToNpoiMergeRange> ListToNpoiMergeRange
        {
            get; set;
        }
    }

    /// <summary>
    /// 导出列处理 Time 2017-5-8
    /// </summary>
    public class ListToNpoiColumnList
    {
        /// <summary>
        /// 导出列
        /// </summary>
        public Dictionary<string, bool> ColumnAll
        {
            get; set;
        }
        /// <summary>
        /// 使用列名称匹配
        /// </summary>
        public ListToNpoiHelper.ListToNpoiEnum Describe
        {
            get; set;
        }
    }

    /// <summary>
    /// 合并范围
    /// </summary>
    public class ListToNpoiMergeRange
    {
        /// <summary>
        /// 
        /// </summary>
        public ListToNpoiMergeRange()
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="startName"></param>
        /// <param name="endName"></param>
        public ListToNpoiMergeRange(string startName, string endName)
        {
            StartName = startName;
            EndName = endName;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="startName"></param>
        public ListToNpoiMergeRange(string startName)
        {
            StartName = startName;
            EndName = startName;
        }
        /// <summary>
        /// 
        /// </summary>
        public string StartName
        {
            get; set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string EndName
        {
            get; set;
        }
    }
}
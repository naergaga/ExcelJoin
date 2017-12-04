using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelJoin.Models;
using OfficeOpenXml;

namespace ExcelJoin.Providers.Epplus
{
    public class SheetProvider : ISheetProvider
    {
        private ExcelWorksheet sheet;
        private bool headTitle = false;
        private bool detectColSpan = false;

        public SheetProvider(ExcelWorksheet sheet, bool headTitle = false, bool detectColSpan=true)
        {
            this.sheet = sheet;
            this.headTitle = headTitle;
            this.detectColSpan = detectColSpan;
        }

        public Sheet GetSimple()
        {
            var sheetItem = new Sheet { Name = sheet.Name };
            return sheetItem;
        }

        public Sheet Get(int identityIndex)
        {
            var sheetItem = new Sheet { Name = sheet.Name, Rows = new List<Row>() };
            var colLength = sheet.Dimension.Columns;
            int rowIndex;
            if (detectColSpan || headTitle)
            {
                var a = DetectData(headTitle);
                rowIndex = a.dataIndex;
                sheetItem.Columns = a.titleList;
            }
            else { rowIndex = 1; }

            //获取每一行
            var rowNum = sheet.Dimension.Rows;
            for (; rowIndex <= rowNum; rowIndex++)
            {
                var rowItem = new Row { Identity = sheet.GetValue(rowIndex, identityIndex).ToString(), Data = new List<ColumnData>() };
                //获取每一列
                for (int colIndex = 1; colIndex <= colLength; colIndex++)
                {
                    var d = new ColumnData { Index = colIndex };
                    d.Value = sheet.GetValue(rowIndex, colIndex);
                    rowItem.Data.Add(d);
                }

                sheetItem.Rows.Add(rowItem);
            }
            return sheetItem;
        }

        /// <summary>
        /// 从第一行开始，当有一行 每一列都有数据时 返回这一行Index
        /// 如果有列名，返回列名集合
        /// </summary>
        /// <returns></returns>
        private (int dataIndex, List<Column> titleList) DetectData(bool headTitle = false)
        {
            List<Column> titleList = headTitle ? new List<Column>() : null;
            for (int ri = 1; ri <= sheet.Dimension.Rows; ri++)
            {
                var fullRow = true;
                for (int ci = 1; ci <= sheet.Dimension.Columns; ci++)
                {
                    var cellValue = sheet.GetValue(ri, ci);
                    if (cellValue == null)
                    {
                        //标记这行不行，退出列循环
                        fullRow = false;
                        break;
                    }
                    //cellValue不为null，有列名
                    titleList?.Add(new Column { Name = cellValue.ToString() });
                }
                if (fullRow)
                {
                    //如果有列名，index 前进1行
                    if (headTitle) return (ri + 1, titleList);
                    return (ri, null);
                }
                titleList?.Clear();
            }
            return (0, null); //没有完整的一行
        }
    }
}

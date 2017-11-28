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

        public SheetProvider(ExcelWorksheet sheet, bool headTitle = false)
        {
            this.sheet = sheet;
            this.headTitle = headTitle;
        }

        public Sheet GetSimple()
        {
            var sheetItem = new Sheet { Name = sheet.Name };
            if (!headTitle)
            {
                return sheetItem;
            }

            sheetItem.Columns = new List<Column>();
            var colLength = sheet.Dimension.Columns;
            for (int i = 1; i <= colLength; i++)
            {
                var name = sheet.Cells[1, i].Value.ToString();
                sheetItem.Columns.Add(new Column { Name = name });
            }

            return sheetItem;
        }

        public Sheet Get(int identityIndex)
        {
            var sheetItem = new Sheet { Name = sheet.Name,Rows = new List<Row>() };
            var colLength = sheet.Dimension.Columns;
            int rowIndex;
            if (headTitle)
            {
                sheetItem.Columns = new List<Column>();
                for (int i = 1; i <= colLength; i++)
                {
                    var name = sheet.Cells[1, i].Value.ToString();
                    sheetItem.Columns.Add(new Column { Name = name });
                }
                rowIndex = 2;
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

    }
}

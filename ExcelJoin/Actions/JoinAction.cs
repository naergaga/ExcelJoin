using ExcelJoin.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin.Actions
{
    public class ExportConfig
    {
        public bool DateTimeIsHourMinute { get; set; }
    }

    public class JoinAction
    {
        public ExportConfig Config { get; set; } = new ExportConfig { DateTimeIsHourMinute = true };

        public JoinAction() { }

        public JoinAction(ExportConfig config)
        {
            Config = config;
        }

        class ResultRow
        {
            public List<ColumnData> Data1 { get; set; }
            public List<ColumnData> Data2 { get; set; }
        }

        public void Export(Sheet sheet1, Sheet sheet2, string outpath, string sheetName, bool headTitle1 = false, bool headTitle2 = false)
        {
            var newFile = new FileInfo(outpath);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(outpath);
            }
            //var query = from s1row in sheet1.Rows
            //            join s2row in sheet2.Rows
            //            on s1row.Identity equals s2row.Identity
            //            select new { Identity = s1row.Identity, Data1 = s1row.Data, Data2 = s2row.Data };
            //TODO: 重名
            //TODO: LEFT JOIN or INNER JOIN
            var query = sheet1.Rows.Select(s1row =>
            {
                return new ResultRow
                {
                    Data1 = s1row.Data,
                    Data2 = sheet2.Rows.FirstOrDefault(t => s1row.Identity==t.Identity)?.Data
                };
            });

            var list = query.ToList();

            using (var package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                int rowIndex = 1;
                //write column title
                int colIndex = 1;
                if (headTitle1)
                {
                    for (int i = 0; i < sheet1.Columns.Count; i++, colIndex++)
                    {
                        worksheet.Cells[rowIndex, colIndex].Value = sheet1.Columns[i].Name;
                    }
                }
                if (headTitle2)
                {
                    for (int i = 0; i < sheet2.Columns.Count; i++, colIndex++)
                    {
                        worksheet.Cells[rowIndex, colIndex].Value = sheet2.Columns[i].Name;
                    }
                }

                if (headTitle1 || headTitle2)
                {
                    rowIndex += 1;
                }

                //write column data
                for (int i = 0; i < list.Count; i++, rowIndex++)
                {
                    var rowData = list[i];
                    var data1 = rowData.Data1;
                    var data2 = rowData.Data2;
                    var col = 1;

                    for (int i2 = 0; i2 < data1.Count; i2++, col++)
                    {
                        SetCell(worksheet.Cells[rowIndex, col], data1[i2].Value);
                    }

                    if (data2 != null)
                        for (int i2 = 0; i2 < data2.Count; i2++, col++)
                        {
                            SetCell(worksheet.Cells[rowIndex, col], data2[i2].Value);
                        }

                }
                package.Save();
            }
        }

        private void SetCell(ExcelRange cellRange, object rawValue)
        {
            if (rawValue == null) return;
            cellRange.Value = rawValue;

            var type = rawValue.GetType();
            if (type == typeof(DateTime))
            {
                if (Config.DateTimeIsHourMinute)
                    cellRange.Style.Numberformat.Format = "h:mm";
                else
                    cellRange.Style.Numberformat.Format = "yyyy/m/d h:mm";
            }
        }

        //private object GetValue(Object rawValue)
        //{
        //    var type = rawValue.GetType();
        //    if (type == typeof(DateTime) && Config.DateTimeIsHourMinute )
        //    {
        //        return ((DateTime)rawValue).TimeOfDay;
        //    }
        //    return rawValue;
        //}
    }
}

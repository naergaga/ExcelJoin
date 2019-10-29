using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using ExcelJoin.Models;

namespace ExcelJoin.Providers.EDR
{
    public class EDRSheetProvider
    {
        public Sheet Get(string path, int sheetNum, int colNum)
        {
            Sheet sheet = null;
            int sheetIndex = sheetNum - 1, identityCol = colNum - 1;
            var stream = File.Open(path, FileMode.Open);
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var dataSet = reader.AsDataSet();
                var table = dataSet.Tables[sheetIndex];
                int index = 0;

                while (reader.NextResult() && !(index == sheetIndex))
                {
                    index++;
                }
                sheet = new Sheet { Name = reader.Name, Index = index++, Columns = new List<Column>(), Rows = new List<Row>() };
                bool titleGot = false;

                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    var rowItem = table.Rows[rowIndex];
                    if (titleGot)
                    {
                        Row row = new Row { Data = new List<ColumnData>() };
                        for (int col = 0; col < table.Columns.Count; col++)
                        {
                            var obj = rowItem[col];
                            row.Data.Add(new ColumnData { Index = col, Value = obj });
                        }
                        row.Identity = row.Data.FirstOrDefault(t => t.Index == identityCol).Value;
                        sheet.Rows.Add(row);
                        continue;
                    }
                    var fullLine = true;
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        var obj = rowItem[col];
                        if (obj != null)
                        {
                            sheet.Columns.Add(new Column { Name = obj.ToString() });
                        }
                        else
                        {
                            fullLine = false;
                            break;
                        }
                    }
                    if (fullLine)
                    {
                        titleGot = true;
                        continue;
                    }
                    sheet.Columns.Clear();
                }

            }
            stream.Dispose();
            return sheet;
        }

    }
}

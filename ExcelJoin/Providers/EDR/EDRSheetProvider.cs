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
            int sheetIndex = sheetNum - 1,identityCol = colNum-1;
            var stream = File.Open(path, FileMode.Open);
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int index = 0;

                while (reader.NextResult() && !(index == sheetIndex)) { index++; }
                sheet = new Sheet { Name = reader.Name, Index = index++, Columns = new List<Column>(), Rows = new List<Row>() };
                object obj = null;
                string valueStr = null;
                bool titleGot=false;
                while (reader.Read())
                {
                    if (titleGot)
                    {
                        Row row = new Row {Data = new List<ColumnData>() };
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            obj = reader.GetValue(i);
                            row.Data.Add(new ColumnData { Index = i, Value = obj });
                        }
                        row.Identity = row.Data.FirstOrDefault(t => t.Index == identityCol).Value;
                        sheet.Rows.Add(row);
                        continue;
                    }
                    var fullLine = true;
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if ((obj = reader.GetValue(i)) != null && (valueStr = obj as string) != null)
                        {
                            sheet.Columns.Add(new Column { Name = valueStr });
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

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
    public class EDRBookProvider
    {
        public Book GetSimple(string path)
        {
            Book book = null;

            var stream = File.Open(path, FileMode.Open);
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                book = new Book();
                book.Sheets = new List<Sheet>();
                int index = 1;
                do
                {
                    var sheet = new Sheet { Name = reader.Name, Index = index++, Columns = new List<Column>() };
                    object obj = null;
                    string valueStr = null;
                    while (reader.Read())
                    {
                        var fullLine = true;
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if ((obj = reader.GetValue(i)) != null && (valueStr=obj as string)!=null)
                            {
                                sheet.Columns.Add(new Column { Name = valueStr });
                            }else
                            {
                                fullLine = false;
                                break;
                            }
                        }
                        if (fullLine)
                        {
                            break;
                        }
                        sheet.Columns.Clear();
                    }
                    book.Sheets.Add(sheet);
                } while (reader.NextResult());
            }
            stream.Dispose();
            return book;
        }
    }
}

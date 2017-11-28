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
    public class JoinAction
    {
        public JoinAction()
        {

        }

        public void Export(Sheet sheet1, Sheet sheet2, string outpath, string sheetName)
        {
            var newFile = new FileInfo(outpath);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(outpath);
            }
            var query = from rs1 in sheet1.Rows
                        join rs2 in sheet2.Rows
                        on rs1.Identity equals rs2.Identity
                        select new { Identity=rs1.Identity, Data1=rs1.Data,Data2=rs2.Data,Columns1=sheet1.Columns,Columns2 =sheet2.Columns};
            var list = query.ToList();

            using (var package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                int rowIndex = 1;
                for (int i = 0; i < list.Count; i++,rowIndex++)
                {
                    var rowData = list[i];
                    var data1 = rowData.Data1;
                    var data2 = rowData.Data2;
                    var col = 1;
                    
                    for (int i2 = 0; i2 < data1.Count; i2++,col++)
                    {
                        worksheet.Cells[rowIndex, col].Value = data1[i2].Value;
                    }

                    for (int i2 = 0; i2 < data2.Count; i2++, col++)
                    {
                        worksheet.Cells[rowIndex, col].Value = data2[i2].Value;
                    }

                }
                package.Save();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin.Models
{
    public class Sheet
    {
        public string Name { get; set; }
        public List<Row> Rows { get; set; }
    }
}

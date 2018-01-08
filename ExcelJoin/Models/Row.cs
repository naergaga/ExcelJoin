using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin.Models
{
    public class Row
    {
        public dynamic Identity { get; set; }
        public List<ColumnData> Data { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin
{
    public class MainWindowModel
    {
        public bool HeadTitle1 { get; set; }
        public bool HeadTitle2 { get; set; }
        public int ColumnIndex1 { get; set; } = 1;
        public int ColumnIndex2 { get; set; } = 1;
    }
}

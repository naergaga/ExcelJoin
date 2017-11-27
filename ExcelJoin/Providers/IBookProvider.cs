using ExcelJoin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelJoin.Providers
{
    public interface IBookProvider
    {
        Book GetSimple();
    }
}

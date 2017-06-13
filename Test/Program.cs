using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Lazy.IO.Excel;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var rows = ExcelService.Read(@"D:\Project.Begin\SVN\Project.RMMV\data-pre\属性.xlsx", "属性关系");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lazy.IO.Excel
{
    public static class ExcelService
    {
        /// <summary>
        /// return: Table[y][x]
        /// </summary>
        /// <param name="pathName"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataRowCollection Read(string pathName, string sheetName = "Sheet1")
        {
            string strConn = $"Data Source={pathName};";
            DataTable dt = new DataTable();

            FileInfo file = new FileInfo(pathName);
            if (!file.Exists) { throw new FileNotFoundException(pathName); }

            switch (file.Extension)
            {
                // https://www.connectionstrings.com
                // "HDR=No" means "Do not ignore caption line, or the result will not contain caption line in dt.Rows"
                case ".xls":
                    strConn += "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
                    break;
                case ".xlsx":
                    strConn += "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=No;IMEX=1'";
                    break;
                default:
                    strConn += "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Text;HDR=No;IMEX=1'";
                    break;
            }

            new OleDbDataAdapter(
                $"SELECT * FROM [{sheetName}$]",
                new OleDbConnection(strConn)).Fill(dt);

            return dt.Rows;
        }
    }
}

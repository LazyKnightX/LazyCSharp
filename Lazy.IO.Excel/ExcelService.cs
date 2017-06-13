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
        /// https://www.connectionstrings.com
        /// "HDR=No" means "Do not ignore caption line, or the result will not contain caption line in dt.Rows"
        /// </summary>
        private static string MakeConnectionString(string pathName)
        {
            string connectionString = $"Data Source={pathName};";

            FileInfo file = new FileInfo(pathName);
            if (!file.Exists) { throw new FileNotFoundException(pathName); }

            switch (file.Extension)
            {
                // https://www.connectionstrings.com
                // "HDR=No" means "Do not ignore caption line, or the result will not contain caption line in dt.Rows"
                case ".xls":
                    connectionString += "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
                    break;
                case ".xlsx":
                    connectionString += "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=No;IMEX=1'";
                    break;
                default:
                    connectionString += "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Text;HDR=No;IMEX=1'";
                    break;
            }

            return connectionString;
        }

        /// <summary>
        /// return: Table[row][col]
        /// </summary>
        /// <param name="pathName"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataRowCollection Read(string pathName, string sheetName = "Sheet1")
        {
            string strConn = MakeConnectionString(pathName);
            DataTable dt = new DataTable();

            new OleDbDataAdapter(
                $"SELECT * FROM [{sheetName}$]",
                new OleDbConnection(strConn)).Fill(dt);

            return dt.Rows;
        }

        public static Dictionary<string, object[][]> ReadAll(string pathName)
        {
            Dictionary<string, object[][]> sheets = new Dictionary<string, object[][]>();

            string strConn = MakeConnectionString(pathName);
            DataTable dt = new DataTable();

            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable excel = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            foreach (DataRow sheetInfo in excel.Rows)
            {
                string sheetName = sheetInfo["TABLE_NAME"].ToString().Replace("'", "");

                dt.Clear();
                new OleDbDataAdapter($"SELECT * FROM [{sheetName}]", conn).Fill(dt);
                string sheetClearName = sheetName.Remove(sheetName.Length - 1, 1); // Remove last "$"

                object[][] output = new object[dt.Rows.Count][];

                for (int y = 0; y < dt.Rows.Count; y++)
                {
                    DataRow row = dt.Rows[y];
                    output[y] = new object[dt.Rows[0].ItemArray.Length];
                    row.ItemArray.CopyTo(output[y], 0);
                }
                sheets.Add(sheetClearName, output);
            }

            return sheets;
        }
    }
}

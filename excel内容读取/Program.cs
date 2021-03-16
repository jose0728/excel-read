using OfficeOpenXml;
using System;
using System.Data;
using System.IO;

namespace excel内容读取
{
    class Program
    {
        static void Main(string[] args)
        {
            DataSet vSet = new DataSet();
            string path = System.Environment.CurrentDirectory + "\\upload\\工作簿1.xlsx";
            if (File.Exists(path))
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    ExcelPackage excel = new ExcelPackage(stream);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

                    if (worksheet.Dimension == null)
                    {
                        Console.WriteLine("excel内容为空");
                        return;
                    }

                    int maxColumns = worksheet.Dimension.End.Column;
                    int maxRows = worksheet.Dimension.End.Row;

                    DataTable vTable = new DataTable();

                    for (int i = 1; i <= maxColumns; i++)
                    {
                        var cell = worksheet.Cells[1, i].Value;
                        if (cell != null)
                        {
                            DataColumn column = new DataColumn(cell.ToString());
                            vTable.Columns.Add(column);
                        }
                    }

                    for (int i = 2; i <= maxRows; i++)
                    {
                        DataRow row = vTable.NewRow();
                        int columnIndex = 0;
                        for (int j = 1; j <= maxColumns; j++)
                        {
                            row[columnIndex++] = worksheet.Cells[i, j].Value;
                        }
                        vTable.Rows.Add(row);
                    }

                    vSet.Tables.Add(vTable);
                }

                var table = vSet.Tables[0];

                var colNames = table.Columns;
                foreach (var j in colNames)
                {
                    Console.Write($"{j.ToString()} ");
                }
                Console.WriteLine();

                for (int item = 0; item < table.Rows.Count; item++)
                {
                    var rowData = table.Rows[item].ItemArray;
                    foreach (var i in rowData)
                    {
                        Console.Write($"{i.ToString()} ");
                    }
                    Console.WriteLine();
                }
            }

        }
    }
}

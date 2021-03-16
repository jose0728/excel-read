using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace excel操作
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        #region 业务代码
        /// <summary>
        /// 读取excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="useHeaderRow"></param>
        /// <returns></returns>
        public DataSet ReadExcelContent(string path, bool useHeaderRow = true)
        {
            DataSet ds;
            var filePath = Path.GetFullPath(path);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到文件！");
            }

            var extension = Path.GetExtension(filePath).ToLower();
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IExcelDataReader reader = null;
                if (extension == ".xls")
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (extension == ".xlsx")
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                else if (extension == ".csv")
                {
                    reader = ExcelReaderFactory.CreateCsvReader(stream);
                }

                if (reader == null)
                    return null;

                using (reader)
                {
                    ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        UseColumnDataType = false,
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = useHeaderRow // 第一行包含列名
                        }
                    });
                }
            }
            return ds;
        }

        /// <summary>
        /// 读取excel方法2
        /// </summary>
        /// <param name="path"></param>
        /// <param name="notNullColums"></param>
        /// <param name="orgname"></param>
        /// <returns></returns>
        public DataSet ReadExcelContent2(string path, List<string> notNullColums, bool orgname = false)
        {
            DataSet ds = new DataSet();
            var filePath = Path.GetFullPath(path);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到文件！");
            }

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var copyStream = new MemoryStream();
                stream.CopyTo(copyStream);

                using (copyStream)
                {
                    ExcelPackage package = new ExcelPackage(copyStream);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    if (worksheet.Dimension == null)
                    {
                        return null;
                    }

                    int maxColumnNum = worksheet.Dimension.End.Column;//最大列
                    int maxRowNum = worksheet.Dimension.Rows;//最大行

                    int minColumnNum = worksheet.Dimension.Start.Column;//最小列
                    int minRowNum = worksheet.Dimension.Start.Row;//最小行

                    DataTable vTable = new DataTable();

                    List<int> nullColumnNums = new List<int>();
                    for (int j = 1; j <= maxColumnNum; j++)
                    {
                        var cellName = worksheet.Cells[1, j].Value;
                        if (cellName != null)
                        {
                            var tempColumnName = cellName.ToString().Replace("\n", "").Trim();

                            var tempColumn = tempColumnName;
                            if (!orgname)
                            {
                                tempColumn = tempColumnName.TrimStart('*');
                            }

                            if (tempColumnName.StartsWith("*"))
                            {
                                notNullColums.Add(tempColumn);
                            }

                            DataColumn vColumn = new DataColumn(tempColumn, typeof(string));
                            vTable.Columns.Add(vColumn);
                        }
                        else
                        {
                            nullColumnNums.Add(j);
                        }
                    }

                    // excel行数是从1开始的
                    for (int n = 2; n <= maxRowNum; n++)
                    {
                        DataRow vRow = vTable.NewRow();
                        int rowIndex = 0;

                        for (int m = 1; m <= maxColumnNum; m++)
                        {
                            if (!nullColumnNums.Contains(m))
                            {
                                vRow[rowIndex] = worksheet.Cells[n, m].Value;
                                rowIndex++;
                            }
                        }
                        vTable.Rows.Add(vRow);
                    }
                    ds.Tables.Add(vTable);
                }
            }
            return ds;
        }
        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            if ((FileUpload1.PostedFile != null) && (FileUpload1.PostedFile.ContentLength > 0))
            {
                string fn = System.IO.Path.GetFileName(FileUpload1.PostedFile.FileName);
                string SaveLocation = Server.MapPath("upload") + "\\" + fn;
                try
                {
                    FileUpload1.PostedFile.SaveAs(SaveLocation);
                    FileUploadStatus.Text = "The file has been uploaded.";
                    ReadExcelContent2(SaveLocation, new List<string>());
                    ReadExcelContent(SaveLocation);
                }
                catch (Exception ex)
                {
                    FileUploadStatus.Text = "Error: " + ex.Message;
                }
            }
            else
            {
                FileUploadStatus.Text = "Please select a file to upload.";
            }
        }

        protected void Button2_Click(object sender, EventArgs e)
        {

        }
    }
}
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateExcel
{
    class Program
    {
        static string ConnectionString;
        static void Main(string[] args)
        {
            ConnectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;


            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                using (SqlDataAdapter adapter = new SqlDataAdapter(string.Format("SELECT * FROM UsersDetails"), connection))
                {
                    DataTable data = new DataTable();
                    adapter.Fill(data);
                    createExcel(data);
                }
            }
        }

        public static void createExcel(DataTable dataTable)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook workBook;
            Microsoft.Office.Interop.Excel.Worksheet workSheet;
            Microsoft.Office.Interop.Excel.Range cellRange;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                workBook = excel.Workbooks.Add(Type.Missing);


                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                workSheet.Name = "UserDetails";

                int rowcount = 1;

                foreach (DataRow datarow in dataTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {

                        if (rowcount == 2)
                        {
                            workSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
                            workSheet.Cells.Font.Color = System.Drawing.Color.Black;

                        }

                        workSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                    }

                }

                // use this for styling
                cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowcount, dataTable.Columns.Count]];
                cellRange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = cellRange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                workBook.SaveAs(ConfigurationManager.AppSettings["ExcelSaveLocation"]);
                workBook.Close();
                excel.Quit();
                Console.WriteLine("Successfully Create Excel File");
            }
            catch (Exception ex)
            {
               Console.WriteLine(ex.Message);

            }
            finally
            {
                workSheet = null;
                cellRange = null;
                workBook = null;
            }
        }
    }
}

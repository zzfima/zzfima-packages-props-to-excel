using System;
using System.Data;

namespace Logic
{
    public class ConvertPropsFile
    {
        public void GenerateExcel(string packagesPropsPath, string destinationExcelPath)
        {
            Microsoft.Office.Interop.Excel.Range cellRange;

            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            var workBook = excel.Workbooks.Open(destinationExcelPath);
            var newWorksheet = workBook.Worksheets.Add();

            newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
            newWorksheet.Name = "Nugets List " + DateTime.Now.Millisecond;

            newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[1, 8]].Merge();
            newWorksheet.Cells[1, 1] = "Nuget Packages List";
            newWorksheet.Cells.Font.Size = 15;


            int rowcount = 2;

            var exportToExcel = ExportToExcel();

            foreach (DataRow datarow in exportToExcel.Rows)
            {
                rowcount += 1;
                for (int i = 1; i <= exportToExcel.Columns.Count; i++)
                {

                    if (rowcount == 3)
                    {
                        newWorksheet.Cells[2, i] = exportToExcel.Columns[i - 1].ColumnName;
                        newWorksheet.Cells.Font.Color = System.Drawing.Color.Black;
                    }

                    newWorksheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                    if (rowcount > 3)
                    {
                        if (i == exportToExcel.Columns.Count)
                        {
                            if (rowcount % 2 == 0)
                            {
                                cellRange = newWorksheet.Range[newWorksheet.Cells[rowcount, 1], newWorksheet.Cells[rowcount, exportToExcel.Columns.Count]];
                            }
                        }
                    }
                }
            }

            cellRange = newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[rowcount, exportToExcel.Columns.Count]];
            cellRange.EntireColumn.AutoFit();
            Microsoft.Office.Interop.Excel.Borders border = cellRange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            cellRange = newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[2, exportToExcel.Columns.Count]];

            workBook.Save();
            workBook.Close();
            excel.Quit();
        }

        private System.Data.DataTable ExportToExcel()
        {
            System.Data.DataTable table = new System.Data.DataTable();

            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Nuget Name", typeof(string));
            table.Columns.Add("Version", typeof(string));
            table.Columns.Add("License", typeof(string));

            table.Rows.Add(1, "KLA.FA.SecsSerializer", "1.2.3", "MIT");
            table.Rows.Add(2, "DependencyInjection", "1.2.3", "MIT");
            table.Rows.Add(3, "Microsoft.CodeAnalysis.NetAnalyzers", "1.2.3", "MIT");
            table.Rows.Add(4, "KLA.Infrastructure.KLogger", "1.2.3", "MIT");
            table.Rows.Add(5, "AutoFixture", "1.2.3", "MIT");
            table.Rows.Add(6, "AutoFixture.AutoMoq", "1.2.3", "MIT");

            return table;
        }
    }
}

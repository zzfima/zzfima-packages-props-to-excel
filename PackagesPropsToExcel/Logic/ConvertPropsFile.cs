using System;
using System.Data;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Logic
{
    public class ConvertPropsFile
    {
        public void GenerateExcel(string packagesPropsPath, string destinationExcelPath)
        {
            var exportToExcel = ExportToExcel(packagesPropsPath);

            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook workBook = null;
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                workBook = excel.Workbooks.Open(destinationExcelPath);

                var newWorksheet = workBook.Worksheets.Add();
                Microsoft.Office.Interop.Excel.Range cellRange;
                newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                newWorksheet.Name = "Nugets List " + DateTime.Now.Millisecond;

                newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[1, 8]].Merge();
                newWorksheet.Cells[1, 1] = "Nuget Packages List";
                newWorksheet.Cells.Font.Size = 15;


                int rowcount = 2;

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
            }
            finally
            {
                workBook?.Save();
                workBook?.Close();
                excel?.Quit();
            }
        }

        private System.Data.DataTable ExportToExcel(string packagesPropsPath)
        {
            System.Data.DataTable table = new System.Data.DataTable();

            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Nuget Name", typeof(string));
            table.Columns.Add("Version", typeof(string));
            table.Columns.Add("License", typeof(string));

            Project packagesProps = null;
            System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(Project)/*, xRoot*/);
            using (System.IO.StreamReader file = new System.IO.StreamReader(packagesPropsPath))
            {
                packagesProps = (Project)reader.Deserialize(file);
            }

            var cnt = 1;
            foreach (var packageReference in packagesProps.ItemGroup.PackageReferences)
            {
                table.Rows.Add(cnt++, packageReference.Update, packageReference.Version, "MIT");
            }

            return table;
        }
    }
}

using NuGet;
using System;
using System.Data;
using System.IO;
using System.Linq;
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
            var isFileMissing = false;
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook workBook = null;
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                if (File.Exists(destinationExcelPath))
                {
                    workBook = excel.Workbooks.Open(destinationExcelPath);
                    isFileMissing = false;
                }
                else
                {
                    workBook = excel.Workbooks.Add(Type.Missing);
                    isFileMissing = true;
                }
                
                var newWorksheet = workBook.Worksheets.Add();
                Microsoft.Office.Interop.Excel.Range cellRange;
                newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                newWorksheet.Name = "Nugets List " + workBook.Sheets.Count;

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

                        if (exportToExcel.Columns[i - 1].ColumnName.Equals("License URL"))
                            newWorksheet.Hyperlinks.Add(newWorksheet.Cells[rowcount, i],
                                datarow[i - 1].ToString(),
                                Type.Missing,
                                datarow[i - 1].ToString(),
                                datarow[i - 1].ToString());
                        else
                            newWorksheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == exportToExcel.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    cellRange = newWorksheet.Range[newWorksheet.Cells[rowcount, 1],
                                        newWorksheet.Cells[rowcount, exportToExcel.Columns.Count]];
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
                if (isFileMissing)
                    workBook?.SaveAs(destinationExcelPath);
                else
                    workBook?.Save();

                workBook?.Close();
                excel?.Quit();

                workBook = null;
                excel = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private System.Data.DataTable ExportToExcel(string packagesPropsPath)
        {
            System.Data.DataTable table = new System.Data.DataTable();

            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Nuget Name", typeof(string));
            table.Columns.Add("Version", typeof(string));
            table.Columns.Add("License URL", typeof(string));

            Project packagesProps = null;
            var reader = new XmlSerializer(typeof(Project));
            using (var file = new StreamReader(packagesPropsPath))
            {
                packagesProps = (Project)reader.Deserialize(file);
            }

            IPackageRepository repo = PackageRepositoryFactory.Default.CreateRepository("https://packages.nuget.org/api/v2");
            var cnt = 1;
            foreach (var packageReference in packagesProps.ItemGroup.PackageReferences)
            {
                var package = (from p in repo.FindPackagesById(packageReference.Update)
                               where p.Version == new SemanticVersion(packageReference.Version)
                               select p).FirstOrDefault();

                table.Rows.Add(cnt++, packageReference.Update, packageReference.Version, package?.LicenseUrl);
            }

            return table;
        }
    }
}

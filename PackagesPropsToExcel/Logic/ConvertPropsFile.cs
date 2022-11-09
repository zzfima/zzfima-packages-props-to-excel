using NuGet;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Logic
{
    public class ConvertPropsFile
    {
        public async Task GenerateExcel(string packagesPropsPath, string destinationExcelPath)
        {
            var isFileMissing = false;
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook workBook = null;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                if (packagesPropsPath == null || packagesPropsPath.Equals(string.Empty))
                {
                    throw new NullReferenceException("Packages.Props File can not be null or empty");
                }

                if (!File.Exists(packagesPropsPath))
                {
                    throw new FileNotFoundException("File " + packagesPropsPath + " not found");
                }

                if (File.Exists(destinationExcelPath))
                {
                    if (IsFileLocked(new FileInfo(destinationExcelPath)))
                    {
                        throw new AccessViolationException("File " + destinationExcelPath + " is opened, can not be modified");
                    }

                    workBook = excel.Workbooks.Open(destinationExcelPath);
                    isFileMissing = false;
                }
                else
                {
                    workBook = excel.Workbooks.Add(Type.Missing);
                    isFileMissing = true;
                }

                var exportToExcel = await RetrieveNugetsData(
                    packagesPropsPath,
                    "https://packages.nuget.org/api/v2",
                    "https://kla-cpg-nuget.adcorp.kla-tencor.com/nuget");

                await Task.Run(() =>
                {
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
                });
            }
            catch (Exception e)
            {
                throw e;
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

        private bool IsFileLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }

        private async Task<DataTable> RetrieveNugetsData(string packagesPropsPath, string mainRepoPath, string secondaryRepoPath)
        {
            var table = new DataTable();

            InitTableColumns(table);

            Project packagesProps = null;
            packagesProps = GetPackagesPropsProject(packagesPropsPath);


            var mainRepo = PackageRepositoryFactory.Default.CreateRepository(mainRepoPath);
            var secaondaryRepo = PackageRepositoryFactory.Default.CreateRepository(secondaryRepoPath);
            var cnt = 1;
            foreach (var packageReference in packagesProps.ItemGroup.PackageReferences)
            {
                var package = await Task.Run(() =>
                {
                    var p = GetIPackage(mainRepo, packageReference);
                    if (p == null)
                        p = GetIPackage(secaondaryRepo, packageReference);
                    return p;
                });

                var license = "not found";
                if (package != null && package.LicenseUrl != null)
                {
                    license = package.LicenseUrl.ToString();
                }

                table.Rows.Add(cnt++, packageReference.Update, packageReference.Version, license);
            }

            return table;
        }

        private static IPackage GetIPackage(IPackageRepository mainRepo, PackageReference packageReference)
        {
            return (from p in mainRepo.FindPackagesById(packageReference.Update)
                    where p.Version == new SemanticVersion(packageReference.Version)
                    select p).FirstOrDefault();
        }

        private static Project GetPackagesPropsProject(string packagesPropsPath)
        {
            Project packagesProps;
            var reader = new XmlSerializer(typeof(Project));
            using (var file = new StreamReader(packagesPropsPath))
            {
                packagesProps = (Project)reader.Deserialize(file);
            }

            return packagesProps;
        }

        private void InitTableColumns(DataTable table)
        {
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Nuget Name", typeof(string));
            table.Columns.Add("Version", typeof(string));
            table.Columns.Add("License URL", typeof(string));
        }
    }
}

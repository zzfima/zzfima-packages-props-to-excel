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

        private System.Data.DataTable ExportToExcel(string packagesPropsPath)
        {
            System.Data.DataTable table = new System.Data.DataTable();

            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Nuget Name", typeof(string));
            table.Columns.Add("Version", typeof(string));
            table.Columns.Add("License", typeof(string));

            //XmlRootAttribute xRoot = new XmlRootAttribute();
            //xRoot.ElementName = "ItemGroup";
            //xRoot.IsNullable = true;

            System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(Project)/*, xRoot*/);
            System.IO.StreamReader file = new System.IO.StreamReader(packagesPropsPath);
            Project packagesProps = (Project)reader.Deserialize(file);
            file.Close();
            //var packagesProps = File.OpenRead(packagesPropsPath);
            //var xml = XDocument.Load(packagesPropsPath);
            //XElement booksFromFile = XElement.Load(packagesPropsPath);
            //using (XmlReader reader = XmlReader.Create(packagesPropsPath))
            //{
            //    while (reader.Read())
            //    {
            //        if (reader.IsStartElement())
            //        {
            //            //return only when you have START tag  
            //            switch (reader.Name.ToString())
            //            {
            //                case "Name":
            //                    Console.WriteLine("Name of the Element is : " + reader.ReadString());
            //                    break;
            //                case "Location":
            //                    Console.WriteLine("Your Location is : " + reader.ReadString());
            //                    break;
            //            }
            //        }
            //        Console.WriteLine("");
            //    }
            //}


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

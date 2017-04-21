using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChangeHeaders
{
    class UpdateHeader
    {
        public static string Run(string FilePath, DirectoryInfo outputDir)
        {
            //FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
            //if (newFile.Exists)
            //{
            //    newFile.Delete();  // ensures we create a new workbook
            //    newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
            //}

            //Console.WriteLine("Reading column 2 of {0}", FilePath);
            //Console.WriteLine();

            var newFilePath = ConvertToXLSX(FilePath);
            FileInfo existingFile = new FileInfo(newFilePath);

            
            //byte[] file = File.ReadAllBytes(newFilePath);
            //using (MemoryStream existingFile = new MemoryStream(file))

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SUMMARY SHEET");

                //find the bottom-most header row
                int HeaderBotRow = 1;
                for (int i = 1; i <= 100; i++)
                {
                    var cellVal = worksheet.Cells[i, 1].Text;
                    if (cellVal == "Line #")
                    {
                        HeaderBotRow = i;
                        break;
                    }

                }

                // find the top of the header
                int HeaderTopRow = 1;
                for (int i = HeaderBotRow; i >= 1; i--)
                {
                    var cellVal = worksheet.Cells[i, 1].Text;
                    if (cellVal == "")
                    {
                        HeaderTopRow = i;

                    }
                    else if (i != HeaderBotRow)
                    {
                        break;
                    }

                }


                // unMerge header cells and replicate the value across
                var mergedCells = worksheet.MergedCells.ToList();
                foreach (var mC in mergedCells)
                {
                    var startCellRow = worksheet.Cells[mC].Start.Row;
                    //make sure we are only looking in the header
                    if(startCellRow<6 || startCellRow > 9){
                        continue;
                    }

                    var startCellCol = worksheet.Cells[mC].Start.Column;
                    var endCellCol = worksheet.Cells[mC].End.Column;

                    var txt = worksheet.Cells[startCellRow, startCellCol].Text;

                    worksheet.Cells[mC].Merge = false;
                    worksheet.Cells[mC].Value = txt;
                }
                //find end of table
                int lastCol = 1;
                for(int i=1; i<= 500; i++)
                {
                    if(i > 1 && worksheet.Cells[HeaderBotRow, i].Text == ""
                        && worksheet.Cells[HeaderBotRow, i+1].Text == "") //check for two blanks in a row
                    {
                        lastCol = i-1;
                        break;
                    }
                }

                //starting at the 2nd to bottom header row, copy cells to the right for "fake" merged fields
                var secHeaderBotRow = HeaderBotRow - 1;
                if (secHeaderBotRow >= 1)
                {
                    bool firstNonBlankHit = false;
                    for (int i=1; i <= lastCol; i++)
                    {
                        var cellVal = worksheet.Cells[secHeaderBotRow, i].Text;
                        if (!firstNonBlankHit && cellVal == "")
                        {

                        }
                        else
                        {
                            firstNonBlankHit = true;
                        }

                        if (firstNonBlankHit)
                        {
                            if(cellVal == "")
                            {
                                var prevCellVal = worksheet.Cells[secHeaderBotRow, i-1].Text;
                                worksheet.Cells[secHeaderBotRow, i].Value = prevCellVal;
                            }

                        }
                        
                    }
                }


                // trickle down header values to the Bottom header row
                for (int i = 1; i <= lastCol; i++)
                {
                    for (int j = HeaderBotRow; j >= HeaderTopRow; j--)
                    {
                        if (j - 1 >= HeaderTopRow && worksheet.Cells[j, i].Text == "" && worksheet.Cells[j - 1, i].Text != "")
                        {
                            worksheet.Cells[j, i].Value = worksheet.Cells[j - 1, i].Text;
                            worksheet.Cells[j - 1, i].Value = "";
                        }
                    }
                }

                // Insert a header row
                worksheet.InsertRow(HeaderBotRow + 1, 1);
                var newHeaderRow = HeaderBotRow + 1;

                //Fill header row with concatenated values of rows above
                for (int i = 1; i <= lastCol; i++)
                {
                    for (int j = HeaderBotRow; j >= HeaderTopRow; j--)
                    {
                        if (worksheet.Cells[j, i].Text == "")
                        {
                            break;
                        }
                        else
                        {
                            if (j == HeaderBotRow)
                            {
                                worksheet.Cells[newHeaderRow, i].Value = worksheet.Cells[j, i].Text;
                            }
                            else {
                                worksheet.Cells[newHeaderRow, i].Value = worksheet.Cells[j, i].Text + " " + worksheet.Cells[newHeaderRow, i].Text;
                            }
                        }
                    }
                }

                //remove everything above the new header row
                for (int i = 1; i < newHeaderRow; i++)
                {
                    worksheet.DeleteRow(1, 1);
                }
                worksheet.View.UnFreezePanes();
                worksheet.View.FreezePanes(1, 2);


                // output the formula in row 5
                //Console.WriteLine("\tCell({0},{1}).Formula={2}", 5, 3, worksheet.Cells[5, 3].Formula);
                //Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 5, 3, worksheet.Cells[5, 3].FormulaR1C1);

                package.Save();
            } // the using statement automatically calls Dispose() which closes the package.

            
            Console.WriteLine();
            Console.WriteLine("UpdateHeader complete");
            Console.WriteLine();

            return newFilePath;
        }

        public void ConvertToXLSXDir(String filesFolder)
        {
            var files = Directory.GetFiles(filesFolder);

            var app = new Microsoft.Office.Interop.Excel.Application();

            foreach (var file in files)
            {
                var wb = app.Workbooks.Open(file);
                wb.SaveAs(Filename: file + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                wb.Close();
            }
            app.Quit();
        }

        public static string ConvertToXLSX(String FilePath)
        {
            //byte[] file = File.ReadAllBytes(FilePath);
            //using (MemoryStream existingFile = new MemoryStream(file)) ;

            var app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false;
            var wb = app.Workbooks.Open(FilePath);
            wb.SaveAs(Filename: FilePath + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return FilePath + "x";
        }
    }
}

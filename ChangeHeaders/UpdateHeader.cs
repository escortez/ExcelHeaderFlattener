using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Packaging;
using Microsoft.Office.Interop.Excel;

namespace ChangeHeaders
{
    class UpdateHeader
    {




        public static string Run(string FilePath, DirectoryInfo outputDir)
        {

            var sheetNames = ConfigurationManager.AppSettings["sheets"].Split(',').Select(y=>y.Trim());

            string[] files = new string[] { };
            if (String.IsNullOrEmpty(FilePath))
            {
                string directoryPath = ConfigurationManager.AppSettings["directory"];
                string fileKeyword = "*" + ConfigurationManager.AppSettings["file_name_keyword"] + "*";
                files = Directory.GetFiles(directoryPath, fileKeyword, SearchOption.AllDirectories).Where(x=> !x.EndsWith("_mod.xlsx") ).ToArray();
            }
            else
            {
                files[0] = FilePath;
            }

            foreach (var fileName in files)
            {

                string newFilePath = fileName;
                //if (fileName.EndsWith(".xls"))
                //{ 
                    newFilePath = ConvertToXLSX(fileName);
                //} else if (fileName.EndsWith(".xlsx"))
                //{
                //    newFilePath = fileName.Replace(".xlxs", "_mod.xlsx");

                //}
                //else
                //{
                //    // don't proceed if the file wasn't xls or xlsx
                //    continue;
                //}

                

                FileInfo existingFile = new FileInfo(newFilePath);


                //byte[] file = File.ReadAllBytes(newFilePath);
                //using (MemoryStream existingFile = new MemoryStream(file))

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    var worksheets = package.Workbook.Worksheets.Where(x => sheetNames.Contains(x.Name));

                    foreach (var worksheet in worksheets)
                    {
                        //find the bottom-most header row
                        int HeaderBotRow = GetBotHeaderRow(worksheet, "Line #");

                        // find the top of the header
                        int HeaderTopRow = GetTopHeaderRow(worksheet, HeaderBotRow);

                        //find end of table
                        int lastCol = FindLastCol(worksheet, HeaderBotRow);

                        // unMerge header cells and replicate the value across
                        var mergedCells = worksheet.MergedCells.ToList();
                        foreach (var mC in mergedCells)
                        {
                            var startCellRow = worksheet.Cells[mC].Start.Row;
                            //make sure we are only looking in the header
                            if (startCellRow < HeaderTopRow || startCellRow > HeaderBotRow)
                            {
                                continue;
                            }

                            var startCellCol = worksheet.Cells[mC].Start.Column;
                            var endCellCol = worksheet.Cells[mC].End.Column;

                            var txt = worksheet.Cells[startCellRow, startCellCol].Text;

                            worksheet.Cells[mC].Merge = false;
                            worksheet.Cells[mC].Value = txt;
                        }


                        //starting at the 2nd to bottom header row, copy cells to the right for "fake" merged fields
                        var secHeaderBotRow = HeaderBotRow - 1;
                        if (secHeaderBotRow >= 1)
                        {
                            bool firstNonBlankHit = false;
                            for (int i = 1; i <= lastCol; i++)
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
                                    if (cellVal == "")
                                    {
                                        var prevCellVal = worksheet.Cells[secHeaderBotRow, i - 1].Text;
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


                        var newHeaderRow = 1;
                        if (worksheet.Name != "AUDIT SHEET")
                        {
                            //// Insert a header row
                            int rowAddCount = 1;
                            newHeaderRow = HeaderBotRow + 1;
                            worksheet.InsertRow(newHeaderRow, rowAddCount);
                        }

                        //using (var rng = worksheet.Cells[newHeaderRow,1,newHeaderRow,lastCol])
                        //{

                        //    rng.Style.Font.Bold = true;
                        //    rng.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        //    rng.Style.WrapText = true;
                        //    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        //    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    rng.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                        //}


                        //var copyrow = newHeaderRow + rowAddCount;
                        //for (var i = 0; i < rowAddCount; i++)
                        //{
                        //    var row = newHeaderRow + i;
                        //    worksheet.Cells[String.Format("{0}:{0}", copyrow)].Copy(worksheet.Cells[String.Format("{0}:{0}", row)]);
                        //    worksheet.Row(row).StyleID = worksheet.Row(copyrow).StyleID;
                        //}
                        //May not be needed but cant hurt
                        //worksheet.Cells.Worksheet.Workbook.Styles.UpdateXml();

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

                        worksheet.View.UnFreezePanes();
                        worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.None;
                        if (worksheet.Name != "AUDIT SHEET")
                        {
                            //remove everything above the new header row
                            for (int i = 1; i < newHeaderRow; i++)
                            {
                                worksheet.DeleteRow(1, 1);
                            }
                        }
                        else
                        {
                            //hack because there is a bug with deleting rows
                            worksheet.Cells[2, 1, HeaderBotRow, lastCol].Value = "";
                            //worksheet.DeleteRow(1, 1);

                        }


                        worksheet.View.FreezePanes(1, 2);

                    } // end loop through worksheets // the using statement automatically calls Dispose() which closes the package.
                    package.Save();

                } //end package using stmnt ( for a particular file)
            } //end files loop 
            
            Console.WriteLine();
            Console.WriteLine("UpdateHeader complete");
            Console.WriteLine();

            return string.Join(",", files);
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

            string newFileName = "";
            if (FilePath.EndsWith(".xls"))
            {
                newFileName = FilePath.Replace(".xls", "_mod.xlsx");
            }
            else if (FilePath.EndsWith(".xlsx"))
            {
                newFileName = FilePath.Replace(".xlsx", "_mod.xlsx");
            }
            
            wb.SaveAs(Filename: newFileName, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return newFileName;
        }

      

        public static int FindLastCol(ExcelWorksheet worksheet, int HeaderBotRow)
        {
            int lastCol = 1;
            for (int i = 1; i <= 500; i++)
            {
                if (i > 1 && worksheet.Cells[HeaderBotRow, i].Text == ""
                    && worksheet.Cells[HeaderBotRow, i + 1].Text == "") //check for two blanks in a row
                {
                    lastCol = i - 1;
                    break;
                }
            }
            return lastCol;
        }

        public static int GetBotHeaderRow(ExcelWorksheet worksheet, string guide)
        {
            int HeaderBotRow = 1;
            for (int i = 1; i <= 100; i++)
            {
                var cellVal = worksheet.Cells[i, 1].Text;
                if (cellVal == guide)
                {
                    HeaderBotRow = i;
                    break;
                }
            }
            return HeaderBotRow;
        }
        public static int GetTopHeaderRow(ExcelWorksheet worksheet, int HeaderBotRow)
        {
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

            return HeaderTopRow;
        }

    }
}

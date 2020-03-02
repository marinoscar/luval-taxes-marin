using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.taxes.app
{
    public class MergeExcelFiles
    {
        public static void DoMerge(IEnumerable<FileInfo> files, FileInfo output, int startColumn, int endColumn)
        {
            if (output.Exists) output.Delete();
            using (var outputPackage = new ExcelPackage(output))
            {
                using (var outputSheet = outputPackage.Workbook.Worksheets.Add("output"))
                {
                    var outputRowCount = 1;
                    foreach (var inputFile in files)
                    {
                        if (!inputFile.Exists) continue;
                        using (var excelPackage = new ExcelPackage(inputFile))
                        {
                            var rowCount = files.First() == inputFile ? 1 : 2;
                            using (var sheet = excelPackage.Workbook.Worksheets.First())
                            {
                                var isEof = string.IsNullOrWhiteSpace(Convert.ToString(sheet.Cells[rowCount, startColumn].Value));
                                while (!isEof)
                                {
                                    for (int i = startColumn; i < endColumn; i++)
                                    {
                                        outputSheet.Cells[outputRowCount, i].Value = sheet.Cells[rowCount, i].Value;
                                    }
                                    rowCount++;
                                    outputRowCount++;
                                    isEof = string.IsNullOrWhiteSpace(Convert.ToString(sheet.Cells[rowCount, startColumn].Value));
                                }
                            }
                        }
                    }
                    outputPackage.Save();
                }
            }
        }
    }
}

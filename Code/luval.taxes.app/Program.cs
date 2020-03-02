using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.taxes.app
{
    class Program
    {
        static void Main(string[] args)
        {
            var dir = new DirectoryInfo(@"C:\Users\ch489gt\Google Drive\Work\EY\Personal-Taxes-2019\Timesheets");
            var files = dir.GetFiles("*.xls*", SearchOption.AllDirectories)
                .Where(i => i.Name != "result.xlsx" && !i.Name.StartsWith("~")).ToList();
            MergeExcelFiles.DoMerge(files, new FileInfo(Path.Combine(dir.FullName, "result.xlsx")), 1, 35);
        }
    }
}

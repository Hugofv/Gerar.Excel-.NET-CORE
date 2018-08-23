
using System.IO;
using System.Net.Http.Headers;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Gerar.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string sWebRootFolder = @"/home/cogtive/Desktop/";
            string sFileName = @"demo.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.Add("Equipamento 1");

                //First add the headers
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Gender";
                worksheet.Cells[1, 4].Value = "Salary (in $)";

                worksheet.Cells[1, 1, 1, 4].Style.Font.Bold = true;
                worksheet.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

                //Add values
                worksheet.Cells["A2"].Value = 1000;
                worksheet.Cells["B2"].Value = "Jon";
                worksheet.Cells["C2"].Value = "M";
                worksheet.Cells["D2"].Value = 5000;

                worksheet.Cells["A3"].Value = 1001;
                worksheet.Cells["B3"].Value = "Graham";
                worksheet.Cells["C3"].Value = "M";
                worksheet.Cells["D3"].Value = 10000;

                worksheet.Cells["A4"].Value = 1002;
                worksheet.Cells["B4"].Value = "Jenny";
                worksheet.Cells["C4"].Value = "F";
                worksheet.Cells["D4"].Value = 5000;

                worksheet = workbook.Worksheets.Add("Equipamento 2");

                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Gender";
                worksheet.Cells[1, 4].Value = "Salary (in $)";

                worksheet.Cells[1, 1, 1, 4].Style.Font.Bold = true;
                worksheet.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

                //Add values
                worksheet.Cells["A2"].Value = 1000;
                worksheet.Cells["B2"].Value = "Jon";
                worksheet.Cells["C2"].Value = "M";
                worksheet.Cells["D2"].Value = 5000;

                worksheet.Cells["A3"].Value = 1001;
                worksheet.Cells["B3"].Value = "Graham";
                worksheet.Cells["C3"].Value = "M";
                worksheet.Cells["D3"].Value = 10000;

                worksheet.Cells["A4"].Value = 1002;
                worksheet.Cells["B4"].Value = "Jenny";
                worksheet.Cells["C4"].Value = "F";
                worksheet.Cells["D4"].Value = 5000;

                package.SaveAs(file);
            }
        }
    }
}

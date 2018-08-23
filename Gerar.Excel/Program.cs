using System;
using System.IO;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Gerar.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            const string folderRoot = @"d:\";
            const string fileName = @"demo.xlsx";
            var file = new FileInfo(Path.Combine(folderRoot, fileName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(folderRoot, fileName));
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.Add("Equipamento 1");

                //First add the headers
                worksheet.Cells[5, 2].Value = "Titulo 1";
                worksheet.Cells[5, 3].Value = "Titulo 2";
                worksheet.Cells[5, 4].Value = "Titulo 3";
                worksheet.Cells[5, 5].Value = "Titulo 4";

                worksheet.Cells[5, 2, 5, 5].Style.Font.Bold = true;
                worksheet.Cells[5, 2, 5, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

                for (var i = 0; i < 10; i++)
                {
                    var numero = new Random();

                    worksheet.Cells[i + 6, 2].Value = "Valor: " + numero.Next();
                    worksheet.Cells[i + 6, 3].Value = "Valor: " + numero.Next();
                    worksheet.Cells[i + 6, 4].Value = "Valor: " + numero.Next();
                    worksheet.Cells[i + 6, 5].Value = "Valor: " + numero.Next();
                }

                worksheet = workbook.Worksheets.Add("Equipamento 2");

                worksheet.Cells[5, 2].Value = "Titulo 1";
                worksheet.Cells[5, 3].Value = "Titulo 2";
                worksheet.Cells[5, 4].Value = "Titulo 3";
                worksheet.Cells[5, 5].Value = "Titulo 4";

                worksheet.Cells[5, 2, 5, 5].Style.Font.Bold = true;
                worksheet.Cells[5, 2, 5, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Double;

                for (var i = 0; i < 10; i++)
                {
                    var numero = new Random();

                    worksheet.Cells[i + 6, 2].Value = "Valor: " + numero.Next();
                    worksheet.Cells[i + 6, 3].Value = "Valor: " + numero.Next();
                    worksheet.Cells[i + 6, 4].Value = "Valor: " + numero.Next();
                    worksheet.Cells[i + 6, 5].Value = "Valor: " + numero.Next();
                }

                package.SaveAs(file);
            }
        }
    }
}

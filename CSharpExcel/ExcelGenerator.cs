using System.Collections.Generic;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CSharpExcel
{
    class ExcelGenerator
    {
        public byte[] Generate(List<Person> students)
        {
            using var excelPackage = new ExcelPackage();

            var worksheet = excelPackage.Workbook.Worksheets.Add("Students");

            var headerParameters = new List<string> { "Имя: ", "Фамилия: ", "Возраст: ", "Телефон: " };

            var i = 1;
            foreach (var e in headerParameters)
            {
                worksheet.Cells[1, i].Value = e;
                i++;
            }

            var row = 2;

            foreach (var e in students)
            {
                worksheet.Cells[row, 1].Value = e.Name;
                worksheet.Cells[row, 2].Value = e.Surname;
                worksheet.Cells[row, 3].Value = e.Age;
                worksheet.Cells[row, 4].Value = e.PhoneNumber;

                row++;
            }

            worksheet.Cells[1, 1, row, headerParameters.Count].AutoFitColumns();

            var header = worksheet.Cells[1, 1, 1, headerParameters.Count];

            header.Style.Fill.PatternType = ExcelFillStyle.Solid;
            header.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#A5A5A5"));
            header.Style.Font.Bold = true;

            var cells = worksheet.Cells[2, 1, students.Count + 1, headerParameters.Count];

            cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cells.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#EEEEEE"));

            worksheet.Cells[1, 1, students.Count + 1, headerParameters.Count].Style.Border
                .BorderAround(ExcelBorderStyle.Double);

            worksheet.Cells[1, 1, students.Count, headerParameters.Count].Style.Border.Bottom.Style =
                ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, students.Count + 1, headerParameters.Count - 1].Style.Border.Right.Style =
                ExcelBorderStyle.Thin;

            return excelPackage.GetAsByteArray();
        }
    }
}

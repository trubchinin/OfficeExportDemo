using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ReportPlugin.Abstractions;

namespace ReportPlugin.Excel
{
    public class ExcelReportGenerator : IReportGenerator
    {
        public string FormatName => "Excel";

        public void Generate(string filePath, List<(int Id, string FullName, string Category)> items)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Readers");

            ws.Cell(1, 1).Value = "Id";
            ws.Cell(1, 2).Value = "ПІБ";
            ws.Cell(1, 3).Value = "Категорія";

            for (int i = 0; i < items.Count; i++)
            {
                ws.Cell(i + 2, 1).Value = items[i].Id;
                ws.Cell(i + 2, 2).Value = items[i].FullName;
                ws.Cell(i + 2, 3).Value = items[i].Category;
            }

            var range = ws.Range(1, 1, items.Count + 1, 3);
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder  = XLBorderStyleValues.Thin;

            wb.SaveAs(filePath);
        }
    }
}

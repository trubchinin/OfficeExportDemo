using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportPlugin.Abstractions;

namespace ReportPlugin.Word
{
    public class WordReportGenerator : IReportGenerator
    {
        public string FormatName => "Word";

        public void Generate(string filePath, List<(int Id, string FullName, string Category)> items)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
            using var doc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body;

            body.AppendChild(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(new Text("Звіт по читачах"))
            ));

            var table = new Table(new TableProperties(
                new TableBorders(
                    new TopBorder    { Val = BorderValues.Single },
                    new BottomBorder { Val = BorderValues.Single },
                    new LeftBorder   { Val = BorderValues.Single },
                    new RightBorder  { Val = BorderValues.Single },
                    new InsideHorizontalBorder { Val = BorderValues.Single },
                    new InsideVerticalBorder   { Val = BorderValues.Single }
                )
            ));

            // Шапка
            var header = new TableRow();
            header.AppendCell("Id").AppendCell("ПІБ").AppendCell("Категорія");
            table.AppendChild(header);

            foreach (var (Id, FullName, Category) in items)
            {
                var row = new TableRow();
                row.AppendCell(Id.ToString())
                   .AppendCell(FullName)
                   .AppendCell(Category);
                table.AppendChild(row);
            }

            body.AppendChild(table);
            main.Document.Save();
        }
    }

    // --- Підтримковий метод-розширення ---
    internal static class TableRowExtensions
    {
        public static TableRow AppendCell(this TableRow row, string text)
        {
            row.Append(new TableCell(new Paragraph(new Run(new Text(text)))));
            return row;
        }
    }
}

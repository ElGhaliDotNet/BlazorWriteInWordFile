using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace blazor.write.word.file.Services;

public class WordService
{
    public byte[] CreateWordFile(string content)
    {
        using (MemoryStream memoryStream = new MemoryStream())
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(memoryStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                // Add a main document part
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Add content to the document
                body.AppendChild(new Paragraph(new Run(new Text(content))));

                mainPart.Document.Save();
            }

            return memoryStream.ToArray();
        }
    }

    private TableCell CreateTableCell(string text)
    {
        return new TableCell(new Paragraph(new Run(new Text(text))));
    }

    public byte[] CreateWordTable()
    {
        using (MemoryStream memoryStream = new MemoryStream())
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(memoryStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainDocumentPart = wordprocessingDocument.AddMainDocumentPart();
                mainDocumentPart.Document = new Document();
                Body body = mainDocumentPart.Document.AppendChild(new Body());
                Paragraph titleParagraph = new Paragraph(
                    new Run(
                        new Text("Table Title: Sample Data Table")
                    )
                )
                {
                    ParagraphProperties = new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center },
                        new SpacingBetweenLines() { Before = "200", After = "200" }
                    )
                };
                body.AppendChild(titleParagraph);
                Table table = new Table();
                TableProperties tableProperties = new TableProperties(
                     new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct }, // 100% width
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                        )
                );

                table.AppendChild(tableProperties);

                TableRow headerRow = new TableRow();
                headerRow.Append(
                    CreateTableCell("Column 1"),
                    CreateTableCell("Column 2"),
                    CreateTableCell("Column 3")
                );
                table.AppendChild(headerRow);

                for (int i = 0; i < 50; i++)
                {
                    TableRow dataRow = new TableRow();
                    dataRow.Append(
                        CreateTableCell($"Row {i + 1}, Cell 1"),
                        CreateTableCell($"Row {i + 1}, Cell 2"),
                        CreateTableCell($"Row {i + 1}, Cell 3")
                    );
                    table.AppendChild(dataRow);
                }
                body.AppendChild(table);

                mainDocumentPart.Document.Save();
            }
            return memoryStream.ToArray();
        }
    }
}

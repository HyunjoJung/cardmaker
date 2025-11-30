using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace BusinessCardMaker.E2ETests;

/// <summary>
/// Helper methods for creating test files (Excel and PowerPoint)
/// </summary>
public static class TestHelpers
{
    /// <summary>
    /// Creates a test Excel file with employee data in memory
    /// </summary>
    public static byte[] CreateTestEmployeeExcel()
    {
        using var memoryStream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Employees"
            });

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            // Add header row (including custom fields)
            var headerRow = new Row { RowIndex = 1 };
            var headers = new[]
            {
                "Name", "Name (English)", "Company", "Department", "Position",
                "Position (English)", "Email", "Mobile", "Phone", "Extension", "Fax",
                "LinkedIn", "Hobby"  // Custom fields for testing
            };

            for (uint i = 0; i < headers.Length; i++)
            {
                headerRow.Append(new Cell
                {
                    CellReference = GetCellReference(1, i + 1),
                    DataType = CellValues.String,
                    CellValue = new CellValue(headers[i])
                });
            }
            sheetData.Append(headerRow);

            // Add sample employee rows (including custom fields)
            AddEmployeeRow(sheetData, 2, "John Doe", "John Doe", "Test Company", "Development", "Team Leader",
                          "Team Leader", "john@test.com", "010-1234-5678", "02-1234-5678", "1234", "02-1234-5679",
                          "linkedin.com/in/johndoe", "Golf");
            AddEmployeeRow(sheetData, 3, "Jane Smith", "Jane Smith", "Test Company", "Sales", "Senior Executive",
                          "Senior Account Executive", "jane@test.com", "010-9876-5432", "", "", "",
                          "linkedin.com/in/janesmith", "Reading");
            AddEmployeeRow(sheetData, 4, "Bob Johnson", "Bob Johnson", "Test Company", "Marketing", "Staff",
                          "Staff", "bob@test.com", "010-1111-2222", "", "", "",
                          "", "Photography");

            worksheetPart.Worksheet.Save();
        }

        return memoryStream.ToArray();
    }

    /// <summary>
    /// Creates a simple test PowerPoint template with placeholders
    /// </summary>
    public static byte[] CreateTestPowerPointTemplate()
    {
        using var memoryStream = new MemoryStream();
        using (var document = PresentationDocument.Create(memoryStream, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(new A.TransformGroup()),
                        CreateTextBox(2, "{name}", 100, 100),
                        CreateTextBox(3, "{position}", 100, 200),
                        CreateTextBox(4, "{email}", 100, 300),
                        CreateTextBox(5, "LinkedIn: {linkedin}", 100, 400),
                        CreateTextBox(6, "Hobby: {hobby}", 100, 500)
                    )
                )
            );

            var slideIdList = presentationPart.Presentation.AppendChild(new SlideIdList());
            var slideId = new SlideId { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart) };
            slideIdList.Append(slideId);

            presentationPart.Presentation.Save();
        }

        return memoryStream.ToArray();
    }

    private static void AddEmployeeRow(SheetData sheetData, uint rowIndex, params string[] values)
    {
        var row = new Row { RowIndex = rowIndex };

        for (uint i = 0; i < values.Length; i++)
        {
            row.Append(new Cell
            {
                CellReference = GetCellReference(rowIndex, i + 1),
                DataType = CellValues.String,
                CellValue = new CellValue(values[i])
            });
        }

        sheetData.Append(row);
    }

    private static string GetCellReference(uint rowIndex, uint columnIndex)
    {
        string columnName = "";
        while (columnIndex > 0)
        {
            var modulo = (columnIndex - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnIndex = (columnIndex - modulo) / 26;
        }
        return columnName + rowIndex;
    }

    private static P.Shape CreateTextBox(uint id, string text, long x, long y)
    {
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = $"TextBox{id}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = 3000000L, Cy = 1000000L })),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.Text { Text = text }))));
    }
}

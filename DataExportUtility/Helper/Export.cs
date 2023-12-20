using System.Data;
using ClosedXML.Excel;
using DinkToPdf;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

namespace DEU;
public static class Export
{
    public static void ExportToExcelFile(DataTable dataTable)
    {
        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        string filePath = $@"D:\Practice\Practice\data-export-utility\File_{timestamp}.xlsx"; // Using timestamp in the file name

        string directoryPath = Path.GetDirectoryName(filePath);

        if (!Directory.Exists(directoryPath))
        {
            // Directory doesn't exist, create it
            Directory.CreateDirectory(directoryPath);
        }

        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        worksheet.Cell(1, 1).InsertTable(dataTable); // Insert DataTable into Excel
                                                     // Save the Excel file

        workbook.SaveAs(filePath);

        Console.WriteLine("Data exported to Excel successfully!");
    }

    public static void ExportToPdfFile(DataTable dataTable)
    {
        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        string filePath = $@"D:\Practice\Practice\data-export-utility\File_{timestamp}.pdf"; // Using timestamp in the file name

        PdfDocument pdfDoc = new PdfDocument(new PdfWriter(filePath));
        Document doc = new Document(pdfDoc);

        // Adding a title
        Paragraph title = new Paragraph("PDF Report");
        title.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
        title.SetFontSize(18);
        doc.Add(title);

        // Adding empty line after the title
        doc.Add(new Paragraph("\n"));

        // Creating a table
        Table table = new Table(dataTable.Columns.Count);
        // Set table alignment to center
        table.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);

        // Add headers to the table
        foreach (DataColumn column in dataTable.Columns)
        {
            table.AddHeaderCell(new Cell().Add(new Paragraph(column.ColumnName)));
        }

        // Add data to the table
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object cellValue in row.ItemArray)
            {
                table.AddCell(new Cell().Add(new Paragraph(cellValue.ToString())));
            }
        }

        // Add the table to the document
        doc.Add(table);

        // Close the document
        doc.Close();

        Console.WriteLine($"PDF generated at: {filePath}");
    }

  public static void ExportToCSVFile(DataTable dataTable)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string filePath = $@"D:\Practice\Practice\data-export-utility\File_{timestamp}.txt"; // Using timestamp in the file name

            using (StreamWriter writer = new StreamWriter(filePath))
            {
                // Write headers
                foreach (DataColumn column in dataTable.Columns)
                {
                    writer.Write($"{column.ColumnName},");
                }
                writer.WriteLine(); // New line after headers

                // Write data
                foreach (DataRow row in dataTable.Rows)
                {
                    foreach (object cellValue in row.ItemArray)
                    {
                        writer.Write($"{cellValue},");
                    }
                    writer.WriteLine(); // New line for each row
                }
            }

            Console.WriteLine($"CSV file generated at: {filePath}");
        }
}

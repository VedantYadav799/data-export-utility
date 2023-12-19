using System.Data;
using ClosedXML.Excel;

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


    }
}


using System;
using ClosedXML.Excel;
using System.IO;

public static class Helper
{
    public static void WriteToExcel(string seriesName, string trend, string channel, string hpvType)
    {
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AnalysisResults.xlsx");

        var workbook = new XLWorkbook();
        IXLWorksheet worksheet;

        if (File.Exists(filePath))
        {
            workbook = new XLWorkbook(filePath);
            worksheet = workbook.Worksheet(1);
        }
        else
        {
            worksheet = workbook.Worksheets.Add("Results");
            // Başlık satırını ekle
            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "Channel";
            worksheet.Cell(1, 3).Value = "HPV Type";
            worksheet.Cell(1, 4).Value = "Result";
        }

        int lastRow = worksheet.LastRowUsed().RowNumber();
        int newRow = lastRow + 1;

        worksheet.Cell(newRow, 1).Value = seriesName;
        worksheet.Cell(newRow, 2).Value = channel;
        worksheet.Cell(newRow, 3).Value = hpvType;
        worksheet.Cell(newRow, 4).Value = trend;

        workbook.SaveAs(filePath);
    }
}

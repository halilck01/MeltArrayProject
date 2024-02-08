using ClosedXML.Excel;
using MeltArrayProject2;
using System.Collections.Generic;

public static class Helper
{
    public static void SaveResultsToExcel(string filePath, List<AnalysisResult> results)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Results");
            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "Channel";
            worksheet.Cell(1, 3).Value = "HPV Type";
            worksheet.Cell(1, 4).Value = "Result";

            int row = 2;
            foreach (var result in results)
            {
                worksheet.Cell(row, 1).Value = result.SeriesName;
                worksheet.Cell(row, 2).Value = result.Channel;
                worksheet.Cell(row, 3).Value = result.HpvType;
                worksheet.Cell(row, 4).Value = result.Trend;
                row++;
            }

            workbook.SaveAs(filePath);
        }
    }
}

using ClosedXML.Excel;
using MeltArrayProject2;
using System.Collections.Generic;
using System.Linq;

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
            worksheet.Cell(1, 5).Value = "Peak Value";

            int row = 2;
            foreach (var result in results)
            {
                worksheet.Cell(row, 1).Value = result.SeriesName;
                worksheet.Cell(row, 2).Value = result.Channel;
                worksheet.Cell(row, 3).Value = result.HpvType;
                worksheet.Cell(row, 4).Value = result.Trend;
                worksheet.Cell(row, 5).Value = result.PeakValue;
                
                row++;
            }

            var firstDataRow = 2;
            var lastDataRow = row - 1;
            var range = worksheet.Range(firstDataRow, 1, lastDataRow, worksheet.ColumnsUsed().Count());
            range.Sort(1, XLSortOrder.Ascending);

            workbook.SaveAs(filePath);
        }
    }
}

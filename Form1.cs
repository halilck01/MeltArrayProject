using ClosedXML.Excel;
using MeltArrayProject2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace deneme2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeChart();
        }
        private void InitializeChart()
        {
            chart1.Series.Clear(); chart1.ChartAreas.Clear(); chart1.ChartAreas.Add(new ChartArea());
            chart2.Series.Clear(); chart2.ChartAreas.Clear(); chart2.ChartAreas.Add(new ChartArea());
            chart3.Series.Clear(); chart3.ChartAreas.Clear(); chart3.ChartAreas.Add(new ChartArea());
            chart4.Series.Clear(); chart4.ChartAreas.Clear(); chart4.ChartAreas.Add(new ChartArea());
        }
        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx|Tüm Dosyalar (*.*)|*.*";
            openFileDialog.Title = "Excel Dosyası Seç";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = openFileDialog.FileName;
            }
        }
        private void txtFileName_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtFileName.Text))
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx|Tüm Dosyalar (*.*)|*.*";
                openFileDialog.Title = "Excel Dosyası Seç";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFileName.Text = openFileDialog.FileName;
                }
            }
            else
            {

            }
        }
        private void btnTransfer_Click(object sender, EventArgs e)
        {
            IXLWorkbook workbook = new XLWorkbook();

            try
            {
                workbook = new XLWorkbook(txtFileName.Text);
            }
            catch
            {
                MessageBox.Show("Lütfen Excel dosyasını kapatıp tekrar deneyiniz!!!");
            }

            foreach (var worksheet in workbook.Worksheets)
            {
                switch (worksheet.Name)
                {
                    case "FAM":
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV("FAM", "HPV33", 43.50, 47.00, chart1);
                        AnalyzeSeriesHPV("FAM", "HPV58", 49.00, 56.50, chart1);
                        AnalyzeSeriesHPV("FAM", "HPV52", 58.50, 64.00, chart1);
                        AnalyzeSeriesHPV("FAM", "HPV59", 65.00, 70.50, chart1);
                        chart1.Visible = true;
                        label1.Visible = true;
                        break;
                    case "HEX":
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV("HEX", "HPV68", 47.00, 51.00, chart2);
                        AnalyzeSeriesHPV("HEX", "HPV35", 54.00, 59.00, chart2);
                        AnalyzeSeriesHPV("HEX", "IntCo", 63.00, 68.50, chart2);
                        chart2.Visible = true;
                        label2.Visible = true;
                        break;
                    case "ROX":
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV("ROX", "HPV45", 47.50, 51.50, chart3);
                        AnalyzeSeriesHPV("ROX", "HPV18", 57.50, 63.00, chart3);
                        AnalyzeSeriesHPV("ROX", "HPV16", 64.00, 68.00, chart3);
                        AnalyzeSeriesHPV("ROX", "HPV31", 69.00, 73.00, chart3);
                        chart3.Visible = true;
                        label3.Visible = true;
                        break;
                    case "Cy5":
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV("Cy5", "HPV66", 46.00, 49.50, chart4);
                        AnalyzeSeriesHPV("Cy5", "HPV56", 57.00, 61.50, chart4);
                        AnalyzeSeriesHPV("Cy5", "HPV39", 63.00, 68.00, chart4);
                        AnalyzeSeriesHPV("Cy5", "HPV51", 68.50, 74.00, chart4);
                        chart4.Visible = true;
                        label4.Visible = true;
                        break;
                    default:
                        break;
                }
            }

            string resultsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AnalysisResults.xlsx");
            Helper.SaveResultsToExcel(resultsFilePath, analysisResults);

            analysisResults.Clear();
        }
        private void ReadExcel(IXLWorksheet worksheet)
        {
            var xValues = worksheet.Column("B").CellsUsed().Skip(1)
                           .Select(cell => cell.GetValue<double>()).ToArray();

            var lastColumnIndex = worksheet.ColumnsUsed().Last().ColumnNumber();
            var seriesNames = new List<string>();
            for (int colIndex = 3; colIndex <= lastColumnIndex; colIndex++)
            {
                var seriesName = worksheet.Cell(1, colIndex).GetValue<string>();
                seriesNames.Add(seriesName);
            }

            for (int colIndex = 3; colIndex <= lastColumnIndex; colIndex++)
            {
                int seriesNameIndex = colIndex - 3;

                var series = new System.Windows.Forms.DataVisualization.Charting.Series(seriesNames[seriesNameIndex])
                {
                    ChartType = SeriesChartType.Line
                };

                var yValues = worksheet.Column(colIndex).CellsUsed().Skip(1)
                               .Select(cell => cell.GetValue<double>()).ToArray();
                for (int i = 0; i < yValues.Length; i++)
                {
                    series.Points.AddXY(xValues[i], yValues[i]);
                }
                if (worksheet.Name == "FAM")
                {
                    chart1.Series.Add(series);
                }
                if (worksheet.Name == "HEX")
                {
                    chart2.Series.Add(series);
                }
                if (worksheet.Name == "ROX")
                {
                    chart3.Series.Add(series);
                }
                if (worksheet.Name == "Cy5")
                {
                    chart4.Series.Add(series);
                }
            }

            if (worksheet.Name == "FAM")
            {
                chart1.ChartAreas[0].AxisX.Minimum = xValues.Min();
                chart1.ChartAreas[0].AxisX.Maximum = xValues.Max();
            }
            if (worksheet.Name == "HEX")
            {
                chart2.ChartAreas[0].AxisX.Minimum = xValues.Min();
                chart2.ChartAreas[0].AxisX.Maximum = xValues.Max();
            }
            if (worksheet.Name == "ROX")
            {
                chart3.ChartAreas[0].AxisX.Minimum = xValues.Min();
                chart3.ChartAreas[0].AxisX.Maximum = xValues.Max();
            }
            if (worksheet.Name == "Cy5")
            {
                chart4.ChartAreas[0].AxisX.Minimum = xValues.Min();
                chart4.ChartAreas[0].AxisX.Maximum = xValues.Max();
            }
        }

        #region AnalyzeSeriesHPVMetot
        private List<AnalysisResult> analysisResults = new List<AnalysisResult>();
        private void AnalyzeSeriesHPV(string channel, string hpvType, double xLowerBound, double xUpperBound, Chart selectedChart)
        {
            double threshold = 20.00;

            foreach (var series in selectedChart.Series)
            {
                double lowerBoundY = double.NaN;
                double upperBoundY = double.NaN;
                double maxBetweenY = double.MinValue;
                double minYBetweenBounds = double.MaxValue;
                int lowerBoundIndex = -1;
                int upperBoundIndex = -1;

                for (int i = 0; i < series.Points.Count; i++)
                {
                    double xValue = series.Points[i].XValue;
                    double yValue = series.Points[i].YValues[0];
                    if (xValue == xLowerBound) { lowerBoundY = yValue; lowerBoundIndex = i; }
                    else if (xValue == xUpperBound) { upperBoundY = yValue; upperBoundIndex = i; }
                }

                if (lowerBoundIndex != -1 && upperBoundIndex != -1)
                {
                    for (int i = lowerBoundIndex + 1; i < upperBoundIndex; i++)
                    {
                        double yValue = series.Points[i].YValues[0];
                        maxBetweenY = Math.Max(maxBetweenY, yValue);
                        minYBetweenBounds = Math.Min(minYBetweenBounds, yValue);
                    }
                }

                bool isPositive = maxBetweenY > lowerBoundY && maxBetweenY > upperBoundY && (maxBetweenY - lowerBoundY > threshold || maxBetweenY - upperBoundY > threshold);

                if (isPositive)
                {
                    analysisResults.Add(new AnalysisResult
                    {
                        SeriesName = series.Name,
                        Trend = "Positive",
                        Channel = channel,
                        HpvType = hpvType,
                        PeakValue = maxBetweenY.ToString()
                    });
                }
            }
        }
        #endregion     
    }
}

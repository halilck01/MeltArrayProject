﻿using ClosedXML.Excel;
using System;
using System.Collections.Generic;
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
                        // FAM sayfası için işlemler
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV33();
                        AnalyzeSeriesHPV58();
                        AnalyzeSeriesHPV52();
                        AnalyzeSeriesHPV59();
                        chart1.Visible = true;
                        break;
                    case "HEX":
                        // VIC sayfası için işlemler
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV68();
                        AnalyzeSeriesHPV35();
                        AnalyzeSeriesIC();
                        chart2.Visible = true;
                        break;
                    case "ROX":
                        // ROX sayfası için işlemler
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV45();
                        AnalyzeSeriesHPV18();
                        AnalyzeSeriesHPV16();
                        AnalyzeSeriesHPV31();
                        chart3.Visible = true;
                        break;
                    case "Cy5":
                        // Cy5 sayfası için işlemler
                        ReadExcel(worksheet);
                        AnalyzeSeriesHPV66();
                        AnalyzeSeriesHPV56();
                        AnalyzeSeriesHPV39();
                        AnalyzeSeriesHPV51();
                        chart4.Visible = true;
                        break;
                    default:
                        break;
                }
            }
        }
        private void ReadExcel(IXLWorksheet worksheet)
        {
            // X ekseninde kullanılacak değerleri B sütunundan oku
            var xValues = worksheet.Column("B").CellsUsed().Skip(1) // İlk satırı atla
                           .Select(cell => cell.GetValue<double>()).ToArray();

            // Seri adlarını 1. satırdan C sütunundan itibaren doğrudan oku
            var lastColumnIndex = worksheet.ColumnsUsed().Last().ColumnNumber(); // Son kullanılan sütunun numarası
            var seriesNames = new List<string>();
            for (int colIndex = 3; colIndex <= lastColumnIndex; colIndex++) // Son sütunu da dikkate al
            {
                var seriesName = worksheet.Cell(1, colIndex).GetValue<string>();
                seriesNames.Add(seriesName);
            }

            // Her bir seri için döngü
            for (int colIndex = 3; colIndex <= lastColumnIndex; colIndex++) // Son sütunu da dikkate al
            {
                int seriesNameIndex = colIndex - 3; // C sütunu için 0, D için 1, vs.

                var series = new System.Windows.Forms.DataVisualization.Charting.Series(seriesNames[seriesNameIndex])
                {
                    ChartType = SeriesChartType.Line
                };

                // Y değerlerini oku ve seriye ekle
                var yValues = worksheet.Column(colIndex).CellsUsed().Skip(1) // İlk satırı atla
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

        #region Metots
        #region AnalyzeSeriesFAMMetots
        private void AnalyzeSeriesHPV33()
        {
            double xLowerBound = 43.75;
            double xUpperBound = 47.00;

            foreach (var series in chart1.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "FAM";
                string hpvType = "HPV33";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV58()
        {
            double xLowerBound = 49.00;
            double xUpperBound = 56.74;

            foreach (var series in chart1.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "FAM";
                string hpvType = "HPV58";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV52()
        {
            double xLowerBound = 58.40;
            double xUpperBound = 64.23;

            foreach (var series in chart1.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "FAM";
                string hpvType = "HPV52";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV59()
        {
            double xLowerBound = 64.79;
            double xUpperBound = 70.54;

            foreach (var series in chart1.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "FAM";
                string hpvType = "HPV59";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        #endregion

        #region AnalyzeSeriesVICMetots
        private void AnalyzeSeriesHPV68()
        {
            double xLowerBound = 46.78;
            double xUpperBound = 50.91;

            foreach (var series in chart2.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "HEX";
                string hpvType = "HPV68";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel , hpvType);
            }
        }
        private void AnalyzeSeriesHPV35()
        {
            double xLowerBound = 54.15;
            double xUpperBound = 58.90;

            foreach (var series in chart2.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "HEX";
                string hpvType = "HPV35";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesIC()
        {
            double xLowerBound = 63.03;
            double xUpperBound = 68.26;

            foreach (var series in chart2.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "HEX";
                string hpvType = "Internal Control";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        #endregion

        #region AnalyzeSeriesROXMetots
        private void AnalyzeSeriesHPV45()
        {
            double xLowerBound = 47.64;
            double xUpperBound = 51.44;

            foreach (var series in chart3.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "ROX";
                string hpvType = "HPV45";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV18()
        {
            double xLowerBound = 57.73;
            double xUpperBound = 62.98;

            foreach (var series in chart3.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "ROX";
                string hpvType = "HPV18";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV16()
        {
            double xLowerBound = 64.00;
            double xUpperBound = 68.15;

            foreach (var series in chart3.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "ROX";
                string hpvType = "HPV16";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV31()
        {
            double xLowerBound = 69.19;
            double xUpperBound = 73.03;

            foreach (var series in chart3.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "ROX";
                string hpvType = "HPV31";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        #endregion

        #region AnalyzeSeriesCY5Metots
        private void AnalyzeSeriesHPV66()
        {
            double xLowerBound = 45.97;
            double xUpperBound = 49.54;

            foreach (var series in chart4.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "Cy5";
                string hpvType = "HPV66";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV56()
        {
            double xLowerBound = 56.98;
            double xUpperBound = 61.42;

            foreach (var series in chart4.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "Cy5";
                string hpvType = "HPV56";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV39()
        {
            double xLowerBound = 63.21;
            double xUpperBound = 67.83;

            foreach (var series in chart4.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "Cy5";
                string hpvType = "HPV39";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        private void AnalyzeSeriesHPV51()
        {
            double xLowerBound = 68.68;
            double xUpperBound = 73.81;

            foreach (var series in chart4.Series)
            {
                bool increaseDetected = false;
                bool decreaseDetectedAfterIncrease = false;

                double lastYValue = double.NaN;

                foreach (var point in series.Points)
                {
                    double currentX = point.XValue;
                    double currentY = point.YValues[0];

                    // X ekseninde belirli bir aralıkta mı kontrol et
                    if (currentX >= xLowerBound && currentX <= xUpperBound)
                    {
                        // Artış tespit et
                        if (!double.IsNaN(lastYValue) && currentY > lastYValue)
                        {
                            increaseDetected = true;
                        }

                        // Artıştan sonra azalış tespit et
                        if (increaseDetected && !double.IsNaN(lastYValue) && currentY < lastYValue - 5.00)
                        {
                            decreaseDetectedAfterIncrease = true;
                            break;
                        }

                        lastYValue = currentY;
                    }
                }
                string channel = "Cy5";
                string hpvType = "HPV51";
                string trend = "Negatif";
                if (increaseDetected && decreaseDetectedAfterIncrease)
                {
                    trend = "Pozitif";
                }

                Helper.WriteToExcel(series.Name, trend, channel, hpvType);
            }
        }
        #endregion

        #endregion     
    }
}
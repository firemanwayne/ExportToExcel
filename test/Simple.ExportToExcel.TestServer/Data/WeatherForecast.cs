using Simple.ExportToExcel.Attributes;
using System;

namespace Simple.ExportToExcel.TestServer.Data
{
    [SpreadSheet("Weather Data", typeof(WeatherForecast))]
    public class WeatherForecast
    {
        [SpreadSheetColumn("Date", 0)]
        public DateTime Date { get; set; }

        [SpreadSheetColumn("Temp. Celsius", 1)]
        public int TemperatureC { get; set; }

        [SpreadSheetColumn("Temp. Fahrenheit", 2)]
        public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);

        [SpreadSheetColumn("Summary", 3)]
        public string Summary { get; set; }
    }
}

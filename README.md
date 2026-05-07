# Simple Blazor Export-To-Excel Component

Simple Blazor Component that will export data as an Excel spreadsheet (.xlsx).

**Dependencies:**

- [NPOI](https://github.com/nissl-lab/npoi) 2.8.0

**Requirements:**

- .NET 10.0

**Installation:**

Register the required services in your `Program.cs`:

```csharp
builder.Services.AddExportToExcel();
```

**Component Parameters:**

| Parameter | Type | Description |
|---|---|---|
| `TValue` | generic | Type of object being exported |
| `RequestDelegate` | `Func<IEnumerable<TValue>>` | Method that retrieves the data to export |
| `ButtonText` | `object` | Text displayed on the export button |
| `ButtonClass` | `string` | CSS class for the export button (default: `btn btn-outline-success`) |
| `ButtonStyle` | `string` | Inline CSS style for the export button |
| `ButtonTitle` | `string` | Tooltip text on the export button (default: `Export To Excel`) |

The component renders color and alignment pickers for both the header and body rows, and a filename input. The spreadsheet is generated and downloaded when the form is submitted.

**Example:**

```razor
<ExcelExport ButtonText="Export Weather Data"
             ButtonClass="btn btn-outline-success"
             TValue="WeatherForecast"
             RequestDelegate="GetData" />

@code {
    WeatherForecast[] forecasts;

    [Inject] WeatherForecastService ForecastService { get; set; }

    protected override async Task OnInitializedAsync()
    {
        forecasts = await ForecastService.GetForecastAsync(DateTime.Now);
    }

    IEnumerable<WeatherForecast> GetData() => forecasts ?? Enumerable.Empty<WeatherForecast>();
}
```

**Column Mapping:**

By default, columns are generated from properties decorated with `[Display(Name="...")]`. Properties without the attribute fall back to the property name.

For explicit control over column order and names, decorate your class with `[SpreadSheet]` and properties with `[SpreadSheetColumn]`:

```csharp
[SpreadSheet("WeatherForecast", typeof(WeatherForecast))]
public class WeatherForecast
{
    [SpreadSheetColumn("Date", 0)]
    public DateTime Date { get; set; }

    [SpreadSheetColumn("Temperature (C)", 1)]
    public int TemperatureC { get; set; }

    [SpreadSheetColumn("Summary", 2)]
    public string Summary { get; set; }
}
```

**Contributions**

Thank you to Andre' Rizzato for discovering an implementation issue and providing a working solution.

**Change Log**

[Version 1.1.0.1]
- Changed how cell value is set so that the correct type is used. Type integer was causing formatting warnings in the exported Excel file
- Updated to .NET 10.0
- Migrated to Central Package Management (CPM)
- Fixed concurrency bug where static cell style fields in `BodyBuilder` could corrupt styles across concurrent exports
- Fixed `async void` on export handler to properly propagate exceptions
- Fixed `IndexOutOfRangeException` when list sub-properties or reflected properties lack a `[Display]` attribute
- Fixed foreground color selection not being applied to the generated spreadsheet

[Version 1.1.0.0]
- Changed parameters in ExcelExport component
  - `CssClass` => `ButtonClass`
  - `ReportName` => removed
  - `Title` => `ButtonTitle`
  - `ButtonStyle` => added for inline styling
- Changed styling of input fields
- Added validation to filename input field
- Updated to .NET 6.0

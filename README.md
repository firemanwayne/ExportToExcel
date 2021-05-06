# Simple Blazor Export-To-Excel Component

Simple Blazor Component that will Export Data in a Spreadsheet format

**Dependencies:**

-NPOI Contributors version: 2.5.1 https://github.com/nissl-lab/npoi

**Implementation:**

Start by calling the registry method when you are configuring your services:

```csharp
 services.AddExportToExcel();
```

Here is an example of using the ExportToExcel component:

```csharp
[Parameter] ButtonText = Sets the text that will be displayed on the Export Button
[Parameter] CssClass = Sets the Css Class for the Export Button
[Parameter] ReportName = The name that you would like the report to be saved as
[Parameter] TValue = Type of object that is being used
[Parameter] RequestDelegate = Func<IEnumerable<TValue>> method that will retrive a list of the specified TValue
```

**Example of usage:**

```html

<ExcelExport ButtonText=ExportButtonText
             CssClass="btn btn-outline-success"
             ReportName=@ExcelSpreadSheetName
             TValue=WeatherForecast
             RequestDelegate=GetData />

@code{

    string ExportButtonText = "Export Weather Data";
    string ExcelSpreadSheetName = $"CurrentWeatherTable{DateTime.Now:ddmmyyyyhhmmss}";   

    WeatherForecast[] forecasts;

    [Inject] WeatherForecastService ForecastService { get; set; }

    protected override async Task OnInitializedAsync()
    {
        forecasts = await ForecastService.GetForecastAsync(DateTime.Now);
    }   

    /// <summary>
    /// Data retrieval method
    /// </summary>
    /// <returns></returns>
    IEnumerable<WeatherForecast> GetData() => forecasts ?? Enumerable.Empty<WeatherForecast>();
     
}

```

**Contributions**

Thank you to Andre' Rizzato for discovering an implementation issue and providing a working solution.

# Simple Blazor Export-To-Excel Component

Simple Blazor Component that will Export Data in a Spreadsheet format

**Dependencies:**

-NPOI Contributors version: 2.5.1 https://github.com/nissl-lab/npoi

**Implementation:**

Start by adding services.AddExportToExcel() method when you are configuring your services:

```csharp
 services.AddExportToExcel();
```

Here is an example of using the ExportToExcel component:

```csharp
[Parameter] ButtonText = Sets the text that will be displayed on the Export Button
[Parameter] ButtonStyle = Sets the Css Style for the Export Button
[Parameter] ButtonClass = Sets the Css Class for the Export Button
[Parameter] ButtonTitle = Displays text when you hover over button
[Parameter] TValue = Type of object that is being exported
[Parameter] RequestDelegate = Func<IEnumerable<TValue>> method that will retrive a list of the specified TValue
```

**Example of usage:**

```html

 <ExcelExport ButtonText=@ButtonText
              ButtonClass="btn btn-outline-success btn-excel"
              TValue=WeatherForecast
              RequestDelegate=GetData />

@code{

    string ButtonText = "Export Weather Data";       

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

**Change Log**

[Version 1.1.0.1]
- Changed how cell value is set so that the correct type is used. Type integer was causing formatting warnings to show in the exported excel file

[Version 1.1.0.0]
 - Changed parameters in ExcelExport component
    - CssClass => ButtonClass
    - ReportName => Has been removed
    - Title => Button Title
    - ButtonStyle => Added so you can add additional styling

 - Changed styling of input fields
 - Added validation to Filename input field
 - Updated to .Net 6.0
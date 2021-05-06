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

```html

<ExcelExport 
ButtonText=UserButtonText
CssClass="btn btn-outline-success" 
ReportName=@UserReportName
TValue="UserSpreadSheet"
RequestDelegate="ExportUserRequest" />

```

**Contributions**

Thank you to Andre' Rizzato for discovering an implementation issue and providing a working solution.

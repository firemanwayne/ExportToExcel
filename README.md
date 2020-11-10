# Blazor Export-To-Excel Component

Blazor Component that will Export Data in a Spreadsheet format

Dependencies:

-NPOI Contributors version: 2.5.1 https://github.com/nissl-lab/npoi

**Implementation:**

Start by calling the registry method when you are configuring your services:

```csharp
 services.AddExportToExcel();
```

Once added to your start up process you can use the following Components to configure the excel document before exporting it:

```html
<ExcelColorSelector OnColorSelected="HandleHeaderColorSelected" />

<ExcelHorizontalAlignmentSelector OnSelected="HandleHeaderHorizontalAlignmentSelected" />

<ExcelVerticalAlignmentSelector OnSelected="HandleHeaderVerticalAlignmentSelected" />

```

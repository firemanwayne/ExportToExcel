# Blazor Export-To-Excel Component

Blazor Component that will Export Data in a Spreadsheet format

**Dependencies:**

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

The event callbacks in the components are used to set the values of your Export Style objects:

 Configures the style of the header:
```csharp
 void HandleHeaderColorSelected(StyleColorSelectedEventArgs a)
 {
     HeaderStyle.SetForegroundColor(a);
 }

 void HandleHeaderVerticalAlignmentSelected(VerticalAlignmentChangedEventArgs a)
 {
     HeaderStyle.SetVerticalAlignment(a);
 }

 void HandleHeaderHorizontalAlignmentSelected(HorizontalAlignmentChangedEventArgs a)
 {
     HeaderStyle.SetHorizontalAlignment(a);
 }
 ```
 Configures the style of the body:

 ```csharp
 void HandleBodyHorizontalAlignmentSelected(HorizontalAlignmentChangedEventArgs a)
 {
     BodyStyle.SetHorizontalAlignment(a);
 }

 void HandleBodyColorSelected(StyleColorSelectedEventArgs a)
 {
     BodyStyle.SetForegroundColor(a);
 }

 void HandleBodyVerticalAlignmentSelected(VerticalAlignmentChangedEventArgs a)
 {
     BodyStyle.SetVerticalAlignment(a);
 }
```
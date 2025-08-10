## Power Query (M Code)
These queries provide a reusable template for connecting to SharePoint sites and lists defined by variables within the workbook itself. To implement this code, developers should create two named ranges in their Excel workbook: "siteUrl" containing the SharePoint site URL, and "custList" containing the exact name of the customer list as it appears in SharePoint. The query automatically handles SharePoint authentication, filters to the specified list, and expands the customer data into a flat table structure suitable for Power BI integration or Excel analysis and visualization. 

### Opportunities Data Connection
```m
let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="oppList"]}[Content]{0}[Column1],
    Source = SharePoint.Tables(siteUrl, [Implementation=null, ApiVersion=15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = listName)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Id", "Title"}),
    #"Expanded Items" = Table.ExpandTableColumn(#"Removed Columns", "Items", {"Id", "Title", "Status", "Stage", "Amount", "Probability", "CustomerId", "Close", "NextMilestone", "NextMilestoneDate", "Comment Log", "Customer"}, {"Id", "Title", "Status", "Stage", "Amount", "Probability", "CustomerId", "Close", "NextMilestone", "NextMilestoneDate", "Comment Log", "Customer"})
in
    #"Expanded Items"
```
### Customers Data Connection
```m
let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="custList"]}[Content]{0}[Column1],
    Source = SharePoint.Tables(siteUrl, [Implementation=null, ApiVersion=15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = listName)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Id", "Title"}),
    #"Expanded Items" = Table.ExpandTableColumn(#"Removed Columns", "Items", {"Id", "Title", "PrimaryContact", "PrimaryContactTitle", "AlternateContact", "AlternateContactTitle", "AlternateContact2", "AlternateContact2Title", "Website", "Modified", "Created"}, {"Id", "Title", "PrimaryContact", "PrimaryContactTitle", "AlternateContact", "AlternateContactTitle", "AlternateContact2", "AlternateContact2Title", "Website", "Modified", "Created"}),
    #"Expanded Website" = Table.ExpandRecordColumn(#"Expanded Items", "Website", {"Url"}, {"Url"})
in
    #"Expanded Website"
```

## Excel
### Formula Patterns

#### High Complexity Formulas
- **Dynamic Arrays**: Extensive use of UNIQUE, FILTER, SORTBY, and TRANSPOSE
- **Nested Lookups**: Multiple XLOOKUP functions with complex criteria
- **Structured References**: Heavy use of Excel table references (e.g., `View[[#This Row],[Id]]`)
- **Complex Text Operations**: Date formatting and URL generation

#### Primary Formula Patterns
1. **Data Lookup and Display**: XLOOKUP patterns for retrieving related data
2. **Dynamic Filtering**: FILTER and UNIQUE combinations for data analysis
3. **URL Generation**: HYPERLINK formulas for SharePoint integration
4. **Date Calculations**: Complex fiscal year and quarter calculations
5. **Data Visualization Support**: Array formulas for chart data preparation

#### External Dependencies
- SharePoint Lists (Opportunities, Customers)
- Named ranges for configuration settings
- External data connections for real-time updates


### Function Inventory

The spreadsheet uses 21 unique Excel functions:

# Excel Functions Introduction Timeline

| Function | Usage | Description | When Introduced |
|----------|-------|-------------|-----------------|
| **Dynamic Array Functions (Office 365/2021)** |
| XLOOKUP | High | Advanced lookup function | 2020 (Excel for Office 365 version 2001/Build 12430.20184) |
| UNIQUE | Medium | Returns unique values | December 2019 (Excel for Office 365 version 1911/Build 12228.20332) |
| FILTER | Medium | Filters data arrays | December 2019 (Excel for Office 365 version 1911/Build 12228.20332) |
| SORTBY | Medium | Sorts data by criteria | December 2019 (Excel for Office 365 version 1911/Build 12228.20332) |
| LET | Low | Defines variables in formulas | December 2019 (Excel for Office 365 version 1911/Build 12228.20332) |
| **Modern Functions (Excel 2007-2019)** |
| SWITCH | Low | Multiple condition evaluation | Excel 2016 (2015) |
| **Classic Functions (Excel 97-2007)** |
| HYPERLINK | High | Creates clickable links | Excel 97 (1997) |
| SUBTOTAL | Medium | Subtotal calculations for dynamic filtering | Excel 97 (1997) |
| TRANSPOSE | Medium | Transposes arrays | Excel 1.0 (1987) - Available since early Excel versions |
| **Legacy Functions (Excel 1.0-95)** |
| IF | High | Conditional logic | Excel 1.0 (1987) - Core function since inception |
| MONTH | Medium | Extracts month from date | Excel 1.0 (1987) - Core date function |
| YEAR | Medium | Extracts year from date | Excel 1.0 (1987) - Core date function |
| RIGHT | Medium | Text extraction | Excel 1.0 (1987) - Core text function |
| CEILING | Low | Rounds up to nearest integer | Excel 1.0 (1987) - Core math function |
| ROUNDUP | Low | Rounds numbers up | Excel 1.0 (1987) - Core math function |
| TEXT | Low | Text formatting | Excel 1.0 (1987) - Core text function |
| CONCATENATE | Low | Text concatenation | Excel 4.0 (1992) |
| SUM | High | Summation | Excel 1.0 (1987) - Core function since inception |
| COUNT | Low | Count functions | Excel 1.0 (1987) - Core function since inception |
| MAX | Low | Maximum value | Excel 1.0 (1987) - Core function since inception |
| MIN | Low | Minimum value | Excel 1.0 (1987) - Core function since inception |

## Key Observations:

### **Revolutionary Period (2019-2020)**
The introduction of Dynamic Array functions in December 2019 represents the most significant Excel function advancement in decades. Functions like XLOOKUP, UNIQUE, FILTER, SORTBY, and LET fundamentally changed how Excel handles data analysis and lookup operations.

### **Foundation Era (1987-1992)**
Most basic mathematical, text, and date functions (SUM, COUNT, IF, YEAR, MONTH, RIGHT, etc.) were included in Excel's earliest versions, establishing the core functionality that users still rely on today.

### **Feature Expansion (1997-2016)**
Functions like HYPERLINK and SUBTOTAL were added during Excel's growth period, while SWITCH represents more recent logical function improvements.

### **Compatibility Notes:**
- **XLOOKUP and Dynamic Array functions** are only available in Excel for Office 365, Excel 2021, and Excel for the web
- **Older functions** (SUM, COUNT, IF, etc.) work across all Excel versions
- **TRANSPOSE** has been enhanced over time but maintains backward compatibility

### Named Ranges Inventory

| Name | Type | Reference | Sheet | Hidden | Description |
|------|------|-----------|-------|---------|-------------|
| calendarType | User Defined | Pipeline!$B$2 | Global | No | Calendar type setting |
| custList | User Defined | Settings!$E$2 | Global | No | Customer list reference |
| dateRange | User Defined | Pipeline!$D$2 | Global | No | Date range setting |
| fqStart | User Defined | Settings!$B$1 | Global | No | Fiscal quarter start |
| oppList | User Defined | Settings!$E$3 | Global | No | Opportunities list reference |
| pipelineType | User Defined | Pipeline!$B$3 | Global | No | Pipeline type setting |
| probabilityThreshold | User Defined | Settings!$B$2 | Global | No | Probability threshold setting |
| siteUrl | User Defined | Settings!$E$1 | Global | No | SharePoint site URL |
| _xlchart.v2.0 | Chart Range | graphics!$A$11:$A$16 | Global | Yes | Chart data series 1 |
| _xlchart.v2.1 | Chart Range | graphics!$C$11:$C$16 | Global | Yes | Chart data series 2 |
| ExternalData_1 | External Data | 'Customers'!$A$1:$K$31 | Customers | Yes | Customer Sharepoint List |
| ExternalData_2 | External Data | Opportunities!$A$1:$L$37 | Opportunities | Yes | Opportunity Sharepoint List |
| ExternalData_3 | External Data | Pipeline!$M$7:$M$43 | Pipeline | Yes | Pipeline View |
| Slicer_Close_Quarter | Slicer | #N/A | Global | No | Close quarter slicer |
| Slicer_Stage | Slicer | #N/A | Global | No | Stage slicer |



/* Old Sharepoint.Tables version
let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="oppList"]}[Content]{0}[Column1],
    Source = SharePoint.Tables(siteUrl, [Implementation=null, ApiVersion=15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = listName)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Id", "Title"}),
    #"Expanded Items" = Table.ExpandTableColumn(#"Removed Columns", "Items", {"Id", "Title", "Status", "Stage", "Amount", "Probability", "CustomerId", "Close", "NextMilestone", "NextMilestoneDate", "Comment Log", "Customer"}, {"Id", "Title", "Status", "Stage", "Amount", "Probability", "CustomerId", "Close", "NextMilestone", "NextMilestoneDate", "Comment Log", "Customer"})
in
    #"Expanded Items"

*/

// Rest API Version

let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="oppList"]}[Content]{0}[Column1],
    
    // Use CustomerIdId to get the lookup ID value
    restUrl = siteUrl & "_api/web/lists/getbytitle('" & listName & "')/items?$select=ID,Title,OpportunityName,Status,OpportunityStage,Amount,Probability,Close,NextMilestoneDate,NextMilestone,CustomerIdId,CommentLog,OpportunityOwner",
    
    Source = Json.Document(Web.Contents(restUrl, [
        Headers=[
            #"Accept" = "application/json;odata=verbose"
        ]
    ])),
    
    Results = Source[d][results],
    
    #"Converted to Table" = Table.FromList(Results, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", 
        {"ID", "Title", "OpportunityName", "Status", "OpportunityStage", "Amount", "Probability", "Close", "NextMilestoneDate", "NextMilestone", "CustomerIdId", "CommentLog", "OpportunityOwner"}, 
        {"ID", "Title", "OpportunityName", "Status", "Stage", "Amount", "Probability", "Close", "NextMilestoneDate", "NextMilestone", "CustomerId", "Comment Log", "OpportunityOwner"}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Expanded Column1", {{"Close", each Text.BeforeDelimiter(_, "T"), type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Text Before Delimiter",{{"Close", type date}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Changed Type", {{"NextMilestoneDate", each Text.BeforeDelimiter(_, "T"), type text}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Extracted Text Before Delimiter1",{{"NextMilestoneDate", type date}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"Title"})
in
    #"Removed Columns"
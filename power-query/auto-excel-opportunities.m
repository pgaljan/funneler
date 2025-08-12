// Rest API Version

let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="oppList"]}[Content]{0}[Column1],
    
    // Include the recurring revenue fields in the $select clause
    restUrl = siteUrl & "_api/web/lists/getbytitle('" & listName & "')/items?$select=ID,Title,OpportunityName,Status,OpportunityStage,Amount,Probability,Close,NextMilestoneDate,NextMilestone,CustomerIdId,CommentLog,OpportunityOwner,RecurringRevenueModel,Recurrences,StartDate",
    
    Source = Json.Document(Web.Contents(restUrl, [
        Headers=[
            #"Accept" = "application/json;odata=verbose"
        ]
    ])),
    
    Results = Source[d][results],
    
    #"Converted to Table" = Table.FromList(Results, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", 
        {"ID", "Title", "OpportunityName", "Status", "OpportunityStage", "Amount", "Probability", "Close", "NextMilestoneDate", "NextMilestone", "CustomerIdId", "CommentLog", "OpportunityOwner", "RecurringRevenueModel", "Recurrences", "StartDate"}, 
        {"ID", "Title", "OpportunityName", "Status", "Stage", "Amount", "Probability", "Close", "NextMilestoneDate", "NextMilestone", "CustomerId", "Comment Log", "OpportunityOwner", "RecurringRevenueModel", "Recurrences", "StartDate"}),
    
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Expanded Column1", {{"Close", each Text.BeforeDelimiter(_, "T"), type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Text Before Delimiter",{{"Close", type date}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Changed Type", {{"NextMilestoneDate", each Text.BeforeDelimiter(_, "T"), type text}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Extracted Text Before Delimiter1",{{"NextMilestoneDate", type date}}),
    
    // Handle StartDate conversion if it exists
    #"Extracted Text Before Delimiter2" = Table.TransformColumns(#"Changed Type1", {{"StartDate", each if _ <> null then Text.BeforeDelimiter(Text.From(_), "T") else null, type text}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Extracted Text Before Delimiter2",{{"StartDate", type date}}),
    
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type2",{"Title"})
in
    #"Removed Columns"
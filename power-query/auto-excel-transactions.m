let
    Source = Opportunities,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"ID", "OpportunityID"}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Renamed Columns",null,"One-time",Replacer.ReplaceValue,{"RecurringRevenueModel"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Replaced Value1",{"OpportunityID", "Amount", "Close", "CustomerId", "RecurringRevenueModel", "Recurrences", "StartDate"}),
    #"Added Conditional Column" = Table.AddColumn(#"Removed Other Columns", "calcStartDate", each if [StartDate] = null then [Close] else [StartDate]),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Conditional Column",{"StartDate"}),
    #"Added Conditional Column1" = Table.AddColumn(#"Removed Columns1", "calcRecurrences", each if [Recurrences] = null then 1 else if [RecurringRevenueModel] = "One-time" then 1 else [Recurrences]),
    #"Removed Columns2" = Table.RemoveColumns(#"Added Conditional Column1",{"Close", "Recurrences", "Amount"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns2",null,"One-time",Replacer.ReplaceValue,{"RecurringRevenueModel"}),
    
    // Function to calculate date increment based on recurring model
    GetDateIncrement = (model as text, occurrenceNumber as number) =>
        let
            increment = if model = "Monthly" then occurrenceNumber - 1
                       else if model = "Quarterly" then (occurrenceNumber - 1) * 3
                       else if model = "Semi-Annually" then (occurrenceNumber - 1) * 6
                       else if model = "Annually" then (occurrenceNumber - 1) * 12
                       else if model = "One-time" then occurrenceNumber - 1 // Assuming monthly for usage-based
                       else 0 // One-time
        in
            increment,
    
    // Expand each row based on calcRecurrences
    #"Expanded Rows" = Table.ExpandListColumn(
        Table.AddColumn(#"Replaced Value", "OccurrenceList", 
            each List.Numbers(1, [calcRecurrences], 1)
        ), "OccurrenceList"
    ),
    
    // Calculate TransactionDate for each occurrence
    #"Added Transaction Date" = Table.AddColumn(#"Expanded Rows", "TransactionDate", 
        each Date.AddMonths([calcStartDate], GetDateIncrement([RecurringRevenueModel], [OccurrenceList]))
    ),
    
    // Select and reorder final columns
    #"Selected Columns" = Table.SelectColumns(#"Added Transaction Date", {"OpportunityID", "CustomerId", "TransactionDate"}),
    #"Sorted Rows" = Table.Sort(#"Selected Columns",{{"TransactionDate", Order.Ascending}})
in
    #"Sorted Rows"
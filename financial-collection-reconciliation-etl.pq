let
    // -------------------------------------------------------------
    // STEP 1: Pehle se maujood table se Manual Inputs (CM NO, Current Remarks) read karna
    // Taake Refresh par aapka likha hua data ghayab na ho
    // -------------------------------------------------------------
    ExistingTable = try Excel.CurrentWorkbook(){[Name="Master_Collection_Table"]}[Content] otherwise null,
    PreservedData = if ExistingTable <> null and List.Contains(Table.ColumnNames(ExistingTable), "CM NO") and List.Contains(Table.ColumnNames(ExistingTable), "Current Remarks") then
        Table.TransformColumnTypes(
            Table.SelectColumns(ExistingTable, {"Complaint No", "CM NO", "Current Remarks"}),
            {{"Complaint No", type text}, {"CM NO", type text}, {"Current Remarks", type text}}
        )
    else
        #table(type table [Complaint No = text, CM NO = text, Current Remarks = text], {}),

    // -------------------------------------------------------------
    // STEP 2: Main Folder se Data Pull & Transform (Clean & Optimized)
    // -------------------------------------------------------------
    Source = Folder.Files("C:\Users\PC\Desktop\COMPLAINT_DETAIL_COLLECTION"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"Region Name", type text}, {"City Name", type text}, {"Branch Name", type text}, {"PRODUCT", type text}, {"Payment Type", type text}, {"Invoice Type", type text}, {"STORE TYPE", type text}, {"Job Type", type text}, {"Complaint Type", type text}, {"Closed Date", type date}, {"Complaint No", Int64.Type}, {"Customer Name", type text}, {"Technician Name", type text}, {"Model Name", type text}, {"Service Warranty", Int64.Type}, {"Other Warranty Charges", Int64.Type}, {"Gas Warranty Charges", Int64.Type}, {"Part Warranty", Int64.Type}, {"Total Warranty", Int64.Type}, {"Service Cash", Int64.Type}, {"Other Charges Cash", Int64.Type}, {"Gas Charges Cash", Int64.Type}, {"Part Cash", Int64.Type}, {"Advance Receipt", Int64.Type}, {"Gross Income", Int64.Type}, {"Less Advance", Int64.Type}, {"Discount", Int64.Type}, {"Net Collection", Int64.Type}, {"CM_NO", type any}}),
    
    // Sirf zaroori columns rakhna
    #"Removed Other Columns" = Table.SelectColumns(#"Changed Type", {"Branch Name", "Complaint Type", "Closed Date", "Complaint No", "Technician Name", "Net Collection"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Other Columns", {{"Complaint No", type text}}),

    // -------------------------------------------------------------
    // STEP 3: Manual Collection Query ke sath Merge (Direct Join - No redundant steps)
    // -------------------------------------------------------------
    #"Merged Queries" = Table.NestedJoin(#"Changed Type1", {"Complaint No"}, Manual_Collection_Query, {"Complaint No"}, "Manual_Data", JoinKind.LeftOuter),
    #"Expanded Manual_Data" = Table.ExpandTableColumn(#"Merged Queries", "Manual_Data", {"Amount Received"}, {"Amount Received"}),
    #"Changed Type Manual" = Table.TransformColumnTypes(#"Expanded Manual_Data", {{"Amount Received", Int64.Type}}),

    // Reconciliation Status Logic
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type Manual", "Reconciliation Status", each 
        if [Amount Received] = null then "Collection Pending" 
        else if [Amount Received] = [Net Collection] then "Matched" 
        else "Discrepancy", 
        type text
    ),
    
    // Net Collection <> 0 Filter
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([Net Collection] <> 0)),

    // -------------------------------------------------------------
    // STEP 4: CM NO & Current Remarks ko Link karna (Complaint No ki bunyad par)
    // -------------------------------------------------------------
    #"Merged Preserved Data" = Table.NestedJoin(#"Filtered Rows", {"Complaint No"}, PreservedData, {"Complaint No"}, "Preserved_Inputs", JoinKind.LeftOuter),
    #"Expanded Preserved Inputs" = Table.ExpandTableColumn(#"Merged Preserved Data", "Preserved_Inputs", {"CM NO", "Current Remarks"}, {"CM NO", "Current Remarks"})
in
    #"Expanded Preserved Inputs"
Attribute VB_Name = "TimeCounterNew2"
'Personnalized data types.

Type NeolaneTask
    Label As String
    StartDate As Date
    EndDate As Date
    Duration As Double
    BusyStatus As String
    InvoicingCategory As String
    Category As String
End Type
    
Type NeolaneBill
    BillingStartDate As Date
    BillingEndDate As Date
    Tasks() As NeolaneTask
End Type

'Returns the invoicing status of a task
Function GetInvoicingStatus(intInputBusyStatus As Integer, strInputCategories As String)
    'Task is to invoice only if its status is "Busy" or "Out Of Office".
    Select Case intInputBusyStatus
        Case 0 To 1
            GetInvoicingStatus = "Do not invoice"
        Case 2 To 3
            GetInvoicingStatus = "To invoice"
        Case Else
            GetInvoicingStatus = "Cannot get invoicing status"
    End Select
End Function

'Returns the busy status text of a task
Function GetBusyStatus(intInputBusyStatus As Integer)
    'Default Outlook meanings.
    Select Case intInputBusyStatus
        Case 0
            GetBusyStatus = "Free"
        Case 1
            GetBusyStatus = "Tentative"
        Case 2
            GetBusyStatus = "Busy"
        Case 3
            GetBusyStatus = "Out of office"
        Case Else
            GetBusyStatus = "Status unknown"
    End Select
End Function


'Procedure to list all Outlook tasks between two dates.
' Input : datInputStartDate As Date, datInputEndDate As Date
Sub ListTasks()
    
    Dim datStart, datEnd As Date
    Dim objCalendar As Outlook.Folder
    Dim objItems As Outlook.Items
    Dim objAppt As Outlook.AppointmentItem
    Dim strRestriction As String
    Dim intLineCount As Integer
    Dim objExcel As Object
    Dim objBook As Object
    Dim objSheet As Object
    
    'Start a new workbook in Excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.Add
    Set objSheet1 = objBook.Worksheets(1)
    Set objSheet2 = objBook.Worksheets(2)

    Set objCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set objItems = objCalendar.Items

    'Include recurring appointments.
    objItems.IncludeRecurrences = True

    'Format dates as Outlook espects them.
    'datStart = Format(datInputStartDate, "ddddd h:nn AMPM")
    'datEnd = Format(DateAdd("d", 1, datInputEndDate), "ddddd h:nn AMPM")
    datStart = Format("26/11/2011", "ddddd h:nn AMPM")
    datEnd = Format("24/12/2011", "ddddd h:nn AMPM")
    
    'Construct a filter.
    strRestriction = "[Start] >= '" & datStart & "' AND [End] <= '" & datEnd & "'"

    'Restrict the Items collection according to the previous filter.
    Set objItems = objItems.Restrict(strRestriction)
    
    'Sort the Items collection
    'objItems.Sort "[Start]"
        

    objSheet1.Range("A1:G1").Value = Array("StartDate", "EndDate", "Label", "Duration", "Categories", "Busy Status", "Invoicing")
    
    'Initialize task count.
    intLineCount = 1
    
    'Place items in our object.
    For Each objAppt In objItems
        'Add item in Excel
        intLineCount = intLineCount + 1
        objSheet1.Range("A" & intLineCount & ":G" & intLineCount).Value = Array(objAppt.Start, objAppt.End, objAppt.Subject, objAppt.Duration / 60, objAppt.Categories, GetBusyStatus(objAppt.BusyStatus), GetInvoicingStatus(objAppt.BusyStatus, objAppt.Categories))
    Next
    
    objSheet1.Range("A1:G" & intLineCount).EntireColumn.AutoFit

    objSheet1.ListObjects.Add(xlSrcRange, Range("A1:G" & intLineCount), , xlYes).Name = "ListTasks"
    objSheet1.ListObjects("ListTasks").TableStyle = "TableStyleLight2"
    
    Dim objTable As PivotTable
    
    ' Create the PivotTable object.
    objBook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:="ListTasks", _
            Version:=xlPivotTableVersion14 _
        ).CreatePivotTable _
            TableDestination:="Sheet2!R1C1", _
            TableName:="PivotTableBilling", _
            DefaultVersion:=xlPivotTableVersion14
    Set objTable = objSheet2.PivotTables("PivotTableBilling")
    With objTable.PivotFields("Categories")
        .Orientation = xlRowField
        .Position = 1
        .PivotItems("Holiday").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With objTable.PivotFields("Label")
        .Orientation = xlRowField
        .Position = 2
    End With
    With objTable.PivotFields("Busy Status")
        .Orientation = xlPageField
        .Position = 1
        .PivotItems("Free").Visible = False
        .PivotItems("Tentative").Visible = False
    End With
    With objTable.PivotFields("Invoicing")
        .Orientation = xlColumnField
    End With
    'objTable.AddDataField objTable.PivotFields("Duration"), "Duration", xlSum
    objTable.AddDataField objTable.PivotFields("Duration"), "Duration of tasks", xlSum

    'Show the Excel sheet.
    objExcel.Visible = True

End Sub

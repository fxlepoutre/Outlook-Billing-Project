Attribute VB_Name = "TimeCounterNew"
'Personnalized data types.

Type NeolaneTask
    Label As String
    StartDate As Date
    EndDate As Date
    Duration As Double
    BusyStatus As String
    InvoicingCategory As String
End Type

Type NeolaneCustomerBill
    Customer As String
    TimeToBill As Double
    TimeToBillToString As String
    Tasks() As NeolaneTask
    Billable As Boolean
End Type
    
Type NeolaneBill
    BillingStartDate As Date
    BillingEndDate As Date
    Customers() As NeolaneCustomerBill
End Type

'Returns the invoicing status of a task
Function GetInvoicingStatus(intInputBusyStatus As Integer)
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

'Finds the total duration of some Taskss, based on the view and the category.
Function FindTotalDuration(sInputCategory As String, datInputStartDate As Date, datInputEndDate As Date)

    Dim datStart, datEnd As Date
    Dim objCalendar As Outlook.Folder
    Dim objItems As Outlook.Items
    Dim objAppt As Outlook.AppointmentItem
    Dim strRestriction As String
    Dim dblDuration As Double

    Set objCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set objItems = objCalendar.Items
    
    'Include recurring appointments.
    objItems.IncludeRecurrences = True

    'Format dates as Outlook espects them.
    datStart = Format(datInputStartDate, "ddddd h:nn AM/PM")
    datEnd = Format(datInputEndDate, "ddddd h:nn AM/PM")
    
    'Construct a filter.
    strRestriction = "[Categories] = '" & sInputCategory & "'AND [Start] >= '" & datStart & "'AND [End] <= '" & datEnd & "'"
    'Debug.Print strRestriction
    'Restrict the Items collection.
    Set objItems = objItems.Restrict(strRestriction)
    
    'Sort and print the final results.
    objItems.Sort "[Start]"
    dblDuration = 0
    For Each objAppt In objItems
        dblDuration = dblDuration + objAppt.Duration
        'Debug.Print objAppt.Start, objAppt.Duration, objAppt.End, objAppt.Categories, objAppt.Subject
    Next
    'Debug.Print "Total duration: ", Format(-(-Int(dblDuration / 60)), "#0") & " hours "; Format(dblDuration / 60 / 24, "nn") & " minutes"
    FindTotalDuration = dblDuration

End Function

Sub ListTasks(ByRef nltInputNeolaneCustomerBill As NeolaneCustomerBill, datInputStartDate As Date, datInputEndDate As Date)
    
    Dim datStart, datEnd As Date
    Dim objCalendar As Outlook.Folder
    Dim objItems As Outlook.Items
    Dim objAppt As Outlook.AppointmentItem
    Dim strRestriction As String
    Dim dblDuration As Double
    Dim intTaskCount As Integer
    
    Set objCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set objItems = objCalendar.Items
    
    'Include recurring appointments.
    objItems.IncludeRecurrences = True

    'Format dates as Outlook espects them.
    datStart = Format(datInputStartDate, "ddddd h:nn AMPM")
    datEnd = Format(DateAdd("d", 1, datInputEndDate), "ddddd h:nn AMPM")
    
    'Construct a filter.
    strRestriction = "[Categories] = '" & nltInputNeolaneCustomerBill.Customer & "'AND [Start] >= '" & datStart & "'AND [End] <= '" & datEnd & "'"

    'Restrict the Items collection according to the previous filter.
    Set objItems = objItems.Restrict(strRestriction)
    
    'Sort the Items collection
    objItems.Sort "[Start]"
    
    'Initialize category count.
    intTaskCount = 0
    
    'Place items in our object.
    For Each objAppt In objItems
        'Add Tasks to item
        intTaskCount = intTaskCount + 1
        ReDim Preserve nltInputNeolaneCustomerBill.Tasks(1 To intTaskCount)
        nltInputNeolaneCustomerBill.Tasks(intTaskCount).Duration = objAppt.Duration
        nltInputNeolaneCustomerBill.Tasks(intTaskCount).StartDate = objAppt.Start
        nltInputNeolaneCustomerBill.Tasks(intTaskCount).EndDate = objAppt.End
        nltInputNeolaneCustomerBill.Tasks(intTaskCount).Label = objAppt.Subject
        nltInputNeolaneCustomerBill.Tasks(intTaskCount).BusyStatus = GetBusyStatus(objAppt.BusyStatus)
        nltInputNeolaneCustomerBill.Tasks(intTaskCount).InvoicingCategory = GetInvoicingStatus(objAppt.BusyStatus)
    Next
    
End Sub


'Checks the validity of a view, and sets a NeolaneBill object passed by reference with start and end dates if view is valid.
Function CheckViewValidity(ByRef nltBill As NeolaneBill)
    Dim strSearchString As String
    Dim objView As Outlook.View
    Dim datStartDate As Date
    Dim datEndDate As Date

    CheckViewValidity = True
    'Check if the current view is a table view.
    If Application.ActiveExplorer.CurrentView.ViewType = olTableView Then
        'Set the view to the active view.
        Set objView = Application.ActiveExplorer.CurrentView
        'Debug.Print "Current view : ", objView.Name, objView.Filter
        'Search the start date of the filter.
        strSearchString = """urn:schemas:calendar:dtstart"" >= '"
        If InStr(objView.Filter, strSearchString) > 0 Then
            nltBill.BillingStartDate = CDate(Mid(objView.Filter, InStr(objView.Filter, strSearchString) + Len(strSearchString), 10))
        Else
            MsgBox "The current view should be filtered with a start date, before which it does not include any items (start date on or after dd/mm/yyyy). Please use ""View Settings"" in ribbon, ""Filter..."" section, ""Advanced"" tab."
            CheckViewValidity = False
            Exit Function
        End If
        'Debug.Print "Start date:", datStartDate
        'Search the end date of the filter.
        strSearchString = """urn:schemas:calendar:dtend"" <= '"
        If InStr(objView.Filter, strSearchString) > 0 Then
            nltBill.BillingEndDate = CDate(Mid(objView.Filter, InStr(objView.Filter, strSearchString) + Len(strSearchString), 10))
        Else
            MsgBox "The current view should be filtered with an end date, after which it does not include any items (end date on or before dd/mm/yyyy).  Please use ""View Settings"" in ribbon, ""Filter..."" section, ""Advanced"" tab."
            CheckViewValidity = False
            Exit Function
        End If
        'Debug.Print "End date:", datEndDate
    Else
        MsgBox "The current view should be a calendar listed view. Please change to Calendar View and/or use ""Change View"" in the ribbon."
        CheckViewValidity = False
    End If
End Function

'Procedure to create and fill a bill.
Sub FillBills()
    Dim nltBill As NeolaneBill
    Dim objNameSpace As NameSpace
    Dim objCategory As Category
    Dim intCatCount As Integer
    Dim dblDuration As Double

    'Check if current view is a valid view and initializes Billing object if it is.
    If CheckViewValidity(nltBill) <> True Then
        Exit Sub
    Else
        'Obtain a NameSpace object reference.
        Set objNameSpace = Application.GetNamespace("MAPI")
        'Check if the Categories collection for the Namespace contains one or more Category objects.
        If objNameSpace.Categories.Count > 0 Then
            'Initialize category count.
            intCatCount = 0
            'Enumerate the Categories collection.
            For Each objCategory In objNameSpace.Categories
                'Remove any unwanted category (here only "0- Personnal" for instance)
                If objCategory.Name <> "0- Personnal" Then
                    'Count total duration
                    dblDuration = FindTotalDuration(objCategory.Name, nltBill.BillingStartDate, nltBill.BillingEndDate)
                    'Only categories that have some work in the month are billable.
                    If dblDuration > 0 Then
                        'Add category to bill.
                        intCatCount = intCatCount + 1
                        ReDim Preserve nltBill.Customers(1 To intCatCount)
                        nltBill.Customers(intCatCount).Customer = objCategory.Name
                        ListTasks nltBill.Customers(intCatCount), nltBill.BillingStartDate, nltBill.BillingEndDate
                        nltBill.Customers(intCatCount).TimeToBill = dblDuration
                        nltBill.Customers(intCatCount).TimeToBillToString = Format(-(-Int(dblDuration / 60)), "#0") & " hours " & Format(dblDuration / 60 / 24, "nn") & " minutes"
                        nltBill.Customers(intCatCount).Billable = True
                    End If
                End If
            Next
            'To transfer to Excel : http://support.microsoft.com/kb/247412
        Else
            MsgBox "Error! No categories found."
        End If
    End If
End Sub
